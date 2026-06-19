/**
 * Actor Apify - Robô Sitrack (Equipamentos em Depósito), serverless.
 *
 * Fluxo:
 *  1) Restaura o histórico (historico.json + relatorios/equip/) do Key-Value store.
 *  2) Loga no portal Sitrack (usuário + senha) e exporta a Lista de Equipamentos em CSV.
 *  3) Roda gerar_relatorio.py e gerar_email.py (geram o xlsx e o corpo_email.html).
 *  4) Salva no Key-Value store: REPORT_XLSX, REPORT_HTML, RESUMO.
 *  5) Persiste o histórico atualizado (STATE) para a próxima execução.
 *
 * O n8n depois lê REPORT_XLSX / REPORT_HTML / RESUMO e envia o email via Gmail.
 */
import { Actor } from 'apify';
import { chromium } from 'playwright';
import { execFileSync } from 'node:child_process';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

// ---------------- Configuração (igual ao robô local) ----------------
const URL_PORTAL = 'https://www.sitrack.com.br/portal/portugues/index.php?lang=pt';
const PASTA_RELAT = path.join(__dirname, 'relatorios');
const PASTA_EMAIL = path.join(__dirname, 'email');
const PASTA_EQUIP = path.join(PASTA_RELAT, 'equip');
const PASTA_TMPL = path.join(PASTA_RELAT, 'Template');
const HIST_JSON = path.join(PASTA_RELAT, 'historico.json');

const ESPERA_MS = 600;
const ESPERA_LENTA_MS = 2200;
const TIMEOUT_LOGIN_MS = 90 * 1000;
const TIMEOUT_DOWNLOAD_MS = 10 * 60 * 1000;
const STATE_KEY = 'STATE'; // tar.gz com historico.json + equip/
// Envio via N8N (mesmo endpoint MCP usado pelo enviar_email.py)
const N8N_MCP_URL_DEFAULT = 'https://sbrgui.app.n8n.cloud/mcp/805ad43f-af5f-4718-8c83-5fbd433a863f';
const N8N_TOOL_DEFAULT = 'Call_Enviar_email_report';
// --------------------------------------------------------------------

await Actor.init();

const input = (await Actor.getInput()) || {};
const USUARIO = input.sitrackUser || process.env.SITRACK_USER || 'financeiro2@sillion.com.br';
const SENHA = input.sitrackPass || process.env.SITRACK_PASS;
const HEADLESS = input.headless !== false;
const RESET_STATE = input.resetState === true;
const DESTINATARIO = input.emailReport || process.env.EMAIL_REPORT || 'guilherme.santos@sillion.com.br';
const N8N_MCP_URL = input.n8nMcpUrl || N8N_MCP_URL_DEFAULT;
const N8N_TOOL = input.n8nTool || N8N_TOOL_DEFAULT;
const ENVIAR_EMAIL = input.enviarEmail !== false;

// --- Envio via endpoint MCP do N8N (JSON-RPC sobre HTTP, com retry) ---
async function mcpPost(url, sid, body, expectResponse = true) {
    const headers = { 'Content-Type': 'application/json', 'Accept': 'application/json, text/event-stream' };
    if (sid) headers['mcp-session-id'] = sid;
    const r = await fetch(url, { method: 'POST', headers, body: JSON.stringify(body) });
    const newSid = r.headers.get('mcp-session-id') || sid;
    if (r.status >= 400) throw new Error('HTTP ' + r.status + ': ' + (await r.text()).slice(0, 200));
    if (!expectResponse) return { sid: newSid, data: null };
    const text = await r.text();
    const ct = r.headers.get('content-type') || '';
    let data = null;
    if (ct.includes('text/event-stream')) {
        for (const line of text.split(/\r?\n/)) {
            if (line.startsWith('data:')) { try { data = JSON.parse(line.slice(5).trim()); break; } catch { /* segue */ } }
        }
    } else {
        try { data = JSON.parse(text); } catch { data = text; }
    }
    return { sid: newSid, data };
}

async function enviarViaN8N({ assunto, html, fileName, xlsxBase64 }) {
    const TENT = 4, ESPERA = 15000;
    let ultimo;
    for (let t = 1; t <= TENT; t++) {
        try {
            const init = await mcpPost(N8N_MCP_URL, null, {
                jsonrpc: '2.0', id: 1, method: 'initialize',
                params: { protocolVersion: '2024-11-05', capabilities: {}, clientInfo: { name: 'sitrack-apify', version: '1.0' } },
            });
            await mcpPost(N8N_MCP_URL, init.sid, { jsonrpc: '2.0', method: 'notifications/initialized', params: {} }, false);
            const { data } = await mcpPost(N8N_MCP_URL, init.sid, {
                jsonrpc: '2.0', id: 2, method: 'tools/call',
                params: {
                    name: N8N_TOOL,
                    arguments: {
                        input: JSON.stringify({
                            email: DESTINATARIO, email_assunto: assunto, email_html: html,
                            fileName, data: xlsxBase64,
                        }),
                    },
                },
            });
            const txt = (((data && data.result && data.result.content) || [{}])[0] || {}).text || JSON.stringify(data || '');
            if (txt.includes('"SENT"') || txt.includes('threadId')) { console.log('[OK] Email enviado via N8N.'); return true; }
            if (txt.toLowerCase().includes('error')) throw new Error('N8N retornou erro: ' + txt.slice(0, 300));
            console.log('[OK] Enviado. Resposta:', txt.slice(0, 200));
            return true;
        } catch (e) {
            ultimo = e;
            console.log(`[tentativa ${t}/${TENT}] falhou: ${e.message}`);
            if (t < TENT) await new Promise((res) => setTimeout(res, ESPERA));
        }
    }
    throw new Error('Falha ao enviar via N8N: ' + (ultimo ? ultimo.message : 'desconhecido'));
}

if (!SENHA) {
    throw new Error('Senha do Sitrack não informada (input.sitrackPass / secret).');
}

const store = await Actor.openKeyValueStore();

// Garante a estrutura de pastas que os scripts Python esperam.
for (const d of [PASTA_RELAT, PASTA_EMAIL, PASTA_EQUIP, PASTA_TMPL]) {
    fs.mkdirSync(d, { recursive: true });
}

// 1) Restaurar histórico (STATE) do Key-Value store ----------------------------
if (!RESET_STATE) {
    const state = await store.getValue(STATE_KEY); // Buffer (tar.gz) ou null
    if (state && Buffer.isBuffer(state) && state.length > 0) {
        const tgz = path.join(__dirname, 'state.tgz');
        fs.writeFileSync(tgz, state);
        try {
            execFileSync('tar', ['xzf', tgz, '-C', PASTA_RELAT], { stdio: 'inherit' });
            console.log('[state] Histórico restaurado do Key-Value store.');
        } catch (e) {
            console.log('[state] Falha ao extrair histórico (seguindo sem):', e.message);
        }
        fs.rmSync(tgz, { force: true });
    } else {
        console.log('[state] Nenhum histórico salvo ainda (primeira execução).');
    }
} else {
    console.log('[state] resetState=true: ignorando histórico salvo.');
}

// 1.1) Semear histórico passado (só datas que ainda não existem no live) -------
try {
    const seedHist = path.join(__dirname, 'seed', 'historico.json');
    if (fs.existsSync(seedHist)) {
        let live = [];
        if (fs.existsSync(HIST_JSON)) {
            try { live = JSON.parse(fs.readFileSync(HIST_JSON, 'utf8')); } catch { live = []; }
        }
        const datas = new Set(live.map((r) => r.data));
        const seed = JSON.parse(fs.readFileSync(seedHist, 'utf8'));
        let add = 0;
        for (const r of seed) { if (r && r.data && !datas.has(r.data)) { live.push(r); add++; } }
        live.sort((a, b) => String(a.data || '').localeCompare(String(b.data || '')));
        fs.writeFileSync(HIST_JSON, JSON.stringify(live, null, 2));
        console.log(`[seed] historico.json: +${add} dia(s) do seed (total ${live.length}).`);
    }
} catch (e) {
    console.log('[seed] Falha ao semear histórico:', e.message);
}

// 1.2) Semear snapshots equip passados (só os que ainda não existem) -----------
//      Necessário para a "Análise de movimentações" comparar com o dia anterior.
try {
    const seedEquipDir = path.join(__dirname, 'seed', 'equip');
    if (fs.existsSync(seedEquipDir)) {
        let add = 0;
        for (const fn of fs.readdirSync(seedEquipDir)) {
            if (!fn.endsWith('.json')) continue;
            const dest = path.join(PASTA_EQUIP, fn);
            if (!fs.existsSync(dest)) { fs.copyFileSync(path.join(seedEquipDir, fn), dest); add++; }
        }
        console.log(`[seed] equip: +${add} snapshot(s) do seed.`);
    }
} catch (e) {
    console.log('[seed] Falha ao semear equip:', e.message);
}

// 2) Login + exportação do CSV -------------------------------------------------
const dataStr = new Date().toLocaleDateString('en-CA', { timeZone: 'America/Sao_Paulo' }); // YYYY-MM-DD
let csvPath;

const browser = await chromium.launch({ headless: HEADLESS });
const context = await browser.newContext({ acceptDownloads: true });
const page = await context.newPage();

const etapa = async (msg, ms = ESPERA_MS) => {
    if (msg) console.log('  - ' + msg);
    await page.waitForTimeout(ms);
};

try {
    console.log('Abrindo portal Sitrack...');
    await page.goto(URL_PORTAL, { waitUntil: 'domcontentloaded' });
    await etapa('Portal aberto', ESPERA_LENTA_MS);

    // Login com usuário + senha (sem perfil persistente no container).
    const jaLogado = () => page.url().includes('/site5/mainframe');
    if (!jaLogado()) {
        try {
            await page.getByRole('textbox', { name: 'Usuario' }).fill(USUARIO);
            console.log('  - Usuário preenchido');
        } catch (e) { console.log('  - Campo Usuário não encontrado:', e.message); }

        // Preenche senha se o campo já estiver visível na mesma tela...
        const passField = page.locator('input[type="password"]').first();
        if (await passField.count() > 0) {
            try { await passField.fill(SENHA); console.log('  - Senha preenchida (tela única)'); } catch (e) {}
        }

        await page.getByRole('button', { name: 'Login' }).click().catch(() => {});
        await etapa('Login (1) clicado', ESPERA_LENTA_MS);

        // ...ou em uma segunda etapa (tela de senha após o usuário).
        if (!jaLogado()) {
            const pass2 = page.locator('input[type="password"]').first();
            if (await pass2.count() > 0) {
                try {
                    await pass2.fill(SENHA);
                    console.log('  - Senha preenchida (2ª etapa)');
                    await page.getByRole('button', { name: 'Login' }).click().catch(() => {});
                    await etapa('Login (2) clicado', ESPERA_LENTA_MS);
                } catch (e) { console.log('  - Erro ao preencher senha (2):', e.message); }
            }
        }

        await page.waitForURL('**/site5/mainframe/**', { timeout: TIMEOUT_LOGIN_MS });
    }
    await etapa('Login confirmado', ESPERA_LENTA_MS);

    // Consultas > Listas > Lista Equipamentos
    await page.getByTitle('Consultas').click();
    await etapa('Menu Consultas');
    await page.getByTitle('Listas').click();
    await etapa('Submenu Listas');
    await page.getByRole('link', { name: 'Lista Equipamentos' }).click();
    await etapa('Lista Equipamentos aberta', ESPERA_LENTA_MS);

    const f = page.locator('iframe[name="tab3"]').contentFrame();

    // Filtro de Dados: busca = "2"
    await f.locator('[id="HS_interface[0][search][element]"]').click();
    await etapa('Campo de busca focado');
    await f.locator('[id="HS_interface[0][search][element]"]').fill('2');
    await etapa('Busca preenchida com "2"');
    await f.locator('[id="HS_interface[0][search][element]"]').press('Enter');
    await etapa('Busca aplicada (Enter)', ESPERA_LENTA_MS);

    // Tipo de Holder -> Deposito
    await f.locator('[id="HFimage[0][holdertipo]"]').click();
    await etapa('Tipo de Holder expandido');
    await f.getByRole('cell', { name: 'Deposito', exact: true }).click();
    await etapa('Deposito selecionado');
    await f.locator('[id="HFspanControlBox[0][depositotipo]"]').getByRole('img').click();
    await etapa('Caixa Deposito (clique 1)');
    await f.locator('[id="HFspanControlBox[0][depositotipo]"]').getByRole('img').click();
    await etapa('Caixa Deposito (clique 2)');

    // Aba "Filtro de Campos"
    await f.getByText('Filtro de Campos').click();
    await etapa('Aba Filtro de Campos', ESPERA_LENTA_MS);

    const ATRIB_ESTADO_A = 'Atributos    x Canal x Modelo x Versión x Script ID x Script x Controlador x Prestadora x Firmware x Alimentación x Equipamento Principal?  -   Observação  -   Número de Serie  -   Fecha de compra  -   Cliente (ID)  -   Cliente (Nome)  -   Cliente (Estado)  -   Cliente (Documento)  -   Cliente (Fecha Cambio Estado)  -   Ejecutivo de Cuentas  -   Compañía de Seguro  -   Cliente (Cont ID)  -   Holder (Nome)  -   Holder (Placa)  -   Holder (Chasis)  -   Holder (Marca / Modelo)  -   Holder (Marca / Modelo [viejo])  -   MMSI  -   Última Posição  -   Última Revisión  -   Última Instalação  -   Último Movimiento Equipamento (Fecha)  -   Último Movimiento Equipamento (Origen)  -   Último Movimiento Equipamento (Destino)  -   Ubicación de Instalação Equipamento  -   Prestador Satelital N°';
    const ATRIB_ESTADO_B = 'Atributos    x Canal x Modelo x Versión x Script ID x Script x Controlador x Prestadora x Firmware x Alimentación x Equipamento Principal? x Observação x Número de Serie x Fecha de compra x Cliente (ID) x Cliente (Nome) x Cliente (Estado) x Cliente (Documento) x Cliente (Fecha Cambio Estado) x Ejecutivo de Cuentas x Compañía de Seguro x Cliente (Cont ID) x Holder (Nome) x Holder (Placa) x Holder (Chasis) x Holder (Marca / Modelo) x Holder (Marca / Modelo [viejo]) x MMSI x Última Posição x Última Revisión x Última Instalação x Último Movimiento Equipamento (Fecha) x Último Movimiento Equipamento (Origen) x Último Movimiento Equipamento (Destino) x Ubicación de Instalação Equipamento x Prestador Satelital N°';

    await f.getByRole('cell', { name: ATRIB_ESTADO_A, exact: true }).getByRole('img').click();
    await etapa('Toggle atributos (clique 1)', ESPERA_LENTA_MS);
    await f.getByRole('cell', { name: ATRIB_ESTADO_B, exact: true }).getByRole('img').click();
    await etapa('Toggle atributos (clique 2)', ESPERA_LENTA_MS);

    const colunas = [
        'Modelo',
        'Cliente (ID)',
        'Cliente (Nome)',
        'Cliente (Estado)',
        'Última Posição',
        'Último Movimiento Equipamento (Fecha)',
        'Último Movimiento Equipamento (Origen)',
        'Último Movimiento Equipamento (Destino)',
    ];
    for (const col of colunas) {
        await f.getByRole('cell', { name: col, exact: true }).click();
        await etapa('Coluna selecionada: ' + col);
    }

    await f.locator('[id="HFspanControlBox[1][parametros]"]').getByRole('img').click();
    await etapa('Caixa parametros (clique 1)', ESPERA_LENTA_MS);
    await f.locator('[id="HFspanControlBox[1][parametros]"]').getByRole('img').click();
    await etapa('Caixa parametros (clique 2)', ESPERA_LENTA_MS);

    // Ver em Arquivo -> CSV (dispara o download)
    await f.getByText('Ver em Arquivo').click();
    await etapa('Ver em Arquivo', ESPERA_LENTA_MS);

    console.log('Solicitando exportação CSV. Aguardando o arquivo ser gerado...');
    const downloadPromise = page.waitForEvent('download', { timeout: TIMEOUT_DOWNLOAD_MS });
    await f.getByRole('cell', { name: 'CSV' }).click({ modifiers: ['Alt'] });
    const download = await downloadPromise;

    const baseNome = download.suggestedFilename() || 'listadoEquipos.csv';
    csvPath = path.join(PASTA_RELAT, dataStr + '_' + baseNome);
    await download.saveAs(csvPath);
    console.log('CSV salvo em: ' + csvPath);
    await etapa('Download concluído', ESPERA_LENTA_MS);
} finally {
    await context.close();
    await browser.close();
}

if (!csvPath || !fs.existsSync(csvPath)) {
    throw new Error('Exportação falhou: CSV do dia não foi gerado.');
}

// 3) Rodar os scripts Python ---------------------------------------------------
console.log('Rodando gerar_relatorio.py ...');
execFileSync('python3', [path.join(__dirname, 'gerar_relatorio.py'), csvPath], {
    cwd: __dirname, stdio: 'inherit',
});
console.log('Rodando gerar_email.py ...');
execFileSync('python3', [path.join(__dirname, 'gerar_email.py'), csvPath], {
    cwd: __dirname, stdio: 'inherit',
});

// 4) Ler resultados e salvar no Key-Value store --------------------------------
const resumoPath = path.join(PASTA_EMAIL, 'resumo.json');
const resumo = JSON.parse(fs.readFileSync(resumoPath, 'utf8'));
const xlsxPath = path.join(PASTA_TMPL, resumo.arquivo_xlsx);
const htmlPath = path.join(PASTA_EMAIL, 'corpo_email.html');

const xlsxBuf = fs.readFileSync(xlsxPath);
const htmlStr = fs.readFileSync(htmlPath, 'utf8');

await store.setValue('REPORT_XLSX', xlsxBuf, {
    contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
});
await store.setValue('REPORT_HTML', htmlStr, { contentType: 'text/html; charset=utf-8' });
await store.setValue('RESUMO', resumo); // JSON

console.log('[ok] REPORT_XLSX / REPORT_HTML / RESUMO gravados no Key-Value store.');

// 4.1) Persistir histórico (STATE) ANTES de enviar o email -------------------
//      Garante que a cadeia diária (06-18 vira base de 06-19) NÃO quebra mesmo
//      que o envio do email falhe (N8N fora do ar, etc.).
try {
    // Poda: mantém só o snapshot equip MAIS RECENTE (evita encher o storage).
    // A comparação diária só precisa do dia anterior; a tabela histórica vem do
    // historico.json (pequeno, acumulado). Então 1 snapshot equip basta.
    try {
        const KEEP = 1;
        const eq = fs.readdirSync(PASTA_EQUIP).filter((f) => f.endsWith('.json')).sort();
        const remover = eq.slice(0, Math.max(0, eq.length - KEEP));
        for (const f of remover) fs.rmSync(path.join(PASTA_EQUIP, f), { force: true });
        if (remover.length) console.log('[poda] equip: removidos ' + remover.length + ', mantido ' + (eq.length - remover.length) + ' (mais recente).');
    } catch (e) { console.log('[poda] equip falhou:', e.message); }

    const tgz = path.join(__dirname, 'state_out.tgz');
    const alvos = [];
    if (fs.existsSync(HIST_JSON)) alvos.push('historico.json');
    if (fs.existsSync(PASTA_EQUIP)) alvos.push('equip');
    if (alvos.length) {
        execFileSync('tar', ['czf', tgz, '-C', PASTA_RELAT, ...alvos], { stdio: 'inherit' });
        const buf = fs.readFileSync(tgz);
        await store.setValue(STATE_KEY, buf, { contentType: 'application/gzip' });
        fs.rmSync(tgz, { force: true });
        console.log('[state] Histórico persistido para a próxima execução.');
    }
} catch (e) {
    console.log('[state] Falha ao persistir histórico:', e.message);
}

// 4.2) Enviar o email via N8N (mesmo endpoint do enviar_email.py) -------------
if (ENVIAR_EMAIL) {
    const assunto = resumo.assunto || ('Relatório Diário - Equipamentos em Depósito (' + resumo.data + ')');
    await enviarViaN8N({
        assunto,
        html: htmlStr,
        fileName: resumo.arquivo_xlsx,
        xlsxBase64: xlsxBuf.toString('base64'),
    });
} else {
    console.log('[info] enviarEmail=false: pulando envio (apenas KV store).');
}

// Resumo desta execução no dataset (auditoria)
await Actor.pushData({
    data: resumo.data,
    total: resumo.total,
    std: resumo.std,
    cameras: resumo.cameras,
    anterior: resumo.anterior ?? null,
    arquivo_xlsx: resumo.arquivo_xlsx,
    csv: path.basename(csvPath),
    geradoEm: new Date().toISOString(),
});

await Actor.exit();
