# Robô Sitrack na nuvem (Apify) — Deploy

Automação 100% na nuvem, **sem PC**: o Apify loga no Sitrack, exporta o CSV, roda os
scripts Python (xlsx + corpo do email), persiste o histórico e dispara o email via N8N.

```
Apify Scheduler (08:45, seg–sex)
  └─ Actor "robo-sitrack-equipamentos"
       1. restaura histórico (Key-Value store)
       2. login Sitrack (usuário + senha) → exporta CSV
       3. gerar_relatorio.py + gerar_email.py  (xlsx + corpo_email.html)
       4. salva REPORT_XLSX / REPORT_HTML / RESUMO no Key-Value store
       5. envia email via N8N (Call_Enviar_email_report)  →  guilherme.santos@sillion.com.br
       6. persiste histórico atualizado para o dia seguinte
```

## Conteúdo desta pasta
- `main.js` — o actor (Playwright + chamadas Python + envio via N8N).
- `gerar_relatorio.py`, `gerar_email.py` — os mesmos scripts de análise (não alterar).
- `Dockerfile` — Node + Playwright + Python (pandas/openpyxl).
- `package.json` — dependências Node (apify, playwright).
- `.actor/actor.json`, `.actor/input_schema.json` — definição e inputs do actor.

## Deploy (uma vez)

1. **Instale o Apify CLI** e faça login:
   ```bash
   npm install -g apify-cli
   apify login          # cole seu token (Apify Console → Settings → Integrations)
   ```
2. **Suba o actor** (dentro desta pasta `apify-actor`):
   ```bash
   apify push
   ```
   Isso cria/atualiza o actor e faz o build da imagem Docker na nuvem.

   > Alternativa sem CLI: no Apify Console → **Create new Actor** → "From source / Git",
   > ou suba esta pasta como um zip. O importante é manter a estrutura de arquivos.

3. **Configure o input/secrets** do actor (Console → seu actor → Input):
   - `sitrackUser`: usuário do portal.
   - `sitrackPass`: **senha** — clique no cadeado para guardar como **secret**.
   - `emailReport`: destinatário (default já é o Guilherme).
   - `enviarEmail`: deixe `true` em produção.

## Primeira execução (validar)

1. Rode com **`headless: false`** e **`enviarEmail: false`** para conferir o login e a
   exportação sem disparar email. Acompanhe pelo **Live View** do Apify.
2. Se o login falhar, ajuste o trecho de login em `main.js` (campos de usuário/senha) —
   os seletores de navegação vêm do codegen do robô e devem funcionar; o passo da senha
   foi adicionado (preenche `input[type="password"]`).
3. Quando o CSV exportar e os logs mostrarem `REPORT_XLSX/HTML gravados`, rode de novo
   com `enviarEmail: true` e confirme o email.

## Agendar (sem PC)

No Console → **Schedules** → New schedule:
- Cron: `45 8 * * 1-5` (08:45, seg–sex)
- **Timezone: America/Sao_Paulo**
- Ação: rodar o actor `robo-sitrack-equipamentos` com o input salvo.

Pronto: roda sozinho todo dia útil, sem máquina local.

## Observações importantes
- **Histórico**: fica no Key-Value store do actor (chave `STATE`). É o que permite o
  comparativo "dia anterior". Não apague. Para reinicializar, rode uma vez com
  `resetState: true`.
- **Login**: como o container é novo a cada run, a senha é usada toda vez (por isso o
  secret). Como o portal é só usuário+senha (sem captcha/2FA), isso funciona.
- **Envio**: reaproveita o endpoint N8N que já funciona (`Call_Enviar_email_report`),
  então **não precisa** do workflow do Drive. Pode arquivar/excluir o
  "Relatorio Depositos Sitrack - Envio" (Drive) que criamos antes, ou deixá-lo desativado.
- **Custo**: o Apify cobra por uso de computação; uma run diária curta é barata, mas
  confira seu plano.
- Se o Sitrack mudar o layout, os seletores em `main.js` podem precisar de ajuste (igual
  ao robô local).
```
