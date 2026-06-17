# Subir o actor pro GitHub e ligar no Apify (auto-build)

Isto resolve o problema do OneDrive (o Apify passa a buildar direto do GitHub) e
deixa as próximas atualizações automáticas: todo `git push` rebuilda o actor.

## 0. Pré-requisito
- Tenha o **Git** instalado: https://git-scm.com/download/win
- (Opcional, recomendado) **GitHub CLI**: https://cli.github.com/  → depois `gh auth login`

## 1. Copie a pasta para FORA do OneDrive
Importante: evita o bug de "arquivo só na nuvem" e conflito do OneDrive com o `.git`.

No PowerShell:
```powershell
robocopy "C:\Users\sitra\OneDrive - SITRACK SERVIÇOS TECNOLOGICOS LTDA\Área de Trabalho\Projetos\Robo Sillion\Robo Sillion\apify-actor" "C:\robo\apify-actor" /E
cd C:\robo\apify-actor
```
Confira que os equip vieram inteiros (cada ~1,6 MB):
```powershell
dir seed\equip
```

## 2. Inicialize o repositório
```powershell
git init
git add .
git commit -m "Robo Sitrack - actor Apify (extracao + analise + envio)"
git branch -M main
```

## 3. Crie o repositório no GitHub e faça o push

### Opção A — com GitHub CLI (mais fácil)
```powershell
gh repo create apify-actor --private --source=. --remote=origin --push
```

### Opção B — pelo site
1. Acesse https://github.com/new → nome: **apify-actor** → Private → **Create repository** (não marque README/gitignore).
2. Copie a URL do repo e rode:
```powershell
git remote add origin https://github.com/SEU_USUARIO/apify-actor.git
git push -u origin main
```

## 4. Ligue o Apify ao repositório
No Apify Console → seu actor (ou Create new Actor) → **Link a Git repository** → GitHub →
escolha **apify-actor** → branch `main` → pasta `/` (raiz).
- Marque para **buildar automaticamente** a cada push.
- Rode um **Build** inicial.

A partir daí: mudou o código → `git push` → o Apify rebuilda sozinho. Sem `apify push`,
sem OneDrive no meio.

## 5. Confirme
Rode o actor e veja no Log:
- `[seed] historico.json: +N dia(s)`
- `[seed] equip: +N snapshot(s) do seed.`
- `total: ... | anterior: 2026-06-16`
