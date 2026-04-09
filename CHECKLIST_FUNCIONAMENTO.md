# Checklist de Funcionamento

Data da auditoria: 09/04/2026
Build auditada: `20260409r02`

## Resultado geral

- [x] Estrutura principal carregando sem arquivos ausentes
- [x] Shell principal com painel de acoes rapidas fixo
- [x] Tema 3D novo referenciado corretamente
- [x] Service worker alinhado com a build atual
- [x] Entradas `index.html` e `launch.html` com links versionados atualizados
- [x] Layout validado visualmente em desktop
- [x] Layout validado visualmente em mobile

## Checklist validado

- [x] `app-shell.html` referencia `modern.css`, `theme-3d.css`, `manifest.webmanifest` e os `app-core-*.js` com a build `20260409r02`
- [x] `theme-3d.css` existe fisicamente em `assets/css/`
- [x] `index.html` abre a entrada publica atual e redireciona para `launch.html?v=20260409r02`
- [x] `launch.html` registra o `service-worker.js?v=20260409r02`
- [x] `index.html` registra o `service-worker.js?v=20260409r02`
- [x] `service-worker.js` usa `CACHE_NAME = cysy-log360-v20260409r02`
- [x] `service-worker.js` faz precache de `theme-3d.css` e `theme-3d.css?v=20260409r02`
- [x] `service-worker.js` atualiza fallbacks versionados de `index.html`, `launch.html` e `app-shell.html`
- [x] `config.app.buildTag` foi alinhado para `20260409r02`
- [x] O fallback de build em `app-core-2.js` foi alinhado para `20260409r02`
- [x] O topo do shell mantém os 4 botoes principais: `Sincronizar`, `Atualizar`, `Permissoes`, `Sair`
- [x] A aba `Sincronizacao` continua com a central de sincronizacao e os utilitarios movidos
- [x] O botao `Instalar app` continua acessivel pelo `installManager`
- [x] O botao `Som` continua acessivel pelo `soundManager`
- [x] `Backup`, `Restaurar` e `Limpar cache e historico` continuam expostos na interface
- [x] O shell foi inspecionado visualmente em viewport desktop
- [x] O shell foi inspecionado visualmente em viewport mobile
- [x] `index.html` foi inspecionado visualmente
- [x] `launch.html` foi inspecionado visualmente

## Correcao aplicada durante a auditoria

- [x] Corrigido risco de cache inconsistente em PWA instalada causado pela inclusao de `theme-3d.css` sem atualizacao do precache/versionamento

## Observacoes

- Esta auditoria validou estrutura, wiring, cache/versionamento e comportamento visual local.
- Fluxos dependentes de rede externa, APIs remotas e sincronizacao real com servidor nao foram exercitados ponta a ponta neste ambiente local.
