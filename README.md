# Monitoramento de Entregas SP — Estrutura Profissional

Projeto reorganizado para uso local no navegador com foco em manutenção, clareza e evolução futura.

## Estrutura

```text
monitoramento-jt-profissional/
├── assets/
│   ├── css/
│   ├── img/
│   └── js/
│       ├── core/
│       └── modules/
├── docs/
├── pages/
├── tools/
├── index.html
└── README.md
```

## Melhorias aplicadas

- Estrutura em pastas por responsabilidade.
- Correção da navegação entre Dashboard e Insucessos.
- Separação do CSS específico de insucessos.
- Botão para limpar dados locais em ambas as páginas.
- Verificação de dependências externas ao carregar a página.
- Head, caminhos e scripts reorganizados com `defer`.
- Base pronta para futura migração para app web com backend.

## Como abrir

### Opção 1 — abrir direto no navegador
Abra `index.html`.

### Opção 2 — recomendado
Execute o servidor local incluído em `tools/start-local-server.bat`.

Isso reduz problemas de clipboard, CORS e comportamento inconsistente do `file://`.

## Dependências externas

Atualmente o projeto continua usando CDN para:
- XLSX
- Chart.js
- html2canvas

Se quiser a próxima etapa ideal é empacotar essas libs localmente em `vendor/`.

## Próximos passos recomendados

1. Externalizar regras de negócio fixas para um arquivo de configuração.
2. Unificar parsing de datas e cabeçalhos em uma camada central.
3. Adicionar relatório de auditoria da importação.
4. Migrar para versão web com autenticação e banco quando quiser subir isso online.
