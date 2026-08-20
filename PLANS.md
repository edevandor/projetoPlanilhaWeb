# Plano de implementação

## Fase 1 — Pipeline mínimo funcional

### MC-1.1 — Fonte e contrato
- [x] Copiar a planilha para `data/` sem alterar o original.
- [x] Registrar o contrato em `PROJECT_SPEC.md`.
- [x] Registrar as regras de governança em `AGENTS.md`.

### MC-1.2 — Extração da MOSTRUARIO
- [x] Recalcular uma cópia temporária com LibreOffice.
- [x] Ler somente a aba `MOSTRUARIO` com valores calculados.
- [x] Interromper com erro se a aba ou os registros não existirem.

### MC-1.3 — HTML estático
- [x] Gerar `site/index.html` autocontido.
- [x] Formatar preços em reais.
- [x] Adicionar busca client-side.

### MC-1.4 — Publicação
- [x] Declarar a dependência e o LibreOffice no GitHub Actions.
- [x] Gerar o site no CI.
- [ ] Publicar via GitHub Pages — depende de conectar o repositório e habilitar Pages.

## Critério de fechamento da fase

Uma execução real contra `data/MODELO_PRECIFICACAO_REVISADA.xlsx` termina com código
zero e produz `site/index.html` contendo registros da aba `MOSTRUARIO`.
