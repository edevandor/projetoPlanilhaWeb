# Planília Web

Pipeline mínimo para publicar a aba `MOSTRUARIO` da planilha de preços como um
site HTML estático no GitHub Pages.

## Arquitetura

```text
MODELO_PRECIFICACAO_REVISADA.xlsx
        │
        ├─ LibreOffice headless: recalcula fórmulas
        │
        └─ scripts/build_site.py
                │
                └─ site/index.html
                        │
                        └─ GitHub Pages
```

O pipeline não recria a matemática da planilha. Ele recalcula uma cópia
temporária, lê os valores finais da aba `MOSTRUARIO` e gera uma única página
autocontida com busca.

## Execução local

```bash
python3 -m venv .venv
. .venv/bin/activate
python -m pip install -e .
python scripts/build_site.py \
  --input data/MODELO_PRECIFICACAO_REVISADA.xlsx \
  --output site
```

Depois, abrir `site/index.html` no navegador.

## GitHub Pages

O workflow em `.github/workflows/publish.yml` instala as dependências, executa
o pipeline e publica o site. No GitHub, habilitar Pages usando `GitHub Actions`
como fonte de build/deploy.

## Atualização

Substituir o XLSX em `data/`, fazer commit e push. O workflow reconstrói a página
automaticamente.

## Atenção de segurança

O site publicado pelo GitHub Pages pode ser público. Não colocar no repositório
nem publicar preços se a política da empresa exigir acesso privado.
