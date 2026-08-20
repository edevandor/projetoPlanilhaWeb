# Planília Web

Gerador local de site estático a partir da aba `MOSTRUARIO` da planilha de
precificação.

Este repositório guarda só:

- o gerador (`scripts/build_site.py`)
- a configuração do projeto (`pyproject.toml`, `.gitignore`)
- o workflow (`.github/workflows/publish.yml`)
- a documentação de uso (`README.md`)

A planilha fonte não é versionada. Ela deve existir só na máquina que executa o
pipeline.

## Fluxo

```text
planilha local
   ↓
LibreOffice headless recalcula fórmulas
   ↓
scripts/build_site.py lê só a aba MOSTRUARIO
   ↓
gera site/index.html
```

## Execução local

1. Coloque a planilha na máquina, por exemplo em:

```text
data/MODELO_PRECIFICACAO_REVISADA.xlsx
```

2. Crie o ambiente e instale as dependências:

```bash
python3 -m venv .venv
. .venv/bin/activate
python -m pip install -e .
```

3. Gere o site:

```bash
python scripts/build_site.py \
  --input data/MODELO_PRECIFICACAO_REVISADA.xlsx \
  --output site
```

4. Abra `site/index.html`.

## GitHub Pages

O workflow está no repositório como contrato de publicação. Ele só faz sentido
se a planilha estiver disponível no ambiente que executar o build.

## Segurança

O repositório não deve conter a planilha fonte. Só o HTML derivado deve ser
publicado ou versionado.
