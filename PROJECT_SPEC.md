# Planília Web — Especificação

## Objetivo

Publicar uma página web estática para consulta dos preços presentes na aba
`MOSTRUARIO` da planilha de precificação revisada.

## Fonte

`data/MODELO_PRECIFICACAO_REVISADA.xlsx`

A aba `MOSTRUARIO` é uma view de apresentação e contém fórmulas que dependem
de outras abas. O pipeline deve recalcular a workbook com LibreOffice antes da
extração, para que o HTML receba valores finais e não fórmulas.

## Saída

`site/index.html`, contendo:

- cabeçalho da aba `MOSTRUARIO`;
- uma linha por produto preenchido;
- valores monetários das colunas de preço formatados em reais;
- busca textual no navegador;
- nenhuma dependência externa em tempo de execução.

## Fora do escopo

- editar preços pelo navegador;
- reproduzir as fórmulas em Python;
- publicar as abas de cálculo ou parâmetros;
- autenticação, banco de dados, API ou servidor próprio;
- sincronização direta com Excel Online ou Google Sheets.

## Publicação

O GitHub Actions executará o gerador e publicará o diretório `site/` no GitHub
Pages. A página será atualizada quando o XLSX for alterado no repositório.
