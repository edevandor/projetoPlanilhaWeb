# Governança do projeto

## Escopo

Este projeto faz uma única transformação: lê a aba `MOSTRUARIO` de
`data/MODELO_PRECIFICACAO_REVISADA.xlsx` e gera `site/index.html`.

## Regras

- O XLSX é a fonte da verdade; não duplicar a matemática da precificação em Python.
- O pipeline deve recalcular a workbook com LibreOffice antes de ler os valores finais.
- O HTML é estático e não deve depender de API, banco, CDN ou backend.
- Não alterar a planilha-fonte durante a geração; o recálculo ocorre em diretório temporário.
- Não incluir no site as abas `PARAMETROS`, `PRODUTOS`, `CALCULO` ou `SAIDA`.
- Manter o código mínimo: uma rotina de geração e uma interface HTML autocontida.
- Dados exibidos no site devem vir exclusivamente da aba `MOSTRUARIO`.

## Verificação

A validação mínima é executar o pipeline com a planilha real e conferir:

- o processo termina com código zero;
- `site/index.html` existe;
- o HTML contém os cabeçalhos da aba `MOSTRUARIO`;
- o número de produtos reportado é maior que zero;
- a busca client-side está presente.
