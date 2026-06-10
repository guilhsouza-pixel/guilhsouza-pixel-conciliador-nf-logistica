# Conciliação de Notas Fiscais e Embalagens

Aplicativo web para conciliar duas planilhas Excel: **Notas Enviadas** e **Notas Recebidas**.

## O que o aplicativo faz

- Lê arquivos `.xlsx` e `.xls`
- Identifica automaticamente colunas de NF, série, fornecedor, CNPJ, data, valor, quantidade e item
- Compara as bases usando `CNPJ + NF + Série`
- Quando não existe CNPJ, usa `Fornecedor + NF + Série`
- Mostra resumo executivo com indicadores
- Lista notas enviadas e não recebidas
- Lista notas recebidas e não enviadas
- Lista divergências de valor, quantidade, data e item
- Identifica duplicidades
- Exporta relatório em Excel com várias abas

## Rodar localmente

```bash
npm install
npm run dev
```

## Publicar na Vercel

Use as configurações:

```text
Framework Preset: Vite
Build Command: npm run build
Output Directory: dist
```

## Observação

Todo o processamento ocorre no navegador. Nenhum arquivo é salvo em servidor.
