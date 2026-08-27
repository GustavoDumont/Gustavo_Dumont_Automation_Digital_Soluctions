# Dicionário de dados do dashboard

O dashboard procura cabeçalhos equivalentes e normaliza variações de escrita. Os principais campos presentes na fonte operacional são:

- **Fornecedor**: nome do fornecedor ou beneficiário.
- **Código**: código do fornecedor.
- **NF / Fatura**: número da nota fiscal, fatura ou identificação do lançamento.
- **Emissão**: data de emissão.
- **Município**: local associado ao documento.
- **Centro**: centro organizacional.
- **Centro de custo**: classificação de custo.
- **Valor**: valor financeiro do registro.
- **Requisição**: número da requisição.
- **Cotação**: número da cotação.
- **Pedido / Contrato**: referência do pedido ou contrato.
- **Aprovações**: situação da aprovação.
- **MIGO**: documento ou etapa SAP relacionada ao recebimento.
- **MIR7**: documento ou etapa SAP relacionada à fatura.
- **FV60 / F47**: referência de lançamento financeiro quando aplicável.
- **Data de pagamento**: data prevista ou realizada.
- **Forma de pagamento**: boleto, depósito, adiantamento ou outra forma.
- **Observação**: contexto operacional e tratativas.

## Qualidade esperada

- uma linha por registro operacional;
- cabeçalhos na primeira linha da aba;
- valores como números, sem texto adicional;
- datas reconhecíveis pelo Excel;
- identificadores mantidos como texto quando zeros à esquerda forem relevantes;
- observações sem dados desnecessários ou sensíveis.
