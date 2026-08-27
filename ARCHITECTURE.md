# Arquitetura

## Visão geral

A solução é composta por quatro camadas funcionais.

### 1. Experiência

O Power Apps **CaP - Contas a Pagar** oferece uma interface única para consulta e atuação sobre duas listas. O dashboard HTML oferece uma visão analítica complementar para arquivos exportados.

### 2. Dados operacionais

- lista vinculada ao fluxo de aprovação de faturas;
- lista de cadastro e gerenciamento;
- planilha operacional usada pelo dashboard;
- relatório ME2N extraído do SAP.

### 3. Automação

Fluxos do Power Automate são executados diariamente para identificar pendências e enviar lembretes baseados nas listas e no ME2N.

### 4. Sistema de origem

O SAP permanece como sistema corporativo de origem para pedidos e documentos relacionados ao processo financeiro.

## Fronteiras

O dashboard HTML:

- processa os dados no navegador;
- não possui backend;
- não grava no SAP;
- não substitui o Power Apps;
- não executa os fluxos do Power Automate;
- não autentica o usuário por conta própria.

## Dependências atuais

A alimentação ME2N depende de atualização manual duas vezes por mês. A conexão direta com o SAP é uma hipótese de evolução e requer desenho específico de integração e segurança.
