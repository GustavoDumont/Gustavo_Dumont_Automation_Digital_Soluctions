# Arquitetura

O ecossistema é modular. Cada aplicativo possui entrada, processamento e saída próprios, permitindo implantação gradual.

## Camadas

### Aquisição
- download de relatórios de pendências no portal autorizado;
- leitura do PDF da folha de pagamento;
- recebimento de documentos em pastas de sondagem.

### Transformação
- renomeação e padronização dos relatórios;
- extração dos registros do efetivo;
- normalização de nomes;
- OCR e correspondência aproximada;
- classificação documental.

### Entrega
- relatórios Excel de pendências;
- base Excel do efetivo;
- documentos organizados por colaborador e categoria;
- relatórios CSV de execução e exceções.

## Princípios de evolução

- módulos independentes;
- configuração externa para ambientes e seletores;
- origem preservada sempre que possível;
- modo de simulação antes de movimentar documentos;
- idempotência para reprocessamentos;
- logs sem dados pessoais desnecessários;
- revisão humana para correspondências ambíguas.
