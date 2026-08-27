# Operação extraordinária de backup

## Contexto

Com o encerramento do contrato com a UNICO People, os dados armazenados na plataforma precisavam ser preservados antes da indisponibilidade do serviço. A demanda envolvia mais de 150 GB e prazo total de apenas 7 dias.

## Estratégia executada

Gustavo Freitas Gomes Dumont e Phillipe Castro, antigo integrante de TI, prepararam 5 computadores dedicados na sala de TI. As máquinas executaram versões empacotadas da automação em paralelo, sem interrupção, por aproximadamente 5 dias.

A divisão do trabalho por máquinas permitiu transformar uma extração sequencial em uma operação paralela. Os registros foram organizados localmente, conferidos e depois disponibilizados no SharePoint corporativo.

## Linha do tempo resumida

```text
Prazo disponível: 7 dias
Preparação: divisão das filas, máquinas, pastas e credenciais autorizadas
Execução: 5 computadores, operação contínua por 5 dias
Fechamento: consolidação, conferência e disponibilização no SharePoint
Volume preservado: mais de 150 GB
```

## Fatores críticos

- estabilidade de sessão e navegador;
- detecção de arquivos concluídos;
- retomada após falhas;
- prevenção de duplicidades;
- divisão inequívoca das filas;
- espaço em disco e conectividade;
- acompanhamento contínuo;
- integridade dos arquivos;
- transferência final para o repositório corporativo.

## Resultado

A operação foi concluída dentro da janela disponível. O prazo apertado foi atendido por meio de paralelização, dedicação de infraestrutura e acompanhamento próximo, preservando o acervo antes do encerramento do acesso.
