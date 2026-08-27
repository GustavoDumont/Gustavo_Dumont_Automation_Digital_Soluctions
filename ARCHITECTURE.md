# Arquitetura

## V1

A primeira versão recebe três arquivos distintos, cria mapas de correspondência por colaborador e combina férias a vencer, programação, efetivo e histórico. A interface concentra gestão operacional, alertas e análise.

## V6

A V6 utiliza relatórios complementares do TOTVS e um arquivo opcional de salvamento. A consolidação ocorre no navegador.

```text
Excel TOTVS -> File API -> leitor XLSX -> normalização -> modelo em memória
                                                    |-> Gantt
                                                    |-> Pessoas
                                                    |-> Períodos
                                                    |-> Férias
                                                    `-> Histórico
```

## Decisões técnicas

- aplicação estática e portável;
- processamento client-side;
- ausência de backend e banco de dados;
- identidade visual incorporada no HTML;
- leitor XLSX incorporado para reduzir dependências;
- estado inicial vazio;
- arquivo de salvamento como mecanismo explícito de persistência;
- `index.html` aponta para a versão V6.

## Limites

A aplicação não autentica usuários, não controla autorização sobre os relatórios importados e não atualiza o TOTVS. As fontes continuam sendo responsáveis pela governança dos dados.
