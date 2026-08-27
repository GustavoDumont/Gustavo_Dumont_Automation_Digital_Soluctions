# Pontos de extensão

## Novas regras
Crie um arquivo JSON em `configs/` usando `between` para conteúdo delimitado por duas âncoras ou `regex` para padrões estruturados.

## Novos formatos
Implemente um leitor em `read_text()` e mantenha a etapa de extração independente do formato.

## Integração com outras automações
A função `process()` retorna o caminho do relatório. Ela pode ser chamada por outro script, por um agendador ou por uma automação de fluxo. Para produção, adicione fila, idempotência, trilha de auditoria e validação humana.

## Perfis específicos
Datas de treinamento, códigos de documentos, registros profissionais e nomes podem ser tratados como configurações. Quando a regra exigir lógica contextual, crie um extrator dedicado sem alterar o pipeline de leitura, relatório e saída.
