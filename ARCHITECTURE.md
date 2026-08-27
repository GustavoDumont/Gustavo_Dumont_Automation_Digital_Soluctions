# Arquitetura do GTS

## Camada operacional

O Power Apps oferece a interface de trabalho por contrato. A aplicação manipula duas representações complementares mantidas em listas do SharePoint: cadastro do colaborador e registros individualizados de treinamento.

## Camada de automação

Fluxos do Power Automate executam verificações recorrentes e preparam alertas para os responsáveis.

## Camada de referência

A planilha Matriz de Treinamentos organiza obrigatoriedades, catálogo, validades, listas de validação e consultas.

## Camada analítica

O arquivo index.html importa exportações autorizadas, normaliza nomes e datas, consolida origens internas e externas, calcula alertas e conformidade e produz visualizações e relatórios.

## Publicação

GitHub hospeda o código-fonte estático. Cloudflare disponibiliza a aplicação web. O Power Apps direciona o usuário para essa camada através da ação Dashboard.

## Fronteiras de segurança

A aplicação HTML não implementa autorização sobre os dados importados. Os controles de acesso permanecem nas fontes corporativas e na distribuição dos arquivos.
