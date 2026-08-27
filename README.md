# Ecossistema de Acompanhamento QSMS

Conjunto de aplicações HTML para acompanhamento de processos de Qualidade, Saúde, Segurança e Meio Ambiente. Os painéis transformam exportações em Excel de listas do SharePoint em indicadores, filtros, gráficos, tabelas e análises executadas diretamente no navegador.

A arquitetura foi escolhida para acelerar a criação e a evolução de soluções gerenciais, reduzir dependências de infraestrutura de BI e permitir que o próprio usuário atualize a visualização em poucos segundos, selecionando uma exportação recente da lista correspondente.

> **Privacidade:** os dados importados são processados no navegador. Esta branch não contém planilhas operacionais, credenciais nem dados de colaboradores. O comportamento de rede de bibliotecas externas deve ser considerado conforme a seção de dependências.

## Problema resolvido

A manutenção de painéis tradicionais pode envolver publicação de modelos, gerenciamento de acessos, contas ativas, licenciamento para distribuição e ciclos mais burocráticos de alteração. Para necessidades operacionais que exigiam rapidez, flexibilidade visual e atualização simples, foi adotado um modelo de aplicações HTML independentes.

Cada painel concentra em um único arquivo:

- interface;
- regras de transformação;
- importação de Excel ou CSV, conforme a implementação;
- indicadores e cálculos;
- filtros e pesquisa;
- gráficos e tabelas;
- identidade visual;
- recursos de impressão ou exportação disponíveis no painel.

## Fluxo de uso

```text
Lista do SharePoint
        |
        v
Exportação para Excel
        |
        v
Abertura do HTML no navegador
        |
        v
Seleção do arquivo exportado
        |
        v
Processamento local em JavaScript
        |
        v
KPIs, filtros, gráficos, tabelas e análises
```

O usuário não precisa renomear o arquivo exportado para utilizar os painéis compatíveis. O objetivo é abrir o HTML, importar a exportação e iniciar a análise.

## Aplicações incluídas

### Dashboard de agendamento de treinamentos

Arquivo: [`dashboard_agendamento_treinamentos_concremat.html`](dashboard_agendamento_treinamentos_concremat.html)

Painel para acompanhamento dos treinamentos agendados, com importação de dados, filtros, indicadores, visualizações e tabela detalhada.

### Cronograma de eventos de Segurança do Trabalho

Arquivo: [`dashboard_cronograma_eventos_seguranca_trabalho.html`](dashboard_cronograma_eventos_seguranca_trabalho.html)

Aplicação para acompanhar o cronograma de eventos de Segurança do Trabalho e facilitar a leitura do planejamento ao longo do período.

### Cronograma de eventos com metas

Arquivo: [`dashboard_cronograma_eventos_seguranca_trabalho_metas.html`](dashboard_cronograma_eventos_seguranca_trabalho_metas.html)

Evolução do cronograma com recursos de comparação entre realização e metas para categorias de eventos, permitindo alternar as referências apresentadas na visualização.

### Central de Controle de Inspeções

Arquivo: [`central_controle_inspecoes_normalizacao_selecao_inspetores_final_corrigido.html`](central_controle_inspecoes_normalizacao_selecao_inspetores_final_corrigido.html)

Central para consolidação e análise de inspeções. Inclui múltiplas entradas de dados, indicadores, visões por perguntas e acompanhamento por inspetores, com mecanismos de seleção e comparação com metas.

### Dashboard de controle de chamados

Arquivo: [`dashboard_controle_chamados_concremat.html`](dashboard_controle_chamados_concremat.html)

Painel para monitorar chamados, prazos, situações e distribuição das demandas, reunindo filtros, indicadores, análises e tabela detalhada.

### GTS, Gestão de Treinamentos de Segurança

Arquivo: [`index.html`](index.html)

O arquivo `index.html` não é apenas mais um dashboard deste ecossistema. Ele é o aplicativo **GTS, Gestão de Treinamentos de Segurança**, um projeto próprio que também é apresentado em outra branch deste repositório.

Ele foi mantido nesta branch porque participa do ecossistema de QSMS e demonstra uma evolução da mesma estratégia de aplicações HTML. Entretanto, sua documentação principal deve apontar para a branch específica do GTS, pois o aplicativo possui escopo mais amplo e fluxos de importação e exportação mais complexos.

Entre os recursos identificados no GTS estão:

- importação do arquivo principal de treinamentos;
- importação simultânea de arquivos de colaboradores;
- detecção de tabelas e planilhas por nomes e cabeçalhos;
- leitura de pastas e subpastas quando suportada pelo navegador;
- consolidação de treinamentos internos e externos;
- visão unificada por colaborador;
- pesquisa por treinamento;
- auditoria e conformidade com matriz de obrigatoriedade;
- alertas de vencimento;
- relatórios para auditoria;
- geração e download de planilhas Excel;
- regras de correspondência aproximada entre nomes de treinamentos.

> Depois de criar a branch do GTS, substitua `LINK_DA_BRANCH_GTS` abaixo pelo endereço definitivo:
>
> **Projeto completo:** [Acessar a branch do GTS](LINK_DA_BRANCH_GTS)

## Arquitetura

As aplicações seguem uma abordagem client-side:

```text
HTML + CSS + JavaScript
          |
          +-- File API do navegador
          +-- leitor de planilhas XLSX
          +-- transformação e normalização
          +-- regras de negócio
          +-- SVG, Canvas ou elementos HTML para visualizações
          +-- renderização de KPIs e tabelas
          `-- impressão ou exportação, quando implementada
```

### Características da solução

- execução sem backend próprio;
- abertura direta no navegador;
- atualização sob demanda pelo usuário;
- distribuição simples por arquivo;
- rápida adaptação de layout e regras;
- estado inicial sem dados operacionais;
- filtros gerados após a importação;
- design responsivo;
- processamento local dos arquivos selecionados;
- identidade visual consistente entre os painéis.

## Tecnologias

- HTML5
- CSS3
- JavaScript
- File API
- SheetJS / XLSX
- SVG e recursos nativos do navegador
- Excel como formato de interoperabilidade
- SharePoint como origem operacional das exportações

## Como executar

Não há instalação por `pip` ou `npm` para os arquivos desta branch.

1. Baixe ou clone a branch.
2. Abra o HTML desejado em um navegador moderno.
3. Exporte a lista correspondente do SharePoint para Excel.
4. Use o seletor de arquivo do painel.
5. Aguarde a confirmação da importação.
6. Utilize filtros, indicadores, gráficos e tabelas.

Exemplo no PowerShell:

```powershell
start dashboard_agendamento_treinamentos_concremat.html
```

Também é possível servir a pasta localmente:

```powershell
python -m http.server 8000
```

Depois, acesse `http://localhost:8000` e escolha o arquivo. Para o GTS, abra `http://localhost:8000/index.html`.

## Dependências e funcionamento offline

Os dashboards usam JavaScript para interpretar arquivos Excel. Parte dos arquivos incorpora recursos no próprio HTML. O `index.html` carrega o leitor XLSX 0.18.5 por CDN e também importa fontes do Google, portanto pode depender de acesso à internet para esses recursos na forma atual.

Para uma distribuição totalmente offline do GTS, uma melhoria futura é incorporar ou fornecer localmente a biblioteca XLSX e utilizar fontes do sistema, respeitando as licenças aplicáveis.

## Privacidade e segurança

- não publique exportações reais das listas;
- não inclua nomes, e-mails, matrículas, CPFs, contratos ou registros de QSMS no repositório;
- use somente dados sintéticos em demonstrações;
- feche o painel ou recarregue a página para descartar o estado em memória;
- revise relatórios baixados antes de compartilhá-los;
- confirme se o navegador ou extensões instaladas estão autorizados para tratar os dados;
- não apresente o processamento local como controle de acesso;
- não use GitHub Pages com arquivos operacionais reais;
- faça revisão de dependências externas antes de uso em ambientes restritos.

## Uso de inteligência artificial

Ferramentas de IA generativa foram utilizadas para acelerar a prototipação, a evolução visual e a implementação de funcionalidades. A escolha dos indicadores, a interpretação dos processos, as regras de negócio, a experiência de importação, os testes e a validação dos resultados permaneceram sob responsabilidade do desenvolvedor.

O benefício central dessa abordagem foi reduzir o ciclo entre uma necessidade operacional e uma nova versão utilizável do painel, sem abrir mão da revisão humana do código e das regras aplicadas aos dados.

## Por que esta arquitetura foi escolhida

### Rapidez de entrega

Um único arquivo pode ser ajustado e redistribuído rapidamente, favorecendo ciclos curtos de melhoria.

### Baixa fricção para o usuário

A atualização dos dados ocorre por seleção de uma exportação recente, sem exigir edição de consultas ou conhecimento de ferramentas de BI.

### Distribuição leve

Os painéis não dependem de um servidor de aplicação próprio nem de uma publicação individual para cada usuário.

### Flexibilidade visual

HTML, CSS e JavaScript permitem adaptar telas, filtros e interações às necessidades de cada processo.

### Limites conscientemente assumidos

Esta arquitetura não substitui todos os cenários de BI. Governança centralizada, atualização automática, grande escala, versionamento de dados, controle de acesso no nível do relatório e fonte única corporativa podem justificar outras plataformas. A solução foi aplicada a processos em que velocidade, simplicidade e autonomia eram prioritárias.

## Estrutura plana da branch

Todos os arquivos ficam na mesma pasta, conforme o padrão deste repositório:

```text
README.md
NOTICE.md
SECURITY.md
.gitignore
dashboard_agendamento_treinamentos_concremat.html
dashboard_cronograma_eventos_seguranca_trabalho.html
dashboard_cronograma_eventos_seguranca_trabalho_metas.html
central_controle_inspecoes_normalizacao_selecao_inspetores_final_corrigido.html
dashboard_controle_chamados_concremat.html
index.html
```

## Melhorias futuras

- padronizar um módulo interno de importação entre os painéis;
- incorporar o leitor XLSX para operação totalmente offline;
- criar arquivos de exemplo com dados sintéticos;
- adicionar testes automatizados para cabeçalhos e normalização;
- registrar versão e data em cada dashboard;
- gerar inventário dos campos esperados por aplicação;
- aprimorar acessibilidade e navegação por teclado;
- documentar compatibilidade entre navegadores;
- criar validação explícita contra planilhas incompatíveis;
- separar formalmente o GTS em sua branch própria e manter aqui apenas a referência cruzada.

## Competências demonstradas

- aplicações web client-side;
- importação e transformação de planilhas;
- visualização de dados sem backend;
- modelagem de indicadores operacionais;
- UX para usuários não técnicos;
- integração prática com exportações do SharePoint;
- uso responsável de IA no desenvolvimento;
- criação rápida de soluções para QSMS;
- análise de dados estruturados no navegador;
- evolução de dashboards para aplicações mais completas, como o GTS.

## Autor

**Gustavo Freitas Gomes Dumont**

Ecossistema profissional de acompanhamento QSMS apresentado como parte de um portfólio técnico.
