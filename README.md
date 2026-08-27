# Gestão de Férias | Consolidador TOTVS

Sistema web criado para suprir a ausência de uma visão gerencial de férias após a migração da base de dados da Concremat para o TOTVS.

A solução transforma relatórios operacionais em uma experiência visual de acompanhamento, com gráfico de Gantt, tabelas consolidadas, filtros e rastreabilidade. A evolução do projeto seguiu uma decisão clara de produto: reduzir complexidade de entrada e tornar o controle mais funcional para quem realmente atua no processo.

## Contexto

A mudança para o TOTVS atendeu à operação da base de dados, mas a análise inicial de requisitos não contemplou um relatório visual específico para gestão de férias. Isso criou uma lacuna entre a disponibilidade dos dados e a capacidade dos responsáveis de interpretar períodos, programações, saldos e sobreposições com rapidez.

O projeto nasceu para preencher essa lacuna sem exigir uma nova infraestrutura. Os relatórios são importados diretamente no navegador e processados localmente pela aplicação HTML.

## Evolução do produto

### Primeira versão: integração ampla

O arquivo [`Gestao de Ferias.html`](Gestao%20de%20Ferias.html) representa a primeira abordagem. Essa versão reúne três fontes:

- planilha de férias;
- efetivo consolidado;
- histórico de férias.

A versão original oferece dashboard, férias a vencer, programação, ficha individual, alertas para gestores, preparação de mensagens para Outlook, filtros por contrato, indicadores de urgência, análise de sobreposições e exportação CSV.

A primeira versão comprovou o valor do controle integrado, mas também mostrou o custo operacional de depender de várias bases e de regras de conciliação entre arquivos.

### Versão consolidada V6: menos fontes, mais utilidade

O arquivo [`Gestao_de_Ferias_Consolidador_v6.html`](Gestao_de_Ferias_Consolidador_v6.html) é a evolução orientada à experiência do usuário. A aplicação trabalha somente com relatórios originados no TOTVS, eliminando a dependência de bases paralelas.

Na interface atual, o usuário pode importar:

- **Férias a Vencer**, para períodos aquisitivos, limites, saldos, status e programações;
- **Listagem de Férias**, para datas de início e fim, pagamento, dias, abono, situação e construção do Gantt.

A V6 também permite abrir e exportar um arquivo de salvamento para preservar o conjunto de pessoas acompanhadas e o histórico de uso do painel.

> “Usar apenas os dados do TOTVS” significa usar somente fontes extraídas do sistema corporativo. A interface atual aceita dois relatórios TOTVS complementares, que podem ser usados individualmente ou em conjunto.

## Funcionalidades da V6

- inicia sem dados, evitando demonstrações fictícias ou informações desatualizadas;
- importa diretamente arquivos Excel exportados do TOTVS;
- mantém o leitor XLSX dentro do próprio HTML;
- informa sucesso ou erro de importação;
- consolida registros de pessoas presentes em uma ou nas duas fontes;
- permite marcar as pessoas que devem ser acompanhadas;
- apresenta Gantt anual com linha do dia atual;
- diferencia programações marcadas, pagas, finalizadas e outras situações;
- filtra por ano, lotação, cargo e colaborador;
- permite ordenar o Gantt pela data inicial;
- apresenta tabelas de pessoas, períodos aquisitivos e férias;
- oferece filtros por coluna;
- mostra fonte e disponibilidade dos dados de cada pessoa;
- exporta e reabre um arquivo de salvamento;
- preserva histórico de importações;
- possui menu lateral compacto para uso em notebook;
- segue a identidade visual Concremat, com azul institucional e destaque em laranja.

## Fluxo de uso

```text
TOTVS
  |-- Relatório Férias a Vencer
  `-- Relatório Listagem de Férias
                 |
                 v
        Importação no navegador
                 |
                 v
      Normalização e consolidação
                 |
        +--------+---------+
        |                  |
        v                  v
  Gantt anual        Tabelas de apoio
        |                  |
        +--------+---------+
                 |
                 v
     acompanhamento pelos responsáveis
```

## Princípio de produto

A evolução não buscou adicionar o maior número possível de telas. O objetivo foi remover dependências, reduzir etapas e fortalecer as visualizações que sustentam a tomada de decisão.

```text
V1: mais fontes + mais conciliações + mais recursos operacionais
V6: fontes oficiais TOTVS + fluxo simplificado + Gantt mais preciso
```

A principal melhoria foi transformar um consolidado tecnicamente amplo em uma ferramenta mais direta, previsível e fácil de manter.

## Arquivos da branch

```text
README.md
ARCHITECTURE.md
CHANGELOG.md
SECURITY.md
NOTICE.md
.gitignore
index.html
Gestao_de_Ferias_Consolidador_v6.html
Gestao de Ferias.html
```

- `index.html`: cópia da V6 pronta para GitHub Pages ou Cloudflare Pages.
- `Gestao_de_Ferias_Consolidador_v6.html`: versão atual com nome original.
- `Gestao de Ferias.html`: versão histórica, preservada para demonstrar a evolução do produto.

## Como executar

A aplicação é estática. Abra `index.html` em um navegador moderno ou sirva a pasta localmente:

```powershell
python -m http.server 8000
```

Depois acesse:

```text
http://localhost:8000
```

## Publicação

A branch pode ser publicada diretamente por GitHub Pages ou Cloudflare Pages, sem etapa de build. Defina a raiz do projeto como diretório de saída.

## Privacidade

Não inclua relatórios reais do TOTVS no repositório. Esses arquivos podem conter nomes, matrículas, cargos, lotações e informações funcionais. O processamento local no navegador reduz exposição técnica, mas não substitui os controles organizacionais de acesso e compartilhamento.

## Uso de inteligência artificial

A inteligência artificial foi utilizada como apoio à prototipação, refinamento de interface, revisão de regras e documentação. A definição do problema, a arquitetura, as regras de negócio, a validação e a evolução orientada aos usuários permaneceram sob responsabilidade humana.

## Resultado

O projeto recuperou uma capacidade gerencial que não estava disponível no novo fluxo do TOTVS. A solução permite visualizar programação, períodos e concentração de férias de forma mais rápida do que na planilha original, sem criar uma nova base corporativa e sem exigir instalação.

## Autor

**Gustavo Freitas Gomes Dumont**
