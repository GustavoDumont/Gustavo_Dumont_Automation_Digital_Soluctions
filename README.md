# Contas a Pagar | Gestão, Alertas e SLAs

Ecossistema desenvolvido para facilitar a gestão, a atuação diária, os alertas e o acompanhamento de SLAs da equipe de Contas a Pagar.

A solução combina Power Apps, listas corporativas, fluxos automatizados e dados periódicos do SAP. O objetivo não é apenas registrar faturas, mas organizar o processo de ponta a ponta, gerar visibilidade sobre pendências e reduzir riscos operacionais.

## Contexto

A rotina de Contas a Pagar envolve aprovações, cadastros, lançamentos, vencimentos, pedidos, documentos SAP, formas de pagamento e tratativas com diferentes responsáveis. Quando essas informações ficam distribuídas entre planilhas, mensagens e controles manuais, aumentam os riscos de atraso, perda de contexto, duplicidade e atuação fora do SLA.

O projeto foi criado para centralizar esse acompanhamento e transformar dados operacionais em ações priorizadas.

## Componentes do sistema

### CaP - Contas a Pagar

Aplicativo em Power Apps que opera duas listas integradas:

1. **Lista do fluxo de aprovação de faturas**
   - recebe e acompanha a tramitação das faturas;
   - registra o estágio de aprovação;
   - fornece a origem do processo que alimenta a operação.

2. **Lista de cadastro e gerenciamento**
   - concentra o acompanhamento operacional;
   - registra informações necessárias para atuação da equipe;
   - permite atualização e consulta pelo mesmo Power Apps.

As duas listas são operadas por uma interface única, reduzindo alternância entre fontes e tornando a experiência mais consistente.

### Power Automate

Fluxos recorrentes são executados diariamente para:

- identificar itens que exigem atuação;
- enviar lembretes das listas;
- reforçar prazos e SLAs;
- reduzir dependência de acompanhamento manual;
- enviar alertas relacionados ao relatório ME2N.

### Relatório ME2N do SAP

O relatório ME2N é atualizado duas vezes por mês pela equipe e utilizado pelos fluxos de lembrete. Essa camada amplia o controle para pedidos e informações que ainda dependem da extração periódica do SAP.

### Dashboard HTML

O arquivo [`Contas a Pagar.html`](Contas%20a%20Pagar.html) oferece uma camada analítica local para a planilha operacional.

Principais recursos identificados:

- inicialização sem dados;
- importação de Excel e CSV;
- leitor XLSX interno, sem dependência de internet;
- filtros por mês, fornecedor, centro, situação, centro de custo, forma de pagamento e período;
- pesquisa na base detalhada;
- indicadores de valor total, lançamentos, valor pago, valor vencido e ticket médio;
- análise de valores por mês;
- distribuição por situação financeira;
- maiores fornecedores;
- análise por centro de custo;
- formas de pagamento;
- distribuição por município;
- leitura rápida com insights;
- tabela detalhada;
- exportação CSV;
- identidade visual Concremat com destaque em laranja;
- barra lateral compacta e visível em notebook.

## Arquitetura funcional

```text
                       +-----------------------------+
                       | Fluxo de aprovação de       |
                       | faturas                     |
                       +--------------+--------------+
                                      |
                                      v
+---------------------+     +---------+----------+     +----------------------+
| Usuários da equipe  |<--->| CaP - Contas a     |<--->| Lista de cadastro e  |
| de Contas a Pagar   |     | Pagar / Power Apps |     | gerenciamento        |
+---------------------+     +---------+----------+     +----------+-----------+
                                      |                           |
                                      +-------------+-------------+
                                                    |
                                                    v
                                      +-------------+-------------+
                                      | Power Automate            |
                                      | lembretes diários e SLAs  |
                                      +-------------+-------------+
                                                    ^
                                                    |
                                      +-------------+-------------+
                                      | Relatório ME2N do SAP     |
                                      | atualização 2 vezes/mês   |
                                      +---------------------------+
```

O dashboard HTML funciona como uma camada complementar de análise. O dashboard não substitui o Power Apps nem atualiza o SAP.

## Ganhos observados

Mesmo em fase recente de uso, a solução já demonstra ganhos em:

- produtividade da equipe;
- centralização das informações;
- priorização de pendências;
- segurança do processo;
- rastreabilidade das atuações;
- redução de lembretes manuais;
- acompanhamento de vencimentos;
- visibilidade dos SLAs;
- padronização do cadastro;
- redução da dependência de controles individuais.

## Evolução proposta

Após a conclusão dos testes e a estabilização das regras, existe potencial para conexão direta com o SAP.

A integração futura poderá reduzir a dependência da atualização manual do ME2N e melhorar a frequência dos dados. Essa etapa deve ser tratada como evolução de arquitetura, sujeita a validações de segurança, perfis de acesso, integração, governança e disponibilidade das interfaces corporativas.

```text
Estado atual:
SAP -> extração ME2N -> atualização pela equipe -> Power Automate

Possível evolução:
SAP -> integração controlada -> camada de dados -> Power Apps / Power Automate
```

A conexão direta com o SAP é uma possibilidade de expansão, não uma funcionalidade implementada nesta versão.

## Arquivos da branch

```text
README.md
ARCHITECTURE.md
CHANGELOG.md
DATA_DICTIONARY.md
SECURITY.md
NOTICE.md
.gitignore
index.html
Contas a Pagar.html
```

- `index.html`: cópia do dashboard pronta para hospedagem estática.
- `Contas a Pagar.html`: dashboard com o nome original.
- `DATA_DICTIONARY.md`: referência dos principais campos aceitos pelo dashboard.

Os pacotes exportados do Power Apps, as definições dos fluxos do Power Automate, as listas e os dados operacionais não estão incluídos nesta branch.

## Como executar o dashboard

Abra `index.html` em um navegador moderno. Também é possível servir a pasta localmente:

```powershell
python -m http.server 8000
```

Acesse:

```text
http://localhost:8000
```

O dashboard começa vazio. Importe diretamente uma planilha Excel exportada da fonte operacional ou um arquivo CSV compatível.

## Publicação

O dashboard é estático e pode ser publicado sem build no GitHub Pages ou Cloudflare Pages. Antes de qualquer publicação, considere que o projeto descreve uma solução corporativa e pode não ser adequado para exposição pública integral.

## Privacidade

Não publique a planilha operacional, exportações das listas, relatórios ME2N ou arquivos com dados reais. Esses materiais podem conter fornecedores, códigos, notas fiscais, valores, pedidos, documentos SAP, responsáveis, municípios, observações e datas de pagamento.

## Uso de inteligência artificial

A inteligência artificial foi utilizada como apoio à prototipação, refinamento da interface, revisão de regras e documentação. A definição do problema, as regras do processo, as decisões de arquitetura, a validação e a implantação permaneceram sob responsabilidade humana.

## Resultado

A solução cria uma camada de gestão sobre um processo sensível e recorrente. A combinação entre aplicação, listas, automações e visualização analítica melhora a capacidade de agir antes do vencimento, acompanhar responsabilidades e preservar o histórico necessário para uma operação mais segura.

## Autor

**Gustavo Freitas Gomes Dumont**
