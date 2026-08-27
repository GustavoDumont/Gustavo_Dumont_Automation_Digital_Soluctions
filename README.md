# Download de Pendências do NACT

Ecossistema de automações em desenvolvimento para apoiar a regularização documental da equipe responsável pelo atendimento ao cliente NACT Vale.

O projeto reúne ferramentas que atuam em três frentes complementares:

1. download e tratamento de relatórios de pendências;
2. organização de documentos nas pastas dos colaboradores;
3. obtenção e estruturação do efetivo a partir da folha de pagamento.

Embora o ecossistema ainda esteja evoluindo, diferentes soluções já estão em produção e geram impacto real. O processo foi mapeado de ponta a ponta, permitindo desenvolver novos módulos de forma incremental e conectar as saídas de uma ferramenta às entradas da seguinte.

## Visão do processo

```text
Portal de pendências
        |
        v
Download dos relatórios
        |
        v
Padronização e consolidação
        |
        +-----------------------+
        |                       |
        v                       v
Folha de pagamento PDF   Pastas e documentos recebidos
        |                       |
        v                       v
Extração do efetivo      Identificação e organização
        |                       |
        +-----------+-----------+
                    |
                    v
       Base de acompanhamento do NACT
                    |
                    v
        Regularização e evidências
```

## Aplicativos

### 1. Automação EGOT

Arquivo: `egot_automacao.py`

Automação do portal de pendências. O aplicativo abre o ambiente autorizado, permite que o usuário conclua a autenticação, seleciona o cliente configurado, percorre os contratos disponíveis e baixa os relatórios de monitoramento mensal.

O módulo também prepara os relatórios para uso operacional:

- identifica contratos e localidades;
- controla arquivos já processados;
- aguarda a conclusão do download;
- padroniza o nome de saída;
- registra sucesso, falha e interrupção;
- mantém a interface responsiva por meio de uma thread de trabalho;
- permite encerrar a execução de forma controlada.

Padrão de nome proposto:

```text
SIGLA_CONTRATO_LOCALIDADE.xlsx
```

A versão pública não contém URL interna, seletores completos ou credenciais. O adaptador precisa ser configurado no ambiente autorizado.

### 2. Organizador de Arquivos de Sondagem

Arquivo: `organizador_arquivos_sondagem.py`

Ferramenta voltada à organização dos documentos recebidos para regularização. O aplicativo compara nomes de arquivos e textos extraídos com as pastas dos colaboradores, calcula similaridade e encaminha cada documento para a pasta e categoria correspondentes.

Categorias previstas:

```text
ATESTADO
FOLHA DE PONTO
HORAS EXTRAS
PERICULOSIDADE
```

O organizador oferece:

- correspondência aproximada de nomes;
- normalização de acentos, pontuação e separadores;
- leitura de texto de PDFs;
- OCR opcional para PDFs digitalizados e imagens;
- cópia ou movimentação;
- modo de simulação antes de alterar arquivos;
- prevenção de sobrescrita;
- suporte a caminhos longos no Windows;
- relatório CSV das ações e itens não classificados.

### 3. Mapeamento FOPAG

Arquivo: `mapeamento_fopag.py`

Extrator de efetivo a partir do PDF da folha de pagamento. O programa identifica a estrutura visual de cada página e transforma os registros em uma base tabular.

Campos extraídos:

```text
ID/chapa
Nome
Salário
Função
Situação
Contrato
Admissão
Demissão
Página
```

O módulo:

- lê palavras e coordenadas do PDF;
- agrupa palavras em linhas;
- reconhece a chapa por padrão de seis dígitos;
- separa nome e função de acordo com a posição horizontal;
- normaliza situações funcionais;
- localiza admissão, demissão, salário e contrato;
- gera Excel formatado;
- mostra os registros na interface antes do uso.

O extrator depende do layout do relatório. Mudanças na folha exigem revisão das coordenadas e expressões regulares.

## Integração entre os módulos

As ferramentas podem operar independentemente, mas o maior valor surge quando são conectadas:

```text
mapeamento_fopag.py
        |
        v
Lista atualizada do efetivo
        |
        +----------------------------+
        |                            |
        v                            v
egot_automacao.py       organizador_arquivos_sondagem.py
        |                            |
        v                            v
Relatórios de pendência      Documentos organizados
        |                            |
        +-------------+--------------+
                      |
                      v
       Acompanhamento e regularização
```

## Estrutura da branch

Todos os arquivos ficam na raiz, conforme o padrão solicitado:

```text
README.md
ARQUITETURA.md
PROCESSO_MAPEADO.md
ROADMAP.md
SECURITY.md
CHANGELOG.md
requirements.txt
requirements-dev.txt
executar_egot.bat
executar_organizador.bat
executar_fopag.bat
build_executaveis.bat
egоt_config.example.json
egоt_automacao.py
organizador_arquivos_sondagem.py
mapeamento_fopag.py
.gitignore
```

## Instalação

```powershell
py -3 -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements.txt
```

## Execução

```text
executar_egot.bat
executar_organizador.bat
executar_fopag.bat
```

## Estado do projeto

### Em produção

- obtenção automatizada de relatórios de pendências;
- apoio à organização de documentos;
- extração de efetivo a partir da folha;
- geração de bases e relatórios operacionais.

### Em desenvolvimento

- consolidação automática entre efetivo, pendências e documentos;
- regras mais fortes de reconciliação;
- retomada idempotente;
- painel de acompanhamento;
- trilha de auditoria centralizada;
- indicadores de regularização e SLA;
- tratamento de novos tipos documentais.

## Resultado

O ecossistema reduz atividades repetitivas de download, leitura, classificação e organização. Além do ganho de produtividade, cria uma base mais segura para acompanhar pendências, identificar colaboradores no escopo e direcionar documentos para regularização.

## Autor

**Gustavo Freitas Gomes Dumont**
