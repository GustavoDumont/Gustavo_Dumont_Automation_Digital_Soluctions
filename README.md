# Automação de Faturamento e Postagem de Notas Fiscais

Aplicação desktop desenvolvida em Python para automatizar o envio de notas fiscais em um portal corporativo. A solução permite selecionar documentos por uma interface gráfica, autenticar-se no sistema e executar postagens em lote com acompanhamento de progresso.

> **Nota de portfólio:** este repositório contém apenas uma versão demonstrativa, anonimizada e livre de dados corporativos. URLs, credenciais, contratos, contatos e regras internas reais não fazem parte da publicação.

## Visão geral

A postagem de centenas de notas fiscais era uma atividade operacional repetitiva, com alto consumo de tempo e necessidade de equipes dedicadas durante vários dias. Esta aplicação transformou o fluxo em um processo assistido por automação, permitindo que uma única pessoa prepare os arquivos, inicie o processamento e acompanhe a execução enquanto realiza outras atividades.

A solução está em uso operacional há mais de um ano e consolidou a automação como parte do processo de trabalho.

## Problema resolvido

O processo anterior apresentava desafios como:

- grande volume de documentos a serem publicados;
- execução manual de etapas repetitivas no navegador;
- necessidade de conferência e associação entre arquivos;
- dedicação de várias pessoas durante aproximadamente uma semana em períodos de alto volume;
- baixa disponibilidade da equipe para outras atividades durante a postagem.

## Solução desenvolvida

Foi criada uma aplicação desktop que centraliza a preparação e a execução do processo. O usuário informa suas próprias credenciais, seleciona o contrato e a modalidade de envio, escolhe os documentos e inicia a automação.

O sistema organiza os arquivos, abre um navegador compatível, autentica o usuário e conduz as etapas de postagem com esperas dinâmicas, acompanhamento visual e tratamento de falhas comuns de automação web.

## Principais funcionalidades

- interface gráfica construída com Tkinter;
- seleção de documentos PDF e XML;
- associação automática de pares PDF/XML por identificadores presentes nos nomes dos arquivos;
- suporte a dois fluxos de envio configuráveis;
- seleção de contrato e parâmetros operacionais;
- autenticação automatizada no portal;
- automação do navegador com Selenium WebDriver;
- detecção de sessão expirada e tentativa de reautenticação;
- controle ajustável de velocidade entre as ações;
- barra de progresso e mensagens de status;
- pré-visualização dos arquivos selecionados;
- processamento em thread separada para preservar a responsividade da interface;
- compatibilidade com Chrome instalado ou distribuição portátil;
- preparação para empacotamento como executável com PyInstaller.

## Fluxo de funcionamento

1. O usuário abre a aplicação.
2. Informa suas próprias credenciais de acesso.
3. Seleciona o contrato e a modalidade de envio.
4. Escolhe os arquivos PDF e, quando exigido pelo fluxo, os XML correspondentes.
5. A aplicação valida os dados obrigatórios e apresenta os arquivos selecionados.
6. O processamento é iniciado em segundo plano.
7. O navegador é aberto e a autenticação é realizada.
8. Cada nota é processada de acordo com as regras do fluxo escolhido.
9. O usuário acompanha o andamento pela barra de progresso e pelas mensagens da interface.

## Arquitetura simplificada

```text
Interface Tkinter
      |
      +-- seleção e validação de arquivos
      +-- parâmetros de processamento
      +-- status e barra de progresso
      |
Thread de processamento
      |
      +-- associação PDF/XML
      +-- extração de identificadores
      +-- controle de tentativas
      |
Selenium WebDriver
      |
      +-- autenticação
      +-- navegação no portal
      +-- preenchimento e envio
      +-- recuperação de sessão
```

## Tecnologias utilizadas

- Python
- Tkinter
- Selenium WebDriver
- Google Chrome / Chromium
- Expressões regulares
- Threading
- PyInstaller

## Decisões técnicas relevantes

### Esperas dinâmicas

A automação utiliza `WebDriverWait` e condições esperadas para reduzir a dependência de pausas fixas e tornar a interação mais resiliente às variações de carregamento das páginas.

### Interface responsiva

O processamento é executado em uma thread separada. Atualizações visuais são encaminhadas à thread principal por meio de `root.after`, evitando alterações inseguras dos componentes do Tkinter.

### Associação de documentos

Os arquivos PDF e XML são relacionados por identificadores numéricos encontrados em seus nomes. Isso reduz a necessidade de seleção manual de pares durante a preparação do lote.

### Execução portátil

A aplicação procura automaticamente o Chrome e o ChromeDriver em locais conhecidos, incluindo recursos empacotados e uma estrutura portátil ao lado do executável.

## Impacto operacional

- substituição de um processo intensivo em trabalho manual por um fluxo automatizado;
- redução da necessidade de uma equipe dedicada durante aproximadamente uma semana para grandes lotes;
- possibilidade de operação por uma única pessoa;
- liberação do usuário para outras atividades durante o processamento;
- uso contínuo em ambiente operacional por mais de um ano;
- padronização das etapas de postagem.

Os números detalhados de volume, tempo e produtividade não são divulgados neste repositório por confidencialidade. As afirmações acima refletem o resultado operacional observado, sem expor informações internas.

## Estrutura sugerida do repositório

```text
Faturamento_Concremat_Gustavo_Dumont/
├── README.md
├── LICENSE
├── .gitignore
├── .env.example
├── requirements.txt
├── src/
│   ├── main.py
│   ├── config.py
│   ├── gui.py
│   ├── processor.py
│   ├── file_matching.py
│   └── browser_setup.py
├── tests/
│   ├── test_file_matching.py
│   └── test_identifiers.py
├── examples/
│   ├── README.md
│   └── arquivos_ficticios/
├── docs/
│   ├── architecture.md
│   └── images/
└── scripts/
    └── build_windows.ps1
```

## Instalação para a versão demonstrativa

### Pré-requisitos

- Python 3.10 ou superior;
- Google Chrome ou Chromium compatível;
- ChromeDriver compatível com a versão do navegador;
- acesso autorizado a um ambiente de demonstração.

### Configuração do ambiente

```bash
python -m venv .venv
```

No Windows:

```powershell
.venv\Scripts\Activate.ps1
```

Instale as dependências:

```bash
pip install -r requirements.txt
```

Execute a aplicação:

```bash
python src/main.py
```

> A versão pública não se conecta ao portal corporativo real. Para demonstração, recomenda-se utilizar páginas locais simuladas ou um ambiente de testes autorizado.

## Segurança e privacidade

Este projeto deve seguir, no mínimo, as seguintes práticas antes de qualquer publicação:

- nenhuma credencial deve permanecer no código-fonte;
- credenciais reais anteriormente utilizadas devem ser revogadas ou alteradas;
- usuários devem informar suas próprias credenciais em tempo de execução;
- senhas e tokens não devem ser registrados em logs;
- URLs internas e identificadores corporativos devem ser substituídos por valores fictícios;
- nomes, e-mails, contratos e dados de clientes ou colaboradores devem ser removidos;
- dados de exemplo devem ser integralmente sintéticos;
- arquivos `.env`, logs e documentos fiscais devem estar no `.gitignore`;
- a publicação deve respeitar as políticas de propriedade intelectual e segurança das organizações envolvidas.

Exemplo de `.env.example`, apenas se uma futura versão realmente precisar de configuração externa:

```dotenv
PORTAL_BASE_URL=https://example.invalid/
CHROME_BINARY=
CHROMEDRIVER_PATH=
```

## Testes recomendados

A versão pública pode demonstrar qualidade de engenharia com testes que não dependam do portal real:

- associação correta entre PDF e XML;
- tratamento de arquivos sem identificador;
- prevenção de associação duplicada;
- validação de campos obrigatórios;
- comportamento diante de listas vazias;
- localização do navegador e do driver;
- testes de interface com uma página HTML local simulada.

## Limitações conhecidas

- automações baseadas na interface de páginas web podem exigir manutenção quando o portal altera seletores ou fluxos;
- o pareamento por números no nome do arquivo deve ser reforçado para evitar correspondências ambíguas;
- a compatibilidade entre navegador e ChromeDriver precisa ser controlada;
- o processamento deve registrar resultados por documento sem armazenar informações sensíveis;
- a publicação aberta requer a substituição integral da integração corporativa por uma demonstração simulada.

## Próximas evoluções

- separar interface, regras de negócio e automação web em módulos independentes;
- adicionar relatório final com sucesso, falha e motivo por documento;
- implementar retomada segura após interrupções;
- adicionar testes automatizados e integração contínua;
- validar previamente duplicidades, formatos e pares de documentos;
- substituir seletores frágeis por estratégias mais estáveis;
- criar modo de demonstração com portal local fictício;
- preparar logs estruturados com mascaramento de dados sensíveis.

## Aprendizados demonstrados

Este projeto evidencia experiência prática em:

- automação de processos operacionais reais;
- desenvolvimento de interfaces desktop;
- automação web com Selenium;
- tratamento de arquivos e expressões regulares;
- concorrência e segurança de atualizações de interface;
- distribuição de aplicações Python para usuários não técnicos;
- manutenção de uma solução utilizada continuamente;
- tradução de regras de negócio em software.

## Confidencialidade

O projeto original foi desenvolvido em contexto profissional. A versão deste portfólio deve apresentar apenas arquitetura, técnicas e dados fictícios, sem disponibilizar código proprietário, informações de acesso, regras confidenciais ou dados das empresas envolvidas.

## Autor

**Gustavo Freitas Gomes Dumont**

Projeto de portfólio baseado em uma automação profissional desenvolvida para otimizar a postagem de notas fiscais.
