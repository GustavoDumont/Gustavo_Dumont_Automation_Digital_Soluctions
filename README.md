# I Love Cmat

Suite desktop portátil para automação e manipulação de arquivos, desenvolvida para apoiar o núcleo administrativo com operações em lote, sem limite explícito de quantidade de arquivos imposto pela aplicação.

A solução foi distribuída como executável portátil para mais de 40 pessoas e reúne, em uma única interface, recursos de PDF, imagens e documentos que complementam ferramentas online e atendem necessidades específicas do ambiente administrativo.

> **Importante:** os códigos desta branch foram preservados exatamente como enviados. A versão `styled` corresponde à edição distribuída em produção. A versão `V2` é experimental e contém funções que ainda não foram liberadas para os usuários.

## Contexto

Atividades de manipulação documental surgem frequentemente em situações urgentes: unir arquivos, reorganizar páginas, converter formatos, aplicar marcas d'água, renomear lotes ou preparar documentos para envio. Ferramentas online podem impor limites, cobrar por determinados recursos ou não oferecer o comportamento específico exigido pelo processo.

O I Love Cmat foi criado como uma central local de produtividade para reduzir essa fricção. O usuário baixa um pacote ZIP compartilhado, extrai os arquivos e executa a aplicação sem precisar trabalhar com o código-fonte.

## Impacto e adoção

- distribuição para mais de 40 pessoas do núcleo administrativo;
- entrega em formato `.exe` portátil;
- centralização de várias rotinas em uma única interface;
- processamento local e em lote;
- ausência de limite explícito de arquivos imposto pela aplicação;
- redução da dependência de serviços online e recursos pagos;
- apoio recorrente em demandas documentais urgentes;
- manutenção de uma versão produtiva separada de uma linha experimental.

Não são divulgados documentos, métricas internas ou dados de usuários.

## Versões incluídas

### Versão distribuída

Arquivo: [`src/production/I_Love_Cmat_styled.py`](src/production/I_Love_Cmat_styled.py)

Esta é a versão convertida em executável e distribuída aos usuários. Ela prioriza recursos já considerados adequados para utilização cotidiana.

Funcionalidades destacadas no código:

- aplicação de marca d'água em PDF;
- união de vários PDFs;
- organização e reordenação de páginas;
- divisão de PDFs por intervalos;
- renomeação de arquivos em lote;
- conversão de imagens para PDF;
- conversão de PDF para imagem;
- conversão de Word para PDF;
- conversão de PDF para Word com OCR;
- interface em modo escuro;
- pré-visualização de páginas;
- processamento em threads para preservar a responsividade.

### Versão experimental de uso pessoal

Arquivo: [`src/experimental/I_Love_Cmat_V2.py`](src/experimental/I_Love_Cmat_V2.py)

Esta versão amplia a suíte com operações de manipulação mais fina. Ela não deve ser apresentada como versão homologada para produção.

Além das bases anteriores, o código contém recursos como:

- rotação de páginas;
- conversão de PDF para Excel;
- conversão de PDF para PowerPoint;
- proteção de PDFs com senha;
- remoção autorizada de senha, mediante conhecimento da senha atual;
- posicionamento interativo de marca d'água;
- edição de PDF com pré-visualização;
- inserção visual de assinatura;
- compressão e descompressão de PDF;
- renomeação assistida por OCR e correspondência aproximada;
- tratamento de caminhos duplicados para evitar sobrescrita acidental.

## Estratégia de estabilidade

A separação entre as versões demonstra uma decisão importante de engenharia:

1. manter uma edição estável para os usuários;
2. testar recursos de maior impacto em uma versão pessoal;
3. evitar liberar manipulações sensíveis antes de atingir confiança suficiente;
4. preservar a continuidade da ferramenta já adotada;
5. preparar uma evolução futura sem comprometer o uso atual.

## Arquitetura resumida

```text
Usuário
  |
  v
Interface CustomTkinter
  |
  +-- operações PDF
  +-- conversões de formato
  +-- OCR e renomeação
  +-- pré-visualização e edição
  |
  v
Processamento local em threads
  |
  +-- pypdf / PyMuPDF
  +-- Pillow / ReportLab
  +-- Tesseract / pdf2image
  +-- openpyxl / python-docx / python-pptx
  |
  v
Arquivos de saída em pasta escolhida pelo usuário
```

Consulte também [`docs/architecture.md`](docs/architecture.md).

## Tecnologias utilizadas

- Python
- CustomTkinter e Tkinter
- pypdf
- PyMuPDF
- Pillow
- ReportLab
- pdf2image
- Tesseract OCR e pytesseract
- RapidFuzz
- pdfplumber
- openpyxl
- python-docx
- docx2pdf
- python-pptx
- PyInstaller

## Estrutura da branch

```text
i-love-cmat/
├── README.md
├── NOTICE.md
├── SECURITY.md
├── .gitignore
├── requirements.txt
├── requirements-development.txt
├── src/
│   ├── production/
│   │   └── I_Love_Cmat_styled.py
│   └── experimental/
│       └── I_Love_Cmat_V2.py
├── docs/
│   ├── architecture.md
│   ├── release-notes.md
│   └── images/
└── examples/
    └── README.md
```

## Instalação para desenvolvimento

```bash
python -m venv .venv
```

No Windows:

```powershell
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

Execute a versão distribuída:

```powershell
python src/production/I_Love_Cmat_styled.py
```

Execute a versão experimental:

```powershell
python src/experimental/I_Love_Cmat_V2.py
```

## Dependências externas e opcionais

Alguns recursos podem depender de componentes instalados ou empacotados separadamente:

- Tesseract OCR e arquivos de idioma;
- Poppler para conversão de páginas de PDF em imagens;
- Microsoft Word para o fluxo `docx2pdf` no Windows;
- LibreOffice como alternativa de conversão em ambientes compatíveis.

A ausência de uma dependência opcional pode afetar apenas as funções relacionadas a ela.

## Distribuição portátil

A versão utilizada pelos usuários foi empacotada como executável portátil e compartilhada dentro de um arquivo ZIP por link corporativo. O repositório não inclui o executável, binários de terceiros ou o pacote distribuído.

Uma estratégia de release recomendada é:

```text
I-Love-Cmat/
├── I_Love_Cmat.exe
├── tesseract/
├── poppler/
└── LEIA-ME.txt
```

A composição exata depende das funções incluídas no build e das licenças dos componentes redistribuídos.

## Cuidados de uso

- mantenha cópia dos arquivos originais antes de operações de edição;
- teste versões experimentais somente com documentos descartáveis ou duplicados;
- revise visualmente os resultados de OCR e conversões;
- utilize a remoção de senha somente em documentos para os quais exista autorização;
- evite interromper a aplicação durante a gravação de arquivos;
- valide espaço em disco antes de processar grandes lotes;
- não disponibilize documentos internos em exemplos públicos.

## Qualidade e decisões técnicas

### Processamento local

A manipulação ocorre no computador do usuário, evitando a necessidade técnica de enviar os documentos a um serviço web para executar as funções do aplicativo.

### Operações em lote

A interface foi desenhada para selecionar múltiplos arquivos e automatizar tarefas repetitivas, característica especialmente relevante para rotinas administrativas.

### Responsividade

Operações demoradas são encaminhadas a threads, reduzindo o bloqueio da interface durante processamentos maiores.

### Dependências opcionais

O código verifica a disponibilidade de bibliotecas e componentes para preservar o funcionamento parcial quando determinado conversor não está disponível.

### Separação entre produção e laboratório

A edição distribuída permanece isolada da versão com funções experimentais, limitando o risco de liberar mudanças ainda não homologadas.

## Roadmap

Uma terceira versão em HTML está planejada para ampliar o acesso por navegador. Antes dessa evolução, devem ser avaliados:

- arquitetura de processamento local ou no servidor;
- privacidade e retenção de documentos enviados;
- limites de tamanho e tempo de processamento;
- autenticação e controle de acesso;
- custos operacionais;
- compatibilidade das bibliotecas Python com o ambiente escolhido;
- filas para tarefas longas;
- descarte seguro dos arquivos temporários;
- observabilidade e tratamento de falhas.

> Um front-end hospedado em Cloudflare não implica, por si só, que todas as operações pesadas de PDF, OCR e conversão possam ser executadas no mesmo ambiente. A arquitetura da versão web deverá ser validada antes da implementação.

## Próximas melhorias recomendadas

- adicionar número de versão e changelog formal ao executável;
- gerar hashes dos pacotes distribuídos;
- assinar digitalmente o executável quando possível;
- criar testes automatizados para operações destrutivas;
- implementar modo de simulação e confirmação antes de sobrescrever arquivos;
- produzir relatório de arquivos processados, ignorados e com erro;
- criar exemplos sintéticos e GIFs de demonstração;
- separar interface, serviços e operações de arquivo em módulos;
- documentar processo reproduzível de build com PyInstaller.

## Competências demonstradas

- desenvolvimento de aplicação desktop utilizada por dezenas de pessoas;
- transformação de necessidades administrativas em produto interno;
- processamento e conversão de documentos;
- OCR e correspondência textual aproximada;
- interface para usuários não técnicos;
- empacotamento e distribuição de aplicações Python;
- gestão de estabilidade entre produção e experimentação;
- manutenção evolutiva baseada no uso real.

## Autor

**Gustavo Freitas Gomes Dumont**

Projeto profissional apresentado como parte de um portfólio técnico.
