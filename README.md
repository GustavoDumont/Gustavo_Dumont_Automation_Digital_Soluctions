# Automação de Geração e Postagem de Notas de Débito

Conjunto de automações desenvolvido para reduzir drasticamente o tempo de emissão e postagem de notas de débito relacionadas a despesas reembolsáveis. A solução cobre dois estágios do processo: geração no SAP e postagem em massa em um portal web.

> **Portfólio anonimizado:** esta branch apresenta arquitetura, evolução técnica e código sanitizado. Credenciais, URLs, empresas, parâmetros SAP, documentos e dados reais foram removidos ou substituídos.

## Impacto operacional

Antes da automação, a emissão e postagem exigiam horas de trabalho repetitivo e participação de uma equipe. O fluxo automatizado passou a:

- gerar notas de débito no SAP em até aproximadamente cinco minutos, conforme o volume e a disponibilidade do ambiente;
- processar múltiplos documentos em lote;
- preencher automaticamente mais de dez campos por nota;
- permitir que uma única pessoa conduza uma atividade antes distribuída entre vários integrantes;
- reduzir erros de digitação e padronizar o preenchimento;
- transformar uma operação de horas em uma execução de poucos minutos.

Os resultados são descritos sem divulgar volumes, valores ou informações internas.

## Evolução da solução

### 1. Geração no SAP com VBScript

O script `src/sap/emitir_notas_de_debito.vbs` lê dados da planilha ativa do Excel e automatiza etapas de criação do documento, classificação contábil, condições, textos, faturamento e salvamento da saída.

Destaques:

- conexão com uma sessão já autorizada do SAP GUI;
- leitura sequencial das linhas do Excel;
- espera pelo término do processamento do SAP;
- tratamento de pop-ups e áreas de venda;
- preenchimento de dados financeiros e contábeis;
- geração e nomeação automática dos arquivos de saída;
- processamento de várias linhas em sequência.

### 2. Postagem web com OCR

A primeira versão Python, `src/web_ocr/app_ocr.py`, interpreta os PDFs com OCR e utiliza os dados extraídos para preencher os formulários do portal.

Tecnologias e técnicas:

- conversão de PDF em imagem com Poppler;
- OCR com Tesseract;
- pré-processamento com OpenCV e NumPy;
- extração de campos com expressões regulares;
- validação de campos obrigatórios e novas tentativas;
- automação web com Selenium;
- interface desktop com CustomTkinter;
- processamento em segundo plano para manter a interface responsiva.

### 3. Postagem web orientada por Excel

A versão mais recente, `src/web_excel/app_excel.py`, utiliza uma planilha estruturada como fonte de dados. A mudança simplifica o fluxo e elimina a dependência do reconhecimento visual para campos críticos.

Essa evolução resolveu especialmente a recorrente confusão entre os caracteres `1` e `7`, comum em determinados documentos processados por OCR. Em vez de tentar mascarar indefinidamente a incerteza do reconhecimento, a arquitetura passou a consumir dados previamente estruturados e verificáveis.

## Por que manter as duas versões

As versões demonstram uma evolução de engenharia:

1. automatização de um processo manual;
2. extração inteligente a partir de documentos não estruturados;
3. identificação de uma limitação real do OCR;
4. substituição por uma fonte estruturada mais confiável;
5. preservação da versão OCR como referência e alternativa para cenários sem planilha.

## Arquitetura

```text
Planilha Excel
   |                     PDFs
   |                      |
   |                OCR e validação
   |                      |
   +----------+-----------+
              |
      Aplicação Python
      CustomTkinter + Selenium
              |
       Portal de demonstração

Planilha Excel
      |
     VBScript
      |
 SAP GUI Scripting
      |
Geração e faturamento
```

Mais detalhes em [`architecture.md`](architecture.md).

## Estrutura da branch

```text
geracao-notas-de-debito/
├── README.md
├── NOTICE.md
├── SECURITY.md
├── .env.example
├── .gitignore
├── requirements.txt
├── requirements-development.txt
├── src/
│   ├── sap/
│   │   └── emitir_notas_de_debito.vbs
│   ├── web_ocr/
│   │   └── app_ocr.py
│   └── web_excel/
│       └── app_excel.py
├── docs/
│   ├── architecture.md
│   └── images/
└── examples/
    └── README.md
```

## Dependências

A versão Excel utiliza:

- Python 3.13.2 no ambiente original;
- Selenium;
- undetected-chromedriver;
- pandas;
- openpyxl;
- CustomTkinter.

A versão OCR acrescenta:

- pdf2image;
- pytesseract;
- OpenCV;
- NumPy;
- Pillow;
- Poppler e Tesseract como dependências externas.

## Instalação

```bash
python -m venv .venv
```

No Windows:

```powershell
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

Para executar a versão baseada em Excel:

```powershell
python src/web_excel/app_excel.py
```

Para executar a versão OCR:

```powershell
python src/web_ocr/app_ocr.py
```

> A versão pública aponta para `https://example.invalid` e não está configurada para um portal real.

## Configuração segura

Copie `.env.example` para `.env` somente no ambiente local. O código lê variáveis do sistema operacional. Não publique credenciais.

```dotenv
APP_EMAIL=
APP_TOKEN=
APP_BASE_URL=https://example.invalid
```

## Formato esperado da planilha

A versão pública preserva os nomes técnicos encontrados na implementação, mas os dados de exemplo devem ser inteiramente fictícios. A planilha deve representar, entre outros, identificadores das partes, número do documento, datas, valor, contrato, pedido, item e referências internas necessárias ao formulário.

Não publique uma planilha real. Crie um modelo sintético depois de confirmar os cabeçalhos aceitos pela versão final.

## Robustez implementada

- esperas explícitas para elementos web;
- repetição controlada em falhas transitórias;
- tratamento de elementos obsoletos ou interceptados;
- ajuste automático à versão principal do Chrome;
- validação de dados ausentes;
- normalização de números, datas e valores monetários;
- processamento em thread separada;
- logs de andamento na interface.

## Segurança e limitações

- SAP GUI Scripting deve ser utilizado apenas em ambiente autorizado;
- seletores web e IDs SAP podem mudar com atualizações dos sistemas;
- automações de interface exigem manutenção e testes de regressão;
- a versão OCR está sujeita a erros de reconhecimento;
- a versão Excel depende da qualidade e do esquema da planilha;
- os parâmetros de negócio e a integração real foram removidos;
- o código público não deve ser tratado como solução pronta para produção.

## Melhorias futuras

- criar uma camada única de validação compartilhada pelas versões OCR e Excel;
- gerar relatório final por documento, com sucesso, falha e causa;
- implementar retomada após interrupções;
- adicionar testes para normalização, datas, moeda e parsing OCR;
- criar portal e planilhas fictícias para demonstração reproduzível;
- modularizar interface, domínio e adaptadores de automação;
- registrar métricas agregadas de tempo e produtividade sem dados sensíveis.

## Competências demonstradas

- automação SAP GUI com VBScript;
- integração entre Excel e sistemas corporativos;
- automação web com Selenium;
- OCR e visão computacional aplicada;
- tratamento de dados estruturados e não estruturados;
- interfaces desktop para usuários não técnicos;
- diagnóstico de falhas e evolução arquitetural;
- transformação de processos operacionais de alto volume.

## Autor

**Gustavo Freitas Gomes Dumont**

Projeto apresentado como estudo técnico anonimizado de uma automação profissional.
