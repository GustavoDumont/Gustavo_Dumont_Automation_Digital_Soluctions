# Backup UNICO People

Projeto de automação e contingência criado para preservar dados corporativos armazenados na UNICO People antes do encerramento do contrato com a plataforma.

A solução teve origem no `App_UNICO.py`, uma aplicação desktop distribuída em formato executável para que a equipe do Núcleo Administrativo pudesse baixar documentos de colaboradores. Quando surgiu a necessidade de backup integral, a base existente foi adaptada para uma operação extraordinária de grande volume.

## Desafio

- mais de **150 GB** de dados;
- apenas **7 dias** para concluir a preservação;
- grande quantidade de registros individuais;
- downloads dependentes da navegação no portal;
- necessidade de organizar, conferir e disponibilizar o acervo;
- risco de perda dos dados após o encerramento do contrato.

## Operação executada

Gustavo Freitas Gomes Dumont e Phillipe Castro, antigo integrante de TI, montaram **5 computadores dedicados** na sala de TI. Os equipamentos executaram versões empacotadas da automação em paralelo e permaneceram funcionando continuamente por **5 dias**.

Após a conclusão, os registros foram consolidados e disponibilizados no SharePoint corporativo. A estratégia permitiu concluir a tarefa dentro do prazo apertado e preservar o acervo necessário.

## Evolução das versões

### App_UNICO

A primeira aplicação já permitia processar vários candidatos, escolher categorias e acompanhar logs. Entre os materiais tratados estavam documentos pessoais, assinaturas, ficha de registro e etiquetas. A aplicação era empacotada em EXE e distribuída aos usuários responsáveis.

### Automação UNICO Multi & Ficha

A versão ampliada adicionou configuração local, repetição controlada, recuperação de sessão, acompanhamento de downloads, pastas por pessoa e modos específicos para documentos, assinaturas e fichas.

### Automação UNICO Assinaturas

A terceira versão foi especializada na extração em massa de assinaturas e envelopes concluídos, permitindo distribuir cargas mais específicas entre as máquinas.

## Arquitetura da operação

```text
Acervo na UNICO People
          |
          v
Lista total dividida em filas exclusivas
          |
   +------+------+------+------+------+
   |      |      |      |      |
  PC 1   PC 2   PC 3   PC 4   PC 5
   |      |      |      |      |
   +------+------+------+------+------+
          |
          v
Pastas organizadas por colaborador
          |
          v
Consolidação e conferência
          |
          v
SharePoint corporativo
```

## Versão GitHub

A branch pública contém uma reconstrução segura da arquitetura:

- interface desktop;
- entrada de pessoas ou identificadores;
- seleção de categorias;
- pasta de saída;
- adapter isolado para o portal;
- relatório consolidado;
- manifests com SHA-256;
- parada controlada;
- scripts para execução e geração de EXE;
- documentação de operação, escalabilidade e segurança.

Por segurança, o método de navegação precisa ser configurado internamente. A versão pública não contém URL real, credenciais ou seletores corporativos.

## Estrutura

```text
app.py
automation.py
storage.py
models.py
OPERACAO_DE_BACKUP.md
ARQUITETURA.md
RUNBOOK.md
README_before_exe.md
SECURITY.md
CHANGELOG.md
```

## Executar

```powershell
py -3 -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements.txt
.venv\Scripts\python.exe src\app.py
```

Ou execute `executar.bat`.

## Gerar EXE

Execute `build_exe.bat`. O arquivo será gerado em `dist\BackupUnicoPeople.exe`.

## Resultado e aprendizados

O projeto demonstrou capacidade de transformar uma automação departamental em uma operação de contingência. O sucesso dependeu da combinação entre software, paralelização, infraestrutura dedicada, divisão de filas, monitoramento e coordenação entre áreas.

O ganho não foi apenas de produtividade. A automação permitiu preservar um acervo corporativo relevante dentro de uma janela crítica, reduzindo o risco de perda de informações após o encerramento do contrato.

## Segurança

As versões históricas analisadas continham credenciais em texto puro. Elas não foram incluídas nesta branch. As senhas correspondentes devem ser substituídas e removidas de versões antigas, executáveis e históricos. Consulte `SECURITY.md`.

## Créditos

- **Gustavo Freitas Gomes Dumont**: desenvolvimento da automação, adaptação das versões e execução da operação.
- **Phillipe Castro**: preparação da infraestrutura, paralelização e suporte técnico à operação de backup.
