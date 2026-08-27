# Buscador de ASO em lote

Primeiro projeto de automação desenvolvido para reduzir o trabalho repetitivo de responsáveis que precisavam baixar muitos Atestados de Saúde Ocupacional, um por colaborador.

O aplicativo recebe uma lista de CPFs, autentica um usuário autorizado no portal da clínica, pesquisa cada registro e solicita o download do documento disponível. O ganho principal aparece em demandas de grande volume, nas quais repetir manualmente a mesma navegação para cada CPF consome tempo e aumenta o risco de omissões.

## Evolução

O projeto teve diferentes versões:

- **executável portátil**, distribuído aos responsáveis para uso sem contato direto com o código;
- **versão para pendrive**, organizada para executar junto de um navegador portátil;
- **versão GitHub**, revisada para portfólio, sem credenciais embutidas e com estrutura reproduzível.

Os arquivos históricos originais não foram publicados porque uma das versões continha credenciais em texto puro e ambas dependiam de detalhes internos do portal. A versão desta branch preserva a lógica do produto, mas remove segredos e adota configurações locais.

## Funcionalidades

- entrada de CPFs, um por linha;
- importação de CPFs a partir de TXT ou CSV;
- normalização, validação básica e remoção de duplicidades;
- autenticação com credenciais informadas em tempo de execução;
- pesquisa sequencial;
- tratamento de registro não encontrado, documento ausente, timeout e erro do navegador;
- pasta de download configurável;
- log visual com CPF mascarado;
- relatório CSV da execução;
- encerramento controlado do navegador;
- opção de execução local, portátil ou empacotada em EXE.

## Estrutura

```text
src/buscador_aso.py       aplicação revisada
executar_portatil.bat     inicializa ambiente local na própria pasta
build_exe.bat             gera o EXE com PyInstaller
requirements.txt          dependência de execução
requirements-dev.txt      dependência adicional para empacotamento
.env.example              exemplo sem dados reais
SECURITY.md               requisitos de privacidade e autorização
CHANGELOG.md              evolução do projeto
LICENSE                    condições de uso
```

## Requisitos

- Windows 10 ou 11;
- Python 3.10 ou superior para executar o código;
- Google Chrome instalado;
- acesso autorizado ao portal;
- compatibilidade dos seletores com a versão atual da página.

O Selenium moderno gerencia automaticamente o driver na maioria dos ambientes.

## Executar pelo código

```powershell
py -3 -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements.txt
.venv\Scripts\python.exe src\buscador_aso.py
```

Informe a URL autorizada na interface. Opcionalmente, defina `ASO_PORTAL_URL` no ambiente.

## Executar como versão portátil

Copie a pasta para um diretório local ou pendrive e execute:

```text
executar_portatil.bat
```

Na primeira execução, o script cria um ambiente virtual dentro da pasta e instala a dependência. Para um pendrive totalmente offline, prepare previamente o ambiente em máquina compatível ou distribua o EXE aprovado internamente.

## Gerar o EXE

```text
build_exe.bat
```

O arquivo será criado em:

```text
dist\BuscadorDeASO.exe
```

O executável não inclui credenciais, CPFs ou documentos.

## Configuração do portal

A automação foi mantida com os seletores conhecidos do protótipo. Portais podem mudar IDs, textos e navegação. Antes do uso, um responsável técnico deve revisar os seletores em `src/buscador_aso.py`.

A versão pública não contorna CAPTCHA, autenticação multifator, bloqueios ou controles anti-automação. Esses mecanismos devem ser respeitados.

## Privacidade e uso autorizado

CPFs e ASOs exigem cuidado elevado. Utilize a ferramenta somente quando houver autorização, finalidade definida, acesso legítimo e armazenamento seguro. Não publique listas de CPFs, documentos baixados, credenciais ou resultados de execução. Consulte `SECURITY.md`.

## Resultado

O projeto demonstrou como uma automação simples pode eliminar uma tarefa repetitiva de alto volume e liberar tempo dos responsáveis para atividades de análise e acompanhamento. Também estabeleceu a base para projetos posteriores com interfaces, tratamento de exceções e foco na experiência do usuário.

## Autor

**Gustavo Freitas Gomes Dumont**
