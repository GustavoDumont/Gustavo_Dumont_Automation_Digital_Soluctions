# Avisos e Recibos de Férias

Aplicação desktop para automatizar, em ambiente autorizado, o download em lote de Avisos e Recibos de Férias no Portal Meu RH.

A operação manual exigia acessar registro por registro, localizar os documentos e solicitar cada download. A automação transforma esse fluxo repetitivo em um processamento por período: o usuário informa as datas, escolhe os documentos desejados e acompanha a execução.

## Origem e evolução

O projeto possuiu diferentes versões. Uma foi preparada para distribuição como executável. Outra passou a trabalhar com mês e ano, ampliando o processamento em lote de Avisos e Recibos. A versão GitHub consolida os dois conceitos em um intervalo definido pelo usuário.

Os anexos disponíveis estavam parciais. Por isso, esta branch contém uma implementação reconstruída e segura, não uma reprodução literal do executável original.

## Funcionalidades

- intervalo com data inicial e final;
- seleção de Avisos, Recibos ou ambos;
- autenticação interativa no navegador;
- pasta de saída configurável;
- execução em segundo plano;
- solicitação de parada;
- tratamento de timeout e falhas do navegador;
- relatório CSV da execução;
- preparação para distribuição em EXE;
- URL configurada em tempo de execução ou variável de ambiente.

## Arquitetura

```text
Usuário informa período e tipo de documento
                    |
                    v
       Navegador abre o Portal Meu RH
                    |
                    v
          Autenticação autorizada
                    |
                    v
       Aplicação do intervalo de datas
                    |
                    v
        Registros encontrados no portal
             |                 |
             v                 v
          Avisos             Recibos
             |                 |
             +--------+--------+
                      v
             Pasta de downloads
                      |
                      v
             Relatório de execução
```

## Estrutura

```text
src/app.py
src/automation.py
src/models.py
docs/CONFIGURACAO_PORTAL.md
legacy/README.md
requirements.txt
executar.bat
build_exe.bat
SECURITY.md
CHANGELOG.md
```

## Configuração necessária

Por segurança, a versão pública não traz a URL corporativa nem os seletores completos do portal. Um responsável técnico deve implementar `list_records()` em `src/automation.py`, conforme `docs/CONFIGURACAO_PORTAL.md`.

A automação não contorna autenticação multifator, CAPTCHA, bloqueios ou controles anti-automação.

## Executar

```powershell
py -3 -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements.txt
.venv\Scripts\python.exe src\app.py
```

Ou execute:

```text
executar.bat
```

## Gerar EXE

```text
build_exe.bat
```

Saída esperada:

```text
dist\AvisosERecibosFerias.exe
```

## Ganho operacional

A solução elimina a repetição de abrir cada registro e baixar manualmente dois documentos. Em períodos com muitos colaboradores, o ganho de produtividade é expressivo e reduz o risco de omissões. O relatório final também facilita a identificação dos itens que precisam de tratamento manual.

## Privacidade

Avisos e recibos podem conter dados pessoais, financeiros e funcionais. Documentos e relatórios não devem ser publicados. Consulte `SECURITY.md`.

## Autor

**Gustavo Freitas Gomes Dumont**
