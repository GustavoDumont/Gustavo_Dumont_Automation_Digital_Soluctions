# Buscador de Dados

Conjunto de aplicativos coringa para localizar informações em documentos, renomear arquivos a partir do conteúdo encontrado e produzir relatórios sobre pastas e arquivos.

O projeto começou como uma automação de OCR. Com a evolução, ganhou versões voltadas a necessidades específicas, como localização de datas de treinamentos e registros profissionais, e versões configuráveis nas quais o próprio usuário define o dado procurado. A mesma base pode ser conectada a outras automações e adaptada a novos formatos documentais.

## Aplicativos

### Buscador de Dados e Renomeador

- lê PDFs, imagens e TXT;
- aproveita texto já existente no PDF;
- aplica OCR quando o conteúdo textual é insuficiente;
- extrai valores entre âncoras;
- extrai padrões por expressão regular;
- copia e renomeia documentos encontrados;
- preserva o arquivo original;
- impede sobrescrita por meio de nomes únicos;
- gera relatório CSV com sucesso, ausência ou erro.

### Mapeador de Pastas

- percorre uma árvore de diretórios;
- lista arquivos e pastas;
- identifica pastas vazias;
- registra tamanho e data de alteração;
- gera relatório CSV.

### Detalhes de Arquivos

- inventaria arquivos recursivamente;
- informa extensão, MIME, tamanho e datas;
- registra permissões de leitura e escrita;
- calcula SHA-256 opcionalmente;
- gera relatório CSV.

## Arquitetura

```text
Documentos -> leitura nativa -> OCR de fallback -> regra configurável
                                              -> relatório
                                              -> cópia renomeada

Pastas -> mapeamento -> relatório estrutural
Arquivos -> metadados/hash -> relatório de inventário
```

## Estrutura

```text
src/app.py
src/extractor.py
src/mapeador_pastas.py
src/detalhes_arquivos.py
src/common.py
configs/
docs/EXTENSAO.md
requirements.txt
executar.bat
build_exe.bat
SECURITY.md
CHANGELOG.md
```

## Instalação

```powershell
py -3 -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements.txt
```

Instale o Tesseract OCR com dados do idioma português ou coloque uma distribuição autorizada em `tesseract/`.

## Executar a interface

```text
executar.bat
```

Escolha a pasta de entrada, uma pasta de saída e uma regra JSON de `configs/`.

## Mapeamento de pastas

```powershell
.venv\Scripts\python.exe src\mapeador_pastas.py "C:\Documentos" "mapa.csv"
```

## Inventário de arquivos

```powershell
.venv\Scripts\python.exe src\detalhes_arquivos.py "C:\Documentos" "detalhes.csv" --hash
```

## Gerar EXE

```text
build_exe.bat
```

## Versatilidade

O pipeline separa leitura, extração, relatório e saída. Isso permite trocar uma regra específica sem reconstruir todo o aplicativo. Exemplos de expansão incluem:

- classificação documental;
- validação de nomenclatura;
- indexação para busca;
- triagem de arquivos;
- integração com robôs de movimentação;
- geração de bases para dashboards;
- detecção de duplicidades por hash;
- auditoria de estruturas de pastas.

## Limitações

OCR não é infalível. Qualidade de imagem, rotação, contraste, idioma e layout afetam a leitura. Os resultados devem ser revisados, principalmente quando usados para renomear ou classificar documentos.

## Privacidade

Documentos e relatórios não devem ser publicados. A branch ignora formatos comuns de entrada e saída. Consulte `SECURITY.md`.

## Autor

**Gustavo Freitas Gomes Dumont**
