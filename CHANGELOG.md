# Histórico

## Versão inicial
- interface Tkinter;
- entrada de vários CPFs;
- autenticação no portal;
- pesquisa e download sequencial de ASOs;
- distribuição em executável.

## Versão portátil
- execução com navegador portátil;
- uso pensado para mídia removível.

## Versão GitHub
- remoção de credenciais e URLs internas do código;
- substituição de mecanismos de evasão por Selenium padrão;
- credenciais solicitadas em tempo de execução;
- pasta de download configurável;
- validação básica e deduplicação de CPFs;
- execução em thread para evitar travamento da interface;
- relatório CSV de resultado;
- logs com CPF mascarado;
- scripts separados para execução portátil e geração de EXE;
- documentação de segurança e uso autorizado.
