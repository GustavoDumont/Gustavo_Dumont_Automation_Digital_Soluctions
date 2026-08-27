# Configuração do portal

A versão pública não contém seletores internos completos. Em `src/automation.py`, implemente `list_records()` no ambiente autorizado.

A implementação deve:

1. aguardar a autenticação interativa;
2. navegar até a área de férias;
3. aplicar data inicial e final;
4. listar somente os registros do intervalo;
5. fornecer localizadores para Aviso e Recibo;
6. manter paginação, idempotência e limites do portal;
7. confirmar o término de cada download antes de prosseguir;
8. registrar falhas sem incluir dados excessivos no log.

Não utilize técnicas para ocultar automação ou contornar controles.
