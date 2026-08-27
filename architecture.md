# Arquitetura e evolução

## Geração SAP

Excel atua como fila de entrada. O VBScript percorre as linhas, controla uma sessão SAP GUI autorizada, preenche os dados e gera a saída.

## Postagem OCR

PDF -> Poppler -> imagem -> OpenCV -> Tesseract -> expressões regulares -> validação -> Selenium.

## Postagem Excel

Excel/CSV -> pandas -> normalização e validação -> Selenium.

## Decisão de evolução

A substituição do OCR por dados estruturados nos campos críticos reduz ambiguidades, incluindo a confusão entre `1` e `7`, e torna o fluxo mais previsível e auditável.
