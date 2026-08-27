# Arquitetura

## Versão desktop

A interface CustomTkinter coordena operações locais implementadas com bibliotecas especializadas. Tarefas demoradas usam threads e os resultados são gravados em pastas escolhidas pelo usuário.

## Linhas de produto

- `production`: base distribuída aos usuários.
- `experimental`: versão pessoal com manipulações mais finas ainda não homologadas.
- `web`: evolução planejada, sujeita a estudo de privacidade, processamento e hospedagem.

## Princípio de segurança

Operações experimentais devem trabalhar com cópias e evitar sobrescrita silenciosa. Conversões e OCR exigem validação humana quando o resultado tiver consequência operacional.
