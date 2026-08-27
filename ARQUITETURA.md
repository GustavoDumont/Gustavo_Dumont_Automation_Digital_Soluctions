# Arquitetura e escalabilidade

```text
Lista de pessoas
      |
      v
Divisão em 5 filas sem sobreposição
      |
      +---- Máquina 1 ----+
      +---- Máquina 2 ----+
      +---- Máquina 3 ----+--> Pastas por pessoa --> conferência --> SharePoint
      +---- Máquina 4 ----+
      +---- Máquina 5 ----+
```

Cada worker deve ter fila própria, diretório de saída próprio e relatório próprio. A consolidação deve verificar contagem, tamanho e SHA-256. Uma execução repetida precisa ser idempotente, pulando arquivos já íntegros e retomando somente itens incompletos.

A versão pública deixa a adaptação do portal em uma camada isolada. Isso reduz acoplamento entre interface, armazenamento e seletores web.
