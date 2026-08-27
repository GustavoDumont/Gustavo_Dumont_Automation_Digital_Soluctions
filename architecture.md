# Arquitetura da solução

```text
Usuário
  |
  v
Interface Tkinter
  |-- seleção e validação de PDF/XML
  |-- parâmetros operacionais
  |-- andamento e mensagens
  v
Thread de processamento
  |-- associação de documentos
  |-- regras do fluxo
  |-- controle de tentativas
  v
Selenium WebDriver
  |-- autenticação autorizada
  |-- navegação
  |-- preenchimento
  `-- envio e recuperação de sessão
```

A versão pública mantém a arquitetura para fins de portfólio, mas usa URL inválida e dados fictícios.
