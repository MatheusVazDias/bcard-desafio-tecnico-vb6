# BCard — Desafio Técnico

## Aplicação desktop desenvolvida em Visual Basic 6 para gerenciamento de transações de cartões de crédito, utilizando SQL Server.

### Funcionalidades
- Cadastro, edição e exclusão de transações
- Bloqueio de edição de transações aprovadas
- Consulta com filtros por cartão, período, valor e status
- Paginação dos resultados para grandes volumes de dados
- Geração de relatório Excel das transações do mês anterior
- Tratamento de erros com mensagens amigáveis e registro em log
- Regras de negócio implementadas também no banco de dados através de Functions, View e Stored Procedure

### Usabilidade

A tela de consulta possui alguns comportamentos para facilitar a utilização:

- 1 clique em uma transação permite selecioná-la para exclusão
- 2 cliques sobre uma transação abrem a tela de edição
- Todos os campos editáveis das telas possuem TabIndex configurado, permitindo navegar pela aplicação utilizando a tecla TAB, sem a necessidade de utilizar o mouse para acessar cada campo

### Tecnologias
- Visual Basic 6
- ADO
- SQL Server
- Excel Automation
- Git / GitHub
- Estrutura
- SQL
  - Scripts do banco, funções, procedures, views e dados de exemplo

VB
-  Projeto e código-fonte VB6
-  Relatorios
    - Exemplo de relatório Excel

### Execução
1. Executar os scripts da pasta SQL no SQL Server, na ordem numérica.
2. Executar o script de dados de exemplo.
3. Ajustar a conexão com o SQL Server no DAO.bas, caso necessário.
4. Abrir VB/BCard.vbp no Visual Basic 6 e executar.

### Observações

O projeto foi desenvolvido buscando manter uma estrutura simples e adequada ao ambiente VB6, com separação entre interface, acesso aos dados e regras implementadas no SQL Server.
