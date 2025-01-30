# Tipos de Commits

O `type` em uma mensagem de commit é responsável por indicar o tipo de alteração ou iteração que está sendo realizada. Abaixo estão os principais tipos de commits, seguindo as regras da convenção:

- **test**: indica qualquer criação ou alteração de códigos de teste.
  - **Exemplo**: Criação de testes unitários.

- **feat**: usado para o desenvolvimento de uma nova funcionalidade (feature) no projeto.
  - **Exemplo**: Acréscimo de um novo serviço, endpoint, etc.

- **refactor**: utilizado quando há uma refatoração de código sem impacto nas regras de negócio.
  - **Exemplo**: Mudanças de código após uma revisão (code review).

- **style**: empregado em mudanças de formatação e estilo que não alteram o sistema.
  - **Exemplo**: Ajustes de indentação, remoção de espaços em branco ou comentários.

- **fix**: indica correção de erros (bugs) no sistema.
  - **Exemplo**: Tratamento de uma função que não está retornando o resultado esperado.

- **chore**: representa mudanças no projeto que não afetam o sistema ou arquivos de teste.
  - **Exemplo**: Atualizar regras do eslint, adicionar arquivos ao `.gitignore`.

- **docs**: utilizado para mudanças na documentação do projeto.
  - **Exemplo**: Atualizar o README ou adicionar informações na documentação da API.

- **build**: indica alterações que afetam o processo de build do projeto ou dependências externas.
  - **Exemplo**: Adicionar ou remover dependências no `npm`.

- **perf**: usado para alterações que melhoram a performance do sistema.
  - **Exemplo**: Substituir `forEach` por `while` para otimização.

- **ci**: indica mudanças nos arquivos de configuração de integração contínua (CI).
  - **Exemplo**: Configurações do CircleCI, Travis, etc.

- **revert**: utilizado para reverter um commit anterior.