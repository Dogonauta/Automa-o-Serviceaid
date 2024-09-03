# Automação de Chamados com Selenium e OpenPyXL

Este projeto automatiza o processo de atendimento, atualização e resolução de chamados em uma plataforma de Service Desk utilizando Selenium para interagir com o site e OpenPyXL para manipular planilhas Excel. A automação é capaz de realizar login, pesquisar chamados, atualizar status, inserir apontamentos e gerar relatórios de chamados resolvidos.

## Funcionalidades

- **Automação de login**: O bot faz login automaticamente na plataforma de Service Desk utilizando as credenciais armazenadas em variáveis de ambiente.
- **Pesquisa de chamados**: O sistema pesquisa chamados na plataforma com base nos dados de uma planilha Excel.
- **Atualização de status**: A automação atualiza o status do chamado conforme os critérios pré-definidos.
- **Inserção de apontamentos**: O bot insere um comentário explicando a resolução do chamado.
- **Geração de relatório**: Os chamados resolvidos são exportados para uma nova planilha Excel.
- **Controle de fluxo**: Dependendo do status do chamado, o script executa diferentes ações, como retomar chamados pendentes ou aceitar solicitações.

## Estrutura do Projeto

### Arquivos principais:

- **tickets.xlsx**: Planilha de entrada contendo a lista de chamados e suas informações.
- **tickets_concluidos.xlsx**: Planilha de saída onde são armazenados os chamados resolvidos pela automação.
- **.env**: Arquivo que contém as variáveis de ambiente para as credenciais do usuário.

### Bibliotecas Utilizadas

- `openpyxl`: Usada para manipulação de arquivos Excel, permitindo a leitura da planilha de chamados e a escrita de relatórios.
- `dotenv`: Utilizada para carregar variáveis de ambiente, garantindo que as credenciais do usuário sejam armazenadas de forma segura.
- `selenium`: Ferramenta principal para a automação do navegador. Utilizada para interagir com o site da plataforma de Service Desk.
- `time`: Utilizada para pausas temporárias no script, permitindo que o site carregue completamente antes da execução das ações.

## Fluxo de Execução do Código

1. **Carregamento das credenciais**: 
   As credenciais de login (`User` e `Senha`) são carregadas do arquivo `.env`.

2. **Carregamento das planilhas**: 
   O script carrega a planilha `tickets.xlsx`, que contém os chamados a serem processados.

3. **Login automático**: 
   A função `logar()` faz login no site da plataforma de Service Desk utilizando Selenium.

4. **Pesquisa e tratamento dos chamados**:
   - Para cada linha da planilha de chamados:
     - O script pesquisa o chamado na plataforma e verifica seu status atual usando a função `testar_status()`.
     - Dependendo do status do chamado, o bot executa diferentes ações:
       - **Pendente Fornecedor**: Retoma e atende o chamado.
       - **Cadastro contábil**: Resolve o chamado.
       - **Aceitação**: Aceita e resolve o chamado.
       - **Em andamento**: Executa a sequência de apontamentos e resolve o chamado.

5. **Inserção de apontamentos e resolução**:
   - A função `atender_chamado()` insere apontamentos e atualiza o status do chamado, quando necessário.
   - Chamados resolvidos recebem a resolução designada pela variável **mensagem_resolucao**

6. **Geração de relatório**:
   - A função `alimentar_nova_planilha()` insere os detalhes dos chamados resolvidos na planilha `tickets_concluidos.xlsx`.

## Estrutura do Código

### Funções

- **logar()**: Executa o login na plataforma usando Selenium.
- **testar_status()**: Verifica o status do chamado e retorna seu valor.
- **pesquisar_chamado()**: Pesquisa um chamado na plataforma usando o número do ticket.
- **atender_chamado()**: Insere apontamentos e atualiza o status do chamado para resolvido.
- **resolver_chamado()**: Resolve o chamado e finaliza o fluxo de atendimento.
- **alimentar_nova_planilha()**: Atualiza a planilha de chamados concluídos com as informações do chamado resolvido.
- **subir_aba()**: Executa o scroll para o topo da página.

### Variáveis Importantes

- **possiveis_status**: Lista de status que são verificados para determinar o tratamento de cada chamado.
- **mensagem_resolucao**: Mensagem padrão inserida no campo de resolução dos chamados.
- **User e Senha**: Credenciais carregadas do arquivo `.env` para login.
