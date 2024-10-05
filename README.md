# Automação de Formulários com Selenium

Este projeto automatiza o preenchimento de formulários web utilizando o Selenium. Os dados são lidos a partir de um ficheiro Excel e as mensagens de resposta são guardadas no mesmo ficheiro após o envio do formulário.

## Requisitos

- **Python 3.x**
- **Selenium**: Para automação do navegador.
- **webdriver-manager**: Para gerir os drivers do navegador.
- **openpyxl**: Para trabalhar com ficheiros Excel.
- **pandas**: Para leitura e manipulação de dados em Excel.
- **dotenv**: Para carregar variáveis de ambiente a partir de um ficheiro `.env`.

## Funcionamento

- O script lê os dados do ficheiro Excel.
- Preenche os campos do formulário automaticamente.
- Envia o formulário e captura a mensagem de resposta.
- Guarda a mensagem no ficheiro Excel na coluna G. """
