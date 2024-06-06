Automação de Coleta e Análise de Dados de Ações


• Descrição do Projeto

Este projeto tem como objetivo automatizar a coleta e análise de dados históricos de ações utilizando a biblioteca Selenium para web scraping e a biblioteca Pandas para manipulação e análise de dados. O código acessa sites financeiros, baixa dados históricos das ações e realiza a extração de informações adicionais como a valorização das ações. Os dados são processados e armazenados em um arquivo Excel para posteriormente serem analisados por uma ferramenta de BI.


• Funcionalidades:

Automação de Download de Dados Históricos de Ações:

Utiliza Selenium para acessar o Yahoo Finance e baixar dados históricos de ações em formato CSV.

Verifica se os arquivos já existem no diretório especificado e substitui os arquivos antigos por novas versões.

• Extração de Dados de Valorização:

Acessa o site Status Invest para extrair informações de valorização das ações (12 meses e mês atual).


• Processamento de Dados:

Lê os arquivos CSV baixados e extrai os dados relevantes.

Aplana listas de dados para facilitar a manipulação e análise.

Cria DataFrames para organizar os dados de forma estruturada.

• Armazenamento em Arquivo Excel:

Concatena os DataFrames e salva os dados em um arquivo Excel com diferentes abas para fácil visualização e análise.


• Bibliotecas Utilizadas

selenium: Para automação de navegação e extração de dados da web.

openpyxl: Para manipulação de arquivos e diretórios.

pandas: Para manipulação e análise de dados.

• link para o dashboard do projeto: https://app.powerbi.com/view?r=eyJrIjoiOTFlY2ZlYzctMTBmMC00ZmNkLWJhNmYtOWMyODUyYzRmOWY0IiwidCI6IjY4NGRkNmEwLWRhMTgtNDc1NC04MjU2LTBmZDcyYzBhMmZhYyJ9
