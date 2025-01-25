# Automação da Coleta de Dados Financeiros

Este projeto automatiza o processo de coleta de dados financeiros para uma lista de tickers de ações (códigos usados na bolsa de valores para identificar e negociar ativos). Ele foi desenvolvido para resolver uma ineficiência operacional recorrente mencionada por um analista de M&A: a tarefa manual de buscar dados financeiros de empresas online, um por um, e inseri-los em uma planilha.

A automação utiliza a biblioteca yfinance, que fornece uma interface para acessar dados financeiros diretamente da API do Yahoo Finance.

O script solicita ao usuário um arquivo Excel contendo uma lista de tickers de ações e, em seguida, extrai automaticamente os dados financeiros relevantes para cada ticker, armazenando-os em um arquivo Excel de saída.

## Funcionalidades

- **Input**: Um arquivo Excel simples contendo uma lista de tickers de ações (um por linha).
- **Extração de Dados**: Automatiza a coleta de dados financeiros (por exemplo, preços das ações, capitalização de mercado, relação P/L, receita, etc.) usando a API do Yahoo Finance.
- **Output**: Um arquivo Excel contendo os dados financeiros coletados para todos os tickers.

## Formato do Arquivo de Entrada

O arquivo Excel de entrada deve conter uma única coluna com os tickers de ações, sem cabeçalhos, como mostrado abaixo:
<div align="center">

|                   |
|-------------------|
|AAPL               |
|MSFT               |
|GOOGL              |
|AMZN               |
|TSLA               |

</div>

Um arquivo de exemplo de entrada está disponível no repositório: `input_example.xlsx`.

## Formato do Arquivo de Saída

O arquivo de saída será um Excel contendo dados financeiros detalhados para cada ticker, incluindo:

- Preço de Abertura
- Preços Máximo e Mínimo do Dia
- Fechamento Anterior
- Volume
- Relação P/L
- Capitalização de Mercado
- EBITDA
- Receita Total
- Setor
- Taxa/Rendimento de Dividendos
- Beta e mais.

Um arquivo de exemplo de saída está disponível no repositório: `output_example.xlsx`.

## Como Utilizar o Executável

1. Clone este repositório:
   ```bash
   git clone https://github.com/LeoAoun/Desafio-Acorn-Advisory.git

2. Navegue para o diretório em que encontra-se o executável:
    ```bash
    cd Desafio-Acorn-Advisory/dist

3. Execute o arquivo `extrair_dados_financeiros.exe`.

# Execução do código (Apenas para Testes)

## Requisitos Para Executar o Script

- Python 3.8 ou superior
- Bibliotecas Python necessárias:
  - `yfinance`
  - `pandas`
  - `tkinter`
  - `openpyxl` 

## Como Executar o Código 

1. Clone este repositório:
   ```bash
   git clone https://github.com/LeoAoun/Desafio-Acorn-Advisory.git

2. Navegue para o diretório do projeto:
    ```bash
    cd Desafio-Acorn-Advisory

3. Instale as dependências necessárias:
    ```bash
    pip install -r requirements.txt

4. Execute o script:
    ```bash
    python extrair_dados_financeiros.py

## Como Criar o Aquivo Executável (Utilize caso tenha feito modificações no código)

1. Execute o script:
    ```bash
    python -m PyInstaller extrair_dados_financeiros.py --onefile

2. O executável será criado na pasta `dist`.