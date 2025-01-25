import yfinance as yf
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

def get_financial_data(tickers):
    # Create an empty list to store the data
    financial_data = []

    for ticker in tickers:
        print(f"Coletando dados para {ticker}...")
        
        # Create a Ticker object for the company
        empresa = yf.Ticker(ticker)
        
        # Get the info of the company
        info = empresa.info
        
        try:
            data = {
                'Ticker': ticker,
                'Preço de Abertura': info.get('regularMarketOpen', 'N/A'),
                'Preço Máximo do Dia': info.get('regularMarketDayHigh', 'N/A'),
                'Preço Mínimo do Dia': info.get('regularMarketDayLow', 'N/A'),
                'Preço de Fechamento Anterior': info.get('regularMarketPreviousClose', 'N/A'),            
                'Volume': info.get('regularMarketVolume', 'N/A'),
                'P/L': info.get('trailingPE', 'N/A'),
                'EBITDA': info.get('ebitda', 'N/A'),
                'Receita': info.get('totalRevenue', 'N/A'),
                'Lucro por Ação': info.get('trailingEps', 'N/A'),
                'Market Cap': info.get('marketCap', 'N/A'),
                'Setor': info.get('sector', 'N/A'),
                'Dividendo (Taxa)': info.get('dividendRate', 'N/A'),
                'Dividendo (Retorno)': info.get('dividendYield', 'N/A'),
                'Beta': info.get('beta', 'N/A'),
                'Projeção Alta': info.get('targetHighPrice', 'N/A'),
                'Projeção Baixa': info.get('targetLowPrice', 'N/A'),
                'Margem Operacional': info.get('operatingMargins', 'N/A'),
                'Margem Bruta': info.get('grossMargins', 'N/A'),
                'Retorno sobre Ativos': info.get('returnOnAssets', 'N/A'),
                'Retorno sobre Patrimônio': info.get('returnOnEquity', 'N/A'),
                'Dívida Total': info.get('totalDebt', 'N/A'),
                'Razão Dívida/Patrimônio': info.get('debtToEquity', 'N/A'),
                'Recomendação Média': info.get('recommendationMean', 'N/A'),
                'Fluxo de Caixa Livre': info.get('freeCashflow', 'N/A'),
            }

            financial_data.append(data)
            
        except KeyError as e:
            print(f"Erro ao coletar dados para {ticker}: {e}")
    
    # Create a DataFrame with the data and return it 
    df = pd.DataFrame(financial_data)
    
    return df

def save_in_excel(df, nome_arquivo):
    df.to_excel(nome_arquivo, index=False)
    print(f"Dados salvos em {nome_arquivo}.")
    messagebox.showinfo("Sucesso", f"Dados salvos em {nome_arquivo}")

def select_file():
    print("Selecione o arquivo excel com os tickers.")
    
    root = tk.Tk()  # Open a file dialog to select the input file
    root.withdraw() # Hide the main window
    
    file_path = filedialog.askopenfilename(
        title="Selecione o arquivo de entrada",
        filetypes=[("Arquivos Excel", "*.xlsx *.xls")]
    )
    
    return file_path

def read_tickers_from_excel(file_path):
    try:
        # Read the Excel file without a header
        tickers_df = pd.read_excel(file_path, header=None) 
        
        # Return the first column as a list of tickers
        return tickers_df.iloc[:, 0].dropna().tolist() 
    
    except Exception as e:
        print(f"Erro ao ler o arquivo {file_path}: {e}")
        messagebox.showerror("Erro", f"Erro ao ler o arquivo: {e}")
        return []

def save_output_file():
    print("Selecione o arquivo de saída.")
    
    output_file_name = filedialog.asksaveasfilename(
        title="Salvar arquivo como",
        defaultextension=".xlsx",
        filetypes=[("Arquivos Excel", "*.xlsx")],
        initialfile="dados_financeiros.xlsx"
    )
    
    if output_file_name:
        return output_file_name
    else:
        print("Nenhum arquivo selecionado para salvar. Encerrando.")
        return None

def main():
    
    # Select the input file
    input_file_name = select_file()
    
    if not input_file_name:
        print("Nenhum arquivo selecionado. Encerrando.")
        return
        
    # Get the tickers from the first column
    tickers = read_tickers_from_excel(input_file_name)
        
    if not tickers:
        print("Nenhum ticker encontrado no arquivo. Encerrando.")
        return
        
    print(f"Tickers a serem processados: {tickers}")

    # Get the financial data for the tickers and create a DataFrame
    df = get_financial_data(tickers)

    # Save the data in an Excel file
    output_file_name = save_output_file()
    
    if output_file_name:
        save_in_excel(df, output_file_name)

if __name__ == "__main__":
    main()