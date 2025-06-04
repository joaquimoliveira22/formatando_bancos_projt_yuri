import pandas as pd
import tkinter as tk
from tkinter import filedialog
import unicodedata
import os

def selecionar_arquivo():
    root = tk.Tk()
    root.withdraw()
    return filedialog.askopenfilename(
        title="Selecione o arquivo",
        filetypes=[("Arquivos TXT/Excel/CSV", "*.txt *.xlsx *.xls *.csv"), ("Todos os arquivos", "*.*")]
    )

def normalizar_texto(texto):
    if not isinstance(texto, str):
        return ""
    texto = unicodedata.normalize('NFKD', texto).encode('ASCII', 'ignore').decode('ASCII')
    return ''.join(c for c in texto if c.isalnum() or c.isspace()).strip().lower()

def criar_nome_arquivo_saida(arquivo, sufixo=""):
    base, ext = os.path.splitext(arquivo)
    novo_nome = f"{base}_{sufixo}.xlsx" if sufixo else f"{base}.xlsx"
    contador = 1
    while os.path.exists(novo_nome):
        novo_nome = f"{base}_{sufixo}_{contador}.xlsx" if sufixo else f"{base}_{contador}.xlsx"
        contador += 1
    return novo_nome

def detectar_delimitador(arquivo):
    with open(arquivo, 'r', encoding='utf-8', errors='ignore') as f:
        linha = f.readline()
        for sep in [',',';','\t','|']:
            if sep in linha:
                return sep
    return None

def carregar_dados(arquivo):
    ext = arquivo.lower().split('.')[-1]
    if ext in ['xls', 'xlsx']:
        df = pd.read_excel(arquivo)
    elif ext == 'csv':
        sep = detectar_delimitador(arquivo) or ','
        df = pd.read_csv(arquivo, sep=sep, encoding='utf-8', on_bad_lines='skip')
    elif ext == 'txt':
        sep = detectar_delimitador(arquivo) or '\t'
        df = pd.read_csv(arquivo, sep=sep, encoding='utf-8', on_bad_lines='skip')
    else:
        raise ValueError("Formato de arquivo não suportado")
    return df

def encontrar_colunas(df):
    variacoes_data = ['data', 'data_mov', 'dataocorrencia', 'data_ocorrencia', 'data movimentacao', 'data_movimentacao']
    variacoes_valor = ['valor', 'valores', 'vlr', 'val', 'montante']
    
    col_data = None
    col_valor = None
    
    for col in df.columns:
        col_norm = normalizar_texto(str(col))
        if any(v in col_norm for v in variacoes_data):
            col_data = col
        if any(v in col_norm for v in variacoes_valor):
            col_valor = col
    
    if col_data is None or col_valor is None:
        raise ValueError("Não foi possível encontrar as colunas Data_Mov e Valor")
    
    return col_data, col_valor

def salvar_data_valor(df, arquivo_origem):
    col_data, col_valor = encontrar_colunas(df)
    df_out = df[[col_data, col_valor]].copy()
    df_out.columns = ['Data_Mov', 'Valor']
    
    # Conversão de datas para o formato dd/mm/aaaa
    df_out['Data_Mov'] = pd.to_datetime(df_out['Data_Mov'], errors='coerce').dt.strftime('%d/%m/%Y')
    
    nome_saida = criar_nome_arquivo_saida(arquivo_origem, "data_valor")
    df_out.to_excel(nome_saida, index=False)
    print(f"Arquivo com Data_Mov e Valor salvo em:\n{os.path.abspath(nome_saida)}")

def main():
    arquivo = selecionar_arquivo()
    if not arquivo:
        print("Nenhum arquivo selecionado.")
        return
    try:
        df = carregar_dados(arquivo)
        salvar_data_valor(df, arquivo)
    except Exception as e:
        print(f"Erro: {e}")

if __name__ == "__main__":
    main()
