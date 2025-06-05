import pandas as pd
import tkinter as tk
from tkinter import filedialog
import unicodedata
import os
import pyexcel as pe

def converter_xls_para_xlsx(caminho):
    if caminho.lower().endswith(".xls"):
        novo_caminho = caminho.replace(".xls", ".xlsx")
        print(f"Convertendo arquivo .xls para .xlsx: {caminho}")
        try:
            pe.save_book_as(file_name=caminho, dest_file_name=novo_caminho)
            return novo_caminho
        except Exception as e:
            print(f" Erro ao converter .xls para .xlsx: {e}")
            return caminho
    return caminho

def selecionar_arquivos():
    root = tk.Tk()
    root.withdraw()
    return filedialog.askopenfilenames(
        title="Selecione um ou mais arquivos",
        filetypes=[("Arquivos Excel/CSV", "*.xlsx *.xls *.csv"), ("Todos os arquivos", "*.*")]
    )

def normalizar_texto(texto):
    if not isinstance(texto, str):
        return ""
    texto = unicodedata.normalize('NFKD', texto).encode('ASCII', 'ignore').decode('ASCII')
    return ''.join(c for c in texto if c.isalnum() or c.isspace()).strip().lower()

def criar_nome_arquivo_saida(arquivo_original, nome_planilha):
    base, ext = os.path.splitext(arquivo_original)
    contador = 1
    while True:
        novo_nome = f"{base}_extraido_{nome_planilha}_{contador}{ext}"
        if not os.path.exists(novo_nome):
            return novo_nome
        contador += 1

def formatar_contabil(valor):
    if pd.isna(valor):
        return ""
    try:
        valor = float(valor)
        return f"{valor:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
    except:
        return str(valor)

def arquivo_parece_html(caminho):
    try:
        with open(caminho, 'rb') as f:
            inicio = f.read(100).lower()
            return b'<html' in inicio or b'<!doctype html' in inicio
    except:
        return False

def extrair_dados(arquivo):
    try:
        if arquivo_parece_html(arquivo):
            print(f"\n Arquivo ignorado (HTML detectado): {arquivo}")
            return

        if arquivo.lower().endswith('.xls'):
            arquivo = converter_xls_para_xlsx(arquivo)

        if arquivo.lower().endswith('.xlsx'):
            xls = pd.ExcelFile(arquivo, engine='openpyxl')
            processar_excel(xls, arquivo)
        elif arquivo.lower().endswith('.csv'):
            processar_csv(arquivo)
        else:
            print(f"\nFormato não suportado: {arquivo}")
    except Exception as e:
        print(f"\n Erro ao abrir {arquivo}: {e}")


def processar_excel(xls, arquivo):
    for sheet_name in xls.sheet_names:
        print(f"\n Processando planilha: {sheet_name}")
        try:
            df = pd.read_excel(arquivo, sheet_name=sheet_name, header=None)
            processar_dataframe(df, arquivo, sheet_name)
        except Exception as e:
            print(f"Erro ao processar planilha {sheet_name}: {e}")

def processar_csv(arquivo):
    print("\n Processando arquivo CSV")
    encodings = ['utf-8', 'latin1', 'iso-8859-1', 'cp1252']
    separadores = [',', ';', '\t']
    for encoding in encodings:
        for sep in separadores:
            try:
                df = pd.read_csv(arquivo, header=None, encoding=encoding, sep=sep)
                print(f"✔️ CSV lido com encoding {encoding} e separador '{sep}'")
                processar_dataframe(df, arquivo, "CSV")
                return
            except:
                continue
    print("Não foi possível ler o arquivo CSV")

def processar_dataframe(df, arquivo, nome_planilha):
    variacoes_cabecalhos = {
        'data': ['data', 'Data da Referência', 'Data'],
        'valor': ['valor', 'valores', 'vlr', 'val', 'Valor'],
        'saldo': ['saldo', 'saldos', 'sld', 'Saldo']
    }

    linha_cabecalho = None
    colunas_originais = []

    for idx, linha in df.iterrows():
        linha_normalizada = [normalizar_texto(str(cell)) for cell in linha.values]
        encontrados = {key: False for key in variacoes_cabecalhos}
        for col in linha_normalizada:
            for key, variacoes in variacoes_cabecalhos.items():
                if any(v in col for v in variacoes):
                    encontrados[key] = True
        if all(encontrados.values()):
            linha_cabecalho = idx
            colunas_originais = [str(cell).strip() for cell in linha.values]
            print(f"✔️ Cabeçalhos encontrados: {colunas_originais}")
            break

    if linha_cabecalho is not None:
        print(f"✔️ Cabeçalhos na linha {linha_cabecalho + 1}")

        if nome_planilha == "CSV":
            df_final = pd.read_csv(arquivo, header=linha_cabecalho)
        else:
            engine = 'xlrd' if arquivo.lower().endswith('.xls') else 'openpyxl'
            df_final = pd.read_excel(arquivo, sheet_name=nome_planilha, header=linha_cabecalho, engine=engine)

        df_final = df_final.dropna(how='all')

        colunas_map = {col: normalizar_texto(col) for col in df_final.columns}
        df_final.rename(columns=colunas_map, inplace=True)

        col_data1 = next((col for col in df_final.columns if 'data' in col), None)
        col_data2 = next((col for col in df_final.columns if 'referencia' in col), None)
        col_valor = next((col for col in df_final.columns if 'valor' in col), None)
        col_saldo = next((col for col in df_final.columns if 'saldo' in col), None)

        if not all([col_data1, col_valor, col_saldo]):
            print(" As colunas esperadas não foram encontradas.")
            return

        if col_data2 and col_data2 != col_data1:
            df_final = df_final[[col_data2, col_data1, col_valor, col_saldo]]
            df_final.columns = ['Data da Referência', 'Data do Extrato', 'Valor', 'Saldo']
        else:
            df_final = df_final[[col_data1, col_valor, col_saldo]]
            df_final.columns = ['Data da Referência', 'Valor', 'Saldo']

        if 'Data da Referência' in df_final.columns:
            df_final['Data da Referência'] = df_final['Data da Referência'].iloc[2:].reset_index(drop=True)
        if 'Data do Extrato' in df_final.columns:
            df_final['Data do Extrato'] = df_final['Data do Extrato'].iloc[2:].reset_index(drop=True)

        df_final['Valor'] = df_final['Valor'].iloc[:-4].reset_index(drop=True)
        df_final['Saldo'] = df_final['Saldo'].iloc[:-4].reset_index(drop=True)

        min_len = min(len(df_final[col]) for col in df_final.columns)
        df_final = df_final.iloc[:min_len]

        df_final['Valor'] = df_final['Valor'].apply(formatar_contabil)
        df_final['Saldo'] = df_final['Saldo'].apply(formatar_contabil)

        for col in ['Data da Referência', 'Data do Extrato']:
            if col in df_final.columns:
                df_final[col] = pd.to_datetime(df_final[col], errors='coerce')
                df_final[col] = df_final[col].apply(lambda x: x.strftime('%d/%m/%Y') if pd.notna(x) else '')

        print(" Primeiras linhas extraídas:")
        print(df_final.head())

        nome_saida = criar_nome_arquivo_saida(arquivo, nome_planilha)
        if nome_planilha == "CSV":
            df_final.to_csv(nome_saida, index=False, encoding='utf-8')
        else:
            from openpyxl import Workbook
            from openpyxl.utils.dataframe import dataframe_to_rows

            wb = Workbook()
            ws = wb.active
            ws.title = 'Dados Extraídos'

            for r_idx, row in enumerate(dataframe_to_rows(df_final, index=False, header=True), 1):
                ws.append(row)
                if r_idx > 1:
                    if 'Valor' in df_final.columns:
                        ws[f'B{r_idx}' if 'Data do Extrato' not in df_final.columns else f'C{r_idx}'].number_format = '#.##0,00_-'
                    if 'Saldo' in df_final.columns:
                        ws[f'C{r_idx}' if 'Data do Extrato' not in df_final.columns else f'D{r_idx}'].number_format = '#.##0,00_-'

            wb.save(nome_saida)

        print(f"Arquivo salvo: {nome_saida}")
    else:
        print(" Não foi possível encontrar os cabeçalhos esperados.")
        print(df.head())

def main():
    arquivos = selecionar_arquivos()
    if arquivos:
        print(f"\n {len(arquivos)} arquivo(s) selecionado(s):\n")
        for arquivo in arquivos:
            print(f"🔹 {arquivo}")
        print("\n Iniciando processamento...\n")
        for arquivo in arquivos:
            extrair_dados(arquivo)
        print("\nTodos os arquivos foram processados.")
    else:
        print("Nenhum arquivo selecionado.")

if __name__ == "__main__":
    main()
