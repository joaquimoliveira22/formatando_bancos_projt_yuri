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
        filetypes=[("Arquivos Excel/CSV", "*.xlsx *.xls *.csv"), ("Todos os arquivos", "*.*")]
    )

def normalizar_texto(texto):
    if not isinstance(texto, str):
        return ""
    texto = unicodedata.normalize('NFKD', texto).encode('ASCII', 'ignore').decode('ASCII')
    return ''.join(c for c in texto if c.isalnum() or c.isspace()).strip().lower()

def criar_nome_arquivo_saida(arquivo_original, sufixo="extraido"):
    base, ext = os.path.splitext(arquivo_original)
    contador = 1
    while True:
        novo_nome = f"{base}_{sufixo}_{contador}.xlsx"
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

def tentar_ler_como_html(caminho_arquivo):
    try:
        dfs = pd.read_html(caminho_arquivo, header=0)
        if dfs:
            print("Arquivo .xls identificado como HTML e lido com sucesso.")
            return dfs[0]
    except Exception as e:
        print(f"Erro ao tentar ler como HTML: {e}")
    return None

def converter_xls_para_xlsx(caminho_arquivo):
    import xlrd
    from openpyxl import Workbook

    print("Convertendo .xls para .xlsx...")
    wb_xls = xlrd.open_workbook(caminho_arquivo)
    pasta = os.path.dirname(caminho_arquivo)
    nome = os.path.splitext(os.path.basename(caminho_arquivo))[0]
    novo_arquivo = os.path.join(pasta, nome + '.xlsx')

    wb_xlsx = Workbook()
    ws_xlsx = wb_xlsx.active
    sheet = wb_xls.sheet_by_index(0)

    for row in range(sheet.nrows):
        ws_xlsx.append(sheet.row_values(row))

    wb_xlsx.save(novo_arquivo)
    print(f"Arquivo convertido para: {novo_arquivo}")
    return novo_arquivo

def extrair_colunas_relevantes(df, arquivo_origem, nome_planilha):
    # Primeiro, garantir que temos um DataFrame
    if isinstance(df, pd.Series):
        df = df.to_frame().T  # Converte Series para DataFrame
    
    # Normalizar nomes de colunas
    try:
        df.columns = [normalizar_texto(str(col)) for col in df.columns]
    except AttributeError:
        print("Erro: Os dados não parecem estar em formato tabular.")
        return None
    
    # Mapear possíveis nomes de colunas para as que queremos
    mapeamento_colunas = {
        'data': ['data', 'datareferencia', 'datadareferencia', 'dataemissao'],
        'valor': ['valor', 'valores', 'vlr', 'val'],
        'saldo': ['saldo', 'saldos', 'sld']
    }
    
    # Encontrar as colunas correspondentes
    colunas_encontradas = {}
    for col in df.columns:
        col_normalizada = normalizar_texto(str(col))
        for tipo, possiveis in mapeamento_colunas.items():
            if any(p in col_normalizada for p in possiveis):
                colunas_encontradas[tipo] = col
                break
    
    # Verificar se encontramos todas as colunas necessárias
    if not all(k in colunas_encontradas for k in ['data', 'valor', 'saldo']):
        print("Não foram encontradas todas as colunas necessárias (Data, Valor, Saldo).")
        print("Colunas disponíveis:", df.columns.tolist())
        return None
    
    # Criar novo DataFrame apenas com as colunas de interesse
    try:
        df_final = df[[colunas_encontradas['data'], colunas_encontradas['valor'], colunas_encontradas['saldo']]].copy()
    except KeyError as e:
        print(f"Erro ao selecionar colunas: {e}")
        return None
    
    # Renomear colunas para nomes padronizados
    df_final.columns = ['Data da Referência', 'Valor', 'Saldo']
    
    # Formatar valores
    df_final['Valor'] = df_final['Valor'].apply(formatar_contabil)
    df_final['Saldo'] = df_final['Saldo'].apply(formatar_contabil)
    
    # Tentar formatar a data
    try:
        df_final['Data da Referência'] = pd.to_datetime(
            df_final['Data da Referência'],
            errors='coerce'
        ).dt.strftime('%d/%m/%Y')
    except Exception as e:
        print(f"Erro ao formatar datas: {e}")
    
    # Remover linhas vazias
    df_final = df_final.dropna(how='all')
    
    # Remover as 4 últimas linhas se houver linhas suficientes
    if len(df_final) > 4:
        df_final = df_final.iloc[:-4]
    else:
        print("Aviso: O dataframe tem 4 ou menos linhas, não foi possível remover as 4 últimas.")
    
    return df_final

def processar_arquivo(arquivo):
    if arquivo.lower().endswith('.xls'):
        df = tentar_ler_como_html(arquivo)
        if df is not None:
            return { "Planilha_HTML": df }, "HTML"
        else:
            try:
                arquivo = converter_xls_para_xlsx(arquivo)
            except Exception as e:
                print(f"Falha ao converter .xls: {e}")
                return None, None
    
    try:
        if arquivo.lower().endswith(('.xlsx', '.xls')):
            xls = pd.ExcelFile(arquivo)
            dfs = {}
            for nome in xls.sheet_names:
                print(f"\nProcessando planilha: {nome}")
                df = pd.read_excel(xls, sheet_name=nome, header=None)
                dfs[nome] = df
            return dfs, "Excel"
        elif arquivo.lower().endswith('.csv'):
            encodings = ['utf-8', 'latin1', 'iso-8859-1', 'cp1252']
            separadores = [',', ';', '\t']
            for encoding in encodings:
                for sep in separadores:
                    try:
                        df = pd.read_csv(arquivo, header=None, encoding=encoding, sep=sep)
                        print(f"CSV lido com encoding {encoding} e separador '{sep}'")
                        return { "CSV": df }, "CSV"
                    except:
                        continue
            print("Não foi possível ler o arquivo CSV.")
            return None, None
        else:
            print("Formato de arquivo não suportado.")
            return None, None
    except Exception as e:
        print(f"Ocorreu um erro: {e}")
        return None, None

def main():
    arquivo = selecionar_arquivo()
    if not arquivo:
        print("Nenhum arquivo selecionado.")
        return
    
    print(f"\nArquivo selecionado: {arquivo}")
    
    # Processar arquivo original
    dados, tipo = processar_arquivo(arquivo)
    if dados is None:
        print("Não foi possível processar o arquivo.")
        return
    
    # Para cada planilha encontrada (pode ser apenas uma no caso de CSV)
    for nome_planilha, df in dados.items():
        print(f"\nProcessando: {nome_planilha}")
        
        # Extrair colunas relevantes
        df_final = extrair_colunas_relevantes(df, arquivo, nome_planilha)
        if df_final is None:
            print(f"Não foi possível extrair dados da planilha {nome_planilha}")
            continue
        
        # Criar novo arquivo XLSX com as colunas extraídas
        nome_saida = criar_nome_arquivo_saida(arquivo, f"extraido_{nome_planilha}")
        df_final.to_excel(nome_saida, index=False)
        print(f"\nDados extraídos e salvos em: {nome_saida}")
        print(f"Total de linhas no arquivo final: {len(df_final)}")
        print("Primeiras linhas do arquivo gerado:")
        print(df_final.head())

if __name__ == "__main__":
    main()