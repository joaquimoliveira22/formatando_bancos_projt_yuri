import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import unicodedata
import os
import re
from decimal import Decimal, InvalidOperation
from dateutil import parser
import xlrd
from openpyxl import Workbook
from datetime import datetime

class ExtratorPlanilha:
    def __init__(self):
        self.root = tk.Tk()
        self.root.withdraw()

    @staticmethod
    def normalizar_texto(texto):
        if not isinstance(texto, str):
            return ""
        texto = unicodedata.normalize('NFKD', texto).encode('ASCII', 'ignore').decode('ASCII')
        return re.sub(r'[^a-zA-Z0-9]', '', texto.lower())

    @staticmethod
    def extrair_valor_numerico(texto):
        if pd.isna(texto) or not str(texto).strip():
            return None

        texto = str(texto).strip()
        padrao_dinheiro = r"""
            (?:R\$\s*)?
            ([-+]?)
            \s*
            (\d{1,3}(?:\.\d{3})*|\d+)
            (?:
            ([.,])
            (\d{1,2})
            )?
        """

        matches = re.finditer(padrao_dinheiro, texto, re.VERBOSE)
        valores = []

        for match in matches:
            sinal, parte_inteira, sep_decimal, parte_decimal = match.groups()

            if '.' in parte_inteira:
                if sep_decimal == ',' or (sep_decimal is None and ',' not in texto):
                    parte_inteira = parte_inteira.replace('.', '')

            numero_str = parte_inteira
            if sep_decimal and parte_decimal:
                parte_decimal = parte_decimal.ljust(2, '0')[:2]
                numero_str += '.' + parte_decimal

            try:
                valor = Decimal(numero_str)
                if sinal == '-':
                    valor = -valor
                valores.append(float(valor))
            except InvalidOperation:
                continue

        if not valores:
            return None
        elif len(valores) == 1:
            return valores[0]
        else:
            return max(valores, key=abs)

    @staticmethod
    def formatar_moeda_brasileira(valor):
        if valor is None or pd.isna(valor):
            return ""

        try:
            valor = float(valor)
            formatado = "{:,.2f}".format(abs(valor))
            if valor < 0:
                return "-" + formatado.replace(",", "X").replace(".", ",").replace("X", ".")
            return formatado.replace(",", "X").replace(".", ",").replace("X", ".")
        except (ValueError, TypeError):
            return str(valor)

    @staticmethod
    def parse_data(data_str):
        try:
            if isinstance(data_str, datetime):
                return data_str.strftime('%d/%m/%Y')

            data_str = str(data_str).strip()
            padroes_data = [
                r'\b\d{2}/\d{2}/\d{4}\b',
                r'\b\d{2}-\d{2}-\d{4}\b',
                r'\b\d{4}-\d{2}-\d{2}\b',
                r'\b\d{8}\b',
                r'\b\d{1,2}/\d{1,2}/\d{2}\b'
            ]

            for padrao in padroes_data:
                match = re.search(padrao, data_str)
                if match:
                    data_str = match.group()
                    data_str = data_str.replace('-', '/')
                    break

            return parser.parse(data_str, dayfirst=True).strftime('%d/%m/%Y')
        except Exception:
            return None

    def selecionar_arquivo(self):
        return filedialog.askopenfilename(
            title="Selecione o arquivo",
            filetypes=[("Arquivos Excel/CSV", "*.xlsx *.xls *.csv"), ("Todos os arquivos", "*.*")]
        )

    def processar_arquivo(self, caminho_arquivo):
        try:
            if caminho_arquivo.endswith('.xls'):
                try:
                    dfs = pd.read_html(caminho_arquivo, header=0)
                    if dfs:
                        print("Leitura HTML feita com sucesso.")
                        return dfs[0]
                except Exception as e:
                    print(f"Falha leitura HTML: {e}")

                try:
                    wb_xls = xlrd.open_workbook(caminho_arquivo)
                    sheet = wb_xls.sheet_by_index(0)
                    data = [sheet.row_values(row) for row in range(sheet.nrows)]
                    df = pd.DataFrame(data)
                    return df
                except Exception as e:
                    print(f"Falha leitura xls com xlrd: {e}")

            elif caminho_arquivo.endswith('.xlsx'):
                return pd.read_excel(caminho_arquivo, header=None)
            elif caminho_arquivo.endswith('.csv'):
                return pd.read_csv(caminho_arquivo, header=None)
        except Exception as e:
            print(f"Erro ao ler o arquivo: {e}")
        return None

    def processar_dataframe(self, df):
        if df is None or df.empty:
            print("DataFrame vazio ou nulo.")
            return None

        try:
            df.columns = [self.normalizar_texto(str(col)) for col in df.iloc[0]]
            df = df.iloc[1:]
        except Exception:
            df.columns = [f"coluna_{i}" for i in range(len(df.columns))]

        col_saldo = next((col for col in df.columns if 'saldo' in col), None)
        col_valor = next((col for col in df.columns if 'valor' in col), None)

        if not col_saldo or not col_valor:
            print("Colunas 'Saldo' ou 'Valor' não encontradas.")
            return None

        df_filtrado = df[[col_valor, col_saldo]].copy()
        df_filtrado.columns = ['Valor', 'Saldo']

        df_filtrado['Valor'] = df_filtrado['Valor'].apply(self.extrair_valor_numerico)
        df_filtrado['Saldo'] = df_filtrado['Saldo'].apply(self.extrair_valor_numerico)

        df_filtrado.dropna(subset=['Valor', 'Saldo'], how='all', inplace=True)

        df_filtrado['Valor'] = df_filtrado['Valor'].apply(self.formatar_moeda_brasileira)
        df_filtrado['Saldo'] = df_filtrado['Saldo'].apply(self.formatar_moeda_brasileira)

        return df_filtrado

    def executar(self):
        caminho = self.selecionar_arquivo()
        if not caminho:
            print("Nenhum arquivo selecionado.")
            return

        df = self.processar_arquivo(caminho)
        df_final = self.processar_dataframe(df)

        if df_final is not None:
            nome_saida = os.path.splitext(caminho)[0] + '_valores_saldos.xlsx'
            df_final.to_excel(nome_saida, index=False)
            messagebox.showinfo("Sucesso", f"Arquivo salvo em: {nome_saida}")
            print(df_final.head())
        else:
            messagebox.showwarning("Aviso", "Não foi possível extrair dados de Valor e Saldo.")

if __name__ == "__main__":
    extrator = ExtratorPlanilha()
    extrator.executar()
