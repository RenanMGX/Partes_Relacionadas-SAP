import os
from getpass import getuser
import pandas as pd
import xlwings as xw
from datetime import datetime
from dateutil.relativedelta import relativedelta
from shutil import copy2
import numpy as nb
from typing import Literal
from .functions import fechar_excel
from tkinter import filedialog

class ExcelData:
    @property
    def df_base(self) -> pd.DataFrame:
        df_base_temp = self.__df_base.replace(nb.nan, "")
        df_base_temp['Data do documento'] = df_base_temp['Data do documento'].astype('datetime64[us]')
        df_base_temp['Data de lançamento'] = df_base_temp['Data de lançamento'].astype('datetime64[us]')
        df_base_temp = df_base_temp[df_base_temp['Empresa'] != ""]
        df_base_temp = df_base_temp[df_base_temp['Divisão'] != ""]
        return df_base_temp
    
    
    def __init__(self, 
                 dados_entrada_path:str|pd.DataFrame,
                 modelo_file:str="MODELO BATCH INPUT.xlsx"
                 ) -> None:
        
        self.__df_base:pd.DataFrame
        if isinstance(dados_entrada_path, str):
            if os.path.exists(dados_entrada_path):
                if dados_entrada_path.endswith(".xlsx"):
                    self.__df_base = pd.read_excel(dados_entrada_path, dtype=str)
                else:
                    raise TypeError(f"é permitido apenas arquivos Excel")
            else:
                raise FileNotFoundError(f"arquivo '{dados_entrada_path=}' não encontrado")
        elif isinstance(dados_entrada_path, pd.DataFrame):
            self.__df_base = dados_entrada_path
        
        self.__modelo_file_path:str = modelo_file
    
    #metodo Principal    
    def alimentar_batch_input(self):
        modelo_file_path = self.__modelo_file_path
        modelo_file_path_copy = modelo_file_path.replace(".xlsx", "_temp.xlsx")
        copy2(modelo_file_path, modelo_file_path_copy)
        
        lista_remover = self.preparar_lista_alimentacao(mod='Remover')
        lista_acrescentar = self.preparar_lista_alimentacao(mod='Acrescentar')
        
        
        app = xw.App(visible=False)
        with app.books.open(modelo_file_path_copy)as wb:
            self._alimentar_celular(wb=wb,sheet="B.I. (ajuste -)",lista_alimentar=lista_remover)
            self._alimentar_celular(wb=wb,sheet="B.I. (ajuste +)",lista_alimentar=lista_acrescentar)
            wb.save(self._caminho_salvar())
        fechar_excel(modelo_file_path_copy)
        os.unlink(modelo_file_path_copy)
    
    def _caminho_salvar(self):
        options = {}
        options['defaultextension'] = ".xlsx"
        options['filetypes'] = [("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*")]
        options['initialfile'] = "MODELO BATCH INPUT.xlsx"
        arquivo_salvar = filedialog.asksaveasfilename(**options)
        if  arquivo_salvar == "":
            return f"C:\\Users\\{getuser()}\\Downloads\\MODELO BATCH INPUT.xlsx"
        return arquivo_salvar
    
    def _alimentar_celular(self, *, wb, sheet:str, lista_alimentar:dict):
        ws = wb.sheets[sheet] 
        #Alimentar Planilha   
        ws.range(f'A2').value = [[x] for x in lista_alimentar["sequencial"]]
        ws.range(f'B2').value = [[x] for x in lista_alimentar["ultimo_dia_mes"]]
        ws.range(f'C2').value = [[x] for x in lista_alimentar["ultimo_dia_mes"]]
        ws.range(f'D2').value = [[x] for x in lista_alimentar["Empresa"]]    
        ws.range(f'E2').value = [[x] for x in lista_alimentar["Divisão"]]
        ws.range(f'F2').value = [[x] for x in lista_alimentar["tipo_documento"]] 
        ws.range(f'G2').value = [[x] for x in lista_alimentar["Texto cabeçalho/Referencia"]]
        ws.range(f'H2').value = [[x] for x in lista_alimentar["Texto cabeçalho/Referencia"]]
        ws.range(f'J2').value = [[x] for x in lista_alimentar["Chave do Lançamento"]] 
        ws.range(f'K2').value = [[x] for x in lista_alimentar["valor"]]
        ws.range(f'L2').value = [[x] for x in lista_alimentar["tipo de conta"]]
        ws.range(f'M2').value = [[x] for x in lista_alimentar["Conta"]]
        ws.range(f'U2').value = [[x] for x in lista_alimentar["texto"]]
    
    def preparar_lista_alimentacao(self, *,
                                   mod:Literal["Remover", "Acrescentar"]
                                   ):
        
        #df_base_menor = df_base[df_base['Montante em moeda interna'] < "0"]
        sequencial = 0
        linha = 1
        ultimo_dia_mes = self.get_ultimo_dia_mes().strftime('%d.%m.%Y')
        lista_alimentar:dict = {
            "sequencial":[],
            "ultimo_dia_mes": [],
            "Empresa": [],
            "Divisão":[],
            "tipo_documento": [],
            "Texto cabeçalho/Referencia": [],
            "Chave do Lançamento": [],
            "valor": [],
            "tipo de conta":[],
            "Conta": [],
            "texto" : []
        }
        for row,dados in self.df_base.iterrows():
            sequencial += 1
            
            #Credito
            lista_alimentar["sequencial"].append(sequencial)
            lista_alimentar["ultimo_dia_mes"].append(ultimo_dia_mes)
            lista_alimentar["Empresa"].append(dados['Empresa'])
            lista_alimentar["Divisão"].append(dados['Divisão'])  
            lista_alimentar["tipo_documento"].append('AB') 
            lista_alimentar["Texto cabeçalho/Referencia"].append('RECLASSIFICAÇÃO PARTES RELACIONAS')
            lista_alimentar["Chave do Lançamento"].append('21') 
            lista_alimentar["valor"].append(dados['Montante em moeda interna']) 
            lista_alimentar["tipo de conta"].append('K') 
            lista_alimentar["Conta"].append(dados['Fornecedor'] if mod == "Remover" else self.tratar_conta(dados['Conta']))
            lista_alimentar["texto"].append(dados['Texto'])
            
            #debito
            lista_alimentar["sequencial"].append(sequencial)
            lista_alimentar["ultimo_dia_mes"].append("")
            lista_alimentar["Empresa"].append("")
            lista_alimentar["Divisão"].append(dados['Divisão'])  
            lista_alimentar["tipo_documento"].append('')
            lista_alimentar["Texto cabeçalho/Referencia"].append('')
            lista_alimentar["Chave do Lançamento"].append('50') 
            lista_alimentar["valor"].append(dados['Montante em moeda interna']) 
            lista_alimentar["tipo de conta"].append('S') 
            lista_alimentar["Conta"].append(self.tratar_conta(dados['Conta']) if mod == "Remover" else dados['Fornecedor'])
            lista_alimentar["texto"].append(dados['Texto']) 
        
        return lista_alimentar         
    
    def get_ultimo_dia_mes(self) -> datetime:
        now = datetime.now()
        date_temp = datetime(year=now.year, month=(now + relativedelta(months=1)).month, day=1)
        return date_temp - relativedelta(days=1)

    def tratar_conta(self, conta):
        conta_temp = str(conta)
        if len(conta_temp) == 10:
            return f"12{conta_temp[2:]}"
        else:
            return conta

if __name__ == "__main__":
    pass
