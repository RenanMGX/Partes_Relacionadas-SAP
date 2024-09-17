import os
from getpass import getuser
import pandas as pd
import xlwings as xw
from xlwings.main import Sheet
from datetime import datetime
from dateutil.relativedelta import relativedelta
from shutil import copy2
import numpy as nb
from typing import Literal
from .functions import fechar_excel, Classific
from .functions import ultimo_dia_mes as obter_ultimo_dia_mes
from tkinter import filedialog
import locale; locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
from PyQt5.QtWidgets import QFileDialog

class ExcelData:
    @property
    def df_base(self) -> pd.DataFrame:
        df_base_temp = self.__df_base.replace(nb.nan, "")
        df_base_temp['Data do documento'] = df_base_temp['Data do documento'].astype('datetime64[us]')
        df_base_temp['Data de lançamento'] = df_base_temp['Data de lançamento'].astype('datetime64[us]')
        df_base_temp = df_base_temp[df_base_temp['Empresa'] != ""]
        df_base_temp = df_base_temp[df_base_temp['Divisão'] != ""]
        return df_base_temp
    
    @property
    def date(self) -> datetime:
        return self.__date    
    
    def __init__(self, *,
                 date:datetime,
                 dados_entrada_path:str|pd.DataFrame,
                 modelo_file:str="MODELO BATCH INPUT.xlsx"
                 ) -> None:
        
        self.__df_base:pd.DataFrame
        if isinstance(dados_entrada_path, str):
            if os.path.exists(dados_entrada_path):
                if dados_entrada_path.endswith(".xlsx"):
                    fechar_excel(dados_entrada_path)
                    self.__df_base = pd.read_excel(dados_entrada_path, dtype=str)
                else:
                    raise TypeError(f"é permitido apenas arquivos Excel")
            else:
                raise FileNotFoundError(f"arquivo '{dados_entrada_path=}' não encontrado")
        elif isinstance(dados_entrada_path, pd.DataFrame):
            self.__df_base = dados_entrada_path
        
        self.__modelo_file_path:str = modelo_file
        self.__date:datetime = date
    
    #metodo Principal    
    def alimentar_batch_input(self):
        file_name_saved = self._caminho_salvar()
        print(file_name_saved)
        modelo_file_path = self.__modelo_file_path
        modelo_file_path_copy = modelo_file_path.replace(".xlsx", "_temp.xlsx")
        copy2(modelo_file_path, modelo_file_path_copy)
        
        lista_intercompany = self._preparar_lista_alimentacao(mod='Intercompany')
        lista_passivo = self._preparar_lista_alimentacao(mod='Passivo')
        
        fechar_excel(modelo_file_path_copy) 
        app = xw.App(visible=False)
        with app.books.open(modelo_file_path_copy)as wb:
            self._alimentar_celular(wb=wb,sheet="B.I. Intercompany",lista_alimentar=lista_intercompany)
            self._alimentar_celular(wb=wb,sheet="B.I. Passivo",lista_alimentar=lista_passivo)
            wb.save(file_name_saved)
        fechar_excel(modelo_file_path_copy)
        os.unlink(modelo_file_path_copy)
        return {"file_name_saved":file_name_saved, "modelo_file_path_copy":modelo_file_path_copy}
    
    def _caminho_salvar(self):
        options = {}
        options['defaultextension'] = ".xlsx"
        options['filetypes'] = [("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*")]
        options['initialfile'] = f"BATCH INPUT {self.date.strftime('%B %Y')}.xlsx"
        arquivo_salvar = filedialog.asksaveasfilename(**options)
        if not arquivo_salvar:
            return f"C:\\Users\\{getuser()}\\Downloads\\MODELO BATCH INPUT.xlsx"
        return arquivo_salvar
        
        # options = QFileDialog.Options()
        # defaultFileName = f"BATCH INPUT {self.date.strftime('%B %Y')}.xlsx"
        # arquivo_salvar, _ = QFileDialog.getSaveFileName(None, "Salvar Arquivo", defaultFileName, "Planilhas Excel (*.xlsx)", options=options)
        # print(arquivo_salvar)
        # if not arquivo_salvar:
        #     raise FileNotFoundError("tela para salvar o arquivo foi encerrada sem selecionar o arquivo!")
        # return arquivo_salvar
    
    def _alimentar_celular(self, *, wb, sheet:str, lista_alimentar:dict):
        ws:Sheet = wb.sheets[sheet] 
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
    
    def _preparar_lista_alimentacao(self, *,
                                   mod:Literal["Intercompany", "Passivo"]
                                   ):
        
        #df_base_menor = df_base[df_base['Montante em moeda interna'] < "0"]
        sequencial = 0
        linha = 1
        ultimo_dia_mes = obter_ultimo_dia_mes(self.date ,forstr='%d.%m.%Y')
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
        
        df:pd.DataFrame
        if mod == 'Intercompany':
            df = self.df_base[self.df_base['Texto'].str.contains('Intercompany', na=False)]
        elif mod == 'Passivo':
            df = self.df_base[~self.df_base['Texto'].str.contains('Intercompany', na=False)]
        else:
            raise Exception(f"incorrect {mod=} selection")
        
        for row,dados in df.iterrows():
            sequencial += 1
            
            montante = Classific(dados['Montante em moeda interna'], sem_negativo=True)
            
            #Linha 1
            lista_alimentar["sequencial"].append(sequencial)
            lista_alimentar["ultimo_dia_mes"].append(ultimo_dia_mes)
            lista_alimentar["Empresa"].append(dados['Empresa'])
            lista_alimentar["Divisão"].append(dados['Divisão'])  
            lista_alimentar["tipo_documento"].append('AB') 
            lista_alimentar["Texto cabeçalho/Referencia"].append('RECLASSIFICAÇÃO PARTES RELACIONAS')
            lista_alimentar["Chave do Lançamento"].append(montante.chave_primaria) 
            lista_alimentar["valor"].append(montante.value) 
            lista_alimentar["tipo de conta"].append(montante.tipo_conta_primeira) 
            lista_alimentar["Conta"].append(dados['Fornecedor'] if montante.tipo_conta_primeira == "K" else self._tratar_conta(dados['Conta']))
            lista_alimentar["texto"].append(dados['Texto'])
            
            #Linha 2
            lista_alimentar["sequencial"].append(sequencial)
            lista_alimentar["ultimo_dia_mes"].append("")
            lista_alimentar["Empresa"].append("")
            lista_alimentar["Divisão"].append(dados['Divisão'])  
            lista_alimentar["tipo_documento"].append('')
            lista_alimentar["Texto cabeçalho/Referencia"].append('')
            lista_alimentar["Chave do Lançamento"].append(montante.chave_secundaria) 
            lista_alimentar["valor"].append(montante.value) 
            lista_alimentar["tipo de conta"].append(montante.tipo_conta_secundaria) 
            lista_alimentar["Conta"].append(dados['Fornecedor'] if montante.tipo_conta_secundaria == "K" else self._tratar_conta(dados['Conta']))
            lista_alimentar["texto"].append(dados['Texto']) 
        
        return lista_alimentar         
    
    # def get_ultimo_dia_mes(self) -> datetime:
    #     now = datetime.now()
    #     date_temp = datetime(year=now.year, month=(now + relativedelta(months=1)).month, day=1)
    #     return date_temp - relativedelta(days=1)

    @staticmethod
    def _tratar_conta(conta):
        conta_temp = str(conta)
        if len(conta_temp) == 10:
            return f"1202{conta_temp[4:]}"
        else:
            return conta

if __name__ == "__main__":
    pass
    #print(ExcelData._tratar_conta("2204010018"))
