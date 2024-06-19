import os
import pandas as pd
import pdb

from Entities.fbl3n import FBL3N
from typing import Literal

class Execution:
    @property
    def modelo_path(self) -> str|FileNotFoundError:
        modelo_path:str = os.path.join(os.getcwd(), self.__modelo_file)
        if not os.path.exists(modelo_path):
            return FileNotFoundError("Não Encontrado!")
        return modelo_path
    
    @property
    def relatorio_path(self) ->str:
        return self.__relatorio_path
    
    @property
    def df(self) -> pd.DataFrame:
        return self.__df
    
    def __init__(self, *,
                 modelo_file:str="MODELO BATCH INPUT.xlsx",
                 ) -> None:
        
        self.__modelo_file:str = modelo_file
        self.__fbl3n:FBL3N = FBL3N()
    
    def execute(self, *,trace_back:bool=True):
        self.__relatorio_path = self.__fbl3n.gerar_relatorio()
        
        if self.relatorio_path.endswith('.xlsx'):
            self.__df:pd.DataFrame = pd.read_excel(self.relatorio_path, dtype=str)
            self.__df['Data do documento'] = self.__df['Data do documento'].astype('datetime64[us]')
            self.__df['Data de lançamento'] = self.__df['Data de lançamento'].astype('datetime64[us]')
        else:
            if trace_back:
                raise Exception(f"o Relatorio {self.relatorio_path=} deve ser .xlsx")
        return self
        
    
if __name__ == "__main__":
    execution = Execution()
    print(execution.execute().df)