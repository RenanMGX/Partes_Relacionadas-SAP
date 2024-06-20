import os
import pandas as pd
import pdb

from Entities.fbl3n import FBL3N
from Entities.excel_data import ExcelData
from typing import Literal
from Entities.functions import fechar_excel

class Execution:
    @property
    def modelo_path(self) -> str|FileNotFoundError:
        modelo_path:str = os.path.join(os.getcwd(), self.__modelo_file)
        if not os.path.exists(modelo_path):
            return FileNotFoundError("Não Encontrado!")
        return modelo_path
        
    
    def __init__(self, *,
                 modelo_file:str="MODELO BATCH INPUT.xlsx",
                 ) -> None:
        
        self.__modelo_file:str = modelo_file
        self.__fbl3n:FBL3N = FBL3N()
    
    def execute(self, *,trace_back:bool=True):
        if isinstance(self.modelo_path, FileNotFoundError):
            raise FileNotFoundError(f"Modelo Batch Input é necessario para continuar")
        
        relatorio_path = self.__fbl3n.gerar_relatorio()
        
        ExcelData(relatorio_path, modelo_file=self.modelo_path).alimentar_batch_input()
        
        os.unlink(relatorio_path)
        fechar_excel("Pasta1", wait=2)
        return True
                 
if __name__ == "__main__":
    execution = Execution()
    
    print(execution.execute())
    