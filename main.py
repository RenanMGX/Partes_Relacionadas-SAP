import os
import pdb

from Entities.fbl3n import FBL3N
from Entities.excel_data import ExcelData
from typing import Literal
from Entities.functions import fechar_excel, relativedelta
from datetime import datetime

class Execution:
    @property
    def modelo_path(self) -> str|FileNotFoundError:
        modelo_path:str = os.path.join(os.getcwd(), self.__modelo_file)
        if not os.path.exists(modelo_path):
            return FileNotFoundError("Não Encontrado!")
        return modelo_path
    
    @property
    def date(self) -> datetime:
        return self.__date

    def __init__(self, *,
                 date:datetime,
                 modelo_file:str="MODELO BATCH INPUT.xlsx",
                 ) -> None:
        
        self.__date:datetime = date
        self.__modelo_file:str = modelo_file
        self.__fbl3n:FBL3N = FBL3N()
    
    def execute(self, *,trace_back:bool=True):
        if isinstance(self.modelo_path, FileNotFoundError):
            raise FileNotFoundError(f"Modelo Batch Input é necessario para continuar")
        
        relatorio_path = self.__fbl3n.gerar_relatorio(date=self.date)
        
        ExcelData(date=self.date,dados_entrada_path=relatorio_path, modelo_file=self.modelo_path).alimentar_batch_input()
        
        os.unlink(relatorio_path)
        fechar_excel("Pasta1", wait=2)
        return True
                 
if __name__ == "__main__":
    date:datetime = datetime.now() - relativedelta(months=2)
    
    execution = Execution(date=date)
    
    print(execution.execute())
    