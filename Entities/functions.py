import xlwings as xw
from xlwings.main import Book
from time import sleep
import traceback
from datetime import datetime
from dateutil.relativedelta import relativedelta

from Entities.dependencies.functions import P


def fechar_excel(caminho:str, *, 
                trace_back:bool=False,
                wait:int=0,
                multiplas_tentativas:bool=False,
                timeout:int=60
                ) -> bool|Exception:
    if wait > 0:
        if isinstance(wait, int):
            sleep(wait)
    
    try:
        for _ in range(timeout):
            for app in xw.apps:
                for app_open in app.books:
                    app_open:Book
                    if app_open.name in caminho:
                        print(P(f"fechou {app_open.name}"))
                        app_open.close()
                        if len(xw.apps) == 0:
                            app.kill()
                        
                        return True
            if not multiplas_tentativas:
                break
            sleep(1)
        return False    
    except Exception as error:
        print(P(f"não foi possivel fechar o {caminho=}\nError: {traceback.format_exc()}"))
        if trace_back:
            raise error
        return error
    
def excel_abertos():
    print("lista de abertos:")
    for app in xw.apps:
        for app_open in app.books:
            app_open:Book
            print(f"{app_open.name} - ainda aberto!")
  
def ultimo_dia_mes(date:datetime=datetime.now(), *, forstr:str=""):
    date_temp = datetime(year=date.year,month=date.month,day=1)
    date_temp = date_temp + relativedelta(months=1)
    date_temp = date_temp - relativedelta(days=1)
    if forstr == "":
        return date_temp
    else:
        return date_temp.strftime(forstr)
    
class Classific:
    @property
    def value(self) -> str:
        return str(self.__value)
    
    @property
    def positivo(self) -> bool:
        if self.__newValue >= 0:
            return True
        else:
            return False
    
    @property
    def chave_primaria(self):
        if self.positivo:
            return '50'
        else:
            return '21'
    
    @property
    def chave_secundaria(self):
        if self.positivo:
            return '21'
        else:
            return '50'
        
    @property
    def tipo_conta_primeira(self):
        if self.positivo:
            return 'S'
        else:
            return 'K'
    
    @property   
    def tipo_conta_secundaria(self):
        if self.positivo:
            return 'K'
        else:
            return 'S'
        
    def __init__(self, value:str|int|float, *, sem_negativo:bool=False) -> None:
        if sem_negativo:
            value_temp = float(value)
            if value_temp < 0:
                value = -float(value)
        self.__value = value
        new_value:float = float(value)
        self.__newValue:int|float = new_value
    

if __name__ == "__main__":
    num = Classific(-100, sem_negativo=True)
    print(num.value)
    print(num.chave_primaria,num.chave_secundaria)
    print(num.tipo_conta_primeira,num.tipo_conta_secundaria)