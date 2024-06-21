import xlwings as xw
from xlwings.main import Book
import os
from time import sleep
import traceback

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
                        print(f"fechou {app_open.name}")
                        app_open.close()
                        if len(xw.apps) == 0:
                            app.kill()
                        
                        return True
            if not multiplas_tentativas:
                break
            sleep(1)
        return False    
    except Exception as error:
        print(f"não foi possivel fechar o {caminho=}\nError: {traceback.format_exc()}")
        if trace_back:
            raise error
        return error
    
def excel_abertos():
    print("lista de abertos:")
    for app in xw.apps:
        for app_open in app.books:
            app_open:Book
            print(f"{app_open.name} - ainda aberto!")
                

if __name__ == "__main__":
    print(len(xw.apps))