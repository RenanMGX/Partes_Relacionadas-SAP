import xlwings as xw
from xlwings.main import Book
import os
from time import sleep

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
                        app_open.close()
                        return True
            if not multiplas_tentativas:
                break
            sleep(1)
        return False    
    except Exception as error:
        print(f"não foi possivel fechar o {caminho=}\nError: {error=}")
        if trace_back:
            raise error
        return error
                

if __name__ == "__main__":
    fechar_excel(r"C:\Users\renan.oliveira\Downloads\relatorio_partes_relacionadas_19062024182439.xlsx", multiplas_tentativas=True)