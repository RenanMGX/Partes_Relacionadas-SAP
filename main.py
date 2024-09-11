import os
import traceback
import sys
from Entities.dependencies.logs import Logs
from Entities.fbl3n import FBL3N
from Entities.excel_data import ExcelData
from typing import Literal
from Entities.functions import fechar_excel, relativedelta
from datetime import datetime
from Entities.interface import Interface, QtWidgets, QtGui
from Entities.dependencies.functions import P, Functions


class Execute(Interface):
    @property
    def modelo_path(self) -> str:
        modelo_path:str = os.path.join(os.getcwd(), self.__modelo_file_name)
        if not os.path.exists(modelo_path):
            from Entities.modelo import modelo
            with open(self.__modelo_file_name, 'wb')as _file:
                _file.write(modelo)
            del modelo            
            #return FileNotFoundError("Não Encontrado!")
        return modelo_path
    
    @property
    def date(self) -> datetime:
        date = self.janela_1_widget_calendario.selectedDate()
        return datetime(date.year(), date.month(), date.day())

    def __init__(self, 
                 #*,
                 #date:datetime,
                 #modelo_file:str="MODELO BATCH INPUT.xlsx",
                 ) -> None:
        
        #self.__date:datetime = date
        self.__modelo_file_name:str = "MODELO BATCH INPUT.xlsx"
        super().__init__(version="1.0") # <--------------------------------- Alterar Versão antes de compilar
        self.setupUi()
        self.__initial_config()
        self.__files_created:dict = {}
        
    def closeEvent(self, event:QtGui.QCloseEvent):
        print(P("Encerrando Script"))
        Functions().fechar_excel("Pasta1")
        if self.__files_created:
                Functions().fechar_excel(str(self.__files_created.get("modelo_file_path_copy"))) 
        event.accept()  # Isso irá fechar a janela
    
    def __initial_config(self):
        self.janela_1_bt_extrair.clicked.connect(self.execute)
        
    def test(self, *, t):
        print(self.date)
    
    def execute(self, *,trace_back:bool=True):
        
        # if isinstance(self.modelo_path, FileNotFoundError):
        #     raise FileNotFoundError(f"Modelo Batch Input é necessario para continuar")
        
        self.showMinimized()
        self.janela_1_label_textoInfor.setText("")
        try:
            relatorio_path = FBL3N().gerar_relatorio(date=self.date)
            
            self.__files_created = ExcelData(date=self.date,dados_entrada_path=relatorio_path, modelo_file=self.modelo_path).alimentar_batch_input()
            
            os.unlink(relatorio_path)
            fechar_excel("Pasta1", wait=2)
            Logs().register(status='Concluido', description="automação encerrou com exito!", exception=None)
        except Exception as error:
            self.janela_1_label_textoInfor.setText(str(error))
            Logs().register(status='Error', description=str(error), exception=traceback.format_exc())
        finally:
            #self.ir_pagina_2()
            self.showNormal()
            
                 
if __name__ == "__main__":
    # date:datetime = datetime.now() - relativedelta(months=0)
    # print(date)
    # response = input("Continuar[S/N]? ")
    # if response.lower() == 's':
    
    #     execution = Execution(date=date)
    
    #     print(execution.execute())
    app = QtWidgets.QApplication(sys.argv)
    ui = Execute()
    ui.show()
    sys.exit(app.exec_())

    