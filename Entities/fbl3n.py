import traceback
from Entities.dependencies.logs import Logs
from .sap import SAPManipulation
from getpass import getuser
from datetime import datetime
from .functions import fechar_excel, excel_abertos, ultimo_dia_mes
import os

class FBL3N(SAPManipulation):
    @property
    def log(self):
        return Logs()
    
    def __init__(self,) -> None:
        super().__init__(using_active_conection=True)
        
    @SAPManipulation.start_SAP
    def gerar_relatorio(self, *,
                        date_partidas_aberto:datetime,
                        date_fechamento:datetime,
                        path:str=os.path.join(os.getcwd(), "relatorios_sap")
                        ) -> str:
        """Executa a transação no sap e gera o relatorio em seguida salva no caminho especificado e retorna o caminho de onde o arquivo foi salvo

        Args:
            path (str, optional): caminho onde será salvo o arquivo. Defaults to f"C:\\Users\\{getuser()}\\Downloads".
            name (str, optional): nome do arquivo que será salvo. Defaults to datetime.now().strftime("relatorio_partes_relacionadas_%d%m%Y%H%M%S.xlsx").

        Raises:
            NotADirectoryError: caso não consiga validar o caminho informado

        Returns:
            str: o caminho de onde o arquivo foi salvo
        """
        
        name:str=datetime.now().strftime("relatorio_partes_relacionadas_%d%m%Y%H%M%S.xlsx")
        
        try:
            if not os.path.exists(path):
                os.makedirs(path)
        except Exception as error:
            self.log.register(status='Error', description=str(error), exception=traceback.format_exc())
            raise NotADirectoryError("erro ao validar caminho!")
        if not name.endswith(".xlsx"):
            name += ".xlsx"
        
        try:
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/n fbl3n"
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById("wnd[0]/tbar[1]/btn[17]").press()
        except:
            raise Exception("O Sap está Fechado!")
        self.session.findById("wnd[1]/usr/txtV-LOW").text = "RAZÃO 2204"
        self.session.findById("wnd[1]/usr/txtENAME-LOW").text = "edias"
        self.session.findById("wnd[1]/tbar[0]/btn[8]").press() 
        self.session.findById("wnd[0]/usr/ctxtPA_STIDA").text = date_partidas_aberto.strftime("%d.%m.%Y")
        self.session.findById("wnd[0]/tbar[1]/btn[8]").press()
        
        #import pdb;pdb.set_trace()
        self.session.findById("wnd[0]/tbar[1]/btn[38]").press()
        
        contador:int = 0
        while True:
            try:
                nome_coluna:str = self.session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_FILTER_CRITERIA:0600/cntlCONTAINER1_FILT/shellcont/shell").GetCellValue(contador,"SELTEXT")
            except:
                break
            if nome_coluna == 'Data de lançamento':
                self.session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_FILTER_CRITERIA:0600/cntlCONTAINER1_FILT/shellcont/shell").currentCellRow = contador
                self.session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_FILTER_CRITERIA:0600/cntlCONTAINER1_FILT/shellcont/shell").selectedRows = str(contador)
                self.session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_FILTER_CRITERIA:0600/cntlCONTAINER1_FILT/shellcont/shell").doubleClickCurrentCell()
                break
            contador+=1
                
        self.session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_FILTER_CRITERIA:0600/btn600_BUTTON").press()
        
        self.session.findById("wnd[2]").sendVKey(2)
        self.session.findById("wnd[3]/usr/cntlOPTION_CONTAINER/shellcont/shell").setCurrentCell(1,"TEXT")
        self.session.findById("wnd[3]/usr/cntlOPTION_CONTAINER/shellcont/shell").selectedRows = "1"
        self.session.findById("wnd[3]/tbar[0]/btn[0]").press()
        self.session.findById("wnd[2]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-HIGH").text = date_fechamento.strftime("%d.%m.%Y")
        self.session.findById("wnd[2]/tbar[0]/btn[0]").press()

        self.session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").contextMenu()
        self.session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectContextMenuItem("&XXL")
        self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
        self.session.findById("wnd[1]/usr/ctxtDY_PATH").text = path
        self.session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = name
        self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
        try:
            self.session.findById("wnd[1]/tbar[0]/btn[12]").press()
        except:
            pass
        
        result:str = os.path.join(path, name)
        
        fechar_excel(result, multiplas_tentativas=True, wait=2)
        
        
        return result

if __name__ == "__main__":
    pass
