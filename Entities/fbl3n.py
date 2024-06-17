from .logs import Log
from .sap import SAPManipulation
from getpass import getuser
from datetime import datetime
import os

class FBL3N(SAPManipulation):
    @property
    def log(self):
        return Log(self.__class__.__name__)
    
    def __init__(self,) -> None:
        super().__init__(using_active_conection=True)
        
    @SAPManipulation.start_SAP
    def gerar_relatorio(self, *, path:str=f"C:\\Users\\{getuser()}\\Downloads", name:str=datetime.now().strftime("relatorio_partes_relacionadas_%d%m%Y%H%M%S.xlsx")) -> str:
        """Executa a transação no sap e gera o relatorio em seguida salva no caminho especificado e retorna o caminho de onde o arquivo foi salvo

        Args:
            path (str, optional): caminho onde será salvo o arquivo. Defaults to f"C:\Users\{getuser()}\Downloads".
            name (str, optional): nome do arquivo que será salvo. Defaults to datetime.now().strftime("relatorio_partes_relacionadas_%d%m%Y%H%M%S.xlsx").

        Raises:
            NotADirectoryError: caso não consiga validar o caminho informado

        Returns:
            str: o caminho de onde o arquivo foi salvo
        """
        try:
            if not os.path.exists(path):
                os.makedirs(path)
        except:
            self.log.register_error()
            raise NotADirectoryError("erro ao validar caminho!")
        if not name.endswith(".xlsx"):
            name += ".xlsx"
        
        self.session.findById("wnd[0]/tbar[0]/okcd").text = "/n fbl3n"
        self.session.findById("wnd[0]").sendVKey(0)
        self.session.findById("wnd[0]/tbar[1]/btn[17]").press()
        self.session.findById("wnd[1]/usr/txtV-LOW").text = "RAZÃO 2204"
        self.session.findById("wnd[1]/usr/txtENAME-LOW").text = "edias"
        self.session.findById("wnd[1]/tbar[0]/btn[8]").press()
        self.session.findById("wnd[0]/usr/ctxtPA_STIDA").text = "31.05.2024"
        self.session.findById("wnd[0]/tbar[1]/btn[8]").press()
        self.session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setCurrentCell (4,"BELNR")
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
        
        return os.path.join(path, name)

if __name__ == "__main__":
    pass