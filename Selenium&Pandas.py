import selenium         as sd
import CreateAFileExcel as cafe
import string           as AphaB

# Importa bibliotecas necessarios para à automação no Google
from   webdriver_manager.chrome          import ChromeDriverManager
from   selenium.webdriver.chrome.service import Service
from   selenium.webdriver.common.by      import By
from   selenium.webdriver.common.keys    import Keys


# Atualiza automaticamente seu webdriver (Sem firewall)
def Driver_AutoInstall():
    servico = Service(ChromeDriverManager().install())
    navegador = sd.webdriver.Chrome(service=servico)

    return navegador

class main():
    def __init__(self):
        # Sites para captura de dados
        self.Site_Info_PlacasDeVideo:str =(
            'https://brotherss.com.br/blog/o-ranking-definitivo-de-placas-de-video/')

        self.PlacasV = []
        self.Titulos = []
        self.lineL:int = 1

    def Pega_Valores_PV(self):
        self.GoogleG = Driver_AutoInstall()
        self.GoogleG.get(self.Site_Info_PlacasDeVideo)
        
        TabelaL = self.GoogleG.find_element(
            By.XPATH, '//*[@id="post-3058"]/div[3]/figure[2]/table/tbody/tr')
        TodasAsLinhas = self.GoogleG.find_element(
            By.XPATH, '//*[@id="post-3058"]/div[3]/figure[2]/table/tbody')

        LinhasG = TodasAsLinhas.find_elements(By.TAG_NAME,'tr')
        Dados = TabelaL.find_elements(By.TAG_NAME,'td')

        # Definer o tamanhos das arrays de acordo com a quantidade de colunas 
        # linhas da tabeça
        self.Titulos = Dados[:] 
        self.PlacasV = LinhasG[:]

        coluna = 1
        f:int = 0
        for coluna in LinhasG:

            Dados = coluna.find_elements(By.TAG_NAME,'td')
            i:int = 0
            for linhas in Dados:
                self.Titulos[i] = linhas.text      
                if i == 0:
                    self.PlacasV[f] = linhas.text
                    print(i, linhas.text)
                    f += 1
                if i == 5:
                    self.Coloca_No_ArquivoExcel(self.Titulos)
                i += 1

        
    def Coloca_No_ArquivoExcel(self, ListNum):
        NameExcel = cafe.instanceL.Cria_Caminho()
        
        for coll in range(len(ListNum)):
            RangeCell = str(AphaB.ascii_uppercase[coll] + str(self.lineL))
            NameExcel.Workbooks.Application.Range(RangeCell).Value = ListNum[coll]

        self.lineL += 1
        if self.lineL == 60:
            NameExcel.ActiveWorkbook.Save()
            


mainInstance = main()
mainInstance.Pega_Valores_PV()