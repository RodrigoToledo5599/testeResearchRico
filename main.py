from Utils.Utils import Utils 
from Services.TakingData import TakingData











if __name__ == "__main__":
    utils = Utils()
    TkData = TakingData("Guia de Investimentos (lista de Trends).docx")


    utils.clearOutputDir()
    fundos = TkData.FetchingDataFromTable()
    
    for i in fundos:
        print(i.text)
    