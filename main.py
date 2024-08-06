from Utils.Utils import Utils 
from Services.TakingData import TakingData
from Services.BuildingTheDocument import BuildingTheDocument
import os





if __name__ == "__main__":
    utils = Utils()
    relatorio_name = os.listdir("Data")
    utils.clearOutputDir() # limpa a saida do programa, deve ser inicializado antes de declarar TkData
    TkData = TakingData(relatorio_name[0])
    TkData.FetchingTableData()
    

    
    buildDoc = BuildingTheDocument(TkData.table_data, TkData.data,TkData.table_data2)
    buildDoc.MontandoODocx()



