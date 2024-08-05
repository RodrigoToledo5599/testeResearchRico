import docx
import os
from pptx import Presentation
from docx import Document
# from docx2pdf import convert

# por enquanto o valor do nome do arquivo será mockado

class TakingData:
    

    def __init__(self, file_name):
        self.file_name = os.path.join(os.getcwd(),f"Data\{file_name}")    

    def FetchingDataFromTable(self):
        doc = Document(self.file_name)
        table = doc.tables[0]

        first_cells_data = []
        first_cells_data_pop_fundos = []

        for row in table.rows:
            first_cells_data.append(row.cells[0])


        for i in range(len(first_cells_data)):
            if(i == 0):
                continue
            first_cells_data_pop_fundos.append(first_cells_data[i])

        return first_cells_data_pop_fundos


            
            

        









    



if __name__ == "__main__":
    dd = TakingData("Guia de Investimentos (lista de Trends).docx")
    # dd.FetchingDataFromTable()