from docx import Document
from docx2pdf import convert
import os



class BuildingTheDocument:

    def __init__(self, table_data,data):
        self.table_data = table_data
        self.data = data
    
    def MontandoODocx(self):

        doc = Document()
        
        dir = os.path.join(os.getcwd(),r"Output")
        print(self.data[1])
        for i in range(len(self.data)):
            doc.add_paragraph(self.table_data[i])
            doc.add_paragraph(self.data[i])
        
        
        
        doc.save(dir+r"\result.docx")