import docx
import os
from pptx import Presentation
from docx import Document
from docx2pdf import convert

# Esse espaço é apenas para eu testar algumas funções


class TestingStuff:
    
    def __init__(self, file_name):
        self.file_name = file_name



    def generateNewDocx(self):
        docxpath = os.getcwd()
        datapath = os.path.join(docxpath,f"Data\{self.file_name}")
        returnpath = os.path.join(docxpath,f"Output\{self.file_name}")

        doc = Document(datapath)        
        
        doc.save(returnpath)

    
    
    def generateNewPttx(self):
        
        pptxpath = os.getcwd()
        modelpath = os.path.join(pptxpath,f"ModeloCarteira\{self.file_name}")
        returnpath = os.path.join(pptxpath,f"Output\{self.file_name}")


        pttx = Presentation(modelpath)
        pttx.save(returnpath)



    def clearOutputDir(self):
        dd = os.getcwd()
        dir = os.path.join(dd,"Output")
        for file in os.listdir(dir):
            caminho_completo = os.path.join(dir, file)
            print("Excluindo: ",caminho_completo)
            os.remove(caminho_completo)

        
    def word2pdf(self):
        dd = os.getcwd()
        dir = os.path.join(dd,f"Data\Guia de Investimentos (lista de Trends).docx")
        outputDir = os.path.join(os.getcwd(),"/Output")


        convert(dir)
        convert(dir, dd+f"\Data\Guia de Investimentos (lista de Trends).pdf")
        convert(outputDir)






if __name__ == "__main__":

    # dd = TestingStuff("Guia de Investimentos (lista de Trends).docx")
    dd = TestingStuff("Energetica.pptx")
    # dd.clearOutputDir()
    dd.word2pdf()

