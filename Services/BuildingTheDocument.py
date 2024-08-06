import docx
import os
from pptx import Presentation
from docx import Document



class BuildingTheDocument:

    def __init__(self, table_data,data):
        self.table_data = table_data
        self.data = data
    
    def MontandoODocx(self):

        doc = Document()
        