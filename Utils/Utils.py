import os



class Utils:


    #remove todos os arquivos da pasta Output
    def __init__(self):
        ...

    def clearOutputDir(self):
        dd = os.getcwd()
        dir = os.path.join(dd,"Output")
        for file in os.listdir(dir):
            file_path = os.path.join(dir, file)
            print("Excluindo: ",file_path)
            os.remove(file_path)