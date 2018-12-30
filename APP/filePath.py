import os.path

class pathFiles():
    def __init__(self,resourcesPath):
        self.resourcesPath = resourcesPath
        self.resourcesXlsx=[]
        for file in os.listdir(self.resourcesPath):
            if file.endswith(".xlsx"):
                self.resourcesXlsx.append(os.path.join(self.resourcesPath, file))


        self.indiceOC = encontrarArchivos(self.resourcesXlsx,'Oc ')
        self.indiceDOD = encontrarArchivos(self.resourcesXlsx,'DoD ')
        self.indiceConsolidado = encontrarArchivos(self.resourcesXlsx,'ConsolidadoOperativo')

        self.OC_FilePath = self.resourcesXlsx[self.indiceOC]
        self.DOD_FilePath = self.resourcesXlsx[self.indiceDOD]
        self.consolidado = self.resourcesXlsx[self.indiceConsolidado]
        print(self.OC_FilePath)
        print(self.DOD_FilePath)
def encontrarArchivos(resourcesXlsx,abrFile):

    for i,file in enumerate(resourcesXlsx):
        if file.find(abrFile) > -1:
            indice = i

    return indice
