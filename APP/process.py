from datetime import datetime
global DOD_FilePath,OC_FilePath
import pandas as pd
import numpy as np
import calendar
import os

def pathConstructor(path):
    global DOD_FilePath,OC_FilePath,consolidado
    DOD_FilePath = path.DOD_FilePath
    OC_FilePath = path.OC_FilePath
    consolidado = path.consolidado

def startproccess(path):
    OC_DataSet = pd.read_excel(OC_FilePath)
    DOD_DataSet = pd.read_excel(DOD_FilePath)
    Dtconsolidado = pd.read_excel(consolidado)
    dataSetFinal = proccess(OC_DataSet,DOD_DataSet)
    year,month = detectYearAndMonth(dataSetFinal)
    dataFinalCleaned = CleanCols(dataSetFinal)
    incidentesSoporte,incidentesTransformacion,cambiosSoporte,cambiosTransformacion,pruebasTerminadas = indicators(dataFinalCleaned)
    Dtconsolidado = dataconsolidated(year,month,incidentesSoporte,cambiosSoporte,Dtconsolidado,'Soporte',pruebasTerminadas)
    Dtconsolidado = dataconsolidated(year,month,incidentesTransformacion,cambiosTransformacion,Dtconsolidado,'Transformacion',pruebasTerminadas)
    ExportarExcel(Dtconsolidado,consolidado)

def indicators(dataFinalCleaned):
    datframeI1 = dataFinalCleaned[dataFinalCleaned['Codigo Cierre'] == 'Ejecutado - Con Incidente']
    datframeI2 = datframeI1[(datframeI1['Carta de certificación'] == 1) | (datframeI1['carta sin pruebas'] == 1) | (datframeI1['Carta de certificación condicionada'] == 1) | (datframeI1['GIT'] == 1)]
    datframeSoporte = datframeI2[datframeI2['Dirección'] == 'Soporte']
    datframeTransformacion = datframeI2[datframeI2['Dirección'] == 'Transformacion']
    cantidadDeCambiosConIncidentesSoporte  = len(datframeSoporte)
    cantidadDeCambiosConIncidentesTransf  = len(datframeTransformacion)

    datframeI2 = dataFinalCleaned[(dataFinalCleaned['Carta de certificación'] == 1) | (dataFinalCleaned['carta sin pruebas'] == 1) | (dataFinalCleaned['Carta de certificación condicionada'] == 1) | (dataFinalCleaned['GIT'] == 1)]
    datframeSoporte = datframeI2[datframeI2['Dirección'] == 'Soporte']
    datframeTransformacion = datframeI2[datframeI2['Dirección'] == 'Transformacion']
    cambiosSoporte = len(datframeSoporte)
    cambiosTransformacion = len(datframeTransformacion)
    pruebasTerminadas = cambiosSoporte + cambiosTransformacion
    return cantidadDeCambiosConIncidentesSoporte,cantidadDeCambiosConIncidentesTransf,cambiosSoporte,cambiosTransformacion,pruebasTerminadas

def dataconsolidated(year,month,cantidadDeCambiosConIncidentes,cambios,Dtconsolidado,direccion,pruebasTerminadas):
    rowToAppend = [year,month,direccion,cantidadDeCambiosConIncidentes,cambios,pruebasTerminadas]

    DataFrameToAppend = pd.DataFrame(columns=['Año','Mes','Direccion','Cambios con incidentes','total cambios','Pruebas terminadas'])
    DataFrameToAppend.loc[0, :] = rowToAppend
    Dtconsolidado = Dtconsolidado.append(DataFrameToAppend, ignore_index=True)
    return Dtconsolidado

def detectYearAndMonth(dtFrame):
    monthCol = dtFrame['Fecha Inicio']
    date = pd.DataFrame(monthCol)
    month = date['Fecha Inicio'].dt.month
    year = date['Fecha Inicio'].dt.year
    year = max(year)
    month = min(month)
    meses = { 0 : 'Diciembre', 1 : 'Enero', 2: 'Febrero',3:'Marzo',
             4:'Abril',5:'Mayo',6:'Junio',7:'Julio',8:'Agosto',9:'Septiembre',
             10:'Octubre',11:'Noviembre',12:'Diciembre' }

    month = meses[month]
    return year,month

def  CleanCols(dataSetFinal):
    teamCol = dataSetFinal['Dirección'].tolist()
    for i in range(len(teamCol)):
        teamCol[i]=teamCol[i].split(' ')[1]

    dataSetFinal['Dirección'] = pd.DataFrame(teamCol)
    return dataSetFinal


def proccess(OC_DataSet,DOD_DataSet):
    OC_DataSet = renombrarColumnaID(OC_DataSet)
    OC_DataSet['ID2']=pd.to_numeric(OC_DataSet['ID2'], errors='coerce')
    OC_DataSet['ID2'] = OC_DataSet['ID2'].fillna(0).astype(np.int64)
    DOD_Final = DOD_DataSet.merge(OC_DataSet, left_on='ID', right_on='ID2', how='inner')
    dataSetFinal = DOD_Final[['ID','Carta de certificación','carta sin pruebas','Carta de certificación condicionada','GIT',
          'Fecha de vencimiento carta de certificación','Aprueba el Paso a Producción','Dueño de producto que aprueba  paso a producción',
          'Orden de  Cambio','Codigo Cierre','Dirección','Fecha Inicio']]
    return dataSetFinal

def renombrarColumnaID(OC_DataSet):
    OC_DataSet = OC_DataSet.rename(columns = {'Código Definición de Terminado':'ID2'})
    return OC_DataSet

def ExportarExcel(dataSetFinal,rutaPath):
    print(rutaPath)
    dataSetFinal.to_excel(rutaPath,index=False)
    print('exportado a {}'.format(rutaPath))
