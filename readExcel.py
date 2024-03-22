# import pandas lib as pd
import numpy as np
import pandas as pd
# import pandas.io.formats.style
import datetime
import os
import re
import shutil

from classDefinition import TotaleSospesiNuovi
from companiesFunction import *
from auxiliaryFunction import getAgencyFromSubagent

try:
    totale_sospesi_vecchi = 0
    totale_sospesi_nuovi = TotaleSospesiNuovi()

    current_working_directory = os.getcwd()

    print("Percorso attuale: ", current_working_directory)

    month_folder = input("\nInserire nome cartella del mese + anno: ")
    # path = r'C:\LUIGI 04052016\AMICONE LUIGI\DATI DAL 31032008 PC PORTATILE\DATI\CONTABILITA\PARTITE REGISTRATE PER CONTABILITA\GENERALI\PARTITE REGISTRATE\FEBBRAIO 2024'

    start_time = time.time()
    # Conversione in UPPER CASE del mese in input inserito in quanto la cartella in cui si trovano tutti i file ha nome (es.) "FEBBRAIO 2024"
    month_folder = month_folder.upper()

    # Calcolo della data dell'ultimo giorno del mese inserito come input in modo tale da non andare oltre tale data nell'inserimento della stringa 'Eseguito' nel foglio 'SOSPESI'
    datetimeSelectedMonthYear = convertMonthYearString_toDatetime(month_folder)
    lastMonthYearDatetime = getLastDatetimeOfAMonth(datetimeSelectedMonthYear)

    partialDir_filesGENERALI    = r'\PARTITE REGISTRATE PER CONTABILITA\GENERALI\2024' + '\\' + month_folder
    partialDir_filesCATTOLICA   = r'\PARTITE REGISTRATE PER CONTABILITA\CATTOLICA\2024' + '\\' + month_folder
    partialDir_filesTUTELA      = r'\PARTITE REGISTRATE PER CONTABILITA\TUTELA LEGALE\2024' + '\\' + month_folder

    finalFileName = 'PRIMA_NOTA_TEST_.xlsx'
    backupFileName = 'BACKUP_PRIMA_NOTA_TEST_.xlsx'
    finalPathName = current_working_directory + '\\' + finalFileName

    # Se esiste gia' un file backupFileName, allora lo elimino
    if(os.path.exists(current_working_directory + '\\' + backupFileName)):
        os.remove(current_working_directory + '\\' + backupFileName)
        print("File ", backupFileName, " rimosso.\n")
    else:
        print("File ", backupFileName, " non trovato.\n")

    # Copia di backup del file PRIMA_NOTA_TEST_.xlsx
    shutil.copyfile(finalPathName, current_working_directory + '\\' + backupFileName)
    print("Creazione del file copia ", backupFileName, " avvenuta con successo.\n")

    fileToManage = '1'

    # Caricamento della list in cui inserire il nome di ogni SUBAGENTE e la corrispettiva AGENZIA di riferimento a partire dai dati presenti nel file Excel 'elenco_collaboratori_agenzia.xlsx'
    getAgencyFromSubagent()

    while fileToManage.isnumeric():
        # GENERALI
        if fileToManage == '1':
            filesGENERALI_toParse = []

            findFilesNotChecked(current_working_directory + partialDir_filesGENERALI + '\\', filesGENERALI_toParse)

            print(*filesGENERALI_toParse, sep='\n')

            newFilesFind = False

            if(filesGENERALI_toParse == []):
                print("Nessun file di GENERALI da analizzare.\n")
            else:
                newFilesFind = True

            startTime_Parsing = time.time()

            for i in range(0, len(filesGENERALI_toParse)):
                pathName_GENERALI = current_working_directory + partialDir_filesGENERALI + '\\' + filesGENERALI_toParse[i]
                readFromGenerali(filesGENERALI_toParse[i], pathName_GENERALI, finalPathName, totale_sospesi_nuovi)
                print("--------------------------------------------------------------------\n")

            endTime_Parsing = time.time()
            executionTime_Parsing = endTime_Parsing - startTime_Parsing

            if(newFilesFind == True):
                print("--------------------------------------------------------------------\n")
                print("Tempo di esecuzione di tutti i files GENERALI = ", int(executionTime_Parsing), " secondi.\n")

            print("--------------------------------------------------------------------\n")

            fileToManage = '2'

        # CATTOLICA
        elif fileToManage == '2':

            filesCATTOLICA_toParse = []

            findFilesNotChecked(current_working_directory + partialDir_filesCATTOLICA + '\\', filesCATTOLICA_toParse)

            print(*filesCATTOLICA_toParse, sep='\n')

            newFilesFind = False

            if(filesCATTOLICA_toParse == []):
                print("Nessun file di CATTOLICA da analizzare.\n")
            else:
                newFilesFind = True

            startTime_Parsing = time.time()

            for i in range(0, len(filesCATTOLICA_toParse)):
                pathName_CATTOLICA = current_working_directory + partialDir_filesCATTOLICA + '\\' + filesCATTOLICA_toParse[i]
                readFromCattolica(filesCATTOLICA_toParse[i], pathName_CATTOLICA, finalPathName, totale_sospesi_nuovi)
                print("--------------------------------------------------------------------\n")

            endTime_Parsing = time.time()
            executionTime_Parsing = endTime_Parsing - startTime_Parsing

            if(newFilesFind == True):
                print("--------------------------------------------------------------------\n")
                print("Tempo di esecuzione di tutti i files CATTOLICA = ", int(executionTime_Parsing), " secondi.\n")

            print("--------------------------------------------------------------------\n")

            fileToManage = '3'

        # TUTELA
        elif fileToManage == '3':

            filesTUTELA_toParse = []

            findFilesNotChecked(current_working_directory + partialDir_filesTUTELA + '\\', filesTUTELA_toParse)

            print(*filesTUTELA_toParse, sep='\n')

            newFilesFind = False

            if(filesTUTELA_toParse == []):
                print("Nessun file di TUTELA LEGALE da analizzare.\n")
            else:
                newFilesFind = True

            startTime_Parsing = time.time()

            # fileName_TUTELA = input("Inserire nome completo del file TUTELA con estensione: ")
            for i in range(0, len(filesTUTELA_toParse)):
                pathName_TUTELA = current_working_directory + partialDir_filesTUTELA + '\\' + filesTUTELA_toParse[i]
                readFromTutela(filesTUTELA_toParse[i], pathName_TUTELA, finalPathName, totale_sospesi_nuovi)
                print("--------------------------------------------------------------------\n")

            endTime_Parsing = time.time()
            executionTime_Parsing = endTime_Parsing - startTime_Parsing

            if(newFilesFind == True):
                print("--------------------------------------------------------------------\n")
                print("Tempo di esecuzione di tutti i files TUTELA LEGALE = ", int(executionTime_Parsing), " secondi.\n")

            print("--------------------------------------------------------------------\n")

            fileToManage = '4'

        elif fileToManage == '4':

            manageSospesi(finalPathName, lastMonthYearDatetime)

            fileToManage = 'end'
            
except Exception as e:
    print("\n\nError: ", e)
    input()

end_time = time.time()
execution_time = end_time - start_time
print("Tempo di esecuzione = ", int(execution_time), " secondi.\n")

print("\nEsecuzione completata.\n")

input()
