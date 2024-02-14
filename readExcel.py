# import pandas lib as pd
import numpy as np
import pandas as pd
import datetime
import os
import re

from companiesFunction import readFromCattolica, readFromGenerali, findFilesNotChecked

current_working_directory = os.getcwd()

partialDir_filesGENERALI    = r'\PARTITE REGISTRATE PER CONTABILITA\GENERALI\PARTITE REGISTRATE\FEBBRAIO 2024'
partialDir_filesCATTOLICA   = r'\PARTITE REGISTRATE PER CONTABILITA\CATTOLICA\PARTITE REGISTRATE\FEBBRAIO 2024'

print("Percorso attuale: ", current_working_directory)

# path = r'C:\LUIGI 04052016\AMICONE LUIGI\DATI DAL 31032008 PC PORTATILE\DATI\CONTABILITA\PARTITE REGISTRATE PER CONTABILITA\GENERALI\PARTITE REGISTRATE\FEBBRAIO 2024'

finalFileName = 'PRIMA_NOTA_TEST_.xlsx'
finalPathName = current_working_directory + '\\' + finalFileName

fileToManage = input("\nScegliere la compagnia di cui effettuare la copia dei dati.\n1. GENERALI\n2. CATTOLICA\n\nPremere numero + INVIO: ")

try:
    while fileToManage.isnumeric():
        # GENERALI
        if fileToManage == '1':
            filesGENERALI_toParse = []

            findFilesNotChecked(current_working_directory + partialDir_filesGENERALI + '\\', filesGENERALI_toParse)

            print(*filesGENERALI_toParse, sep='\n')

            if(filesGENERALI_toParse == []):
                print("\nI dati di tutti i files GENERALI sono stati copiati in PRIMA NOTA.\n")
                print("--------------------------------------------------------------------\n")

            # fileName_GENERALI = input("Inserire nome completo del file GENERALI con estensione: ")
            for i in range(0, len(filesGENERALI_toParse)):
                pathName_GENERALI = current_working_directory + partialDir_filesGENERALI + '\\' + filesGENERALI_toParse[i]

                print("\nPercorso completo del file: ", pathName_GENERALI)

                readFromGenerali(filesGENERALI_toParse[i], pathName_GENERALI, finalPathName)

            print("--------------------------------------------------------------------\n")

        # CATTOLICA
        elif fileToManage == '2':
            filesCATTOLICA_toParse = []

            findFilesNotChecked(current_working_directory + partialDir_filesCATTOLICA + '\\', filesCATTOLICA_toParse)

            print(*filesCATTOLICA_toParse, sep='\n')

            if(filesCATTOLICA_toParse == []):
                print("\nI dati di tutti i files CATTOLICA sono stati copiati in PRIMA NOTA.\n")
                print("--------------------------------------------------------------------\n")

            # fileName_CATTOLICA = input("Inserire nome completo del file CATTOLICA con estensione: ")
            for i in range(0, len(filesCATTOLICA_toParse)):
                pathName_CATTOLICA = current_working_directory + partialDir_filesCATTOLICA + '\\' + filesCATTOLICA_toParse[i]

                print("\nPercorso completo del file: ", pathName_CATTOLICA)

                readFromCattolica(filesCATTOLICA_toParse[i], pathName_CATTOLICA, finalPathName)

            print("--------------------------------------------------------------------\n")
            
        fileToManage = input("\nPremere INVIO per uscire, oppure scegliere un'altra compagnia di cui effettuare la copia dei dati.\n1. GENERALI\n2. CATTOLICA\n\nPremere numero + INVIO oppure solo INVIO per uscire: ")
except Exception as e:
    print("\n\nError: ", e)
    input()

input("\nEsecuzione completata.\n")
