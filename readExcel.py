# import pandas lib as pd
import numpy as np
import pandas as pd
# import pandas.io.formats.style
import datetime
import os
import re

from classDefinition import TotaleSospesiNuovi
from companiesFunction import *

try:
    totale_sospesi_vecchi = 0
    totale_sospesi_nuovi = TotaleSospesiNuovi()

    # Gestione versamenti effettuati
    versamenti_SI_NO = input("\nSono stati eseguiti dei versamenti? [SI]/[NO]: ")
    versamenti_SI_NO = versamenti_SI_NO.upper()

    if(versamenti_SI_NO[0] == 'S'):
        agenzia_versamenti = input("\nPer quale agenzia e' stato eseguito il versamento? Digitare il numero corrispondente all'agenzia e premere INVIO:\n1. RHO\n2. SOMMA LOMBARDO\n3. LEGNANO\n4. GALLARATE\n5. n\n\nAgenzia numero: ")
        importo_versamenti = input("\nInserire l'importo del versamento: ")

        # Sostituisco un eventuale ',' con un '.' per non avere poi un errore con la funzione float()
        importo_versamenti = importo_versamenti.replace(',', '.')
        importo_versamenti = float(importo_versamenti)

    current_working_directory = os.getcwd()

    print("Percorso attuale: ", current_working_directory)

    month_folder = input("\nInserire nome cartella del mese + anno (tutto in maiuscolo): ")
    # path = r'C:\LUIGI 04052016\AMICONE LUIGI\DATI DAL 31032008 PC PORTATILE\DATI\CONTABILITA\PARTITE REGISTRATE PER CONTABILITA\GENERALI\PARTITE REGISTRATE\FEBBRAIO 2024'

    partialDir_filesGENERALI    = r'\PARTITE REGISTRATE PER CONTABILITA\GENERALI\2024' + '\\' + month_folder
    partialDir_filesCATTOLICA   = r'\PARTITE REGISTRATE PER CONTABILITA\CATTOLICA\2024' + '\\' + month_folder
    partialDir_filesTUTELA      = r'\PARTITE REGISTRATE PER CONTABILITA\TUTELA LEGALE\2024' + '\\' + month_folder

    finalFileName = 'PRIMA_NOTA_TEST_.xlsx'
    finalPathName = current_working_directory + '\\' + finalFileName

    # fileToManage = input("\nScegliere la compagnia di cui effettuare la copia dei dati.\n1. GENERALI\n2. CATTOLICA\n3. TUTELA\n\nPremere numero + INVIO: ")
    fileToManage = '1'

    # readSospesiFromExcel(finalPathName)

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

                readFromGenerali(filesGENERALI_toParse[i], pathName_GENERALI, finalPathName, totale_sospesi_nuovi)

            print("--------------------------------------------------------------------\n")

            fileToManage = '2'

        # CATTOLICA
        elif fileToManage == '2':
            print("\nTotale sospesi nuovi dopo GENERALI: ", totale_sospesi_nuovi)

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

                readFromCattolica(filesCATTOLICA_toParse[i], pathName_CATTOLICA, finalPathName, totale_sospesi_nuovi)

            print("--------------------------------------------------------------------\n")

            fileToManage = '3'

        # TUTELA
        elif fileToManage == '3':
            print("\nTotale sospesi nuovi dopo CATTOLICA: ", totale_sospesi_nuovi)

            filesTUTELA_toParse = []

            findFilesNotChecked(current_working_directory + partialDir_filesTUTELA + '\\', filesTUTELA_toParse)

            print(*filesTUTELA_toParse, sep='\n')

            if(filesTUTELA_toParse == []):
                print("\nI dati di tutti i files TUTELA sono stati copiati in PRIMA NOTA.\n")
                print("--------------------------------------------------------------------\n")

            # fileName_TUTELA = input("Inserire nome completo del file TUTELA con estensione: ")
            for i in range(0, len(filesTUTELA_toParse)):
                pathName_TUTELA = current_working_directory + partialDir_filesTUTELA + '\\' + filesTUTELA_toParse[i]


                print("\nPercorso completo del file: ", pathName_TUTELA)

                readFromTutela(filesTUTELA_toParse[i], pathName_TUTELA, finalPathName, totale_sospesi_nuovi)

            print("--------------------------------------------------------------------\n")

            print("\nTotale sospesi nuovi dopo TUTELA LEGALE: ", totale_sospesi_nuovi)
            fileToManage = '4'

        elif fileToManage == '4':
            readSospesiFromExcel(finalPathName)

            fileToManage = 'end'
            
        # fileToManage = input("\nPremere INVIO per uscire, oppure scegliere un'altra compagnia di cui effettuare la copia dei dati.\n1. GENERALI\n2. CATTOLICA\n3. TUTELA\n\nPremere numero + INVIO oppure solo INVIO per uscire: ")
except Exception as e:
    print("\n\nError: ", e)
    input()

input("\nEsecuzione completata.\n")
