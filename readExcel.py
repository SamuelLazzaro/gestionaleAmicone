# import pandas lib as pd
import numpy as np
import pandas as pd
import datetime
import os
import re

from companiesFunction import readFromCattolica, readFromGenerali

current_working_directory = os.getcwd()

print("Percorso attuale: ", current_working_directory)

# path = r'C:\LUIGI 04052016\AMICONE LUIGI\DATI DAL 31032008 PC PORTATILE\DATI\CONTABILITA\PARTITE REGISTRATE PER CONTABILITA\GENERALI\PARTITE REGISTRATE\FEBBRAIO 2024'

finalFileName = 'PRIMA_NOTA_TEST_.xlsx'
finalPathName = current_working_directory + '\\' + finalFileName

fileToManage = input("\nScegliere compagnia di cui effettuare la copia dei dati.\nPremere numero + INVIO:\n1. GENERALI\n2. CATTOLICA\n\n")

while fileToManage.isnumeric():
    # GENERALI
    if fileToManage == '1':
        fileName_GENERALI = input("Inserire nome completo del file GENERALI con estensione: ")
        pathName_GENERALI = current_working_directory + r'\PARTITE REGISTRATE PER CONTABILITA\GENERALI\PARTITE REGISTRATE\FEBBRAIO 2024' + '\\' + fileName_GENERALI

        print("\nPercorso completo del file: ", pathName_GENERALI)

        readFromGenerali(pathName_GENERALI, finalPathName)

    # CATTOLICA
    elif fileToManage == '2':
        fileName_CATTOLICA = input("Inserire nome completo del file CATTOLICA con estensione: ")
        pathName_CATTOLICA = current_working_directory + r'\PARTITE REGISTRATE PER CONTABILITA\CATTOLICA\PARTITE REGISTRATE\FEBBRAIO 2024' + '\\' + fileName_CATTOLICA

        print("\nPercorso completo del file: ", pathName_CATTOLICA)

        readFromCattolica(pathName_CATTOLICA, finalPathName)

    fileToManage = input("\nPremere INVIO per uscire, oppure scegliere un'altra compagnia di cui effettuare la copia dei dati.\nPremere numero + INVIO:\n1. GENERALI\n2. CATTOLICA\n\n")


input("\nEsecuzione completata.\n")
