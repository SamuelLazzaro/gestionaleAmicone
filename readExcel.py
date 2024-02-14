# import pandas lib as pd
import numpy as np
import pandas as pd
import datetime
import os
import re

from companiesFunction import readFromCattolica

current_working_directory = os.getcwd()

print("Percorso attuale: ", current_working_directory)

# path = r'C:\LUIGI 04052016\AMICONE LUIGI\DATI DAL 31032008 PC PORTATILE\DATI\CONTABILITA\PARTITE REGISTRATE PER CONTABILITA\GENERALI\PARTITE REGISTRATE\FEBBRAIO 2024'

print("Percorso del file da cui copiare i dati: ", current_working_directory)
fileName = input("Inserire nome completo del file con estensione: ")

# fileName = r'\REPORT - PARTITE REGISTRATE  - 2024-02-07T184414.309.xls'

src = current_working_directory + r'\PARTITE REGISTRATE PER CONTABILITA\GENERALI\PARTITE REGISTRATE\FEBBRAIO 2024' + '\\' + fileName

print(src)

year_month_day = re.findall('\\d{4}-\\d{2}-\\d{2}', fileName)[0]

year = year_month_day[0:4]
month = year_month_day[5:7]
day = year_month_day[8:10]


findImporto = False
# read by default 1st sheet of an excel file
dataframe1 = pd.read_excel(src, usecols='A,B,H,J')

print('len = ', len(dataframe1))

arr_nr_rif = []
arr_anagrafica = []
arr_importo = []


for i in range(0, len(dataframe1)-1):

    if(dataframe1.iat[i, 1] == 'CONTENITORE'):
        # Se trovo la stringa 'CONTENITORE' mi fermo perche' per ora sono andato troppo oltre, poi sara' da gestire diversamente
        findImporto = False

    if(findImporto and dataframe1.isnull().iat[i, 0] == False and dataframe1.isnull().iat[i, 1] == False and dataframe1.iat[i, 2] == 'BONIFICO' and dataframe1.isnull().iat[i, 3] == False):
        # NON devo salvare le righe di dati vuoti che si trovano all'interno della tabella con i dati da salvare
        # N.B. In questo caso non sto salvando nemmeno la riga con il Totale, tanto me lo ricreo dopo
        # Salvo solamente le righe che hanno 'BONIFICO' nella colonna H del file di partenza
        # Togliere importi negativi
        # Cerco l'indice corrispondente alla ',' nella stringa con l'importo
        # commaIndex = dataframe1.iat[i, 3].find(',')
        # print(dataframe1.iat[i, 3][0:commaIndex], " type: ", type(dataframe1.iat[i, 3][0:commaIndex]))
        # Trasformo solo le cifre intere della stringa con l'importo in un float per poi verificare se tale valore e' > 0, dato che non mi interessa avere anche le cifre decimali per fare tale confronto
        # floatValue = float(dataframe1.iat[i, 3][0:commaIndex])

        # In realta', essendo una stringa, mi basterebbe vedere se il primo carattere e' un '-' (importo negativo) oppure no
        # if(floatValue > 0):
        if(dataframe1.iat[i, 3][0] != '-'):
            arr_nr_rif.append(dataframe1.iat[i, 0])
            arr_anagrafica.append(dataframe1.iat[i, 1])
            arr_importo.append(dataframe1.iat[i, 3])

    if(dataframe1.isnull().iat[i, 0] == False and dataframe1.iat[i, 1] == 'ANAGRAFICA'):
        # Se trovo la stringa 'ANAGRAFICA' vuol dire che dal ciclo successivo inizio a salvare tutti i dati
        findImporto = True

    # print(dataframe1.iloc[i].to_string())
    
total = 0

print('\n\n\n\n')

# for i in range(0, len(arr_importo)):
    # print(float(arr_importo[i]))

    # total += float(arr_importo[i])

# print('Total = ', total)

final_struct = list(zip(arr_importo, arr_nr_rif, arr_anagrafica))

# for i in range(0, len(final_struct)):
#     print(final_struct[i])

df2 = pd.DataFrame(final_struct)

# print(df2.to_string())

# with pd.ExcelWriter("output.xls", engine="openpyxl") as writer:
#     df2.to_excel(writer, index=False, sheet_name="Sheet0")



# PATH = r'C:\Users\s.lazzaro\OneDrive - CUSTOM SPA\Desktop\File_Gigi\PRIMA NOTA DEL 2024 NUOVA GESTIONE  1.xls'
# PATH = r'C:\Users\s.lazzaro\OneDrive - CUSTOM SPA\Desktop\File_Gigi\FILE_PROVA.xlsx'
# wb = xw.Book(PATH)
# sheet = wb.sheets['BONIFICI GENERALI ']

# df3 = sheet['A1:C4'].options(pd.DataFrame, index=False, header=True).value

# print(df3.to_string())

# writer = pd.ExcelWriter(PATH, engine='openpyxl')
# df3.to_excel(writer, sheet_name="BONIFICI GENERALI ", startrow=25)

# writeInFile = r'C:\Users\s.lazzaro\OneDrive - CUSTOM SPA\Desktop\File_Gigi\PRIMA_NOTA_TEST_.xlsx'

resultFileName = 'PRIMA_NOTA_TEST_.xlsx'
writeInFile = current_working_directory + '\\' + resultFileName

dataread = pd.read_excel(writeInFile, sheet_name='BONIFICI GENERALI ', usecols='A')

rowData = 0

print('\n')
print("Data inserita: ", day, '-', month, '-', year)
print('\n')

dateToCompare = datetime.datetime(int(year), int(month), int(day), 0, 0)

for i in range(0, len(dataread)):
    if(dataread.values[i] == dateToCompare):
        # print(dataread.values[i])
        rowData = i+1
        break

# print('rowData = ', rowData)

# writer = pd.ExcelWriter(writeInFile, engine='openpyxl')

with pd.ExcelWriter(writeInFile, engine ="openpyxl", mode='a', if_sheet_exists='overlay') as writer:
    df2.to_excel(writer, index=False, header=False, sheet_name="BONIFICI GENERALI ", startrow=rowData+1, startcol=1)


input("Press Enter to finish...")
