import numpy as np
import pandas as pd
import datetime
import re

def readFromCattolica(pathName_read, fileToWrite):

    sheetNameCattolica = 'BONIFICI CATTOLICA'

    day_month_year = re.findall('\\d{2}_\\d{2}_\\d{4}', pathName_read)[0]

    day = day_month_year[0:2]
    month = day_month_year[3:5]
    year = day_month_year[7:11]
    
    # read sheet 'Incassi' of CATTOLICA excel file
    dataframe1 = pd.read_excel(pathName_read, sheet_name='Incassi', usecols='A,E,H,K')
    print("\n\nLettura file CATTOLICA eseguita correttamente.\n\n")

    # A -> 0 : CONTRAENTE
    # E -> 1 : NUMERO POLIZZA
    # H -> 2 : IMPORTO PREMIO
    # K -> 3 : MODALITA' PAGAMENTO

    cattolica_contraente = []
    cattolica_nr_polizza = []
    cattolica_importo = []

    findData = False

    for i in range(0, len(dataframe1)-1):
        if(findImporto and dataframe1.isnull().iat[i, 0] == False and dataframe1.isnull().iat[i, 1] == False and dataframe1.isnull().iat[i, 2] == False and dataframe1.iat[i, 3].find('Bonifico') != -1):
            # NON devo salvare le righe di dati vuoti che si trovano all'interno della tabella con i dati da salvare
            # N.B. In questo caso non sto salvando nemmeno la riga con il Totale, tanto me lo ricreo dopo
            # Salvo solamente le righe che hanno la sottostringa 'Bonifico' nella colonna K del file di partenza
            # Togliere importi negativi
            # Cerco l'indice corrispondente alla ',' nella stringa con l'importo
            # commaIndex = dataframe1.iat[i, 3].find(',')
            # print(dataframe1.iat[i, 3][0:commaIndex], " type: ", type(dataframe1.iat[i, 3][0:commaIndex]))
            # Trasformo solo le cifre intere della stringa con l'importo in un float per poi verificare se tale valore e' > 0, dato che non mi interessa avere anche le cifre decimali per fare tale confronto
            # floatValue = float(dataframe1.iat[i, 3][0:commaIndex])

            # In realta', essendo una stringa, mi basterebbe vedere se il primo carattere e' un '-' (importo negativo) oppure no
            # if(floatValue > 0):
            if(dataframe1.iat[i, 3][0] != '-'):
                cattolica_contraente.append(dataframe1.iat[i, 0])
                cattolica_nr_polizza.append(dataframe1.iat[i, 1])
                cattolica_importo.append(dataframe1.iat[i, 2])

        if(dataframe1.isnull().iat[i, 0] == False):
            # Se trovo la cella della colonna 'CONTRAENTE' diversa da NaN allora dal ciclo successivo inizio a salvare i dati.
            # In realta', per come e' fatto il file di CATTOLICA, ho solamente l'header come 1^a riga e poi ho subito tutti i dati.
            findImporto = True


    final_Cattolica = list(zip(cattolica_importo, cattolica_nr_polizza, cattolica_contraente))

    final_df = pd.DataFrame(final_Cattolica)

    # Dal file finale vado a leggere tutte le date presenti nel relativo sheet nella colonna 'A'
    datareadCattolica = pd.read_excel(fileToWrite, sheet_name = sheetNameCattolica, usecols='A')

    # riga su file excel PRIMA_NOTA in cui andare a scrivere i vari dati
    rowData = 0

    print('\n')
    print("Data inserita: ", day, '-', month, '-', year)
    print('\n')

    dateToCompare = datetime.datetime(int(year), int(month), int(day), 0, 0)

    for i in range(0, len(datareadCattolica)):
        if(datareadCattolica.values[i] == dateToCompare):
            # print(dataread.values[i])
            rowData = i+1
            break

    # print('rowData = ', rowData)

    # writer = pd.ExcelWriter(writeInFile, engine='openpyxl')

    with pd.ExcelWriter(fileToWrite, engine ="openpyxl", mode='a', if_sheet_exists='overlay') as writer:
        final_df.to_excel(writer, index = False, header = False, sheet_name = sheetNameCattolica, startrow = rowData+1, startcol = 1)

    print("\n\nCopia dei dati di CATTOLICA terminata.\n\n")
    input("Premere un tasto per proseguire...")
