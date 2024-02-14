import numpy as np
import pandas as pd
import datetime
import re

# Funzione per leggere i dati dal file GENERALI e salvarli nel file finale 'fileToWrite'
def readFromGenerali(fileGenerali_read, fileToWrite):
    sheetNameGenerali = 'BONIFICI GENERALI '    # ATTENZIONE allo spazio finale nel sheet name

    year_month_day = re.findall('\\d{4}-\\d{2}-\\d{2}', fileGenerali_read)[0]

    year = year_month_day[0:4]
    month = year_month_day[5:7]
    day = year_month_day[8:10]

    findImporto = False
    # read by default 1st sheet of an excel file
    # ATTENZIONE: per come e' fatto attualmente il file di GENERALI vi sono molte righe prima della tabella con i dati, quindi posso non considerare il parametro 'header' nella read_excel.
    # Dovesse pero' cambiare il formato come quello di CATTOLICA bisognera' probabilmente modificare il funzionamento di fileImporto, in quanto se non viene
    # utilizzato il parametro 'header' vengono letti tutti i dati da dopo la riga di intestazione del file excel.
    dataframe1 = pd.read_excel(fileGenerali_read, usecols='A,B,H,J')

    print("\nLettura file GENERALI eseguita correttamente.\n")

    # A -> 0 : NUMERO POLIZZA
    # B -> 1 : ANAGRAFICA (CONTRAENTE)
    # H -> 2 : MODALITA' PAGAMENTO
    # J -> 3 : IMPORTO

    generali_nr_polizza = []
    generali_anagrafica = []
    generali_importo = []

    for i in range(0, len(dataframe1)):

        if(dataframe1.iat[i, 1] == 'CONTENITORE'):
            # Se trovo la stringa 'CONTENITORE' mi fermo perche' per ora sono andato troppo oltre, poi sara' da gestire diversamente
            findImporto = False

        if(findImporto and dataframe1.isnull().iat[i, 0] == False and dataframe1.isnull().iat[i, 1] == False and dataframe1.iat[i, 2] == 'BONIFICO' and dataframe1.isnull().iat[i, 3] == False):
            # NON devo salvare le righe di dati vuoti che si trovano all'interno della tabella con i dati da salvare
            # N.B. In questo caso non sto salvando nemmeno la riga con il Totale, tanto me lo ricreo dopo
            # Salvo solamente le righe che hanno 'BONIFICO' nella colonna H del file di partenza
            # Togliere importi negativi

            # In realta', essendo una stringa, mi basterebbe vedere se il primo carattere e' un '-' (importo negativo) oppure no
            # if(floatValue > 0):
            condition = False

            if(isinstance(dataframe1.iat[i, 3], str)):
                condition = (dataframe1.iat[i, 3][0] != '-')
            
            elif(isinstance(dataframe1.iat[i, 3], int)):
                condition = (dataframe1.iat[i, 3] > 0)

            elif(isinstance(dataframe1.iat[i, 3], float)):
                condition = (dataframe1.iat[i, 3] > 0.0)

            if(condition):
                generali_nr_polizza.append(dataframe1.iat[i, 0])
                generali_anagrafica.append(dataframe1.iat[i, 1])
                generali_importo.append(dataframe1.iat[i, 3])

        if(dataframe1.isnull().iat[i, 0] == False and dataframe1.iat[i, 1] == 'ANAGRAFICA' and findImporto == False):
            # Se trovo la stringa 'ANAGRAFICA' vuol dire che dal ciclo successivo inizio a salvare tutti i dati
            findImporto = True

        
    # total = 0

    # for i in range(0, len(arr_importo)):
        # print(float(arr_importo[i]))

        # total += float(arr_importo[i])

    # print('Total = ', total)

    final_Generali = list(zip(generali_importo, generali_nr_polizza, generali_anagrafica))

    # for i in range(0, len(final_struct)):
    #     print(final_struct[i])

    df_Generali = pd.DataFrame(final_Generali)

    # PATH = r'C:\Users\s.lazzaro\OneDrive - CUSTOM SPA\Desktop\File_Gigi\PRIMA NOTA DEL 2024 NUOVA GESTIONE  1.xls'
    # PATH = r'C:\Users\s.lazzaro\OneDrive - CUSTOM SPA\Desktop\File_Gigi\FILE_PROVA.xlsx'
    # wb = xw.Book(PATH)
    # sheet = wb.sheets['BONIFICI GENERALI ']

    # df3 = sheet['A1:C4'].options(pd.DataFrame, index=False, header=True).value

    # print(df3.to_string())

    # writer = pd.ExcelWriter(PATH, engine='openpyxl')
    # df3.to_excel(writer, sheet_name="BONIFICI GENERALI ", startrow=25)

    # writeInFile = r'C:\Users\s.lazzaro\OneDrive - CUSTOM SPA\Desktop\File_Gigi\PRIMA_NOTA_TEST_.xlsx'

    datareadGenerali = pd.read_excel(fileToWrite, sheet_name = sheetNameGenerali, usecols='A')

    rowData = 0

    print("\nData inserita: ", day, '-', month, '-', year)

    dateToCompare = datetime.datetime(int(year), int(month), int(day), 0, 0)

    for i in range(0, len(datareadGenerali)):
        if(datareadGenerali.values[i] == dateToCompare):
            # print(dataread.values[i])
            rowData = i+1
            break

    print("Copia e salvataggio dati in esecuzione, attendere ...\n")

    with pd.ExcelWriter(fileToWrite, engine ="openpyxl", mode='a', if_sheet_exists='overlay') as writer:
        df_Generali.to_excel(writer, index = False, header = False, sheet_name = sheetNameGenerali, startrow = rowData+1, startcol = 1)

    print("Copia dei dati di GENERALI terminata.\n")
    input("Premere INVIO per proseguire...\n")


#   -----------------------------------------------------------------------------------------
#   -----------------------------------------------------------------------------------------
#   -----------------------------------------------------------------------------------------
#   -----------------------------------------------------------------------------------------
# Funzione per leggere i dati dal file di CATTOLICA e salvarli nel file finale 'fileToWrite'
def readFromCattolica(pathName_read, fileToWrite):

    sheetNameCattolica = 'BONIFICI CATTOLICA'

    day_month_year = re.findall('\\d{2}_\\d{2}_\\d{4}', pathName_read)[0]

    day = day_month_year[0:2]
    month = day_month_year[3:5]
    year = day_month_year[6:10]
    
    # read sheet 'Incassi' of CATTOLICA excel file
    dataframe1 = pd.read_excel(pathName_read, sheet_name='Incassi', usecols='A,E,H,K')
    # Non avendo inserito il parametro 'header' nella read_excel, la 1^a riga di dataframe1 contiene gia' i dati
    print("\nLettura file CATTOLICA eseguita correttamente.")

    # A -> 0 : CONTRAENTE
    # E -> 1 : NUMERO POLIZZA
    # H -> 2 : IMPORTO PREMIO
    # K -> 3 : MODALITA' PAGAMENTO

    cattolica_contraente = []
    cattolica_nr_polizza = []
    cattolica_importo = []

    for i in range(0, len(dataframe1)):
        if(dataframe1.isnull().iat[i, 0] == False and dataframe1.isnull().iat[i, 1] == False and dataframe1.isnull().iat[i, 2] == False and dataframe1.iat[i, 3].find('Bonifico') != -1):
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
            condition = False

            if(isinstance(dataframe1.iat[i, 3], str)):
                condition = (dataframe1.iat[i, 3][0] != '-')
            
            elif(isinstance(dataframe1.iat[i, 3], int)):
                condition = (dataframe1.iat[i, 3] > 0)

            elif(isinstance(dataframe1.iat[i, 3], float)):
                condition = (dataframe1.iat[i, 3] > 0.0)

            if(condition):
                cattolica_contraente.append(dataframe1.iat[i, 0])
                cattolica_nr_polizza.append(dataframe1.iat[i, 1])
                cattolica_importo.append(dataframe1.iat[i, 2])


    final_Cattolica = list(zip(cattolica_importo, cattolica_nr_polizza, cattolica_contraente))

    df_Cattolica = pd.DataFrame(final_Cattolica)

    # Dal file finale vado a leggere tutte le date presenti nel relativo sheet nella colonna 'A'
    datareadCattolica = pd.read_excel(fileToWrite, sheet_name = sheetNameCattolica, usecols='A')

    # riga su file excel PRIMA_NOTA in cui andare a scrivere i vari dati
    rowData = 0

    print("\nData inserita: ", day, '-', month, '-', year)

    dateToCompare = datetime.datetime(int(year), int(month), int(day), 0, 0)

    for i in range(0, len(datareadCattolica)):
        if(datareadCattolica.values[i] == dateToCompare):
            # print(dataread.values[i])
            rowData = i+1
            break

    print("Copia e salvataggio dati in esecuzione, attendere ...\n")

    with pd.ExcelWriter(fileToWrite, engine ="openpyxl", mode='a', if_sheet_exists='overlay') as writer:
        df_Cattolica.to_excel(writer, index = False, header = False, sheet_name = sheetNameCattolica, startrow = rowData+1, startcol = 1)

    print("Copia dei dati di CATTOLICA terminata.\n")
    input("Premere INVIO per proseguire...\n")
