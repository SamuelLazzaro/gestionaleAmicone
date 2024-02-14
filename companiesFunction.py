import numpy as np
import pandas as pd
import datetime
import re
import os

# Funzione per leggere i dati dal file GENERALI e salvarli nel file finale 'fileToWrite'
def readFromGenerali(fileName_Generali, fileGenerali_read, fileToWrite):
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

    print("Copia dei dati del file ", fileName_Generali, " di GENERALI terminata.\n")
    
    #  Rinomino il file di cui ho appena salvato i dati con la desinenza '_checked'
    renameFileChecked(fileGenerali_read)

    print("File ", fileName_Generali, " rinominato con '_checked' come desinenza.\n")

    # input("Premere INVIO per proseguire...\n")


#   -----------------------------------------------------------------------------------------
#   -----------------------------------------------------------------------------------------
#   -----------------------------------------------------------------------------------------
#   -----------------------------------------------------------------------------------------
# Funzione per leggere i dati dal file di CATTOLICA e salvarli nel file finale 'fileToWrite'
def readFromCattolica(fileName_Cattolica, pathName_read, fileToWrite):

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

    print("Copia dei dati del file ", fileName_Cattolica, " di CATTOLICA terminata.\n")

    #  Rinomino il file di cui ho appena salvato i dati con la desinenza '_checked'
    renameFileChecked(pathName_read)

    print("File ", fileName_Cattolica, " rinominato con '_checked' come desinenza.\n")

    # input("Premere INVIO per proseguire...\n")



#   -----------------------------------------------------------------------------------------
#   -----------------------------------------------------------------------------------------
#   -----------------------------------------------------------------------------------------
#   -----------------------------------------------------------------------------------------
# Funzione per leggere i dati dal file di TUTELA e salvarli nel file finale 'fileToWrite'
def readFromTutela(fileName_Tutela, fileTutela_read, fileToWrite):
    sheetNameTutela = 'BONIFICI TUTELA'

    findImporto = False
    # read by default 1st sheet of an excel file
    # ATTENZIONE: per come e' fatto attualmente il file di GENERALI vi sono molte righe prima della tabella con i dati, quindi posso non considerare il parametro 'header' nella read_excel.
    # Dovesse pero' cambiare il formato come quello di CATTOLICA bisognera' probabilmente modificare il funzionamento di fileImporto, in quanto se non viene
    # utilizzato il parametro 'header' vengono letti tutti i dati da dopo la riga di intestazione del file excel.
    dataframe1 = pd.read_excel(fileTutela_read, usecols='C,H,I,L,M')

    print("\nLettura file TUTELA eseguita correttamente.\n")

    # C -> 0 : NUMERO POLIZZA
    # H -> 1 : MODALITA' PAGAMENTO
    # I -> 2 : IMPORTO
    # L -> 3 : ANAGRAFICA (CONTRAENTE)
    # M -> 4 : DATA

    tutela_nr_polizza = []
    tutela_anagrafica = []
    tutela_importo = []
    tutela_data = []

    for i in range(0, len(dataframe1)):

        if(dataframe1.isnull().iat[i, 0] == True):
            # Se la colonna 'C' del file 'Fondocassa' e' vuota, vuol dire che non c'e' un dato da salvare
            findImporto = False

        if(findImporto and dataframe1.iat[i, 1] == 'BB' and dataframe1.isnull().iat[i, 2] == False):
            # NON devo salvare le righe di dati vuoti che si trovano all'interno della tabella con i dati da salvare
            # N.B. In questo caso non sto salvando nemmeno la riga con il Totale, tanto me lo ricreo dopo
            # Salvo solamente le righe che hanno 'BB' nella colonna H del file di partenza
            # Togliere importi negativi

            # Se l'importo e' una stringa allora controllo che il primo carattere sia diverso da '-', mentre se e' un int o un float deve essere rispettivamente > 0 oppure > 0.0
            condition = False

            if(isinstance(dataframe1.iat[i, 3], str)):
                condition = (dataframe1.iat[i, 3][0] != '-')
            
            elif(isinstance(dataframe1.iat[i, 3], int)):
                condition = (dataframe1.iat[i, 3] > 0)

            elif(isinstance(dataframe1.iat[i, 3], float)):
                condition = (dataframe1.iat[i, 3] > 0.0)

            if(condition):
                tutela_nr_polizza.append(dataframe1.iat[i, 0])
                tutela_anagrafica.append(dataframe1.iat[i, 3])
                tutela_importo.append(dataframe1.iat[i, 2])
                tutela_data.append(dataframe1.iat[i, 4])

        if(dataframe1.isnull().iat[i, 0] == False and findImporto == False):
            # Se trovo la stringa 'ANAGRAFICA' vuol dire che dal ciclo successivo inizio a salvare tutti i dati
            findImporto = True

    final_Tutela = list(zip(tutela_data, tutela_importo, tutela_nr_polizza, tutela_anagrafica))

    df_Tutela = pd.DataFrame(final_Tutela)

    datareadTutela = pd.read_excel(fileToWrite, sheet_name = sheetNameTutela, usecols='A')

    rowData = [[], []]

    # print(*final_Tutela, sep='\n')

    for i in range(0, len(datareadTutela)):
        # Step 1: ricostruire la data da confrontare poi con quella presente nella tabella di BONIFICI TUTELA in PRIMA NOTA
        # day_month_year = re.findall('\\d{2}_\\d{2}_\\d{4}', df_Tutela.iat[i, 0])[0]

        # day = day_month_year[0:2]
        # month = day_month_year[3:5]
        # year = day_month_year[6:10]

        # dateToCompare = datetime.datetime(year, month, day, 0, 0)

        # print(datareadTutela.iat[i, 0])

        if(isinstance(datareadTutela.iat[i, 0], datetime.datetime)):
            # print(dataread.values[i])
            # Se il dato appena letto dal foglio BONIFICI TUTELA in PRIMA NOTA nella colonna 'A' e' una data, vedo se corrisponde ad una delle date di cui ho dei dati da salvare
            for j in range(0, len(df_Tutela)):
                if(datareadTutela.iat[i, 0] == df_Tutela.iat[j, 0]):
                    rowData[0].append(i+1)
                    rowData[1].append(df_Tutela.iat[j, 0])
                    break

    print("Copia e salvataggio dati in esecuzione, attendere ...\n")

    # In rowData ho gli indici delle righe in ordine crescente di data, ma in df_Tutela i vari dati si trovano in ordine decrescente di data, per questo motivo faccio un reverse for loop in modo tale da partire a salvare i dati con data piu' recente (rowData[i] con i = len(rowData)) fino ad arrivare a quelli con data meno recente (rowData[i] con i = 0)

    # print(rowData)

    for i in range(len(rowData[0])-1, -1, -1):

        final_listTutela = []

        for k in range(len(df_Tutela)-1, -1, -1):
            if(rowData[1][i] == df_Tutela.iat[k, 0]):
                # print("\nBefore: ", df_Tutela.values[k, 1:4])
                final_listTutela.append(df_Tutela.values[k, 1:4])
                # print("\nAfter: ", *final_listTutela, sep='\n')

        final_dfTutela = pd.DataFrame(final_listTutela)

        with pd.ExcelWriter(fileToWrite, engine ="openpyxl", mode='a', if_sheet_exists='overlay') as writer:
            final_dfTutela.to_excel(writer, index = False, header = False, sheet_name = sheetNameTutela, startrow = rowData[0][i] + 1, startcol = 1)

    print("Copia dei dati del file ", fileName_Tutela, " di TUTELA terminata.\n")
    
    #  Rinomino il file di cui ho appena salvato i dati con la desinenza '_checked'
    renameFileChecked(fileTutela_read)

    print("File ", fileName_Tutela, " rinominato con '_checked' come desinenza.\n")
            



#   -----------------------------------------------------------------------------------------
#   -----------------------------------------------------------------------------------------
#   -----------------------------------------------------------------------------------------
#   -----------------------------------------------------------------------------------------
# Funzione per trovare tutti i file di cui non sono ancora stati salvati i dati
def findFilesNotChecked(pathName, filesToParse):
    dir_path = pathName

    for root, dirs, files in os.walk(dir_path):
        for file in files: 
    
            # change the extension from '.mp3' to 
            # the one of your choice.
            if file.endswith('checked.xls') == False:
                filesToParse.append(file)
            
    
    # print(*filesToParse, sep='\n')


#   -----------------------------------------------------------------------------------------
#   -----------------------------------------------------------------------------------------
#   -----------------------------------------------------------------------------------------
#   -----------------------------------------------------------------------------------------
# Funzione per rinominare il file di cui sono appena stati salvati i dati: viene aggiunta la desinenza '_checked'
def renameFileChecked(pathFileName):
    renameFile = pathFileName
    indexFileExtension = renameFile.find('.xls')
    renameFile = renameFile[0:indexFileExtension]
    renameFile = renameFile + '_checked.xls'

    os.rename(pathFileName, renameFile)