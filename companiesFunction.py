import numpy as np
import pandas as pd
# import pandas.io.formats.style
import datetime
import re
import os
import copy
import time

from classDefinition import TotaleSospesiNew_Date
from auxiliaryFunction import *

sheetNameSospesi = 'SOSPESI'

DATA_SOSPESI    = int(0)
IMPORTO_SOSPESI = int(1)

ROW_INDEX_TOT_GENERALI = int(2)
ROW_INDEX_TOT_CATTOLICA = int(4)
ROW_INDEX_TOT_TUTELA = int(6)
ROW_INDEX_TOT_UCA = int(8)

dateFormat = "%d/%m/%Y"

def highlight_if_FinConsumo(val):
    """
    Takes a scalar and returns a string with
    the css property 'background-color: yellow' for
    values greater than 80, black otherwise.
    """
    color = 'red' if val.find('Fin. Consumo') else 'white'
    return f'background-color: {color}'


# Funzione per leggere i dati dal file GENERALI e salvarli nel file finale 'fileToWrite'
def readFromGenerali(fileName_Generali, fileGenerali_read, fileToWrite, totale_sospesi_nuovi):
    sheetNameGenerali = 'BONIFICI GENERALI '    # ATTENZIONE allo spazio finale nel sheet name

    # year_month_day = re.findall('\\d{4}-\\d{2}-\\d{2}', fileGenerali_read)[0]

    # year = year_month_day[0:4]
    # month = year_month_day[5:7]
    # day = year_month_day[8:10]

    findImporto = False
    # read by default 1st sheet of an excel file
    # ATTENZIONE: per come e' fatto attualmente il file di GENERALI vi sono molte righe prima della tabella con i dati, quindi posso non considerare il parametro 'header' nella read_excel.
    # Dovesse pero' cambiare il formato come quello di CATTOLICA bisognera' probabilmente modificare il funzionamento di fileImporto, in quanto se non viene
    # utilizzato il parametro 'header' vengono letti tutti i dati da dopo la riga di intestazione del file excel.
    dataframe1 = pd.read_excel(fileGenerali_read, usecols='A,B,C,H,J,K,P')

    print("\nLettura file GENERALI eseguita correttamente.\n")

    # A -> 0 : NUMERO POLIZZA
    # B -> 1 : ANAGRAFICA (CONTRAENTE)
    # C -> 2 : DATA DI REGISTRAZIONE
    # H -> 3 : MODALITA' PAGAMENTO
    # J -> 4 : IMPORTO
    # K -> 5 : COLLABORATORE
    # P -> 6 : PROVVIGIONI
    
    NUM_POLIZZA = int(0)
    ANAGRAFICA = int(1)
    DATA_REGISTRAZIONE = int(2)
    MOD_PAGAMENTO = int(3)
    IMPORTO = int(4)
    COLLABORATORE = int(5)
    PROVVIGIONI = int(6)

    generali_nr_polizza = []
    generali_anagrafica = []
    generali_importo = []

    sospesi_generali_data = []
    sospesi_generali_nr_polizza = []
    sospesi_generali_anagrafica = []
    sospesi_generali_importo = []
    sospesi_generali_agenzia = []
    sospesi_generali_compagnia = []
    sospesi_generali_collaboratore = []
    sospesi_generali_metodo_pagamento = []
    sospesi_generali_pagato = []

    totale_incassi = 0
    totale_provvigioni = 0

    # Data di registrazione
    dateToCompare = dataframe1.iat[0, DATA_REGISTRAZIONE]

    if(isinstance(dateToCompare, datetime.datetime) == False):
        dateToCompare = datetime.datetime.strptime(dateToCompare, dateFormat)

    for i in range(0, len(dataframe1)):

        if(dataframe1.iat[i, 1] == 'CONTENITORE'):
            # Se trovo la stringa 'CONTENITORE' mi fermo perche' per ora sono andato troppo oltre, poi sara' da gestire diversamente
            findImporto = False

        if(findImporto and dataframe1.isnull().iat[i, NUM_POLIZZA] == False and dataframe1.isnull().iat[i, ANAGRAFICA] == False and dataframe1.isnull().iat[i, IMPORTO] == False):
            # NON devo salvare le righe di dati vuoti che si trovano all'interno della tabella con i dati da salvare
            # N.B. In questo caso non sto salvando nemmeno la riga con il Totale, tanto me lo ricreo dopo
            # Salvo solamente le righe che hanno 'BONIFICO' nella colonna H del file di partenza
            # Togliere importi negativi

            # In realta', essendo una stringa, mi basterebbe vedere se il primo carattere e' un '-' (importo negativo) oppure no
            # if(floatValue > 0):
            condition = False

            if(isinstance(dataframe1.iat[i, IMPORTO], str)):
                condition = (dataframe1.iat[i, IMPORTO][0] != '-')
            
            else:
                condition = (dataframe1.iat[i, IMPORTO] > 0) or (dataframe1.iat[i, IMPORTO] > 0.0)

            if(condition):

                # Modifico l'importo del singolo versamento in un float
                importo_versamento = convertToFloat(dataframe1.iat[i, IMPORTO])

                # Modifico l'importo della singola provvigione in un float
                importo_provvigione = convertToFloat(dataframe1.iat[i, PROVVIGIONI])

                # Per il totale degli INCASSI e delle PROVVIGIONI considero tutti i metodi di pagamento tranne che 'MOBILE POS'
                if(dataframe1.iat[i, MOD_PAGAMENTO] != 'MOBILE POS'):
                    totale_incassi += importo_versamento
                    totale_provvigioni += importo_provvigione

                if (dataframe1.iat[i, MOD_PAGAMENTO] == 'BONIFICO'):
                    # BONIFICO con FINANZIAMENTO A CONSUMO GENERALI -> da inserire nei SOSPESI
                    if(dataframe1.iat[i, NUM_POLIZZA].find('Fin. Consumo') != -1):
                        dateString = dateToCompare.strftime(dateFormat)
                        sospesi_generali_data.append(dateString)
                        sospesi_generali_nr_polizza.append(dataframe1.iat[i, NUM_POLIZZA])
                        sospesi_generali_anagrafica.append(dataframe1.iat[i, ANAGRAFICA])
                        sospesi_generali_metodo_pagamento.append(dataframe1.iat[i, MOD_PAGAMENTO])
                        sospesi_generali_importo.append(importo_versamento)
                        agenzia = "AGOS"
                        sospesi_generali_agenzia.append(agenzia)
                        sospesi_generali_compagnia.append('GENERALI')
                        sospesi_generali_collaboratore.append(dataframe1.iat[i, COLLABORATORE])
                        sospesi_generali_pagato.append('No')
                        
                        updateAgencyTotaleSospesi(totale_sospesi_nuovi, importo_versamento, agenzia)
                        continue        # Vado all'iterazione successiva del while loop

                    # BONIFICO GENERALI
                    generali_nr_polizza.append(dataframe1.iat[i, NUM_POLIZZA])
                    generali_anagrafica.append(dataframe1.iat[i, ANAGRAFICA])
                    generali_importo.append(importo_versamento)

                # CONTANTI o ASSEGNO BANCARIO/POSTALE GENERALI
                elif (dataframe1.iat[i, MOD_PAGAMENTO] == 'CONTANTI' or dataframe1.iat[i, MOD_PAGAMENTO].find('ASSEGNO') != -1):
                        dateString = dateToCompare.strftime(dateFormat)
                        sospesi_generali_data.append(dateString)
                        sospesi_generali_nr_polizza.append(dataframe1.iat[i, NUM_POLIZZA])
                        sospesi_generali_anagrafica.append(dataframe1.iat[i, ANAGRAFICA])
                        sospesi_generali_metodo_pagamento.append(dataframe1.iat[i, MOD_PAGAMENTO])
                        sospesi_generali_importo.append(importo_versamento)
                        agenzia = findAgencyFromSubagent(dataframe1.iat[i, COLLABORATORE])
                        sospesi_generali_agenzia.append(agenzia)
                        sospesi_generali_compagnia.append('GENERALI')
                        sospesi_generali_collaboratore.append(dataframe1.iat[i, COLLABORATORE])
                        sospesi_generali_pagato.append('No')
                        
                        updateAgencyTotaleSospesi(totale_sospesi_nuovi, importo_versamento, agenzia)
                

        if(dataframe1.isnull().iat[i, NUM_POLIZZA] == False and dataframe1.iat[i, ANAGRAFICA] == 'ANAGRAFICA' and findImporto == False):
            # Se trovo la stringa 'ANAGRAFICA' vuol dire che dal ciclo successivo inizio a salvare tutti i dati
            findImporto = True

    # Creazione dataframe BONIFICI
    final_BonificiGenerali = list(zip(generali_importo, generali_nr_polizza, generali_anagrafica))
    df_BonificiGenerali = pd.DataFrame(final_BonificiGenerali)

    # Creazione dataframe SOSPESI
    final_SospesiGenerali = list(zip(sospesi_generali_data, sospesi_generali_importo, sospesi_generali_nr_polizza, sospesi_generali_anagrafica, sospesi_generali_agenzia, sospesi_generali_compagnia, sospesi_generali_collaboratore, sospesi_generali_metodo_pagamento, sospesi_generali_pagato))
    df_SospesiGenerali = pd.DataFrame(final_SospesiGenerali)

    # Creazione dataframe INCASSI e PROVVIGIONI GENERALI
    listIncassiProvvigioniGenerali = [["Incassi GENERALI", totale_incassi]]
    listIncassiProvvigioniGenerali.append(["Provvigioni GENERALI", totale_provvigioni])

    df_IncassiProvvigioniGenerali = pd.DataFrame(listIncassiProvvigioniGenerali)

    # Lettura dati presenti nel file excel per i fogli sheetNameGenerali, sheetNameSospesi e sheetNamePrimaNota
    datareadBonifici = pd.read_excel(fileToWrite, sheet_name = sheetNameGenerali, usecols='A')
    datareadSospesi = pd.read_excel(fileToWrite, sheet_name = sheetNameSospesi, usecols='A,B')
    datareadPrimaNota = pd.read_excel(fileToWrite, sheet_name = sheetNamePrimaNota, usecols = 'A')

    BonificiRowData = 0
    SospesiRowData = 0

    # In questo caso e' una datetime.datetime
    newDateToCompare = dateToCompare

    # Ricerca della riga nel foglio 'BONIFICI GENERALI ' in cui andare a salvare i dati corrispondenti alla data newDateToCompare
    for i in range(0, len(datareadBonifici)):
        if(datareadBonifici.values[i] == newDateToCompare):
            # print(dataread.values[i])
            BonificiRowData = i+1
            break

    dateFound = False

    # Ricerca della riga nel foglio 'SOSPESI' in cui andare a salvare i dati corrispondenti alla data newDateToCompare
    for i in range(0, len(datareadSospesi)):
        if(dateFound == True and datareadSospesi.isnull().iat[i, IMPORTO_SOSPESI] == True):
            # Se in precedenza avevo trovato la data corrispondente e la colonna IMPORTO e' vuota, allora devo iniziare a salvare dalla riga i-esima i valori
            # In questo modo non sovrascrivo i dati precedentemente salvati
            SospesiRowData = i
            break
        if(datareadSospesi.iat[i, DATA_SOSPESI] == newDateToCompare):
            dateFound = True

    # Ricerca della riga nel foglio 'PRIMA NOTA' in cui andare a salvare i dati corrispondenti alla data newDateToCompare
    PrimaNotaRowData = findPrimaNotaRow_forIncassiProvvigioni(datareadPrimaNota, newDateToCompare)

    print("Copia e salvataggio dati in esecuzione, attendere ...\n")

    with pd.ExcelWriter(fileToWrite, engine ="openpyxl", mode='a', if_sheet_exists='overlay') as writer:
        df_BonificiGenerali.to_excel(writer, index = False, header = False, sheet_name = sheetNameGenerali, startrow = BonificiRowData+1, startcol = 1)
        df_SospesiGenerali.to_excel(writer, index = False, header = False, sheet_name = sheetNameSospesi, startrow = SospesiRowData+1, startcol = 0)
        df_IncassiProvvigioniGenerali.to_excel(writer, index = False, header = False, sheet_name = sheetNamePrimaNota, startrow = PrimaNotaRowData+ROW_INDEX_TOT_GENERALI, startcol = 2)        # 2 -> 'C' : DESCRIZIONE

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
def readFromCattolica(fileName_Cattolica, pathName_read, fileToWrite, totale_sospesi_nuovi):

    sheetNameCattolica = 'BONIFICI CATTOLICA'

    # day_month_year = re.findall('\\d{2}_\\d{2}_\\d{4}', pathName_read)[0]

    # day = day_month_year[0:2]
    # month = day_month_year[3:5]
    # year = day_month_year[6:10]
    
    # read sheet 'Incassi' of CATTOLICA excel file
    dataframe1 = pd.read_excel(pathName_read, sheet_name='Incassi', usecols='A,E,H,I,K,Y,Z')
    # Non avendo inserito il parametro 'header' nella read_excel, la 1^a riga di dataframe1 contiene gia' i dati
    print("\nLettura file CATTOLICA eseguita correttamente.")

    # A -> 0 : CONTRAENTE
    # E -> 1 : NUMERO POLIZZA
    # H -> 2 : IMPORTO PREMIO
    # I -> 3 : PROVVIGIONI
    # K -> 4 : MODALITA' PAGAMENTO
    # Y -> 5 : DATA FOGLIO CASSA
    # Z -> 6 : COLLABORATORE

    CONTRAENTE      = int(0)
    NUM_POLIZZA     = int(1)
    IMPORTO         = int(2)
    PROVVIGIONI     = int(3)
    MOD_PAGAMENTO   = int(4)
    DATA_FOGLIO_CASSA = int(5)
    COLLABORATORE   = int(6)

    cattolica_contraente = []
    cattolica_nr_polizza = []
    cattolica_importo = []

    sospesi_cattolica_data = []
    sospesi_cattolica_nr_polizza = []
    sospesi_cattolica_anagrafica = []
    sospesi_cattolica_importo = []
    sospesi_cattolica_agenzia = []
    sospesi_cattolica_compagnia = []
    sospesi_cattolica_collaboratore = []
    sospesi_cattolica_metodo_pagamento = []
    sospesi_cattolica_pagato = []

    totale_incassi = 0
    totale_provvigioni = 0

    dateToCompare = dataframe1.iat[0, DATA_FOGLIO_CASSA]

    if(isinstance(dateToCompare, datetime.datetime) == False):
        dateToCompare = datetime.datetime.strptime(dateToCompare, dateFormat)
    

    for i in range(0, len(dataframe1)):
        if(dataframe1.isnull().iat[i, CONTRAENTE] == False and dataframe1.isnull().iat[i, NUM_POLIZZA] == False and dataframe1.isnull().iat[i, IMPORTO] == False):
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

            if(isinstance(dataframe1.iat[i, IMPORTO], str)):
                condition = (dataframe1.iat[i, IMPORTO][0] != '-')
            
            else:
                condition = (dataframe1.iat[i, IMPORTO] > 0) or (dataframe1.iat[i, IMPORTO] > 0.0)

            if(condition):
                # Variabile booleana utilizzata per sapere se aggiornare i valori di PROVVIGIONI e INCASSI nel caso in cui il metodo di pagamento sia diverso da 'MOBILE POS' che per CATTOLICA non so come si chiama
                updateProvvigioniIncassi = False

                # Modifico l'importo del singolo versamento in un float
                importo_versamento = convertToFloat(dataframe1.iat[i, IMPORTO])
                # Modifico l'importo della singola provvigione in un float
                importo_provvigione = convertToFloat(dataframe1.iat[i, PROVVIGIONI])

                # BONIFICI CATTOLICA
                if(dataframe1.iat[i, MOD_PAGAMENTO] == 'Bonifico su CC di Agenzia'):
                    updateProvvigioniIncassi = True
                    cattolica_contraente.append(dataframe1.iat[i, 0])
                    cattolica_nr_polizza.append(dataframe1.iat[i, NUM_POLIZZA])
                    cattolica_importo.append(importo_versamento)

                # SOSPESI CATTOLICA: ASSEGNO BANCARIO, CONTANTI, FINANZIAMENTO AL CONSUMO
                elif(dataframe1.iat[i, MOD_PAGAMENTO].find('Assegno') != -1 or dataframe1.iat[i, MOD_PAGAMENTO] == 'Contante' or dataframe1.iat[i, MOD_PAGAMENTO] == 'Bonifico su CC di Direzione'):
                    updateProvvigioniIncassi = True
                    dateString = dateToCompare.strftime(dateFormat)
                    sospesi_cattolica_data.append(dateString)
                    sospesi_cattolica_nr_polizza.append(dataframe1.iat[i, NUM_POLIZZA])
                    sospesi_cattolica_anagrafica.append(dataframe1.iat[i, CONTRAENTE])
                    sospesi_cattolica_metodo_pagamento.append(dataframe1.iat[i, MOD_PAGAMENTO])
                    sospesi_cattolica_importo.append(importo_versamento)
                    if(dataframe1.iat[i, MOD_PAGAMENTO] == 'Bonifico su CC di Direzione'):
                        agenzia = "AGOS"
                        sospesi_cattolica_agenzia.append(agenzia)
                        updateAgencyTotaleSospesi(totale_sospesi_nuovi, importo_versamento, agenzia)
                    else:
                        agenzia = findAgencyFromSubagent(dataframe1.iat[i, COLLABORATORE])
                        sospesi_cattolica_agenzia.append(agenzia)
                        updateAgencyTotaleSospesi(totale_sospesi_nuovi, importo_versamento, agenzia)

                    sospesi_cattolica_compagnia.append('CATTOLICA')
                    sospesi_cattolica_collaboratore.append(dataframe1.iat[i, COLLABORATORE])
                    sospesi_cattolica_pagato.append('No')

                # Per il totale degli INCASSI e delle PROVVIGIONI considero tutti i metodi di pagamento tranne che 'MOBILE POS' che non so come si chiami su CATTOLICA
                if(updateProvvigioniIncassi == True):
                    totale_incassi += importo_versamento
                    totale_provvigioni += importo_provvigione

                        


    # Creazione dataframe BONIFICI
    final_BonificiCattolica = list(zip(cattolica_importo, cattolica_nr_polizza, cattolica_contraente))
    df_BonificiCattolica = pd.DataFrame(final_BonificiCattolica)

    # Creazione dataframe SOSPESI
    final_SospesiCattolica = list(zip(sospesi_cattolica_data, sospesi_cattolica_importo, sospesi_cattolica_nr_polizza, sospesi_cattolica_anagrafica, sospesi_cattolica_agenzia, sospesi_cattolica_compagnia, sospesi_cattolica_collaboratore, sospesi_cattolica_metodo_pagamento, sospesi_cattolica_pagato))
    df_SospesiCattolica = pd.DataFrame(final_SospesiCattolica)

    # Creazione dataframe INCASSI e PROVVIGIONI GENERALI
    listIncassiProvvigioniCattolica = [["Incassi CATTOLICA", totale_incassi]]
    listIncassiProvvigioniCattolica.append(["Provvigioni CATTOLICA", totale_provvigioni])

    df_IncassiProvvigioniCattolica = pd.DataFrame(listIncassiProvvigioniCattolica)

    # Dal file finale vado a leggere tutte le date presenti nel relativo sheet nella colonna 'A'
    datareadBonifici = pd.read_excel(fileToWrite, sheet_name = sheetNameCattolica, usecols='A')
    datareadSospesi = pd.read_excel(fileToWrite, sheet_name = sheetNameSospesi, usecols='A,B')
    datareadPrimaNota = pd.read_excel(fileToWrite, sheet_name = sheetNamePrimaNota, usecols='A')

    # Converto tutte le date del foglio 'SOSPESI' da formato stringa a formato 'datetime'
    convertStringToDatetime(datareadSospesi, DATA_SOSPESI)

    BonificiRowData = 0
    SospesiRowData = 0

    # In questo caso e' una datetime.datetime
    newDateToCompare = dateToCompare

    # Ricerca numero riga in cui andare a salvare i nuovi record in BONIFICI CATTOLICA
    for i in range(0, len(datareadBonifici)):
        if(datareadBonifici.values[i] == newDateToCompare):
            # print(dataread.values[i])
            BonificiRowData = i+1
            break

    # Ricerca numero riga su cui andare a salvare i nuovi record in SOSPESI senza sovrascrivere quelli precedenti relativi alla stessa data
    dateFound = False

    for i in range(0, len(datareadSospesi)):
        if(dateFound == True and datareadSospesi.isnull().iat[i, IMPORTO_SOSPESI] == True):
            # Se in precedenza avevo trovato la data corrispondente e la colonna IMPORTO e' vuota, allora devo iniziare a salvare dalla riga i-esima i valori
            # In questo modo non sovrascrivo i dati precedentemente salvati
            SospesiRowData = i
            break
        if(datareadSospesi.iat[i, DATA_SOSPESI] == newDateToCompare):
            dateFound = True

    # Ricerca della riga nel foglio 'PRIMA NOTA' in cui andare a salvare i dati corrispondenti alla data newDateToCompare
    PrimaNotaRowData = findPrimaNotaRow_forIncassiProvvigioni(datareadPrimaNota, newDateToCompare)
    
    print("Copia e salvataggio dati in esecuzione, attendere ...\n")

    with pd.ExcelWriter(fileToWrite, engine ="openpyxl", mode='a', if_sheet_exists='overlay') as writer:
        df_BonificiCattolica.to_excel(writer, index = False, header = False, sheet_name = sheetNameCattolica, startrow = BonificiRowData+1, startcol = 1)
        df_SospesiCattolica.to_excel(writer, index = False, header = False, sheet_name = sheetNameSospesi, startrow = SospesiRowData+1, startcol = 0)
        df_IncassiProvvigioniCattolica.to_excel(writer, index = False, header = False, sheet_name = sheetNamePrimaNota, startrow = PrimaNotaRowData+ROW_INDEX_TOT_CATTOLICA, startcol = 2)      # 2 -> 'C' : DESCRIZIONE
    
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
def readFromTutela(fileName_Tutela, fileTutela_read, fileToWrite, totale_sospesi_nuovi):
    sheetNameTutela = 'BONIFICI TUTELA'

    findImporto = False
    # read by default 1st sheet of an excel file
    # ATTENZIONE: per come e' fatto attualmente il file di GENERALI vi sono molte righe prima della tabella con i dati, quindi posso non considerare il parametro 'header' nella read_excel.
    # Dovesse pero' cambiare il formato come quello di CATTOLICA bisognera' probabilmente modificare il funzionamento di fileImporto, in quanto se non viene
    # utilizzato il parametro 'header' vengono letti tutti i dati da dopo la riga di intestazione del file excel.
    dataframe1 = pd.read_excel(fileTutela_read, usecols='C,H,I,J,L,M')

    print("\nLettura file TUTELA eseguita correttamente.\n")

    # C -> 0 : NUMERO POLIZZA
    # H -> 1 : MODALITA' PAGAMENTO
    # I -> 2 : IMPORTO
    # J -> 3 : PROVVIGIONI
    # L -> 4 : ANAGRAFICA (CONTRAENTE)
    # M -> 5 : DATA

    NUM_POLIZZA     = int(0)
    MOD_PAGAMENTO   = int(1)
    IMPORTO         = int(2)
    PROVVIGIONI     = int(3)
    ANAGRAFICA      = int(4)
    DATA            = int(5)

    tutela_nr_polizza = []
    tutela_anagrafica = []
    tutela_importo = []
    tutela_data = []

    sospesi_tutela_nr_polizza = []
    sospesi_tutela_anagrafica = []
    sospesi_tutela_importo = []
    sospesi_tutela_agenzia = []
    sospesi_tutela_compagnia = []
    sospesi_tutela_collaboratore = []
    sospesi_tutela_metodo_pagamento = []
    sospesi_tutela_pagato = []
    sospesi_tutela_data = []

    totale_IncassiProvvigioni = [[datetime.datetime(2000, 1, 1, 0, 0), 0.0, 0.0]]

    TOT_DATA = int(0)
    TOT_INCASSI = int(1)
    TOT_PROVVIGIONI = int(2)

    numberOfDifferentDates = 0

    # Converto tutte le date presenti nei dati caricati dal file 'fogliocassa' di TUTELA LEGALE in formato 'datetime'
    convertStringToDatetime(dataframe1, DATA)

    for i in range(0, len(dataframe1)):

        if(dataframe1.isnull().iat[i, NUM_POLIZZA] == True):
            # Se la colonna 'C' del file 'Fondocassa' e' vuota, vuol dire che non c'e' un dato da salvare
            findImporto = False

        if(findImporto and dataframe1.isnull().iat[i, IMPORTO] == False):
            # NON devo salvare le righe di dati vuoti che si trovano all'interno della tabella con i dati da salvare
            # N.B. In questo caso non sto salvando nemmeno la riga con il Totale, tanto me lo ricreo dopo
            # Salvo solamente le righe che hanno 'BB' nella colonna H del file di partenza
            # Togliere importi negativi

            # Se l'importo e' una stringa allora controllo che il primo carattere sia diverso da '-', mentre se e' un int o un float deve essere rispettivamente > 0 oppure > 0.0
            condition = False

            if(isinstance(dataframe1.iat[i, IMPORTO], str)):
                condition = (dataframe1.iat[i, IMPORTO][0] != '-')
            
            else:
                condition = (dataframe1.iat[i, IMPORTO] > 0) or (dataframe1.iat[i, IMPORTO] > 0.0)

            if(condition):
                updateProvvigioniIncassi = False

                # Modifico l'importo del singolo versamento in un float
                importo_versamento = convertToFloat(dataframe1.iat[i, IMPORTO])
                # Modifico l'importo della singola provvigione in un float
                importo_provvigione = convertToFloat(dataframe1.iat[i, PROVVIGIONI])

                # BONIFICI TUTELA LEGALE
                if(dataframe1.iat[i, MOD_PAGAMENTO] == 'BB'):
                    updateProvvigioniIncassi = True
                    tutela_nr_polizza.append(dataframe1.iat[i, NUM_POLIZZA])
                    tutela_anagrafica.append(dataframe1.iat[i, ANAGRAFICA])
                    tutela_importo.append(importo_versamento)
                    tutela_data.append(dataframe1.iat[i, DATA])

                # SOSPESI TUTELA LEGALE: CONTANTI e ASSEGNO BANCARIO (al momento non c'e' il FINANZIAMENTO AL CONSUMO per TUTELA LEGALE)
                elif(dataframe1.iat[i, MOD_PAGAMENTO] == 'CC' or dataframe1.iat[i, MOD_PAGAMENTO] == 'AB'):
                    updateProvvigioniIncassi = True
                    # Al momento il Finanziamento a Consumo (AGOS) non e' possibile con TUTELA LEGALE
                    dateAs_datetimeType = dataframe1.iat[i, DATA]
                    sospesi_tutela_data.append(dateAs_datetimeType)
                    sospesi_tutela_nr_polizza.append(dataframe1.iat[i, NUM_POLIZZA])
                    sospesi_tutela_anagrafica.append(dataframe1.iat[i, ANAGRAFICA])
                    sospesi_tutela_metodo_pagamento.append(dataframe1.iat[i, MOD_PAGAMENTO])
                    sospesi_tutela_importo.append(importo_versamento)
                    sospesi_tutela_agenzia.append('TUTELA LEGALE')     # sospesi_tutela_agenzia.append(dataframe1.iat[i, COLLABORATORE])
                    sospesi_tutela_compagnia.append('TUTELA')
                    sospesi_tutela_collaboratore.append('')
                    sospesi_tutela_pagato.append('No')
                    
                    updateAgencyTotaleSospesi(totale_sospesi_nuovi, importo_versamento, "TUTELA LEGALE")

                # Per il totale degli INCASSI e delle PROVVIGIONI considero tutti i metodi di pagamento tranne che 'MOBILE POS' che non so come si chiami su CATTOLICA
                if(updateProvvigioniIncassi == True):
                    if(totale_IncassiProvvigioni[0][TOT_DATA] == datetime.datetime(2000, 1, 1, 0, 0)):
                            totale_IncassiProvvigioni = [[dataframe1.iat[i, DATA], importo_versamento, importo_provvigione]]
                    else:
                        if(dataframe1.iat[i, DATA] != totale_IncassiProvvigioni[numberOfDifferentDates][TOT_DATA]):
                            # Se la data che sto analizzando dal file di TUTELA LEGALE e' diversa dall'ultima salvata nella list totale_IncassiProvvigioni, allora sto analizzando una nuova data.
                            # In questo caso stiamo facendo l'assunzione che tutte le date presenti nel file di TUTELA LEGALE siano tutte in ordine crescente/decrescente
                            totale_IncassiProvvigioni.append([dataframe1.iat[i, DATA], importo_versamento, importo_provvigione])
                            numberOfDifferentDates += 1
                            
                        else:
                            totale_IncassiProvvigioni[numberOfDifferentDates][TOT_INCASSI] += importo_versamento
                            totale_IncassiProvvigioni[numberOfDifferentDates][TOT_PROVVIGIONI] += importo_provvigione


        if(dataframe1.isnull().iat[i, NUM_POLIZZA] == False and findImporto == False):
            # Se trovo la stringa 'ANAGRAFICA' vuol dire che dal ciclo successivo inizio a salvare tutti i dati
            findImporto = True

    # Creazione dataframe BONIFICI TUTELA
    final_BonificiTutela = list(zip(tutela_data, tutela_importo, tutela_nr_polizza, tutela_anagrafica))
    df_BonificiTutela = pd.DataFrame(final_BonificiTutela)

    # Creazione dataframe SOSPESI
    final_SospesiTutela = list(zip(sospesi_tutela_data, sospesi_tutela_importo, sospesi_tutela_nr_polizza, sospesi_tutela_anagrafica, sospesi_tutela_agenzia, sospesi_tutela_compagnia, sospesi_tutela_collaboratore, sospesi_tutela_metodo_pagamento, sospesi_tutela_pagato))
    df_SospesiTutela = pd.DataFrame(final_SospesiTutela)

    # Creazione list per TOTALI INCASSI e PROVVIGIONI TUTELA LEGALE
    # listIncassi = list()
    # listProvvigioni = list()

    # for loc_date, tot_incassi, tot_provvigioni in totale_IncassiProvvigioni:
    #     listIncassi.append([loc_date, tot_incassi])
    #     listProvvigioni.append([loc_date, tot_provvigioni])

    # # Creazione dataframe INCASSI e PROVVIGIONI
    # df_IncassiTutela = pd.DataFrame(listIncassi)
    # df_ProvvigioniTutela = pd.DataFrame(listProvvigioni)
    

    datareadTutela = pd.read_excel(fileToWrite, sheet_name = sheetNameTutela, usecols='A')
    datareadSospesi = pd.read_excel(fileToWrite, sheet_name = sheetNameSospesi, usecols='A,B')
    datareadPrimaNota = pd.read_excel(fileToWrite, sheet_name = sheetNamePrimaNota, usecols='A')

    # Converto tutte le date presenti nei dati caricati dal foglio 'SOSPESI' in formato 'datetime'
    convertStringToDatetime(datareadSospesi, DATA_SOSPESI)

    BonificiRowData = [[], []]
    SospesiRowData = [[], []]

    # Creazione matrice con col = 0 : numero riga e col = 1 : data corrispondente per BONIFICI TUTELA
    for i in range(0, len(datareadTutela)):
        # Step 1: ricostruire la data da confrontare poi con quella presente nella tabella di BONIFICI TUTELA in PRIMA NOTA

        if(isinstance(datareadTutela.iat[i, 0], datetime.datetime)):
            # Se il dato appena letto dal foglio BONIFICI TUTELA in PRIMA NOTA nella colonna 'A' e' una data, vedo se corrisponde ad una delle date di cui ho dei dati da salvare
            for j in range(0, len(df_BonificiTutela)):
                if(datareadTutela.iat[i, 0] == df_BonificiTutela.iat[j, 0]):
                    BonificiRowData[0].append(i+1)
                    BonificiRowData[1].append(df_BonificiTutela.iat[j, 0])
                    break


    # Creazione matrice con col = 0 : numero riga e col = 1 : data corrispondente per SOSPESI
    dateFound = False

    for i in range(0, len(datareadSospesi)):
        # Step 1: ricostruire la data da confrontare poi con quella presente nella tabella dello sheet SOSPESI

        if(dateFound == True and datareadSospesi.isnull().iat[i, IMPORTO_SOSPESI] == True):
            SospesiRowData[0].append(i)
            dateFound = False
        
        if(isinstance(datareadSospesi.iat[i, DATA_SOSPESI], datetime.datetime) or isinstance(datareadSospesi.iat[i, DATA_SOSPESI], datetime.date)):
            if(datareadSospesi.iat[i, DATA_SOSPESI] != datareadSospesi.iat[i-1, DATA_SOSPESI]):
                # print(dataread.values[i])
                # Se il dato appena letto dal foglio SOSPESI nella colonna 'A' e' una data, vedo se corrisponde ad una delle date di cui ho dei dati da salvare
                for j in range(0, len(df_SospesiTutela)):
                    if(datareadSospesi.iat[i, DATA_SOSPESI] == df_SospesiTutela.iat[j, DATA_SOSPESI]):
                        dateFound = True
                        SospesiRowData[1].append(df_SospesiTutela.iat[j, DATA_SOSPESI])
                        break


    # Ricerca della riga nel foglio 'PRIMA NOTA' in cui andare a salvare i dati corrispondenti alla data newDateToCompare
    PrimaNotaRowData = []

    for i in range(0, len(totale_IncassiProvvigioni)):
        PrimaNotaRowData.append(findPrimaNotaRow_forIncassiProvvigioni(datareadPrimaNota, totale_IncassiProvvigioni[i][TOT_DATA]))

    print("Copia e salvataggio dati in esecuzione, attendere ...\n")

    # In BonificiRowData ho gli indici delle righe in ordine crescente di data, ma in df_BonificiTutela i vari dati si trovano in ordine decrescente di data, per questo motivo faccio un reverse for loop in modo tale da partire a salvare i dati con data piu' recente (BonificiRowData[i] con i = len(BonificiRowData)) fino ad arrivare a quelli con data meno recente (BonificiRowData[i] con i = 0)

    # Salvataggio dati in BONIFICI TUTELA
    for i in range(len(BonificiRowData[0])-1, -1, -1):

        final_listTutela = []

        for k in range(len(df_BonificiTutela)-1, -1, -1):
            if(BonificiRowData[1][i] == df_BonificiTutela.iat[k, 0]):
                # print("\nBefore: ", df_BonificiTutela.values[k, 1:4])
                final_listTutela.append(df_BonificiTutela.values[k, 1:4])
                # print("\nAfter: ", *final_listTutela, sep='\n')


        final_dfTutela = pd.DataFrame(final_listTutela)

        with pd.ExcelWriter(fileToWrite, engine ="openpyxl", mode='a', if_sheet_exists='overlay') as writer:
            final_dfTutela.to_excel(writer, index = False, header = False, sheet_name = sheetNameTutela, startrow = BonificiRowData[0][i] + 1, startcol = 1)


    # Salvataggio dati in SOSPESI
    for i in range(len(SospesiRowData[0])-1, -1, -1):

        final_listSospesi = []

        for k in range(len(df_SospesiTutela)-1, -1, -1):
            if(SospesiRowData[1][i] == df_SospesiTutela.iat[k, 0]):
                # Converto tutte le date da formato 'datetime' a formato stringa "%d/%m/%Y"
                # df_SospesiTutela.values[k, 0] e' un oggetto di tipo pandas Timestamp
                fromTimestampToDatetime = (df_SospesiTutela.values[k, 0]).date()
                # Converto la datetime.date in una datetime.datetime
                fromTimestampToDatetime = datetime.datetime.combine(fromTimestampToDatetime, datetime.datetime.min.time())
                datetimeString = convertDatetimeValueToString(fromTimestampToDatetime)
                temp_listSospesiTutela = df_SospesiTutela.values[k, 0:8].tolist() # Conversione da pandas dataframe a list
                temp_listSospesiTutela[0] = datetimeString  # Assegno la data come stringa al 1° valore della list
                final_listSospesi.append(temp_listSospesiTutela)
                break


        final_dfSospesi = pd.DataFrame(final_listSospesi)

        with pd.ExcelWriter(fileToWrite, engine ="openpyxl", mode='a', if_sheet_exists='overlay') as writer:
            final_dfSospesi.to_excel(writer, index = False, header = False, sheet_name = sheetNameSospesi, startrow = SospesiRowData[0][i] + 1, startcol = 0)

    
    # Salvataggio TOTALE INCASSI TUTELA LEGALE nel foglio 'PRIMA NOTA'
    for i in range(len(PrimaNotaRowData)-1, -1, -1):
        listIncassiProvvigioniTutela = [["Incassi TUTELA LEGALE", totale_IncassiProvvigioni[i][TOT_INCASSI]]]
        listIncassiProvvigioniTutela.append(["Provvigioni TUTELA LEGALE", totale_IncassiProvvigioni[i][TOT_PROVVIGIONI]])
        
        df_IncassiProvvigioniTutela = pd.DataFrame(listIncassiProvvigioniTutela)

        with pd.ExcelWriter(fileToWrite, engine ="openpyxl", mode='a', if_sheet_exists='overlay') as writer:
            df_IncassiProvvigioniTutela.to_excel(writer, index = False, header = False, sheet_name = sheetNamePrimaNota, startrow = PrimaNotaRowData[i]+ROW_INDEX_TOT_TUTELA, startcol = 2)

            
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


# Legge tutti i dati nel foglio SOSPESI e va a scrivere in tutte le tabelle del foglio PRIMA NOTA i relativi sospesi del giorno
# ATTENZIONE perche' per come e' fatto adesso va a fare questo lavoro per tutti i giorni, anche per quelli che erano gia' stati fatti in precedenza.
# Bisogna quindi ottimizzare il tutto per far eseguire questa funzione solamente per i giorni di cui non sono stati ancora scritti i SOSPESI NUOVI
def readSospesiFromExcel(fileToWrite, lastDatetime):
    sheetNameSospesi = "SOSPESI"

    # Caricamento dei dati dal foglio 'SOSPESI'
    dataSospesiExcel = pd.read_excel(fileToWrite, sheet_name = sheetNameSospesi, usecols='A,B,E,I,K')

    print("Lettura dati dal sheet SOSPESI eseguita con successo.\n")

    # A -> 0 : DATA
    # B -> 1 : IMPORTO
    # E -> 2 : AGENZIA
    # I -> 3 : PAGATO
    # K -> 4 : NOTE

    listSospesiNew = []
    totSospesiNew = TotaleSospesiNew_Date()
    
    DATA = int(0)
    IMPORTO = int(1)
    AGENZIA = int(2)
    PAGATO = int(3)
    NOTE = int(4)

    # Inserisco 45 e non 50 per non saltare troppe righe per precauzione
    SIZEOF_SINGLE_TABLE_SOSPESI = int(45)

    previousDate = datetime.date(2024, 1, 1)    # Data di default

    todayDate = datetime.datetime.today()
    todayDate = todayDate.replace(hour = 0, minute = 0, second = 0, microsecond = 0)

    indexRowExecuted = []   # Lista che contiene tutti gli indici delle righe del foglio 'SOSPESI' in cui andare a scrivere la stringa "Eseguito" per non considerare piu' i dati di quel giorno per la scrittura dei NUOVI SOSPESI nel foglio 'PRIMA NOTA'
    atleastOneDataFound = False

    convertStringToDatetime(dataSospesiExcel, DATA)

    # Analizzo tutti i dati caricati dal foglio 'SOSPESI' per andare a creare i NUOVI SOSPESI per ogni data.
    # Salvo anche gli indici di tutte le righe nel foglio 'SOSPESI' in cui andare poi a scrivere la stringa "Eseguito" in modo tale da non considerare piu' la tabella di quella data nelle future esecuzioni dell'applicazione.
    i = 0
    while(i < len(dataSospesiExcel)):
        # Controllo prima che il dato alla riga i-esima non sia vuoto
        if(dataSospesiExcel.isnull().iat[i, DATA] == False and dataSospesiExcel.iat[i, DATA] != "DATA" and dataSospesiExcel.iat[i, DATA] != "TOTALE" ):
            # Se la riga i-esima ha una data, un importo diverso da nullo, e "NO" nella colonna "Pagato", allora salvo tale riga nel buffer
            if(isinstance(dataSospesiExcel.iat[i, DATA], datetime.datetime) or isinstance(dataSospesiExcel.iat[i, DATA], datetime.date) or isinstance(datetime.datetime.strptime(dataSospesiExcel.iat[i, DATA], "%d/%m/%Y"), datetime.datetime)):
                if(dataSospesiExcel.iat[i, NOTE] == "Eseguito"):
                    # Se trovo la stringa "Eseguito" vado alla tabella successiva, ossia a quella relativa alla data successiva
                    i += SIZEOF_SINGLE_TABLE_SOSPESI - 1    # -1 in quanto alla fine del while loop c'e' un i += 1
                elif(dataSospesiExcel.isnull().iat[i, IMPORTO] == False): # and dataSospesiExcel.iat[i, PAGATO].upper() == "NO"):
                    if(previousDate != dataSospesiExcel.iat[i, DATA]):      # Se la data corrente e' diversa da quella precedente vuol dire che sto salvando un nuovo record della listSospesiNew
                        # if(len(listSospesiNew) == 0):   # Per non fare l'append quanto si ha totSospesiNew tutta a 0 all'inizio
                        if(totSospesiNew.date != datetime.date(2024, 1, 1)):    # Per non fare l'append quanto si ha totSospesiNew tutta a 0 all'inizio
                            listSospesiNew.append(copy.deepcopy(totSospesiNew))

                        # Resetto totSospesiNew
                        totSospesiNew.totRho = 0.0
                        totSospesiNew.totSommaLombardo = 0.0
                        totSospesiNew.totLegnano = 0.0
                        totSospesiNew.totGallarate = 0.0
                        totSospesiNew.totAgos = 0.0
                        totSospesiNew.totTutelaLegale = 0.0

                        if(isinstance(dataSospesiExcel.iat[i, DATA], datetime.datetime) or isinstance(dataSospesiExcel.iat[i, DATA], datetime.date)):
                            totSospesiNew.date = dataSospesiExcel.iat[i, DATA]
                            previousDate = dataSospesiExcel.iat[i, DATA]
                        else:
                            totSospesiNew.date = datetime.datetime.strptime(dataSospesiExcel.iat[i, DATA], dateFormat)
                            previousDate = datetime.datetime.strptime(dataSospesiExcel.iat[i, DATA], dateFormat)
                            
                    atleastOneDataFound = True
                    updateAgencyTotaleSospesi(totSospesiNew, dataSospesiExcel.iat[i, IMPORTO], dataSospesiExcel.iat[i, AGENZIA])
                else:
                    # Non c'e' un importo e non c'e' la stringa "Eseguito", quindi e' una nuova tabella che sto analizzando, quindi salvo la riga in cui poi andare a scrivere la stringa "Eseguito" se e solo se la data della tabella che sto analizzando e' precedente o coincidente con la data attuale
                    if(dataSospesiExcel.iat[i, DATA] <= todayDate and dataSospesiExcel.iat[i, DATA] <= lastDatetime):
                        indexRowExecuted.append(i+1)

        i += 1

    # Aggiungo l'ultimo record alla list dei NUOVI SOSPESI - bug: se non trova nulla va comunque a fare un append di dati vuoti per la data 01/01/2024 andando a scrivere tali dati nel foglio 'PRIMA NOTA'
    if(atleastOneDataFound == True):
        listSospesiNew.append(totSospesiNew)

    strExecuted = list(["Eseguito"])
    df_sospesiExecuted = pd.DataFrame(strExecuted)

    for i in range(0, len(listSospesiNew)):
        writeSospesi_inPrimaNota(listSospesiNew[i], fileToWrite, listSospesiNew[i].date)
        time.sleep(2)

    print("Numero di righe in cui scrivere la stringa 'Eseguito' = ", len(indexRowExecuted), ".\n")

    for i in range(0, len(indexRowExecuted)):
        with pd.ExcelWriter(fileToWrite, engine ="openpyxl", mode='a', if_sheet_exists='overlay') as writer:
            df_sospesiExecuted.to_excel(writer, index = False, header = False, sheet_name = sheetNameSospesi, startrow = indexRowExecuted[i], startcol = 10)    # 10 = 'J' -> NOTE

        print("", end=f"\rNumero di righe mancanti in cui scrivere la stringa 'Eseguito': {len(indexRowExecuted) - i} %")



