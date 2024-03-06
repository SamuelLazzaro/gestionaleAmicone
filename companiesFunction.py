import numpy as np
import pandas as pd
# import pandas.io.formats.style
import datetime
import re
import os

from auxiliaryFunction import findAgencyFromSubagent, updateAgencyTotaleSospesi, writeSospesi_inPrimaNota

sheetNameSospesi = 'SOSPESI'

DATA_SOSPESI    = int(0)
IMPORTO_SOSPESI = int(1)

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

    year_month_day = re.findall('\\d{4}-\\d{2}-\\d{2}', fileGenerali_read)[0]

    year = year_month_day[0:4]
    month = year_month_day[5:7]
    day = year_month_day[8:10]

    findImporto = False
    # read by default 1st sheet of an excel file
    # ATTENZIONE: per come e' fatto attualmente il file di GENERALI vi sono molte righe prima della tabella con i dati, quindi posso non considerare il parametro 'header' nella read_excel.
    # Dovesse pero' cambiare il formato come quello di CATTOLICA bisognera' probabilmente modificare il funzionamento di fileImporto, in quanto se non viene
    # utilizzato il parametro 'header' vengono letti tutti i dati da dopo la riga di intestazione del file excel.
    dataframe1 = pd.read_excel(fileGenerali_read, usecols='A,B,H,J,K,P')

    print("\nLettura file GENERALI eseguita correttamente.\n")

    # A -> 0 : NUMERO POLIZZA
    # B -> 1 : ANAGRAFICA (CONTRAENTE)
    # H -> 2 : MODALITA' PAGAMENTO
    # J -> 3 : IMPORTO
    # K -> 4 : COLLABORATORE
    # P -> 5 : PROVVIGIONI
    
    NUM_POLIZZA = int(0)
    ANAGRAFICA = int(1)
    MOD_PAGAMENTO = int(2)
    IMPORTO = int(3)
    COLLABORATORE = int(4)
    PROVVIGIONI = int(5)

    generali_nr_polizza = []
    generali_anagrafica = []
    generali_importo = []

    sospesi_generali_nr_polizza = []
    sospesi_generali_anagrafica = []
    sospesi_generali_importo = []
    sospesi_generali_agenzia = []
    sospesi_generali_compagnia = []
    sospesi_generali_metodo_pagamento = []
    sospesi_generali_pagato = []

    totale_provvigioni = 0


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
                if(isinstance(dataframe1.iat[i, IMPORTO], str)):
                    dataframe1.iat[i, IMPORTO] = dataframe1.iat[i, IMPORTO].replace('.', '')
                    importo_versamento = dataframe1.iat[i, IMPORTO].replace(',', '.')
                    importo_versamento = float(importo_versamento)
                else:
                    importo_versamento = float(dataframe1.iat[i, IMPORTO])

                if (dataframe1.iat[i, MOD_PAGAMENTO] == 'BONIFICO'):
                    # BONIFICO con FINANZIAMENTO A CONSUMO GENERALI -> da inserire nei SOSPESI
                    if(dataframe1.iat[i, NUM_POLIZZA].find('Fin. Consumo') != -1):
                        sospesi_generali_nr_polizza.append(dataframe1.iat[i, NUM_POLIZZA])
                        sospesi_generali_anagrafica.append(dataframe1.iat[i, ANAGRAFICA])
                        sospesi_generali_metodo_pagamento.append(dataframe1.iat[i, MOD_PAGAMENTO])
                        sospesi_generali_importo.append(importo_versamento)
                        sospesi_generali_agenzia.append('AGOS')
                        sospesi_generali_compagnia.append('GENERALI')
                        sospesi_generali_pagato.append('No')
                        totale_provvigioni += dataframe1.iat[i, PROVVIGIONI]
                        updateAgencyTotaleSospesi(totale_sospesi_nuovi, importo_versamento, agenzia)
                        continue        # Vado all'iterazione successiva del while loop

                    # BONIFICO GENERALI
                    generali_nr_polizza.append(dataframe1.iat[i, NUM_POLIZZA])
                    generali_anagrafica.append(dataframe1.iat[i, ANAGRAFICA])
                    generali_importo.append(importo_versamento)

                # CONTANTI o ASSEGNO BANCARIO/POSTALE GENERALI
                elif (dataframe1.iat[i, MOD_PAGAMENTO] == 'CONTANTI' or dataframe1.iat[i, MOD_PAGAMENTO].find('ASSEGNO') != -1):
                        sospesi_generali_nr_polizza.append(dataframe1.iat[i, NUM_POLIZZA])
                        sospesi_generali_anagrafica.append(dataframe1.iat[i, ANAGRAFICA])
                        sospesi_generali_metodo_pagamento.append(dataframe1.iat[i, MOD_PAGAMENTO])
                        sospesi_generali_importo.append(importo_versamento)
                        agenzia = findAgencyFromSubagent(dataframe1.iat[i, COLLABORATORE])
                        sospesi_generali_agenzia.append(agenzia)
                        sospesi_generali_compagnia.append('GENERALI')
                        sospesi_generali_pagato.append('No')
                        totale_provvigioni += dataframe1.iat[i, PROVVIGIONI]
                        updateAgencyTotaleSospesi(totale_sospesi_nuovi, importo_versamento, agenzia)
                

        if(dataframe1.isnull().iat[i, NUM_POLIZZA] == False and dataframe1.iat[i, ANAGRAFICA] == 'ANAGRAFICA' and findImporto == False):
            # Se trovo la stringa 'ANAGRAFICA' vuol dire che dal ciclo successivo inizio a salvare tutti i dati
            findImporto = True

    
    # Calcolo totale dei sospesi del giorno da sommare al totale dei sospesi precedenti in PRIMA_NOTA
    # totale_sospesi = 0

    # for i in range(0, len(sospesi_generali_importo)):
    #     print(float(sospesi_generali_importo[i]))

    #     totale_sospesi += float(sospesi_generali_importo[i])

    # print('Totale sospesi del giorno = ', totale_sospesi)


    # Creazione dataframe BONIFICI
    final_BonificiGenerali = list(zip(generali_importo, generali_nr_polizza, generali_anagrafica))
    df_BonificiGenerali = pd.DataFrame(final_BonificiGenerali)

    # Creazione dataframe SOSPESI
    final_SospesiGenerali = list(zip(sospesi_generali_importo, sospesi_generali_nr_polizza, sospesi_generali_anagrafica, sospesi_generali_agenzia, sospesi_generali_compagnia, sospesi_generali_metodo_pagamento, sospesi_generali_pagato))
    df_SospesiGenerali = pd.DataFrame(final_SospesiGenerali)

    # Lettura dati presenti nel file excel per i fogli sheetNameGenerali e sheetNameSospesi
    datareadBonifici = pd.read_excel(fileToWrite, sheet_name = sheetNameGenerali, usecols='A')
    datareadSospesi = pd.read_excel(fileToWrite, sheet_name = sheetNameSospesi, usecols='A,B')

    BonificiRowData = 0
    SospesiRowData = 0

    print("\nData inserita: ", day, '-', month, '-', year)

    dateToCompare = datetime.datetime(int(year), int(month), int(day), 0, 0)

    for i in range(0, len(datareadBonifici)):
        if(datareadBonifici.values[i] == dateToCompare):
            # print(dataread.values[i])
            BonificiRowData = i+1
            break

    dateFound = False

    for i in range(0, len(datareadSospesi)):
        if(dateFound == True and datareadSospesi.isnull().iat[i, IMPORTO_SOSPESI] == True):
            # Se in precedenza avevo trovato la data corrispondente e la colonna IMPORTO e' vuota, allora devo iniziare a salvare dalla riga i-esima i valori
            # In questo modo non sovrascrivo i dati precedentemente salvati
            SospesiRowData = i
            break
        if(datareadSospesi.iat[i, DATA_SOSPESI] == dateToCompare):
            dateFound = True

    print("Copia e salvataggio dati in esecuzione, attendere ...\n")
    # df_Generali.style.apply(lambda x: x.map(highlight_if_FinConsumo), axis=None)

    with pd.ExcelWriter(fileToWrite, engine ="openpyxl", mode='a', if_sheet_exists='overlay') as writer:
        df_BonificiGenerali.to_excel(writer, index = False, header = False, sheet_name = sheetNameGenerali, startrow = BonificiRowData+1, startcol = 1)
        df_SospesiGenerali.to_excel(writer, index = False, header = False, sheet_name = sheetNameSospesi, startrow = SospesiRowData+1, startcol = 1)
    
    writeSospesi_inPrimaNota(totale_sospesi_nuovi, fileToWrite, dateToCompare)

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

    day_month_year = re.findall('\\d{2}_\\d{2}_\\d{4}', pathName_read)[0]

    day = day_month_year[0:2]
    month = day_month_year[3:5]
    year = day_month_year[6:10]
    
    # read sheet 'Incassi' of CATTOLICA excel file
    dataframe1 = pd.read_excel(pathName_read, sheet_name='Incassi', usecols='A,E,H,I,K,Z')
    # Non avendo inserito il parametro 'header' nella read_excel, la 1^a riga di dataframe1 contiene gia' i dati
    print("\nLettura file CATTOLICA eseguita correttamente.")

    # A -> 0 : CONTRAENTE
    # E -> 1 : NUMERO POLIZZA
    # H -> 2 : IMPORTO PREMIO
    # I -> 3 : PROVVIGIONI
    # K -> 4 : MODALITA' PAGAMENTO
    # Z -> 5 : COLLABORATORE

    CONTRAENTE      = int(0)
    NUM_POLIZZA     = int(1)
    IMPORTO         = int(2)
    PROVVIGIONI     = int(3)
    MOD_PAGAMENTO   = int(4)
    COLLABORATORE   = int(5)

    cattolica_contraente = []
    cattolica_nr_polizza = []
    cattolica_importo = []

    sospesi_cattolica_nr_polizza = []
    sospesi_cattolica_anagrafica = []
    sospesi_cattolica_importo = []
    sospesi_cattolica_agenzia = []
    sospesi_cattolica_compagnia = []
    sospesi_cattolica_metodo_pagamento = []
    sospesi_cattolica_pagato = []

    totale_provvigioni = 0

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
                # Modifico l'importo del singolo versamento in un float
                if(isinstance(dataframe1.iat[i, IMPORTO], str)):
                    dataframe1.iat[i, IMPORTO] = dataframe1.iat[i, IMPORTO].replace('.', '')
                    importo_versamento = dataframe1.iat[i, IMPORTO].replace(',', '.')
                    importo_versamento = float(importo_versamento)
                else:
                    importo_versamento = float(dataframe1.iat[i, IMPORTO])

                # BONIFICI CATTOLICA
                if(dataframe1.iat[i, MOD_PAGAMENTO] == 'Bonifico su CC di Agenzia'):
                    cattolica_contraente.append(dataframe1.iat[i, 0])
                    cattolica_nr_polizza.append(dataframe1.iat[i, NUM_POLIZZA])
                    cattolica_importo.append(importo_versamento)

                # SOSPESI CATTOLICA: ASSEGNO BANCARIO, CONTANTI, FINANZIAMENTO AL CONSUMO
                elif(dataframe1.iat[i, MOD_PAGAMENTO].find('Assegno') != -1 or dataframe1.iat[i, MOD_PAGAMENTO] == 'Contante' or dataframe1.iat[i, MOD_PAGAMENTO] == 'Bonifico su CC di Direzione'):
                        sospesi_cattolica_nr_polizza.append(dataframe1.iat[i, NUM_POLIZZA])
                        sospesi_cattolica_anagrafica.append(dataframe1.iat[i, CONTRAENTE])
                        sospesi_cattolica_metodo_pagamento.append(dataframe1.iat[i, MOD_PAGAMENTO])
                        sospesi_cattolica_importo.append(importo_versamento)
                        if(dataframe1.iat[i, MOD_PAGAMENTO] == 'Bonifico su CC di Direzione'):
                            sospesi_cattolica_agenzia.append('AGOS')
                            updateAgencyTotaleSospesi(totale_sospesi_nuovi, importo_versamento, "AGOS")
                        else:
                            agenzia = findAgencyFromSubagent(dataframe1.iat[i, COLLABORATORE])
                            sospesi_cattolica_agenzia.append(agenzia)
                            updateAgencyTotaleSospesi(totale_sospesi_nuovi, importo_versamento, agenzia)

                        sospesi_cattolica_compagnia.append('CATTOLICA')
                        sospesi_cattolica_pagato.append('No')

                        totale_provvigioni += dataframe1.iat[i, PROVVIGIONI]
                        


    # Creazione dataframe BONIFICI
    final_BonificiCattolica = list(zip(cattolica_importo, cattolica_nr_polizza, cattolica_contraente))
    df_BonificiCattolica = pd.DataFrame(final_BonificiCattolica)

    # Creazione dataframe SOSPESI
    final_SospesiCattolica = list(zip(sospesi_cattolica_importo, sospesi_cattolica_nr_polizza, sospesi_cattolica_anagrafica, sospesi_cattolica_agenzia, sospesi_cattolica_compagnia, sospesi_cattolica_metodo_pagamento, sospesi_cattolica_pagato))
    df_SospesiCattolica = pd.DataFrame(final_SospesiCattolica)

    # Dal file finale vado a leggere tutte le date presenti nel relativo sheet nella colonna 'A'
    datareadBonifici = pd.read_excel(fileToWrite, sheet_name = sheetNameCattolica, usecols='A')
    datareadSospesi = pd.read_excel(fileToWrite, sheet_name = sheetNameSospesi, usecols='A,B')

    BonificiRowData = 0
    SospesiRowData = 0

    print("\nData inserita: ", day, '-', month, '-', year)

    dateToCompare = datetime.datetime(int(year), int(month), int(day), 0, 0)

    # Ricerca numero riga in cui andare a salvare i nuovi record in BONIFICI CATTOLICA
    for i in range(0, len(datareadBonifici)):
        if(datareadBonifici.values[i] == dateToCompare):
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
        if(datareadSospesi.iat[i, DATA_SOSPESI] == dateToCompare):
            dateFound = True

    print("Copia e salvataggio dati in esecuzione, attendere ...\n")

    with pd.ExcelWriter(fileToWrite, engine ="openpyxl", mode='a', if_sheet_exists='overlay') as writer:
        df_BonificiCattolica.to_excel(writer, index = False, header = False, sheet_name = sheetNameCattolica, startrow = BonificiRowData+1, startcol = 1)
        df_SospesiCattolica.to_excel(writer, index = False, header = False, sheet_name = sheetNameSospesi, startrow = SospesiRowData+1, startcol = 1)

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
    sospesi_tutela_metodo_pagamento = []
    sospesi_tutela_pagato = []
    sospesi_tutela_data = []

    totale_provvigioni = 0

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
                # Modifico l'importo del singolo versamento in un float
                if(isinstance(dataframe1.iat[i, IMPORTO], str)):
                    dataframe1.iat[i, IMPORTO] = dataframe1.iat[i, IMPORTO].replace('.', '')
                    importo_versamento = dataframe1.iat[i, IMPORTO].replace(',', '.')
                    importo_versamento = float(importo_versamento)
                else:
                    importo_versamento = float(dataframe1.iat[i, IMPORTO])

                # BONIFICI TUTELA LEGALE
                if(dataframe1.iat[i, MOD_PAGAMENTO] == 'BB'):
                    tutela_nr_polizza.append(dataframe1.iat[i, NUM_POLIZZA])
                    tutela_anagrafica.append(dataframe1.iat[i, ANAGRAFICA])
                    tutela_importo.append(importo_versamento)
                    tutela_data.append(dataframe1.iat[i, DATA])

                # SOSPESI TUTELA LEGALE: CONTANTI e ASSEGNO BANCARIO (al momento non c'e' il FINANZIAMENTO AL CONSUMO per TUTELA LEGALE)
                elif(dataframe1.iat[i, MOD_PAGAMENTO] == 'CC' or dataframe1.iat[i, MOD_PAGAMENTO] == 'AB'):
                        # Al momento il Finanziamento a Consumo (AGOS) non e' possibile con TUTELA LEGALE
                        sospesi_tutela_data.append(dataframe1.iat[i, DATA])
                        sospesi_tutela_nr_polizza.append(dataframe1.iat[i, NUM_POLIZZA])
                        sospesi_tutela_anagrafica.append(dataframe1.iat[i, ANAGRAFICA])
                        sospesi_tutela_metodo_pagamento.append(dataframe1.iat[i, MOD_PAGAMENTO])
                        sospesi_tutela_importo.append(importo_versamento)
                        sospesi_tutela_agenzia.append('Indefinito')     # sospesi_tutela_agenzia.append(dataframe1.iat[i, COLLABORATORE])
                        sospesi_tutela_compagnia.append('TUTELA')
                        sospesi_tutela_pagato.append('No')
                        
                        totale_provvigioni += dataframe1.iat[i, PROVVIGIONI]
                        updateAgencyTotaleSospesi(totale_sospesi_nuovi, importo_versamento, "TUTELA LEGALE")

        if(dataframe1.isnull().iat[i, NUM_POLIZZA] == False and findImporto == False):
            # Se trovo la stringa 'ANAGRAFICA' vuol dire che dal ciclo successivo inizio a salvare tutti i dati
            findImporto = True

    # Creazione dataframe BONIFICI TUTELA
    final_BonificiTutela = list(zip(tutela_data, tutela_importo, tutela_nr_polizza, tutela_anagrafica))
    df_BonificiTutela = pd.DataFrame(final_BonificiTutela)

    # Creazione dataframe SOSPESI
    final_SospesiTutela = list(zip(sospesi_tutela_data, sospesi_tutela_importo, sospesi_tutela_nr_polizza, sospesi_tutela_anagrafica, sospesi_tutela_agenzia, sospesi_tutela_compagnia, sospesi_tutela_metodo_pagamento, sospesi_tutela_pagato))
    df_SospesiTutela = pd.DataFrame(final_SospesiTutela)

    datareadTutela = pd.read_excel(fileToWrite, sheet_name = sheetNameTutela, usecols='A')
    datareadSospesi = pd.read_excel(fileToWrite, sheet_name = sheetNameSospesi, usecols='A,B')

    BonificiRowData = [[], []]
    SospesiRowData = [[], []]

    # Creazione matrice con col = 0 : numero riga e col = 1 : data corrispondente per BONIFICI TUTELA
    for i in range(0, len(datareadTutela)):
        # Step 1: ricostruire la data da confrontare poi con quella presente nella tabella di BONIFICI TUTELA in PRIMA NOTA

        if(isinstance(datareadTutela.iat[i, 0], datetime.datetime)):
            # print(dataread.values[i])
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
        
        if(isinstance(datareadSospesi.iat[i, 0], datetime.datetime)):
            # print(dataread.values[i])
            # Se il dato appena letto dal foglio BONIFICI TUTELA in PRIMA NOTA nella colonna 'A' e' una data, vedo se corrisponde ad una delle date di cui ho dei dati da salvare
            for j in range(0, len(df_SospesiTutela)):
                if(datareadSospesi.iat[i, 0] == df_SospesiTutela.iat[j, 0]):
                    # SospesiRowData[0].append(i+1)
                    dateFound = True
                    SospesiRowData[1].append(df_SospesiTutela.iat[j, 0])
                    break


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
                # print("\nBefore: ", df_Sospesi.values[k, 1:8])
                final_listSospesi.append(df_SospesiTutela.values[k, 1:8])
                # print("\nAfter: ", *final_listSospesi, sep='\n')

        final_dfSospesi = pd.DataFrame(final_listSospesi)        
        with pd.ExcelWriter(fileToWrite, engine ="openpyxl", mode='a', if_sheet_exists='overlay') as writer:
            final_dfSospesi.to_excel(writer, index = False, header = False, sheet_name = sheetNameSospesi, startrow = SospesiRowData[0][i] + 1, startcol = 1)

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