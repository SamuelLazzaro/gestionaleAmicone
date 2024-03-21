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
sheetNameRimborsi = 'RIMBORSI'

DATA_SOSPESI        = int(0)
DATA_RIMBORSI       = int(0)
IMPORTO_SOSPESI     = int(1)
IMPORTO_RIMBORSI    = int(1)

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

    findImporto = False
    # read by default 1st sheet of an excel file
    # ATTENZIONE: per come e' fatto attualmente il file di GENERALI vi sono molte righe prima della tabella con i dati, quindi posso non considerare il parametro 'header' nella read_excel.
    # Dovesse pero' cambiare il formato come quello di CATTOLICA bisognera' probabilmente modificare il funzionamento di fileImporto, in quanto se non viene
    # utilizzato il parametro 'header' vengono letti tutti i dati da dopo la riga di intestazione del file excel.
    dataframe1 = pd.read_excel(fileGenerali_read, usecols='A,B,C,H,J,K,P')

    print("\nLettura file GENERALI ", fileName_Generali, " eseguita correttamente.\n")

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

    rimborsi_generali_data = []
    rimborsi_generali_nr_polizza = []
    rimborsi_generali_anagrafica = []
    rimborsi_generali_importo = []
    rimborsi_generali_agenzia = []
    rimborsi_generali_compagnia = []
    rimborsi_generali_collaboratore = []
    rimborsi_generali_metodo_pagamento = []
    rimborsi_generali_pagato = []

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

            # Se nella colonna NUMERO POLIZZA trovo la scritta "Restituz. RID agenziale" (es. GENERALI 01/03/2024) salto completamente la riga in quanto e' un caso particolarissimo
            if(dataframe1.iat[i, NUM_POLIZZA].find('Restituz. RID agenziale') != -1):
                continue

            # L'importo della provvigione e dell'incasso puo' essere negativo
            # Modifico l'importo della singola provvigione in un float
            importo_provvigione = convertToFloat(dataframe1.iat[i, PROVVIGIONI])
            importo_incasso = convertToFloat(dataframe1.iat[i, IMPORTO])
            importo_versamento = convertToFloat(dataframe1.iat[i, IMPORTO])
                                                                                                                        
            # Calcolo dei totali degli INCASSI e delle PROVVIGIONI di GENERALI
            # BONIFICO, CONTANTI, ASSEGNO BANCARIO/POSTALE: si INCASSI, si PROVVIGIONI
            # MOBILE POS, VIRTUAL POS: no INCASSI, si PROVVIGIONI
            # ANTICIPO AGENTE: si INCASSI, si PROVVIGIONI
            # FINANZIAMENTO AL CONSUMO: no INCASSI, si PROVVIGIONI
            # REGOLAZIONE SU CONTO COMPENSO: no INCASSI, si PROVVIGIONI prendendo l'importo dell'INCASSO e cambiandolo di segno
            # COMPENSAZIONE: normalmente ha sia importo_incasso sia importo_provvigione = 0.0

            if(dataframe1.iat[i, MOD_PAGAMENTO] == 'BONIFICO' or dataframe1.iat[i, MOD_PAGAMENTO] == 'CONTANTI' or dataframe1.iat[i, MOD_PAGAMENTO].find('ASSEGNO') != -1 or dataframe1.iat[i, MOD_PAGAMENTO] == 'ANTICIPO AGENTE'):
                totale_incassi += importo_incasso
                if(importo_incasso >= 0.0):
                    totale_provvigioni += importo_provvigione
                else:
                    # Se importo_incasso < 0.0 sul file la provvigione comparira' positiva, ma in realta' deve essere negativa. Uso il valore assoluto per evitare problemi in file gia' modificati da Gigi
                    totale_provvigioni += (-abs(importo_provvigione))
            elif(dataframe1.iat[i, MOD_PAGAMENTO] == 'MOBILE POS' or dataframe1.iat[i, MOD_PAGAMENTO] == 'VIRTUAL POS'):
                totale_provvigioni += importo_provvigione
            elif(dataframe1.iat[i, MOD_PAGAMENTO] == 'COMPENSAZIONE'):
                totale_incassi += importo_incasso
                totale_provvigioni += importo_provvigione
            elif(dataframe1.iat[i, MOD_PAGAMENTO] == 'FINANZIAMENTO AL CONSUMO'):
                totale_provvigioni += importo_provvigione
            elif(dataframe1.iat[i, MOD_PAGAMENTO].find('REGOLAZIONE SU CONTO COMPENSO') != -1):
                # Nel caso di "REGOLAZIONE SU CONTO COMPENSO" e "REGOLAZIONE SU CONTO COMPENSO;COMPENSAZIONE", l'importo presente sotto la colonna IMPORTI deve essere cambiato di segno ed aggiunto alle PROVVIGIONI invece che agli INCASSI
                totale_provvigioni += (-importo_incasso)
            else:
                raise Exception("\nTrovato un metodo di pagamento non convenzionale nel file di GENERALI.\n")

            if(importo_versamento > 0.0):
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
                        continue        # Vado all'iterazione successiva del while loop -> NON TOGLIERE
                    else:
                        # BONIFICO GENERALI
                        generali_nr_polizza.append(dataframe1.iat[i, NUM_POLIZZA])
                        generali_anagrafica.append(dataframe1.iat[i, ANAGRAFICA])
                        generali_importo.append(importo_versamento)
                        continue        # Vado all'iterazione successiva del while loop -> NON TOGLIERE

            # Gestione SOSPESI:
            # - BONIFICI (non Finanziamento al consumo): vengono inseriti nei RIMBORSI nel caso in cui abbiano degli importi negativi
            # - CONTANTI o ASSEGNO BANCARIO/POSTALE GENERALI: vengono inseriti nei SOSPESI se hanno degli importi positivi, oppure nei RIMBORSI se hanno degli importi negativi
            # ATTENZIONE:
            if (dataframe1.iat[i, MOD_PAGAMENTO] == 'CONTANTI' or dataframe1.iat[i, MOD_PAGAMENTO].find('ASSEGNO') != -1 or dataframe1.iat[i, MOD_PAGAMENTO] == 'ANTICIPO AGENTE' or (dataframe1.iat[i, MOD_PAGAMENTO] == 'BONIFICO' and dataframe1.iat[i, NUM_POLIZZA].find('Fin. Consumo') == -1 and importo_versamento < 0.0)):
                dateString = dateToCompare.strftime(dateFormat)

                if(importo_versamento >= 0.0):
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
                else:
                    # importo_versamento < 0.0
                    rimborsi_generali_data.append(dateString)
                    rimborsi_generali_nr_polizza.append(dataframe1.iat[i, NUM_POLIZZA])
                    rimborsi_generali_anagrafica.append(dataframe1.iat[i, ANAGRAFICA])
                    rimborsi_generali_metodo_pagamento.append(dataframe1.iat[i, MOD_PAGAMENTO])
                    rimborsi_generali_importo.append(importo_versamento)
                    agenzia = findAgencyFromSubagent(dataframe1.iat[i, COLLABORATORE])
                    rimborsi_generali_agenzia.append(agenzia)
                    rimborsi_generali_compagnia.append('GENERALI')
                    rimborsi_generali_collaboratore.append(dataframe1.iat[i, COLLABORATORE])
                    rimborsi_generali_pagato.append('No')
                    
                

        if(dataframe1.isnull().iat[i, NUM_POLIZZA] == False and dataframe1.iat[i, ANAGRAFICA] == 'ANAGRAFICA' and findImporto == False):
            # Se trovo la stringa 'ANAGRAFICA' vuol dire che dal ciclo successivo inizio a salvare tutti i dati
            findImporto = True


    # Creazione dataframe BONIFICI
    final_BonificiGenerali = list(zip(generali_importo, generali_nr_polizza, generali_anagrafica))
    df_BonificiGenerali = pd.DataFrame(final_BonificiGenerali)

    # Creazione dataframe SOSPESI
    final_SospesiGenerali = list(zip(sospesi_generali_data, sospesi_generali_importo, sospesi_generali_nr_polizza, sospesi_generali_anagrafica, sospesi_generali_agenzia, sospesi_generali_compagnia, sospesi_generali_collaboratore, sospesi_generali_metodo_pagamento, sospesi_generali_pagato))
    df_SospesiGenerali = pd.DataFrame(final_SospesiGenerali)

    # Creazione dataframe RIMBORSI
    final_RimborsiGenerali = list(zip(rimborsi_generali_data, rimborsi_generali_importo, rimborsi_generali_nr_polizza, rimborsi_generali_anagrafica, rimborsi_generali_agenzia, rimborsi_generali_compagnia, rimborsi_generali_collaboratore, rimborsi_generali_metodo_pagamento, rimborsi_generali_pagato))
    df_RimborsiGenerali = pd.DataFrame(final_RimborsiGenerali)

    # Creazione dataframe INCASSI e PROVVIGIONI GENERALI
    listIncassiProvvigioniGenerali = [["Incassi GENERALI", totale_incassi]]
    listIncassiProvvigioniGenerali.append(["Provvigioni GENERALI", totale_provvigioni])

    df_IncassiProvvigioniGenerali = pd.DataFrame(listIncassiProvvigioniGenerali)

    # Lettura dati presenti nel file excel per i fogli sheetNameGenerali, sheetNameSospesi e sheetNamePrimaNota
    datareadBonifici = pd.read_excel(fileToWrite, sheet_name = sheetNameGenerali, usecols = 'A')
    datareadSospesi = pd.read_excel(fileToWrite, sheet_name = sheetNameSospesi, usecols = 'A,B')
    datareadRimborsi = pd.read_excel(fileToWrite, sheet_name = sheetNameRimborsi, usecols = 'A,B')
    datareadPrimaNota = pd.read_excel(fileToWrite, sheet_name = sheetNamePrimaNota, usecols = 'A')

    BonificiRowData = 0
    SospesiRowData = 0
    RimborsiRowData = 0

    # In questo caso e' una datetime.datetime
    newDateToCompare = dateToCompare

    # Ricerca dell'indice di riga nel foglio 'BONIFICI GENERALI ' in cui andare a salvare i dati corrispondenti alla data newDateToCompare
    listDateBonificiGenerali = datareadBonifici.values.tolist()
    BonificiRowData = listDateBonificiGenerali.index([newDateToCompare]) + 1

    dateFoundSospesi = False
    dateFoundRimborsi = False

    # Ricerca della riga nel foglio 'SOSPESI' in cui andare a salvare i dati corrispondenti alla data newDateToCompare
    for i in range(0, len(datareadSospesi)):
        if(dateFoundSospesi == True and datareadSospesi.isnull().iat[i, IMPORTO_SOSPESI] == True):
            # Se in precedenza avevo trovato la data corrispondente e la colonna IMPORTO e' vuota, allora devo iniziare a salvare dalla riga i-esima i valori
            # In questo modo non sovrascrivo i dati precedentemente salvati
            SospesiRowData = i
            break
        if(datareadSospesi.iat[i, DATA_SOSPESI] == newDateToCompare):
            dateFoundSospesi = True

    # Ricerca della riga nel foglio 'RIMBORSI' in cui andare a salvare i dati corrispondenti alla data newDateToCompare
    for i in range(0, len(datareadRimborsi)):
        if(dateFoundRimborsi == True and datareadRimborsi.isnull().iat[i, IMPORTO_RIMBORSI] == True):
            # Se in precedenza avevo trovato la data corrispondente e la colonna IMPORTO e' vuota, allora devo iniziare a salvare dalla riga i-esima i valori
            # In questo modo non sovrascrivo i dati precedentemente salvati
            RimborsiRowData = i
            break
        if(datareadRimborsi.iat[i, DATA_RIMBORSI] == newDateToCompare):
            dateFoundRimborsi = True

    # Ricerca della riga nel foglio 'PRIMA NOTA' in cui andare a salvare i dati corrispondenti alla data newDateToCompare
    PrimaNotaRowData = findPrimaNotaRow_forIncassiProvvigioni(datareadPrimaNota, newDateToCompare)

    print("Copia e salvataggio dei dati del file ", fileName_Generali, " di GENERALI in esecuzione, attendere ...\n")

    with pd.ExcelWriter(fileToWrite, engine ="openpyxl", mode='a', if_sheet_exists='overlay') as writer:
        df_BonificiGenerali.to_excel(writer, index = False, header = False, sheet_name = sheetNameGenerali, startrow = BonificiRowData+1, startcol = 1)
        df_SospesiGenerali.to_excel(writer, index = False, header = False, sheet_name = sheetNameSospesi, startrow = SospesiRowData+1, startcol = 0)
        df_RimborsiGenerali.to_excel(writer, index = False, header = False, sheet_name = sheetNameRimborsi, startrow = RimborsiRowData+1, startcol = 0)
        df_IncassiProvvigioniGenerali.to_excel(writer, index = False, header = False, sheet_name = sheetNamePrimaNota, startrow = PrimaNotaRowData+ROW_INDEX_TOT_GENERALI, startcol = 2)        # 2 -> 'C' : DESCRIZIONE

    print("Copia dei dati del file ", fileName_Generali, " di GENERALI terminata.\n")
    
    #  Rinomino il file di cui ho appena salvato i dati con la desinenza '_checked'
    renameFileChecked(fileGenerali_read)

    print("File ", fileName_Generali, " rinominato con '_checked' come desinenza.\n")



#   -----------------------------------------------------------------------------------------
#   -----------------------------------------------------------------------------------------
#   -----------------------------------------------------------------------------------------
#   -----------------------------------------------------------------------------------------
# Funzione per leggere i dati dal file di CATTOLICA e salvarli nel file finale 'fileToWrite'
def readFromCattolica(fileName_Cattolica, pathName_read, fileToWrite, totale_sospesi_nuovi):

    sheetNameCattolica = 'BONIFICI CATTOLICA'
    
    # read sheet 'Incassi' of CATTOLICA excel file
    dataframe1 = pd.read_excel(pathName_read, sheet_name='Incassi', usecols='A,E,H,I,K,Y,Z')
    # Non avendo inserito il parametro 'header' nella read_excel, la 1^a riga di dataframe1 contiene gia' i dati

    print("\nLettura file CATTOLICA ", fileName_Cattolica, " eseguita correttamente.")

    # A -> 0 : CONTRAENTE
    # E -> 1 : NUMERO POLIZZA
    # H -> 2 : IMPORTO PREMIO
    # I -> 3 : PROVVIGIONI
    # K -> 4 : MODALITA' PAGAMENTO
    # Y -> 5 : DATA FOGLIO CASSA
    # Z -> 6 : COLLABORATORE

    CONTRAENTE          = int(0)
    NUM_POLIZZA         = int(1)
    IMPORTO             = int(2)
    PROVVIGIONI         = int(3)
    MOD_PAGAMENTO       = int(4)
    DATA_FOGLIO_CASSA   = int(5)
    COLLABORATORE       = int(6)


    # Il file di CATTOLICA e' vuoto se len(dataframe1) = 1 e se le celle di NUM_POLIZZA e MOD_PAGAMENTO sono vuote: in tal caso rinomino il file ed esco dalla funzione
    if(len(dataframe1) == 0 or (len(dataframe1) <= 1 and dataframe1.isnull().iat[0, CONTRAENTE] == True and dataframe1.isnull().iat[0, MOD_PAGAMENTO] == True and dataframe1.isnull().iat[0, NUM_POLIZZA] == True)):
        print("File ", fileName_Cattolica, " vuoto.\n")
        #  Rinomino comunque il file con la desinenza '_checked' in modo tale da non analizzarlo piu'
        renameFileChecked(pathName_read)
        print("File ", fileName_Cattolica, " rinominato con '_checked' come desinenza.\n")
        return

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

    rimborsi_cattolica_data = []
    rimborsi_cattolica_nr_polizza = []
    rimborsi_cattolica_anagrafica = []
    rimborsi_cattolica_importo = []
    rimborsi_cattolica_agenzia = []
    rimborsi_cattolica_compagnia = []
    rimborsi_cattolica_collaboratore = []
    rimborsi_cattolica_metodo_pagamento = []
    rimborsi_cattolica_pagato = []

    totale_incassi = 0
    totale_provvigioni = 0

    dateToCompare = dataframe1.iat[0, DATA_FOGLIO_CASSA]

    if(isinstance(dateToCompare, datetime.datetime) == False):
        dateToCompare = datetime.datetime.strptime(dateToCompare, dateFormat)
    

    for i in range(0, len(dataframe1)):
        if(dataframe1.isnull().iat[i, CONTRAENTE] == False and dataframe1.isnull().iat[i, NUM_POLIZZA] == False and dataframe1.isnull().iat[i, IMPORTO] == False):
            # NON devo salvare le righe di dati vuoti che si trovano all'interno della tabella con i dati da salvare

            # Modifico l'importo della singola provvigione in un float
            importo_versamento = convertToFloat(dataframe1.iat[i, IMPORTO])
            importo_incasso = convertToFloat(dataframe1.iat[i, IMPORTO])
            if(importo_incasso >= 0.0):
                importo_provvigione = convertToFloat(dataframe1.iat[i, PROVVIGIONI])
            else:
                # Se l'importo del premio e' < 0, allora anche la provvigione deve essere < 0
                importo_provvigione = -abs(convertToFloat(dataframe1.iat[i, PROVVIGIONI]))

            # Calcolo dei totali degli INCASSI e delle PROVVIGIONI di GENERALI
            # Assegno, Contante, Bonifico su CC di Agenzia, Bonifico su CC di Direzione (Finanziamento al consumo): si INCASSI, si PROVVIGIONI
            # Rid, MPos, Automatico (incasso = 0.0): no INCASSI, si PROVVIGIONI
            if(dataframe1.iat[i, MOD_PAGAMENTO].find('Assegno') != -1 or dataframe1.iat[i, MOD_PAGAMENTO] == 'Contante' or dataframe1.iat[i, MOD_PAGAMENTO].find('Bonifico') != -1):
                totale_incassi += importo_incasso
                totale_provvigioni += importo_provvigione
            elif(dataframe1.iat[i, MOD_PAGAMENTO] == 'Rid' or dataframe1.iat[i, MOD_PAGAMENTO] == 'MPos' or dataframe1.iat[i, MOD_PAGAMENTO] == 'Automatico'):
                totale_provvigioni += importo_provvigione
            else:
                raise Exception("\nTrovato un metodo di pagamento non convenzionale nel file di CATTOLICA.\n")

            # Gestione BONIFICI CATTOLICA (NON Finanziamento al Consumo) con importo > 0.0 -> foglio 'BONIFICI CATTOLICA)
            if(dataframe1.iat[i, MOD_PAGAMENTO] == 'Bonifico su CC di Agenzia' and importo_versamento > 0.0):
                # Importo BONIFICO positivo -> foglio 'BONIFICI CATTOLICA'
                cattolica_contraente.append(dataframe1.iat[i, 0])
                cattolica_nr_polizza.append(dataframe1.iat[i, NUM_POLIZZA])
                cattolica_importo.append(importo_versamento)

            # Foglio 'SOSPESI' deve contenere:
            # - Bonifico su CC di Direzione con importo >= 0.0
            # - Assegno e Contante con importo >= 0.0
            # Foglio 'RIMBORSI' deve contenere:
            # - Bonifico su CC di Agenzia con importo < 0.0
            # - Assegno e Contante con importo < 0.0
            elif((dataframe1.iat[i, MOD_PAGAMENTO] == 'Bonifico su CC di Agenzia' and importo_versamento < 0.0) or (dataframe1.iat[i, MOD_PAGAMENTO] == 'Bonifico su CC di Direzione' and importo_versamento > 0.0) or dataframe1.iat[i, MOD_PAGAMENTO].find('Assegno') != -1 or dataframe1.iat[i, MOD_PAGAMENTO] == 'Contante'):
                dateString = dateToCompare.strftime(dateFormat)
                
                if(importo_versamento >= 0.0):
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
                else:
                    # importo_versamento < 0.0
                    rimborsi_cattolica_data.append(dateString)
                    rimborsi_cattolica_nr_polizza.append(dataframe1.iat[i, NUM_POLIZZA])
                    rimborsi_cattolica_anagrafica.append(dataframe1.iat[i, CONTRAENTE])
                    rimborsi_cattolica_metodo_pagamento.append(dataframe1.iat[i, MOD_PAGAMENTO])
                    rimborsi_cattolica_importo.append(importo_versamento)

                    # Non puo' essere un 'Bonifico su CC di Direzione' con importo_versamento < 0.0
                    agenzia = findAgencyFromSubagent(dataframe1.iat[i, COLLABORATORE])
                    rimborsi_cattolica_agenzia.append(agenzia)
                    rimborsi_cattolica_compagnia.append('CATTOLICA')
                    rimborsi_cattolica_collaboratore.append(dataframe1.iat[i, COLLABORATORE])
                    rimborsi_cattolica_pagato.append('No')



    # Creazione dataframe BONIFICI
    final_BonificiCattolica = list(zip(cattolica_importo, cattolica_nr_polizza, cattolica_contraente))
    df_BonificiCattolica = pd.DataFrame(final_BonificiCattolica)

    # Creazione dataframe SOSPESI
    final_SospesiCattolica = list(zip(sospesi_cattolica_data, sospesi_cattolica_importo, sospesi_cattolica_nr_polizza, sospesi_cattolica_anagrafica, sospesi_cattolica_agenzia, sospesi_cattolica_compagnia, sospesi_cattolica_collaboratore, sospesi_cattolica_metodo_pagamento, sospesi_cattolica_pagato))
    df_SospesiCattolica = pd.DataFrame(final_SospesiCattolica)

    # Creazione dataframe RIMBORSI
    final_RimborsiCattolica = list(zip(rimborsi_cattolica_data, rimborsi_cattolica_importo, rimborsi_cattolica_nr_polizza, rimborsi_cattolica_anagrafica, rimborsi_cattolica_agenzia, rimborsi_cattolica_compagnia, rimborsi_cattolica_collaboratore, rimborsi_cattolica_metodo_pagamento, rimborsi_cattolica_pagato))
    df_RimborsiCattolica = pd.DataFrame(final_RimborsiCattolica)

    # Creazione dataframe INCASSI e PROVVIGIONI GENERALI
    listIncassiProvvigioniCattolica = [["Incassi CATTOLICA", totale_incassi]]
    listIncassiProvvigioniCattolica.append(["Provvigioni CATTOLICA", totale_provvigioni])

    df_IncassiProvvigioniCattolica = pd.DataFrame(listIncassiProvvigioniCattolica)

    # Dal file finale vado a leggere tutte le date presenti nel relativo sheet nella colonna 'A'
    datareadBonifici = pd.read_excel(fileToWrite, sheet_name = sheetNameCattolica, usecols='A')
    datareadSospesi = pd.read_excel(fileToWrite, sheet_name = sheetNameSospesi, usecols='A,B')
    datareadRimborsi = pd.read_excel(fileToWrite, sheet_name = sheetNameRimborsi, usecols='A,B')
    datareadPrimaNota = pd.read_excel(fileToWrite, sheet_name = sheetNamePrimaNota, usecols='A')

    # Converto tutte le date del foglio 'SOSPESI' e 'RIMBORSI' da formato stringa a formato 'datetime'
    # N.B. In realta' questa conversione potrebbe non essere necessaria
    convertStringToDatetime(datareadSospesi, DATA_SOSPESI)
    convertStringToDatetime(datareadRimborsi, DATA_RIMBORSI)

    BonificiRowData = 0
    SospesiRowData = 0
    RimborsiRowData = 0

    # In questo caso e' una datetime.datetime
    newDateToCompare = dateToCompare

    # Ricerca dell'indice di riga in cui andare a salvare i nuovi record in BONIFICI CATTOLICA
    listDateBonificiCattolica = datareadBonifici.values.tolist()
    BonificiRowData = listDateBonificiCattolica.index([newDateToCompare]) + 1

    # Ricerca numero riga su cui andare a salvare i nuovi record in SOSPESI e RIMBORSI senza sovrascrivere quelli precedenti relativi alla stessa data
    dateFoundSospesi = False
    dateFoundRimborsi = False

    for i in range(0, len(datareadSospesi)):
        if(dateFoundSospesi == True and datareadSospesi.isnull().iat[i, IMPORTO_SOSPESI] == True):
            # Se in precedenza avevo trovato la data corrispondente e la colonna IMPORTO e' vuota, allora devo iniziare a salvare dalla riga i-esima i valori
            # In questo modo non sovrascrivo i dati precedentemente salvati
            SospesiRowData = i
            break
        if(datareadSospesi.iat[i, DATA_SOSPESI] == newDateToCompare):
            dateFoundSospesi = True

    for i in range(0, len(datareadRimborsi)):
        if(dateFoundRimborsi == True and datareadRimborsi.isnull().iat[i, IMPORTO_RIMBORSI] == True):
            # Se in precedenza avevo trovato la data corrispondente e la colonna IMPORTO e' vuota, allora devo iniziare a salvare dalla riga i-esima i valori
            # In questo modo non sovrascrivo i dati precedentemente salvati
            RimborsiRowData = i
            break
        if(datareadRimborsi.iat[i, DATA_RIMBORSI] == newDateToCompare):
            dateFoundRimborsi = True

    # Ricerca della riga nel foglio 'PRIMA NOTA' in cui andare a salvare i dati corrispondenti alla data newDateToCompare
    PrimaNotaRowData = findPrimaNotaRow_forIncassiProvvigioni(datareadPrimaNota, newDateToCompare)
    
    print("Copia e salvataggio dati del file ", fileName_Cattolica, " di CATTOLICA in esecuzione, attendere ...\n")

    with pd.ExcelWriter(fileToWrite, engine ="openpyxl", mode='a', if_sheet_exists='overlay') as writer:
        df_BonificiCattolica.to_excel(writer, index = False, header = False, sheet_name = sheetNameCattolica, startrow = BonificiRowData+1, startcol = 1)
        df_SospesiCattolica.to_excel(writer, index = False, header = False, sheet_name = sheetNameSospesi, startrow = SospesiRowData+1, startcol = 0)
        df_RimborsiCattolica.to_excel(writer, index = False, header = False, sheet_name = sheetNameRimborsi, startrow = RimborsiRowData+1, startcol = 0)
        df_IncassiProvvigioniCattolica.to_excel(writer, index = False, header = False, sheet_name = sheetNamePrimaNota, startrow = PrimaNotaRowData+ROW_INDEX_TOT_CATTOLICA, startcol = 2)      # 2 -> 'C' : DESCRIZIONE
    
    print("Copia dei dati del file ", fileName_Cattolica, " di CATTOLICA terminata.\n")

    #  Rinomino il file di cui ho appena salvato i dati con la desinenza '_checked'
    renameFileChecked(pathName_read)

    print("File ", fileName_Cattolica, " rinominato con '_checked' come desinenza.\n")




#   -----------------------------------------------------------------------------------------
#   -----------------------------------------------------------------------------------------
#   -----------------------------------------------------------------------------------------
#   -----------------------------------------------------------------------------------------
# Funzione per leggere i dati dal file di TUTELA e salvarli nel file finale 'fileToWrite'
def readFromTutela(fileName_Tutela, fileTutela_read, fileToWrite, totale_sospesi_nuovi):
    sheetNameTutela = 'BONIFICI TUTELA'

    findImporto = False
    # read by default 1st sheet of an excel file
    dataframe1 = pd.read_excel(fileTutela_read, usecols='C,H,I,J,L,M')

    print("\nLettura file TUTELA ", fileName_Tutela, " eseguita correttamente.\n")

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

    rimborsi_tutela_nr_polizza = []
    rimborsi_tutela_anagrafica = []
    rimborsi_tutela_importo = []
    rimborsi_tutela_agenzia = []
    rimborsi_tutela_compagnia = []
    rimborsi_tutela_collaboratore = []
    rimborsi_tutela_metodo_pagamento = []
    rimborsi_tutela_pagato = []
    rimborsi_tutela_data = []

    defaultDatetime = datetime.datetime(2000, 1, 1, 0, 0)
    totale_IncassiProvvigioni = [[defaultDatetime, 0.0, 0.0]]

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

            # Modifico l'importo della singola provvigione in un float
            importo_versamento = convertToFloat(dataframe1.iat[i, IMPORTO])
            importo_incasso = convertToFloat(dataframe1.iat[i, IMPORTO])
            if(importo_incasso >= 0.0):
                importo_provvigione = convertToFloat(dataframe1.iat[i, PROVVIGIONI])
            else:
                importo_provvigione = -abs(convertToFloat(dataframe1.iat[i, PROVVIGIONI]))

            if(dataframe1.iat[i, MOD_PAGAMENTO] != 'CC' and dataframe1.iat[i, MOD_PAGAMENTO] != 'BB' and dataframe1.iat[i, MOD_PAGAMENTO] != 'AB'):
                raise Exception("\nTrovato metodo di pagamento non convenzionale in TUTELA LEGALE.\n")
            
            if(totale_IncassiProvvigioni[0][TOT_DATA] == defaultDatetime):
                # Primo valore inserito    
                totale_IncassiProvvigioni = [[dataframe1.iat[i, DATA], importo_incasso, importo_provvigione]]
            else:
                if(dataframe1.iat[i, DATA] != totale_IncassiProvvigioni[numberOfDifferentDates][TOT_DATA]):
                    # Se la data che sto analizzando dal file di TUTELA LEGALE e' diversa dall'ultima salvata nella list totale_IncassiProvvigioni, allora sto analizzando una nuova data.
                    # In questo caso stiamo facendo l'assunzione che tutte le date presenti nel file di TUTELA LEGALE siano tutte in ordine crescente/decrescente
                    totale_IncassiProvvigioni.append([dataframe1.iat[i, DATA], importo_incasso, importo_provvigione])
                    numberOfDifferentDates += 1
                else:
                    totale_IncassiProvvigioni[numberOfDifferentDates][TOT_INCASSI] += importo_incasso
                    totale_IncassiProvvigioni[numberOfDifferentDates][TOT_PROVVIGIONI] += importo_provvigione

            # ATTENZIONE: al momento per TUTELA LEGALE non e' previsto il FINANZIAMENTO AL CONSUMO
            
            # BONIFICI TUTELA LEGALE con importo > 0.0 -> foglio 'BONIFICI TUTELA LEGALE'
            if(dataframe1.iat[i, MOD_PAGAMENTO] == 'BB' and importo_versamento >= 0.0):
                tutela_nr_polizza.append(dataframe1.iat[i, NUM_POLIZZA])
                tutela_anagrafica.append(dataframe1.iat[i, ANAGRAFICA])
                tutela_importo.append(importo_versamento)
                tutela_data.append(dataframe1.iat[i, DATA])
                      
            # Foglio 'SOSPESI' deve contenere:
            # - ASSEGNI e CONTANTI con importo >= 0.0
            # Foglio 'RIMBORSI' deve contenere:
            # - ASSEGNI e CONTANTI con importo < 0.0
            # - BONIFICO con importo < 0.0 (N.B. In TUTELA LEGALE al momento non vi e' il FINANZIAMENTO AL CONSUMO)
            elif((dataframe1.iat[i, MOD_PAGAMENTO] == 'BB' and importo_versamento < 0.0) or dataframe1.iat[i, MOD_PAGAMENTO] == 'CC' or dataframe1.iat[i, MOD_PAGAMENTO] == 'AB'):
                # Al momento il Finanziamento a Consumo (AGOS) non e' possibile con TUTELA LEGALE
                dateAs_datetimeType = dataframe1.iat[i, DATA]

                if(importo_versamento >= 0.0):
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
                else:
                    # importo_versamento < 0.0
                    rimborsi_tutela_data.append(dateAs_datetimeType)
                    rimborsi_tutela_nr_polizza.append(dataframe1.iat[i, NUM_POLIZZA])
                    rimborsi_tutela_anagrafica.append(dataframe1.iat[i, ANAGRAFICA])
                    rimborsi_tutela_metodo_pagamento.append(dataframe1.iat[i, MOD_PAGAMENTO])
                    rimborsi_tutela_importo.append(importo_versamento)
                    rimborsi_tutela_agenzia.append('TUTELA LEGALE')     # sospesi_tutela_agenzia.append(dataframe1.iat[i, COLLABORATORE])
                    rimborsi_tutela_compagnia.append('TUTELA')
                    rimborsi_tutela_collaboratore.append('')
                    rimborsi_tutela_pagato.append('No')

        if(dataframe1.isnull().iat[i, NUM_POLIZZA] == False and findImporto == False):
            # Se trovo la stringa 'ANAGRAFICA' vuol dire che dal ciclo successivo inizio a salvare tutti i dati
            findImporto = True

    # Se totale_IncassiProvvigioni ha ancora come primo valore la defaultDate, vuol dire che il file di TUTELA LEGALE era vuoto, quindi rinomino il file con '_checked' ed esco dalla funzione
    if(len(totale_IncassiProvvigioni) == 1 and totale_IncassiProvvigioni[0][TOT_DATA] == defaultDatetime):
        print("File di TUTELA LEGALE vuoto.\n")
        # Rinomino il file di cui ho appena salvato i dati con la desinenza '_checked'
        renameFileChecked(fileTutela_read)
        print("File ", fileName_Tutela, " rinominato con '_checked' come desinenza.\n")
        return
    
    # Creazione dataframe BONIFICI TUTELA
    final_BonificiTutela = list(zip(tutela_data, tutela_importo, tutela_nr_polizza, tutela_anagrafica))
    df_BonificiTutela = pd.DataFrame(final_BonificiTutela)

    # Creazione dataframe SOSPESI
    final_SospesiTutela = list(zip(sospesi_tutela_data, sospesi_tutela_importo, sospesi_tutela_nr_polizza, sospesi_tutela_anagrafica, sospesi_tutela_agenzia, sospesi_tutela_compagnia, sospesi_tutela_collaboratore, sospesi_tutela_metodo_pagamento, sospesi_tutela_pagato))
    df_SospesiTutela = pd.DataFrame(final_SospesiTutela)

    # Creazione dataframe RIMBORSI
    final_RimborsiTutela = list(zip(rimborsi_tutela_data, rimborsi_tutela_importo, rimborsi_tutela_nr_polizza, rimborsi_tutela_anagrafica, rimborsi_tutela_agenzia, rimborsi_tutela_compagnia, rimborsi_tutela_collaboratore, rimborsi_tutela_metodo_pagamento, rimborsi_tutela_pagato))
    df_RimborsiTutela = pd.DataFrame(final_RimborsiTutela)  

    datareadTutela = pd.read_excel(fileToWrite, sheet_name = sheetNameTutela, usecols='A')
    datareadSospesi = pd.read_excel(fileToWrite, sheet_name = sheetNameSospesi, usecols='A,B')
    datareadRimborsi = pd.read_excel(fileToWrite, sheet_name = sheetNameRimborsi, usecols='A,B')
    datareadPrimaNota = pd.read_excel(fileToWrite, sheet_name = sheetNamePrimaNota, usecols='A')

    # Converto tutte le date presenti nei dati caricati dal foglio 'SOSPESI' e 'RIMBORSI' in formato 'datetime'
    convertStringToDatetime(datareadSospesi, DATA_SOSPESI)
    convertStringToDatetime(datareadRimborsi, DATA_RIMBORSI)

    BonificiRowData = [[], []]
    SospesiRowData = [[], []]
    RimborsiRowData = [[], []]

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
    dateSospesiFound = False
    dateRimborsiFound = False

    # Ricerca delle righe da salvare nel foglio 'SOSPESI'
    for i in range(0, len(datareadSospesi)):
        # Step 1: ricostruire la data da confrontare poi con quella presente nella tabella dello sheet SOSPESI

        if(dateSospesiFound == True and datareadSospesi.isnull().iat[i, IMPORTO_SOSPESI] == True):
            SospesiRowData[0].append(i)
            dateSospesiFound = False
        
        if(isinstance(datareadSospesi.iat[i, DATA_SOSPESI], datetime.datetime) or isinstance(datareadSospesi.iat[i, DATA_SOSPESI], datetime.date)):
            if(datareadSospesi.iat[i, DATA_SOSPESI] != datareadSospesi.iat[i-1, DATA_SOSPESI]):
                # print(dataread.values[i])
                # Se il dato appena letto dal foglio SOSPESI nella colonna 'A' e' una data, vedo se corrisponde ad una delle date di cui ho dei dati da salvare
                for j in range(0, len(df_SospesiTutela)):
                    if(datareadSospesi.iat[i, DATA_SOSPESI] == df_SospesiTutela.iat[j, DATA_SOSPESI]):
                        dateSospesiFound = True
                        SospesiRowData[1].append(df_SospesiTutela.iat[j, DATA_SOSPESI])
                        break

    # Ricerca delle righe da salvare nei 'RIMBORSI'
    for i in range(0, len(datareadRimborsi)):
        # Step 1: ricostruire la data da confrontare poi con quella presente nella tabella del foglio 'RIMBORSI'

        if(dateRimborsiFound == True and datareadRimborsi.isnull().iat[i, IMPORTO_RIMBORSI] == True):
            RimborsiRowData[0].append(i)
            dateRimborsiFound = False
        
        if(isinstance(datareadRimborsi.iat[i, DATA_RIMBORSI], datetime.datetime) or isinstance(datareadRimborsi.iat[i, DATA_RIMBORSI], datetime.date)):
            if(datareadRimborsi.iat[i, DATA_RIMBORSI] != datareadRimborsi.iat[i-1, DATA_RIMBORSI]):

                # Se il dato appena letto dal foglio 'RIMBORSI' nella colonna 'A' e' una data, vedo se corrisponde ad una delle date di cui ho dei dati da salvare
                for j in range(0, len(df_RimborsiTutela)):
                    if(datareadRimborsi.iat[i, DATA_RIMBORSI] == df_RimborsiTutela.iat[j, DATA_RIMBORSI]):
                        dateRimborsiFound = True
                        RimborsiRowData[1].append(df_RimborsiTutela.iat[j, DATA_RIMBORSI])
                        break


    # Ricerca della riga nel foglio 'PRIMA NOTA' in cui andare a salvare i dati corrispondenti alla data newDateToCompare
    PrimaNotaRowData = []

    for i in range(0, len(totale_IncassiProvvigioni)):
        PrimaNotaRowData.append(findPrimaNotaRow_forIncassiProvvigioni(datareadPrimaNota, totale_IncassiProvvigioni[i][TOT_DATA]))

    print("Copia e salvataggio dati del file ", fileName_Tutela, " di TUTELA LEGALE in esecuzione, attendere ...\n")

    # In BonificiRowData ho gli indici delle righe in ordine crescente di data, ma in df_BonificiTutela i vari dati si trovano in ordine decrescente di data, per questo motivo faccio un reverse for loop in modo tale da partire a salvare i dati con data piu' recente (BonificiRowData[i] con i = len(BonificiRowData)) fino ad arrivare a quelli con data meno recente (BonificiRowData[i] con i = 0)

    # Salvataggio dati in BONIFICI TUTELA
    if(len(BonificiRowData[0]) > 0):
        with pd.ExcelWriter(fileToWrite, engine ="openpyxl", mode='a', if_sheet_exists='overlay') as writer:
            for i in range(len(BonificiRowData[0])-1, -1, -1):

                final_listTutela = []

                for k in range(len(df_BonificiTutela)-1, -1, -1):
                    if(BonificiRowData[1][i] == df_BonificiTutela.iat[k, 0]):
                        final_listTutela.append(df_BonificiTutela.values[k, 1:4])

                final_dfTutela = pd.DataFrame(final_listTutela)
                final_dfTutela.to_excel(writer, index = False, header = False, sheet_name = sheetNameTutela, startrow = BonificiRowData[0][i] + 1, startcol = 1)


    # Salvataggio dati in SOSPESI
    if(len(SospesiRowData[0]) > 0):
        with pd.ExcelWriter(fileToWrite, engine ="openpyxl", mode='a', if_sheet_exists='overlay') as writer:
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

                final_dfSospesi = pd.DataFrame(final_listSospesi)
                final_dfSospesi.to_excel(writer, index = False, header = False, sheet_name = sheetNameSospesi, startrow = SospesiRowData[0][i] + 1, startcol = 0)


    # Salvataggio dati in RIMBORSI
    if(len(RimborsiRowData[0]) > 0):
        with pd.ExcelWriter(fileToWrite, engine ="openpyxl", mode='a', if_sheet_exists='overlay') as writer:
            for i in range(len(RimborsiRowData[0])-1, -1, -1):

                final_listRimborsi = []
                
                for k in range(len(df_RimborsiTutela)-1, -1, -1):
                    if(RimborsiRowData[1][i] == df_RimborsiTutela.iat[k, 0]):
                        # Converto tutte le date da formato 'datetime' a formato stringa "%d/%m/%Y"
                        # df_RimborsiTutela.values[k, 0] e' un oggetto di tipo pandas Timestamp
                        fromTimestampToDatetime = (df_RimborsiTutela.values[k, 0]).date()
                        # Converto la datetime.date in una datetime.datetime
                        fromTimestampToDatetime = datetime.datetime.combine(fromTimestampToDatetime, datetime.datetime.min.time())
                        datetimeString = convertDatetimeValueToString(fromTimestampToDatetime)
                        temp_listRimborsiTutela = df_RimborsiTutela.values[k, 0:8].tolist() # Conversione da pandas dataframe a list
                        temp_listRimborsiTutela[0] = datetimeString  # Assegno la data come stringa al 1° valore della list
                        final_listRimborsi.append(temp_listRimborsiTutela)

                final_dfRimborsi = pd.DataFrame(final_listRimborsi)
                final_dfRimborsi.to_excel(writer, index = False, header = False, sheet_name = sheetNameRimborsi, startrow = RimborsiRowData[0][i] + 1, startcol = 0)


    # Salvataggio TOTALE INCASSI e PROVVIGIONI TUTELA LEGALE nel foglio 'PRIMA NOTA'
    with pd.ExcelWriter(fileToWrite, engine ="openpyxl", mode='a', if_sheet_exists='overlay') as writer:
        for i in range(len(PrimaNotaRowData)-1, -1, -1):
            listIncassiProvvigioniTutela = [["Incassi TUTELA LEGALE", totale_IncassiProvvigioni[i][TOT_INCASSI]]]
            listIncassiProvvigioniTutela.append(["Provvigioni TUTELA LEGALE", totale_IncassiProvvigioni[i][TOT_PROVVIGIONI]])
            
            df_IncassiProvvigioniTutela = pd.DataFrame(listIncassiProvvigioniTutela)
            df_IncassiProvvigioniTutela.to_excel(writer, index = False, header = False, sheet_name = sheetNamePrimaNota, startrow = PrimaNotaRowData[i]+ROW_INDEX_TOT_TUTELA, startcol = 2)

            
    print("Copia dei dati del file ", fileName_Tutela, " di TUTELA LEGALE terminata.\n")
    
    #  Rinomino il file di cui ho appena salvato i dati con la desinenza '_checked'
    renameFileChecked(fileTutela_read)

    print("File ", fileName_Tutela, " rinominato con '_checked' come desinenza.\n")
            



#   -----------------------------------------------------------------------------------------
#   -----------------------------------------------------------------------------------------
#   -----------------------------------------------------------------------------------------
#   -----------------------------------------------------------------------------------------
# Funzione per trovare tutti i file di cui non sono ancora stati analizzati e salvati i dati
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


#   -----------------------------------------------------------------------------------------
#   -----------------------------------------------------------------------------------------
#   -----------------------------------------------------------------------------------------
#   -----------------------------------------------------------------------------------------
# Legge tutti i dati nel foglio SOSPESI e va a scrivere in tutte le tabelle del foglio PRIMA NOTA i relativi sospesi del giorno
# ATTENZIONE perche' per come e' fatto adesso va a fare questo lavoro per tutti i giorni, anche per quelli che erano gia' stati fatti in precedenza.
# Bisogna quindi ottimizzare il tutto per far eseguire questa funzione solamente per i giorni di cui non sono stati ancora scritti i SOSPESI NUOVI
def manageSospesi(fileToWrite, lastDatetime):
    sheetNameSospesi = "SOSPESI"

    # Caricamento dei dati dal foglio 'SOSPESI'
    dataSospesiExcel = pd.read_excel(fileToWrite, sheet_name = sheetNameSospesi, usecols='A,B,E,I,K')

    print("Lettura dati dal foglio 'SOSPESI' eseguita con successo.\n")

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

    # Inserisco 55 e non 60 per non saltare troppe righe per precauzione
    SIZEOF_SINGLE_TABLE_SOSPESI = int(55)

    defaultDate = datetime.date(2020, 1, 1)    # Data di default
    previousDate = defaultDate

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
                        if(totSospesiNew.date != defaultDate):    # Per non fare l'append quanto si ha totSospesiNew tutta a 0 all'inizio
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
                    if(dataSospesiExcel.iat[i, DATA] < todayDate and dataSospesiExcel.iat[i, DATA] <= lastDatetime):
                        indexRowExecuted.append(i+1)

        i += 1

    # Aggiungo l'ultimo record alla list dei NUOVI SOSPESI - bug: se non trova nulla va comunque a fare un append di dati vuoti per la data 01/01/2024 andando a scrivere tali dati nel foglio 'PRIMA NOTA'
    if(atleastOneDataFound == True):
        listSospesiNew.append(totSospesiNew)

    strExecuted = list(["Eseguito"])
    df_sospesiExecuted = pd.DataFrame(strExecuted)

    # Scrittura dei TOTALI NUOVI SOSPESI nel foglio 'PRIMA NOTA'
    writeSospesi_inPrimaNota(listSospesiNew, fileToWrite)

    if(len(indexRowExecuted) > 0):
        print("Numero di righe in cui scrivere la stringa 'Eseguito' = ", len(indexRowExecuted), ".\n")

        with pd.ExcelWriter(fileToWrite, engine ="openpyxl", mode='a', if_sheet_exists='overlay') as writer:
            for i in range(0, len(indexRowExecuted)):
                df_sospesiExecuted.to_excel(writer, index = False, header = False, sheet_name = sheetNameSospesi, startrow = indexRowExecuted[i], startcol = 10)    # 10 = 'J' -> NOTE
                # print("", end=f"\rNumero di righe mancanti in cui scrivere la stringa 'Eseguito': {len(indexRowExecuted) - i - 1}  ")

        print("Completata scrittura della stringa 'Eseguito' nel foglio 'SOSPESI' per tutte le righe.\n")
    else:
        print("Non vi e' nessuna riga nel foglio 'SOSPESI' in cui dover scrivere la stringa 'Eseguito'.\n")

