import pandas as pd
import datetime

dateFormat = "%d/%m/%Y"

subagentAgency =   [["AMICONE LUIGI",             "GALLARATE"], 
                    ["AMICONE RENZO",             "GALLARATE"],
                    ["LOSCHI SANDRO",             "GALLARATE"],
                    ["CATTANEO SILVIO",           "GALLARATE"],
                    ["ORLANDI MARZIA",            "GALLARATE"],
                    ["AMICONE DEBORAH",           "GALLARATE"],
                    ["MARCHESOLI FRANCESCO",      "RHO"],
                    ["MAZZOCCHI DORIANA",         "RHO"],
                    ["DI VITO VINCENZO",          "RHO"],
                    ["TENCONI GABRIELA",          "LEGNANO"],
                    ["TENCONI FRANCESCA",         "LEGNANO"],
                    ["TALLARINI IVANA",           "SOMMA LOMBARDO"],
                    ["TALLARINI VITTORIO",        "SOMMA LOMBARDO"],
                    ["DE SILVESTRI SERENELLA",    "SOMMA LOMBARDO"],
                    ["SPAGNUOLO STEFANIA",        "SOMMA LOMBARDO"]]

def findAgencyFromSubagent(subagentName):
    NOME = int(0)
    AGENZIA = int(1)    

    for i in range(0, len(subagentAgency)):
        if(isinstance(subagentName, str) == False): # Se non e' una stringa, allora quel campo e' vuoto
            return "Collaboratore non trovato"
        if(subagentName.upper().find(subagentAgency[i][NOME]) != -1):
            break

    return subagentAgency[i][AGENZIA]


def updateAgencyTotaleSospesi(totSospesi, importo, agenzia):
    if(agenzia == "GALLARATE"):
        totSospesi.totGallarate += importo
    elif(agenzia == "RHO"):
        totSospesi.totRho += importo
    elif(agenzia == "LEGNANO"):
        totSospesi.totLegnano += importo
    elif(agenzia == "SOMMA LOMBARDO"):
        totSospesi.totSommaLombardo += importo
    elif(agenzia == "AGOS"):
        totSospesi.totAgos += importo
    elif(agenzia == "TUTELA LEGALE"):
        totSospesi.totTutelaLegale += importo
    else:
        # raise Exception("\nERRORE: agenzia a cui assegnare l'importo del versamento non trovata.\n")
        print("\nAgenzia non trovata.\n")
    

def writeSospesi_inPrimaNota(totSospesi, filePathnameToWrite, dateFromSospesi):
    sheetNamePrimaNota = "PRIMA NOTA"

    # Caricamento di tutti i dati relativi alla colonna 'A' dal foglio 'PRIMA NOTA' -> mi serve per avere tutte le date
    dataExcel = pd.read_excel(filePathnameToWrite, sheet_name = sheetNamePrimaNota, usecols='A')

    print("Lettura sheet PRIMA NOTA eseguita con successo.\n")

    # A -> 0 : DATA
    # D -> 1 : CASSA ENTRATE
    # E -> 2 : CASSA USCITE
    # L -> 3 : TOTALE SOSPESI VECCHI
    # N -> 4 : TOTALE SOSPESI NUOVI

    DATA = int(0)
    CASSA_ENTRATE = int(1)
    CASSA_USCITE = int(2)
    TOT_SOSPESI_OLD = int(3)
    TOT_SOSPESI_NEW = int(4)

    convertStringToDatetime(dataExcel, DATA)

    rowData = 0

    listPrimaNota = [["SOSPESI RHO",        "NUOVI SOSPESI RHO",        float(totSospesi.totRho)          ],
                     ["SOSPESI GALLARATE",  "NUOVI SOSPESI GALLARATE",  float(totSospesi.totGallarate)    ],
                     ["SOSPESI LEGNANO",    "NUOVI SOSPESI LEGNANO",    float(totSospesi.totLegnano)      ],
                     ["SOSPESI SOMMA",      "NUOVI SOSPESI SOMMA",      float(totSospesi.totSommaLombardo)],
                     ["SOSPESI AGOS",       "NUOVI SOSPESI AGOS",       float(totSospesi.totAgos)         ],
                     ["SOSPESI TUTELA",     "NUOVI SOSPESI TUTELA",     float(totSospesi.totTutelaLegale) ]]

    # Creazione dataframe PRIMA NOTA
    df_PrimaNota = pd.DataFrame(listPrimaNota)

    for i in range(0, len(dataExcel)):
        if(dataExcel.values[i][DATA] == 'DATA'):
            try:
                # Controllo se il dato presente nel file e' di tipo datetime, oppure date, oppure se e' una stringa che ha lo stesso formato di una data
                if(isinstance(dataExcel.values[i+1][DATA], datetime.datetime) or isinstance(dataExcel.values[i+1][DATA], datetime.date) or isinstance(datetime.datetime.strptime(dataExcel.values[i+1][DATA], "%d/%m/%Y"), datetime.datetime)):
                    if(isinstance(dataExcel.values[i+1][DATA], str) and isinstance(datetime.datetime.strptime(dataExcel.values[i+1][DATA], "%d/%m/%Y"), datetime.datetime)):
                        dateFromPrimaNota = datetime.datetime.strptime(dataExcel.values[i+1][DATA], "%d/%m/%Y")
                    else:
                        dateFromPrimaNota = dataExcel.values[i+1][DATA]

                    # ATTENZIONE: dateFromPrimaNota e dateFromSospesi devono essere entrambe del tipo datetime.datetime, altrimenti il confronto fallisce
                    # Potrei aggiungere un if con else in errore se type(dateFromPrimaNota) != type(dateFromSospesi)
                    if(dateFromPrimaNota == dateFromSospesi):
                        rowData = i+2   # i -> 'DATA'       i+1 -> es. '01/01/2024'     Devo quindi aggiungere un altro + 1 -> i+2
                        break
                else:
                    # In realta' non funziona ma viene dato il seguente errore: "Error:  time data ' ' does not match format '%d/%m/%Y' ""
                    raise Exception("\nManca la data alla riga ", i+1, " del foglio 'PRIMA NOTA'. Inserire la data mancante ed eseguire nuovamente l'applicazione.\n")  
            except:
                print("ERROR")


    print("Copia e salvataggio dati nel foglio 'PRIMA NOTA' della data ", dateFromSospesi.date(), " in esecuzione, attendere ...\n")
    # df_Generali.style.apply(lambda x: x.map(highlight_if_FinConsumo), axis=None)

    with pd.ExcelWriter(filePathnameToWrite, engine ="openpyxl", mode='a', if_sheet_exists='overlay') as writer:
        df_PrimaNota.to_excel(writer, index = False, columns = [10, 12, 13], header = False, sheet_name = sheetNamePrimaNota, startrow = rowData+37, startcol = 10)    # 10 = 'K'



# Funzione utilizzata per convertire da stringa a 'datetime' il contenuto di una list di input
def convertStringToDatetime(listToConvert, DATE_INDEX):
    for i in range(0, len(listToConvert)):
        # checking if format matches the date
        
        res = True
        
        # using try-except to check for truth value
        try:
            if(isinstance(listToConvert.values[i][DATE_INDEX], str)):
                res = bool(datetime.datetime.strptime(listToConvert.values[i][DATE_INDEX], dateFormat))
                if(res == True):
                    listToConvert.values[i][DATE_INDEX] = datetime.datetime.strptime(listToConvert.values[i][DATE_INDEX], dateFormat)
        except ValueError:
            res = False



        # if(isinstance(listToConvert[i], str) and listToConvert[i] != 'DATA' and listToConvert[i] != 'TOTALE'):
        #     listToConvert[i] = datetime.datetime.strptime(listToConvert[i], "%d/%m/%Y")


# Funzione per convertire le date di una list da 'datetime' a stringa formato '%d/%m/%Y'
def convertDatetimeToString(datetimeToConvert, DATE_INDEX):
    for i in range(0, len(datetimeToConvert)):
        try:
            if(isinstance(datetimeToConvert.values[i][DATE_INDEX], datetime.datetime)):
                datetimeToConvert.values[i][DATE_INDEX] = datetimeToConvert.values[i][DATE_INDEX].strftime(dateFormat)
        except ValueError:
            print("\nErrore in convertDatetimeToString.\n")


def convertDatetimeValueToString(datetimeToConvert):
        try:
            if(isinstance(datetimeToConvert, datetime.datetime)):
                dateString = datetimeToConvert.strftime(dateFormat)

                return dateString
        except ValueError:
            print("\nErrore in convertDatetimeToString.\n")

            
