import pandas as pd
import datetime

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
    

def writeSospesi_inPrimaNota(totSospesi, filePathnameToWrite, dateOfData):
    sheetNamePrimaNota = "PRIMA NOTA"
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

    rowData = 0

    listPrimaNota = [["SOSPESI RHO",        0.0,    "NUOVI SOSPESI RHO",        totSospesi.totRho          ],
                     ["SOSPESI GALLARATE",  0.0,    "NUOVI SOSPESI GALLARATE",  totSospesi.totGallarate    ],
                     ["SOSPESI LEGNANO",    0.0,    "NUOVI SOSPESI LEGNANO",    totSospesi.totLegnano      ],
                     ["SOSPESI SOMMA",      0.0,    "NUOVI SOSPESI SOMMA",      totSospesi.totSommaLombardo],
                     ["SOSPESI AGOS",       0.0,    "NUOVI SOSPESI AGOS",       totSospesi.totAgos         ],
                     ["SOSPESI TUTELA",     0.0,    "NUOVI SOSPESI TUTELA",     totSospesi.totTutelaLegale ]]

    # Creazione dataframe BONIFICI
    df_PrimaNota = pd.DataFrame(listPrimaNota)

    for i in range(0, len(dataExcel)):
        if(dataExcel.values[i][DATA] == dateOfData):
            # print(dataread.values[i])
            rowData = i+1
            break    

    print("Copia e salvataggio dati in esecuzione, attendere ...\n")
    # df_Generali.style.apply(lambda x: x.map(highlight_if_FinConsumo), axis=None)

    with pd.ExcelWriter(filePathnameToWrite, engine ="openpyxl", mode='a', if_sheet_exists='overlay') as writer:
        df_PrimaNota.to_excel(writer, index = False, header = False, sheet_name = sheetNamePrimaNota, startrow = rowData+37, startcol = 10)    # 10 = 'K'




            
            
            
