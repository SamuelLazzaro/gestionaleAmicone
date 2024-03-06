import pandas as pd

# from classDefinition import SubagentAgency

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
        if(subagentName.upper().find(subagentAgency[i][NOME]) != -1):
            break

    return subagentAgency[i][AGENZIA]


def updateAgencyTotaleSospesi(totSospesiNuovi, importo, agenzia):
    if(agenzia == "GALLARATE"):
        totSospesiNuovi.totGallarate += importo
    elif(agenzia == "RHO"):
        totSospesiNuovi.totRho += importo
    elif(agenzia == "LEGNANO"):
        totSospesiNuovi.totLegnano += importo
    elif(agenzia == "SOMMA LOMBARDO"):
        totSospesiNuovi.totSommaLombardo += importo
    elif(agenzia == "AGOS"):
        totSospesiNuovi.totAgos += importo
    elif(agenzia == "TUTELA LEGALE"):
        totSospesiNuovi.totTutelaLegale += importo
    else:
        raise Exception("\nERRORE: agenzia a cui assegnare l'importo del versamento non trovata.\n")
    

def writeSospesi_inPrimaNota(totSospesi, filePathnameToWrite, dateOfData):
    sheetNamePrimaNota = "PRIMA NOTA"
    dataExcel = pd.read_excel(filePathnameToWrite, sheet_name = sheetNamePrimaNota, usecols='A,D,E,L,N')

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

    listPrimaNota = [["SOSPESI RHO",        totSospesi.totRho,              "NUOVI SOSPESI RHO",        0.0],
                     ["SOSPESI GALLARATE",  totSospesi.totGallarate,        "NUOVI SOSPESI GALLARATE",  0.0],
                     ["SOSPESI LEGNANO",    totSospesi.totLegnano,          "NUOVI SOSPESI LEGNANO",    0.0],
                     ["SOSPESI SOMMA",      totSospesi.totSommaLombardo,    "NUOVI SOSPESI SOMMA",      0.0],
                     ["SOSPESI AGOS",       totSospesi.totAgos,             "NUOVI SOSPESI AGOS",       0.0],
                     ["SOSPESI TUTELA",     totSospesi.totTutelaLegale,     "NUOVI SOSPESI TUTELA",     0.0]]

    # Creazione dataframe BONIFICI
    df_PrimaNota = pd.DataFrame(listPrimaNota)

    for i in range(0, len(dataExcel)):
        if(dataExcel.iat[i][DATA] == dateOfData):
            # print(dataread.values[i])
            rowData = i+1
            break    

    print("Copia e salvataggio dati in esecuzione, attendere ...\n")
    # df_Generali.style.apply(lambda x: x.map(highlight_if_FinConsumo), axis=None)

    with pd.ExcelWriter(filePathnameToWrite, engine ="openpyxl", mode='a', if_sheet_exists='overlay') as writer:
        df_PrimaNota.to_excel(writer, index = False, header = False, sheet_name = sheetNamePrimaNota, startrow = rowData+37, startcol = 10)    # 10 = 'K'
