import pandas as pd
import datetime
import os

try:

    current_working_directory = os.getcwd()
    fileToWrite = current_working_directory + '\\testFile.xlsx'

    date = datetime.date(2024, 3, 7)

    print("Data = ", date)

    listDate = list([date])
    
    df = pd.DataFrame(listDate)

    with pd.ExcelWriter(fileToWrite, engine ="openpyxl", mode='a', if_sheet_exists='overlay') as writer:
        df.to_excel(writer, index = False, header = False, startrow = 0, startcol = 0)


except Exception as e:
    print("Error: ", e)
    input()