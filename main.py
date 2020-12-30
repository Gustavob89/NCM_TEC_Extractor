import openpyxl
import mariadb

def UploadDataToMariaDB(tableName, lineNumber, conn):
    # Uploads an NCM, its description and TEC value to a Mysql/MariaDB

    cursor = conn.cursor()

    cellA = "A" + str(lineNumber)  # Which cell will be read
    cellB = "B" + str(lineNumber)
    cellC = "C" + str(lineNumber)

    valueA = sheet[cellA].value  # The value at each cell is placed at a variable
    valueB = sheet[cellB].value
    valueC = sheet[cellC].value

    statement = "INSERT INTO " + tableName +" (NCM, Description, TEC) VALUES (%s, %s, %s)"
    stmTuple = (valueA, valueB, valueC)

    try:
        cursor.execute(statement, stmTuple)
        conn.commit()
    except mariadb.Error as e:
        print(f"Failed to upload data: {e}")
        input()
        exit()


def HasNCMData(cell):
    # Verifies if the value at a cell of a spreadsheet is a NCM (Nomenclatura Comum do Mercosul) value;
    # If it is not, return 0. Otherwise return 1

    currentValue = sheet[cell].value

    # Will verify common characters that indicate that the value is not a NCM
    if currentValue == None:
        return 0

    if str(currentValue).find("-") != -1:
        return 0

    if str(currentValue).find("a") != -1:
        return 0

    if str(currentValue).find("e") != -1:
        return 0

    if str(currentValue).find("i") != -1:
        return 0

    if str(currentValue).find("o") != -1:
        return 0

    if str(currentValue).find("u") != -1:
        return 0

    if str(currentValue).find("N") != -1:
        return 0

    if str(currentValue).find("T") != -1:
        return 0

    return 1  # Value appeard to be a NCM. Return 1


def ConnectToMariaDB(loginFileName):
    # Reads login data from txt file and connects to database;
    # returns conn variable if successful or 0 after fail

    loginInformationFile = open(loginFileName, "r")

    databaseIP = loginInformationFile.readline().rstrip()
    databasePort = loginInformationFile.readline().rstrip()
    databaseName = loginInformationFile.readline().rstrip()
    username0 = loginInformationFile.readline().rstrip()
    password0 = loginInformationFile.readline().rstrip()

    print("Attempting connection to database:", databaseName)
    print(" as:", username0)
    print(" at:", databaseIP + '/' + databasePort)

    try:
        conn = mariadb.connect(
            user=username0,
            password=password0,
            host=databaseIP,
            port=int(databasePort),
            database=databaseName
        )
    except mariadb.Error as e:
        print(f"Error connecting to MariaDB Platform: {e}")
        return 0

    return conn


def VerifyTable(tableName, conn):
    #Verifies if the desiredtable exists and creates it if required

    cursor = conn.cursor()

    statement ="CREATE TABLE IF NOT EXISTS "+ tableName +" (NCM VARCHAR(50), Description VARCHAR(500), TEC VARCHAR(50))"
    cursor.execute(statement)


if __name__ == '__main__':
    # Script utilized to read NCM (Nomenclature Comum do Mercosul)  and TEC (Tarifa Externa Comum do Mercosul)
    # information from an Excel spreadsheet provided by Mercosul and upload the data a Mysql/MariaDB

    # Connect to Mariadb:
    connMariadb = ConnectToMariaDB("LoginInformation.txt")

    if connMariadb == 0:
        print("Errror")
        input()
        exit()

    print("Connected")

    # Verify if the required table exists and create it if it does not
    VerifyTable("ncm_tec_20200707", connMariadb)

    # Open the spreadsheet file. It has to be previously converted from '.xlmx' to '.xlsm'
    wb = openpyxl.load_workbook("TEC_20200707.xlsm")
    sheet = wb["TEC"]# Select the sheet

    lineMin = 504  # First line of the table to be verified
    lineMax = 19268  # Final line of the table to be verified

    for lineCounter in range(lineMin, lineMax):
        # For each line of the table, verify if the term in the "A" column is an NCM value;
        # If it is, upload its data to the concentrated database

        currentCell = "A" + str(lineCounter)

        if HasNCMData(currentCell) == 1:
            UploadDataToMariaDB("ncm_tec_20200707", lineCounter, connMariadb)

    print("Script complete")
    input()