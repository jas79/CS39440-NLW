"""Reads data from Excel files, stores it in MySQL database"""
import os
import xlrd
import MySQLdb

class Ship:
    """Class for storing information on Ship"""
    name: str = "Unknown"
    key: int = 0
    port: str = "Unknown"

    def to_string(self) -> str:
        return ("Name: " + self.name + "\n"
                "ID: " + str(self.key) + "\n"
                "Port: " + self.port + "\n"
                "----------")
    def add(self):
        """Add Ship to Database"""
        if self.check_key() is None:
            C.execute("""INSERT INTO vessels (idvessels, name, port_of_registry) VALUES (%s, %s, %s)""",
                      (self.key, self.name, self.port))

    def fill_missing(self):
        print(self.to_string())
        if self.key == 0 and self.name is not 'Unknown':
            print(self.check_name())
            self.key = self.check_name()[0]

        elif self.name is 'Unknown' and self.key != 0:
            self.name = self.check_key()[1]

        if self.port is 'Unknown' and self.key != 0:
            self.port = self.check_key()[2]

        elif self.port is 'Unknown' and self.name is not 'Unknown':
            self.port = self.check_name()[2]

    def check_key(self):
        C.execute("""SELECT * FROM vessels WHERE idvessels = %s""", (self.key, ))
        result = C.fetchone()
        return result

    def check_name(self):
        C.execute("""SELECT * FROM vessels WHERE name = %s""", (self.name, ))
        result = C.fetchone()
        return result

def directory_search(folder):
    """Recursively searches directories to find files

    Key arguments:
    file -- Folder to search"""
    for result in os.scandir(folder):
        if result.is_file() and result.name.startswith("File"):
            file_reader(result)
        elif result.is_dir():
            directory_search(result)

def file_reader(file):
    """Reads Excel files, adds data to database

    Key arguments:
    file -- Excel file to read"""
    workbook = xlrd.open_workbook(file)
    for worksheet in workbook.sheets():
        try:
            ship = Ship()
            if isinstance(worksheet.cell_value(3, 5), int) or isinstance(worksheet.cell_value(3, 5), float):
                ship.key = int(worksheet.cell_value(3, 5))
                
            if worksheet.cell_value(1, 5) is not '':
                ship.name = str(worksheet.cell_value(1, 5))

            if worksheet.cell_value(5, 5) is not '':
                ship.port = str(worksheet.cell_value(5, 5))
        except IndexError:
            break
        if ship.name is 'Unknown' and ship.key == 0:
            ERROR_FILES.append(file)
        elif ship.name is 'Unknown' or ship.port is 'Unknown' or ship.key == 0:
            MISSING_DATA.append(ship)
        else:
            ship.add()

        print_database()

def print_database():
    C.execute("""SELECT * FROM vessels ORDER BY name ASC""")
    print(C.fetchall())

ERROR_FILES = []
MISSING_DATA = []
PATH = '../ABERSHIP_transcription_vtls004566921'
DB = MySQLdb.connect('db.dcs.aber.ac.uk', 'jas79', 'dragon00js', 'cs39440_17_18_jas79')
C = DB.cursor()

directory_search(PATH)

for ship in MISSING_DATA:
    print(ship.to_string())

for ship in MISSING_DATA:
    if ship.name is 'Unknown' or ship.port is 'Unknown' or ship.key == 0:
        ship.fill_missing()
        MISSING_DATA.remove(ship)
        ship.add()
        
if "broken_files.txt" not in os.listdir():
    WRITE_FILE = os.open("broken_files.txt", os.O_CREAT)
else:
    WRITE_FILE = os.open("broken_files.txt", os.O_APPEND)

for file in ERROR_FILES:
    WRITE_FILE.append(file.name)

print_database()

WRITE_FILE.close()
DB.close()
