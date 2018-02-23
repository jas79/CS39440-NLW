"""Takes data from transcribed Excel files, manipulates them, places them into the database"""
import os
import xlrd
import MySQLdb

class Ship:
    def __init__(self, name="Unknown", key=0, port="Unknown"):
        """Defines the Ship object"""
        self.name = name
        self.key = key
        self.port = port







        

    def add(self):
        """Add Ship to Database"""
        C.execute("""INSERT INTO vessels (idvessels, name, port_of_registry) VALUES (%s, %s, %s)""",
                  (self.key, self.name, self.port),)

    def check_database(self):
        """Checks database for existing Ships"""
        C.execute("""SELECT idvessels FROM vessels WHERE idvessels = %s""", (self.key, ))
        result = C.fetchone()
        return result is not None

    def to_string(self):
        print("Name: "+ self.name)
        print("ID: ", self.key)
        print("Port: "+self.port)
        print("----------")

class Mariner:
    def __init__(self, name, year, age, place, address):
        self.name = name
        self.year = year
        self.age = age
        self.place = place
        self.address = address

    def add(self):
        print(self.name)

class Service:
    def __init__(self, mariner, ship, capacity, start, end, joined, left, reason, signed, notes):
        self.mariner = mariner
        self.ship = ship
        self.capacity = capacity
        self.start = start
        self.end = end
        self.joined = joined
        self.left = left
        self.reason = reason
        self.signed = signed
        self.notes = notes

    def add(self):
        print(self.start + ' ' + self.end)

def directory_search(folder):
    for result in os.scandir(folder):
        if result.is_file() and result.name.startswith("File"):
            file_converter(result)
        elif result.is_dir():
            directory_search(result)

def file_converter(file):
    workbook = xlrd.open_workbook(file)
    for worksheet in workbook.sheets():
        try:
            ship = Ship(worksheet.cell_value(1, 5), int(worksheet.cell_value(3, 5)), worksheet.cell_value(5, 5))
        except ValueError:
            ship = Ship()
            print("Name: "+ worksheet.cell_value(1, 5))
            if worksheet.cell_value(1, 5) is not '':
                ship.name = worksheet.cell_value(1, 5)
            print("ID: ", worksheet.cell_value(3, 5))
            if worksheet.cell_value(1, 5) is not '' and not isinstance(worksheet.cell_value(3, 5), int):
                ship.key = worksheet.cell_value(3, 5)
            print("Port: " + worksheet.cell_value(5, 5))
            if worksheet.cell_value(5, 5) is not '':
                ship.port = worksheet.cell_value(5, 5)
            ship.to_string()
            ERROR_FILES.append(file)
        except IndexError:
            print(worksheet.name)
            ERROR_FILES.append(file)
        if not ship.check_database():
            ship.add()
        for row in range(11, worksheet.nrows):
            if worksheet.cell_value(row, 0) is '':
                break
            else:
                mariner = Mariner(worksheet.cell_value(row, 0),
                                  worksheet.cell_value(row, 1),
                                  worksheet.cell_value(row, 2),
                                  worksheet.cell_value(row, 3),
                                  worksheet.cell_value(row, 4))

ERROR_FILES = []

PATH = '../ABERSHIP_transcription_vtls004566921'

DB = MySQLdb.connect('db.dcs.aber.ac.uk', 'jas79', 'dragon00js', 'cs39440_17_18_jas79')
C = DB.cursor()
print("the code is running")
if "broken_files.txt" not in os.listdir():
    WRITE_FILE = os.open("broken_files.txt", os.O_CREAT)
else:
    WRITE_FILE = os.open("broken_files.txt", os.O_APPEND)

directory_search(PATH)

WRITE_FILE.append(ERROR_FILES)
WRITE_FILE.close()
DB.commit()
DB.close()
