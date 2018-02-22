import os;
import xlrd;
import MySQLdb;

class Ship:
    def __init__(self, name, id, port):
        self.name=name;
        self.id=id;
        self.port=port;

    def add(self):
        c.execute("""INSERT INTO vessels (idvessels, name, port_of_registry) VALUES (%s, %s, %s)""",
                  (self.id,
                   self.name,
                   self.port));

    def check_database(self):
        c.execute("""SELECT idvessels FROM vessels WHERE idvessels = %s""", (self.id,));
        result = c.fetchone();
        if result is None:
            return False
        else:
            return True;

    def to_string(self):
        print(self.name);
        print(self.id);
        print(self.port);

class Mariner:
    def __init__(self, name, year, age, place, address):
        self.name=name;
        self.year=year;
        self.age=age;
        self.place=place;
        self.address=address;

    def add(self):
        print(self.name);

class Service:
    def __init__(self, mariner, ship, capacity, start, end, joined, left, reason, signed, notes):
        self.mariner=mariner;
        self.ship=ship;
        self.capacity=capacity;
        self.start=start;
        self.end=end;
        self.joined=joined;
        self.left=left;
        self.reason=reason;
        self.signed=signed;
        self.notes=notes;

    def add(self):
        print(self.start + ' ' + self.end);

def directory_search(folder):
    for result in os.scandir(folder):
        if result.is_file() and result.name.startswith("File"):
            print(result);
            file_converter(result);
        elif result.is_dir():
            directory_search(result);

def file_converter(file):
    workbook = xlrd.open_workbook(file);
    for worksheet in workbook.sheets():
        try:
            ship = Ship(worksheet.cell_value(1, 5),
                        int(worksheet.cell_value(3, 5)),
                        worksheet.cell_value(5, 5),);
        except ValueError:
            print(file);
            print(ship);
        if not ship.check_database():
            ship.add();
        for row in range(11, worksheet.nrows):
            if worksheet.cell_value(row,0) is '':
                break;
            else:
                mariner = Mariner(worksheet.cell_value(row,0),
                                  worksheet.cell_value(row,1),
                                  worksheet.cell_value(row,2),
                                  worksheet.cell_value(row,3),
                                  worksheet.cell_value(row,4));

path = '../ABERSHIP_transcription_vtls004566921'

db = MySQLdb.connect('db.dcs.aber.ac.uk', 'jas79', 'dragon00js', 'cs39440_17_18_jas79');
c = db.cursor();

directory_search(path);
db.commit();
