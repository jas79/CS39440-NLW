import xlrd
import MySQLdb;

class Ship:
    def __init__(self, name, id, port):
        self.name=name;
        self.id=id;
        self.port=port;

    def add(self):
        

class Mariner:
    def __init__(self, name, year, age, place, address):
        self.name=name;
        self.year=year;
        self.age=age;
        self.place=place;
        self.address=address;

    def add(self):

class Service:
    def __init__(self, mariner, ship):
        self.mariner=mariner;
        self.ship=ship;
        self.capacity='';
        self.start=0;
        self.end=0;
        self.joined='';
        self.left='';
        self.reason='';
        self.signed=False;
        self.notes='';

    def add(self):

file = 'File_1-1_vtls004583057.xlsx'

workbook = xlrd.open_workbook(file);

for worksheet in workbook.sheets():
    ship = Ship();
    ship.add();
    for row in range(11, worksheet.nrows):
        if worksheet.cell_value(row,0) is '':
            break;
        else:
            mariner = Mariner();
            service = Service();
            mariner.add();
            service.add();
