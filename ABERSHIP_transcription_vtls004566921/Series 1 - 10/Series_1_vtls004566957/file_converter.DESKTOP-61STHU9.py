import xlrd;
import mysql.connector;

#conn = mysql.connector.connect(user='jas79',
#                               password='dragon00js',
#                               host='db.dcs.aber.ac.uk',
#                               database='cs39440_17_18_jas79');

class Ship:
    def __init__(self, name, id, port):
        self.name=name;
        self.id=id;
        self.port=port;

class Mariner:
    def __init__(self, name, year, age, place, address):
        self.id=name;
        self.name=name;
        self.year=year;
        self.age=age;
        self.place=place;
        self.address=address;

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

def in_list(item, items):
    for result in items:
        if item.id == result.id:
            return True;
        else:
            return False;

file = 'File_1-1_vtls004583057.xlsx'
add_ship = ("INSERT INTO vessels "
            "(idvessels, name, port_of_registry)"
            "VALUES (%d, %s, %s)");

workbook = xlrd.open_workbook(file);

ships =[];
mariners = [];

for worksheet in workbook.sheets():
    ship = Ship(worksheet.cell_value(1,5), worksheet.cell_value(3,5), worksheet.cell_value(5,5));
    if not in_list(ship, ships):
        ships.append(ship);
    for row in range(11, worksheet.nrows):
        if worksheet.cell_value(row,0) is '':
            break;
        else:
            mariner = Mariner(worksheet.cell_value(row, 0),worksheet.cell_value(row, 1),worksheet.cell_value(row, 2),worksheet.cell_value(row, 3),worksheet.cell_value(row, 4));
            if not in_list(mariner, mariners):
                mariners.append(mariner);
            service = Service(mariner, ship);

for result in ships:
    print(result.name);

for result in mariners:
    print(result.name);
