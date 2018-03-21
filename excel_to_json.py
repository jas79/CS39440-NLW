"""Spreadsheet => JSON"""
## import datetime
import json
import os
import xlrd

NODES = []
LINKS = []
D3_OBJECT = {
    "nodes": NODES,
    "links": LINKS
    }

PATH = str("C:\\Users\\Jake\\OneDrive - Aberystwyth University\\"
           "CS39440 Major Project\\ABERSHIP_transcription_vtls004566921")
JSON_FILE = open("objects.json", 'r+')

MARINER_ID = 1

SHIP = {}
MARINER = {}
LINK = {}

def main() -> (None):
    D3_OBJECT = json.load(JSON_FILE)
    directory_search(PATH)
    json.dump(D3_OBJECT, JSON_FILE)
    return

def ship_init() -> (SHIP):
    """Initialises / resets ship"""
    result = {
        "id": 0,
        "name": "Unknown",
        "port": "Unknown",
        "group": 1
    }
    return result

def link_init(source, target) -> (LINK):
    """Initialises / resets link"""
    result = {
        "source": source["id"],
        "target": target["id"],
    }
    return result

def mariner_init() -> (MARINER):
    """Initialises / resets mariner"""
    result = {
        "id": 0,
        "name": "Unknown",
        "year": 0,
        "age": 0,
        "place_of_birth": "Unknown",
        "home_address": "Unknown",
        "group": 2
    }
    return result

def directory_search(path: str) -> (None):
    """Recursively searches folders to find files
    key arguments:
    path -- Folder to search"""
    for result in os.scandir(path):
        if result.is_file() and result.name.startswith("File"):
            excel_reader(result)
        elif result.is_dir():
            directory_search(result)
    return

def excel_reader(file: str) -> (None):
    """Reads file, adds to dictionary
    key arguments:
    file -- File to read"""
    print(str(file))
    workbook = xlrd.open_workbook(file)
    for worksheet in workbook.sheets():
        ship = ship_init()
        try:
            ship["name"] = str(worksheet.cell_value(1, 5).strip().title())
        except AttributeError as error:
            result = str(worksheet.cell_value(1, 5))
            ship["name"] = result
        try:
            ship["port"] = str(worksheet.cell_value(5, 5).strip().title())
        except AttributeError as error:
            result = str(worksheet.cell_value(5, 5))
            ship["port"] = result
        try:
            ship["id"] = int(worksheet.cell_value(3, 5))
        except ValueError as error:
            ignore()
        if is_unique(ship):
            NODES.append(ship)
            ##print(str(ship))
        for row in range(11, worksheet.nrows):
            if worksheet.cell_value(row, 0) == '':
                break
            else:
                global MARINER_ID
                mariner = mariner_init()
                mariner["name"] = check_blk(worksheet.cell_value(row, 0))
                try:
                    mariner["year"] = int(worksheet.cell_value(row, 1))
                except ValueError as error:
                    ignore()
                try:
                    mariner["age"] = int(worksheet.cell_value(row, 2))
                except ValueError as error:
                    ignore()
                mariner["place_of_birth"] = check_blk(worksheet.cell_value(row, 3))
                mariner["home_address"] = check_blk(worksheet.cell_value(row, 4))
                if is_unique(mariner):
                    mariner["id"] = MARINER_ID
                    NODES.append(mariner)
                    ##print(str(mariner))
                    MARINER_ID += 1
                link = link_init(mariner, ship)
                link = link_add_weight(link)
                LINKS.append(link)
                ##print(str(LINKS))
    return

def check_blk(string: str) -> (str):
    try:
        if string.upper() == "blk".upper():
            return "Unknown"
        else:
            return string.strip().title()
    except AttributeError as error:
        return 0

def ignore() -> (None):
    return None

def link_add_weight(item) -> (LINK):
    for result in LINKS:
        if (item["source"] == result["source"]) and (item["target"] == result["target"]):
            if result["weight"] >= 1:
                item["weight"] = result["weight"] + 1
                return item
    item["weight"] = 1
    return item

def is_unique(item) -> (bool):
    """Returns True if object is unique
    (i.e. does not already exist in list)
    key arguments:
    dict -- object to compare (ship or mariner)
    """
    if item["group"] == 1:
        for result in NODES:
            if item["id"] == result["id"]:
                return False
            elif item["id"] == 0:
                if item["name"] == result["name"]:
                    return False
    else:
        try:
            for result in NODES:
                if ((item["group"] == 2) and
                    (item["name"] == result["name"]) and 
                    (item["year"] == result["year"]) and
                    (item["place_of_birth"] == result["place_of_birth"])):
                    return False
        except KeyError as error:
            return False
    return True

main()
