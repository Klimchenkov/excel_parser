import sys

import sqlite3 as sl
from openpyxl import load_workbook

class ReadExcel:
    
    def __init__(self, path_to_file:str):
        
        self.path_to_file = path_to_file
        self.con = sl.connect('excel_parser.db')
        self._create_db()
        self._parse_file()
        
    def _create_db(self):
        with self.con as con:
            sql = """
            CREATE TABLE IF NOT EXISTS parsed_data (
                id INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT,
                value INTEGER NOT NULL,
                date STRING,
                type STRING,
                company STRING,
                timing STRING
            );
            """
            con.execute(sql)
    
    def _parse_file(self):
        print(self.path_to_file)
        
        



if __name__ == '__main__':
    ReadExcel(sys.argv[1])