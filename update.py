# Auto update script for USTChorus birthday database. It will read from and write to `birthday-table.ods`.
# Before updating, paste the participants table into `input/participants-paste.txt`.
# It is intended for the Mixed and Men's choir setup, and can DEFINITELY NOT handle duplicate names.
# Remember to re-filter after table update if a filter is applied.

import os
from zipfile import ZipFile
from odf import opendocument
import odf.element, odf.table, odf.text
from typing import Optional
import shutil
import time

FILE_ODS = 'birthday-table.ods'
FILE_BACKUP = 'backup/birthday-table_%s.ods'
FILE_PARTICIPANTS = 'input/participants-paste.txt'

class BirthdayTable:
    def __init__(self, filename: str):
        self.xml = opendocument.load(filename)
        self.table: odf.element.Element = self.xml.body.getElementsByType(odf.table.Table)[0]
        
    def write(self, filename: str):
        self.xml.save(filename)
    
    @staticmethod
    def access_row(row: odf.element.Element, index: int):
        cells: list[odf.element.Element] = row.getElementsByType(odf.table.TableCell)
        # if len(cells) <= index:
        #     return ''
        # cell: odf.element.Element = cells[index]
        for cell in cells:
            repeat = int(cell.getAttrNS(cell.get_knownns('table'), 'number-columns-repeated') or '1')
            index -= repeat
            if index < 0:
                return str(cell).strip()
        return ''
    
    def get_row_list(self):
        return self.table.getElementsByType(odf.table.TableRow)[1:]
    
    def find_row_of_name(self, name: str):
        rows = self.get_row_list()
        for row in rows:
            if self.access_row(row, 0) == name:
                return row
        return None

    def remove_row(self, row: odf.element.Element):
        if row.parentNode:
            row.parentNode.childNodes.remove(row)
        else:
            raise ValueError('Row has no parent')
    
    @staticmethod
    def get_row_data(row: odf.element.Element):
        return tuple(
            BirthdayTable.access_row(row, i)
            for i in range(0, 4)
        )

    @staticmethod
    def set_row_data(row: odf.element.Element, data: tuple[str, str, str, str]):
        row.childNodes.clear()
        # (name, birthday, mixedPart, menPart)
        for i in range(0, 4):
            text_wrap = odf.text.P(text=data[i])
            table_cell = odf.table.TableCell()
            table_cell.setAttrNS('office', 'value-type', 'string')
            table_cell.setAttrNS('ns42', 'value-type', 'string')
            if i == 0:
                table_cell.setAttrNS('table', 'style-name', 'ce2')
            elif i == 1:
                table_cell.setAttrNS('table', 'style-name', 'ce7')
                table_cell.setAttrNS('table', 'content-validation-name', 'val1')
            table_cell.appendChild(text_wrap)
            row.appendChild(table_cell)
        
        # (..., isAlive)
        text_wrap = odf.text.P(text=('TRUE' if data[2] != '' or data[3] != '' else 'FALSE'))
        table_cell = odf.table.TableCell()
        table_cell.setAttrNS('office', 'value-type', 'boolean')
        table_cell.setAttrNS('ns42', 'value-type', 'boolean')
        table_cell.appendChild(text_wrap)
        row.appendChild(table_cell)
        
        table_cell = odf.table.TableCell()
        table_cell.setAttrNS('table', 'number-columns-repeated', '16378')
        row.appendChild(table_cell)
        
    def new_row(self) -> odf.element.Element:
        table_row = odf.table.TableRow()
        table_row.setAttrNS('table', 'style-name', 'ro1')
        rows: list[odf.element.Element] = self.table.childNodes
        
        i = 0
        for i in range(len(rows)):
            row = rows[i]
            if row.tagName == 'table:table-row' and self.access_row(row, 0) == '':
                break
        
        rows.insert(i, table_row)
        
        return table_row

def make_table_backup(filename: str, backup_pattern: str):
    backup_file = backup_pattern % (time.strftime("%Y-%m-%d_%H-%M-%S"), )
    os.makedirs(os.path.dirname(backup_file), exist_ok=True)
    shutil.copy(filename, backup_file)

# data = BirthdayTable('birthday-table.ods')
# row = data.new_row()
# data.set_row_data(row, ('Test', '1145', '', ''))
# row = data.new_row()
# data.set_row_data(row, ('Test2', '1145', '', ''))
# row = data.new_row()
# data.set_row_data(row, ('Test3', '1145', '', ''))
# row = data.new_row()
# data.set_row_data(row, ('Test4', '1145', '', ''))
# make_table_backup()
# data.write('birthday-table.ods')

def get_participant_parts(filename: str):
    paste_file_lines = open(filename, 'r', encoding='utf-8').read().split("\n")
    grid = [ s.split("\t") for s in paste_file_lines ]
    num_rows = len(grid)
    num_cols = max(0, *[ len(row) for row in grid ])
    
    def get_from_grid(row_index: int, col_index: int):
        if row_index >= len(grid):
            return ''
        row = grid[row_index]
        if col_index >= len(row):
            return ''
        return row[col_index]

    def is_part_name(text: str):
        if len(text) < 1:
            return False
        return ord('A') <= ord(text[0]) <= ord('Z')

    parts_map: dict[str, tuple[str, str]] = {}
    
    current_group = -1  # -1 = Unknown, 0 = Mixed, 1 = Men
    for col_index in range(0, num_cols):
        current_part = 'N'
        
        for row_index in range(0, num_rows):
            text = get_from_grid(row_index, col_index).strip()
            if text == '': continue
            if is_part_name(text):
                if current_group == -1:
                    current_group = 0
                current_part = text[0]
            else:
                if current_part == 'N':
                    print(f'WARN participant {text} is in unknown part.')
                part_tuple = list(parts_map.get(text, ('', '')))
                if part_tuple[current_group] != '':  # In case if someone have multiple parts
                    part_tuple[current_group] += '/'
                part_tuple[current_group] = current_part
                parts_map[text] = (part_tuple[0], part_tuple[1])
        
        if current_part == 'N' and current_group == 0:
            # We are crossing the boundary between Mixed and Men's choir
            current_group = 1

    return parts_map

print(f'Reading from {FILE_PARTICIPANTS}')
parts_map = get_participant_parts(FILE_PARTICIPANTS)

data = BirthdayTable(FILE_ODS)

for row in data.get_row_list():
    info = list(data.get_row_data(row))
    if info[0] != '':
        # Clear parts data, but always preserve C (conductor)
        if info[2].strip() != 'C': info[2] = ''
        if info[3].strip() != 'C': info[3] = ''
        data.set_row_data(row, (info[0], info[1], info[2], info[3]))

for name in parts_map:
    row = data.find_row_of_name(name)
    if row == None:
        row = data.new_row()
    info = list(data.get_row_data(row))
    info[0] = name
    
    parts = parts_map[name]

    if info[1] == '':
        print(f'WARN alive participant {name} has no birthday info')

    info[2] = parts[0]
    info[3] = parts[1]
    data.set_row_data(row, (info[0], info[1], info[2], info[3]))

for row in data.get_row_list():
    info = data.get_row_data(row)
    if info[0] != '' and info[1] == '' and info[2] == '' and info[3] == '':
        print(f'INFO removed {info[0]} because this entry has no birthday info and is not alive')
        data.remove_row(row)

print(f'Making backup of the original table')
make_table_backup(FILE_ODS, FILE_BACKUP)
print(f'Writing results')
data.write(FILE_ODS)
