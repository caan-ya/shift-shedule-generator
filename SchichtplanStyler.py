import openpyxl as pyx
from openpyxl.styles import Alignment
import string
from SchichtplanUtils import getLettersForRef

def styleSheet(file_name, title_names, sheet_names):
    sheet_name = "Zusammenfassung"
    schichtplan_wb = pyx.load_workbook(file_name)
    schichtplan_sh = schichtplan_wb[sheet_name]

    #Spalten der KW horizontal und vertikal zentrieren
    for col in schichtplan_sh['B:' + getLettersForRef(sheet_names)]:
        for cell in col:
            try:
                cell.alignment = Alignment(horizontal='center', vertical='center')
            except:
                pass

    for idx, col in enumerate(schichtplan_sh.columns):
        max_length = 0
        column = col[0].column_letter # name der spalte
        for cell in col:
            cell.row
            try: # damit keine fehler bei leeren zellen kommen
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        if idx > 0 and idx < len(sheet_names):
            adjusted_width = (max_length + 2) * 1.5
            schichtplan_sh.column_dimensions[column].width = adjusted_width
        if idx == 0 or idx == len(sheet_names):
            adjusted_width = (max_length + 2) * 1.6
            schichtplan_sh.column_dimensions[column].width = adjusted_width


    sheet_names.pop()
    for cell in schichtplan_sh['A']:
        if cell.value in title_names:
            schichtplan_sh.merge_cells("B" + str(cell.row) + ":" + getLettersForRef(sheet_names) + str(cell.row))

    schichtplan_wb.save(file_name)