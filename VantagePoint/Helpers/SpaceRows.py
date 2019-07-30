import openpyxl
import re


# noinspection PyBroadException
from Helpers.Globals import HEADERS


# noinspection PyBroadException
def space_rows(excel, lists, groups):
    wb = openpyxl.load_workbook(excel)
    sheet_names = wb.sheetnames
    sheets = []

    # Make a list of active sheets. :)
    for sheet in wb.worksheets:
        sheets.append(sheet)
    # Remove the first sheet since we don't need to put spaces in something that isnt sorted
    sheets.pop(0)

    idx = -1

    for sheet in sheets:
        # Index for the groups and which list.
        idx += 1
        for col in sheet.columns:
            i = 0
            try:
                column = col[0].value.lower()
            except:
                column = col[0].value

            if column == groups[idx]:
                for index, cell in enumerate(col):
                    if not isinstance(cell.value, str):
                        try:
                            if index >= 2 and cell.value > lists[idx][i]:
                                lists[idx].pop(0)
                                # Since openpyxl cant get a row easily, this will use regex to pick the row index
                                # Out of the string for cell.
                                insert = re.findall(r'\d+', str(cell))
                                sheet.insert_rows(int(insert[0]), 4)
                        except:
                            pass

                    elif cell.value in lists[idx] and cell.value not in HEADERS:
                        if index >= 2 and cell.value != col[index - 1].value:
                            # Since openpyxl cant get a row easily, this will use regex to pick the row index
                            # Out of the string for cell.
                            insert = re.findall(r'\d+', str(cell))
                            sheet.insert_rows(int(insert[0]), 4)
    wb.save(excel)
