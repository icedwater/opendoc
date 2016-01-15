#! /usr/bin/env python

"""
Here's a small script based on a github gist to show diffs
between Excel files. I'll be using this to see if openpyxl
or xlrd is easier to use or they serve different use cases
entirely. Thanks, nmz787.
"""
import xlrd
import sys

def cell_on_line(excel_filename):
    """
    Take an input file and display each cell on its own line.
    """
    excel_file = xlrd.open_workbook(excel_filename)
    if not excel_file:
        raise Exception('The file provided was not an Excel file.')

    worksheets = excel_file.sheet_names()
    for sheet in worksheets:
        current_sheet = excel_file.sheet_by_name(sheet)
        cells = []

        for row in xrange(current_sheet.nrows):
            current_row = current_sheet.row(row)

            if not current_row:
                continue

            row_as_string = "[{}: ({},".format(sheet, row)

            for cell in xrange(current_sheet.ncols):
                s = str(current_sheet.cell_value(row, cell))
                s = s.replace(r"\\", "\\\\")
                s = s.replace(r"\n", " ")
                s = s.replace(r"\r", " ")
                s = s.replace(r"\t", " ")
                if s:
                    cells.append('{} {})] {}\n'.format(row_as_string, cell, s))

        if cells:
            return ''.join(cells)


if __name__ == '__main__':
    filename = 'sample.xlsx'
    print("Opening %s..." % filename)
    print cell_on_line('sample.xlsx')
