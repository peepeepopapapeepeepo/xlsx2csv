import xlrd
import csv
import sys

def csv_from_excel(file_in, sheet='Sheet1', file_out=""):
    try:
        wb = xlrd.open_workbook(file_in)
        sh = wb.sheet_by_name(sheet)
        file_out = '.'.join(file_in.split('.')[:-1] + ['csv']) if file_out == '' else file_out

        print(f'in = {file_in}, sheet = {sheet}')
        print(f'out= {file_out}')

        with open(file_out, mode='w', newline='', encoding='utf-8') as csv_file:
            # produce - aaaa,bbbb,cccc
            wr = csv.writer(csv_file, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
            # produce - "aaaa","bbbb","cccc"
            #wr = csv.writer(csv_file, quoting=csv.QUOTE_ALL)
            for rownum in range(sh.nrows):
                wr.writerow(sh.row_values(rownum))

        return True

    except xlrd.biffh.XLRDError as err:
        print(f'Error: {err}')
    except FileNotFoundError as err:
        print(f'Error: {err}')
    
    return False

if __name__ == "__main__":
    
    if len(sys.argv) == 2:
        csv_from_excel(sys.argv[1])
    elif len(sys.argv) == 3:
        csv_from_excel(sys.argv[1], sys.argv[2])
    elif len(sys.argv) == 4:
        csv_from_excel(sys.argv[1], sys.argv[2], sys.argv[3])
    else:
        print(f'Error: Usage = {sys.argv[0]} <xlsx> <sheet> <csv>')
