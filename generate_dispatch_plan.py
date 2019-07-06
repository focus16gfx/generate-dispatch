from openpyxl import Workbook, load_workbook
import openpyxl
from openpyxl.styles.borders import BORDER_THIN, BORDER_THICK
from openpyxl.styles import Border, Side, PatternFill, Font, Alignment
import datetime
import re
import pyexcel as p
import os
import sys

print('>>Initialising...')
wd = '\\\ATL09FPS01\Accord-Folders\sschmidt\Desktop\Dispatch_script\Dispatch_script'


def style_range(ws, cell_range, border=Border(), fill=None, font=None, alignment=None):
    top = Border(top=border.top)
    left = Border(left=border.left)
    right = Border(right=border.right)
    bottom = Border(bottom=border.bottom)
    first_cell = ws[cell_range.split(":")[0]]
    if alignment:
        ws.merge_cells(cell_range)
        first_cell.alignment = alignment
    rows = ws[cell_range]
    if font:
        first_cell.font = font
    for cell in rows[0]:
        cell.border = cell.border + top
    for cell in rows[-1]:
        cell.border = cell.border + bottom
    for row in rows:
        l = row[0]
        r = row[-1]
        l.border = l.border + left
        r.border = r.border + right
        if fill:
            for c in row:
                c.fill = fill


redFill = PatternFill(start_color='FFFF0000', end_color='FFFF0000', fill_type='solid')
yellowFill = PatternFill(start_color='FFFFFF00', end_color='FFFFFF00', fill_type='solid')
greenFill = PatternFill(start_color='FF00cc00', end_color='FF006600', fill_type='solid')
blueFill = PatternFill(start_color='FF0099ff', end_color='FF0099ff', fill_type='solid')
lightblueFill = PatternFill(start_color='FFC5D9F1', end_color='FFC5D9F1', fill_type='solid')

# --------------------------------------------------------------------------------


ASSGNFill = PatternFill(start_color='FFff0000', end_color='FFff0000', fill_type='solid')
DISPFill = PatternFill(start_color='FFffa500', end_color='FFffa500', fill_type='solid')
ATPICKFill = PatternFill(start_color='FFffff00', end_color='FFffff00', fill_type='solid')
PICKDFill = PatternFill(start_color='FF008000', end_color='FF008000', fill_type='solid')
BKRPICKDFill = PatternFill(start_color='FFff00ff', end_color='FFff00ff', fill_type='solid')
BORDERFill = PatternFill(start_color='FF000000', end_color='FF000000', fill_type='solid')
ENRTEFill = PatternFill(start_color='FF800080', end_color='FF800080', fill_type='solid')
ATCONSFill = PatternFill(start_color='FF0000ff', end_color='FF0000ff', fill_type='solid')
SP4DELFill = PatternFill(start_color='FF4b0082', end_color='FF4b0082', fill_type='solid')
SP4OBFill = PatternFill(start_color='FFee82ee', end_color='FFee82ee', fill_type='solid')
SPTLDFill = PatternFill(start_color='FF00ffff', end_color='FF00ffff', fill_type='solid')
thin_border = Border(bottom=Side(border_style=BORDER_THIN, color='00000000', ))
thin_allborder = Border(bottom=Side(border_style=BORDER_THIN, color='00000000'),
                        top=Side(border_style=BORDER_THIN, color='00000000'),
                        left=Side(border_style=BORDER_THIN, color='00000000'),
                        right=Side(border_style=BORDER_THIN, color='00000000'))
thick_border = Border(bottom=Side(border_style=BORDER_THICK, color='00000000'),
                      top=Side(border_style=BORDER_THICK, color='00000000'),
                      left=Side(border_style=BORDER_THICK, color='00000000'),
                      right=Side(border_style=BORDER_THICK, color='00000000'))
print('>>Script started!')

'''take filename as input and convert xls to xlsx'''
og_filename = input('**Report Excel file name: ')
#og_filename = 'BC to AB OB PLANNER'
#is_import_v = input('**Want to import probills starting with \'V\'?(Y/N): ')
is_import_v = 'Y'
old_filename = input('**Old Dispatch plan file name: ')
#old_filename = 'DISPATCH_PLAN Thursday_July_04_2019 t_1314'

if old_filename == '':
    old_filename = None

if is_import_v.lower() == 'y':
    is_import_v = True
elif is_import_v.lower() == 'n':
    is_import_v = False
else:
    print(f'>>Invalid input. Please use either Y for Yes or N for No')

try:
    p.save_book_as(file_name=wd + '\\' + og_filename + '.xls',
                   dest_file_name=f'{wd}\\xlsx_{og_filename}.xlsx')
except FileNotFoundError:
    try:
        p.save_book_as(file_name=wd + '\\' + og_filename + '.xlsx',
                       dest_file_name=f'{wd}\\xlsx_{og_filename}.xlsx')
    except FileNotFoundError:
        print(f'>>Incorrect file name. Check again.')
        sys.exit()

old_ws = ''

if old_filename:
    old_wb = load_workbook(filename=wd + '\\' + old_filename + '.xlsx')
    old_ws = old_wb.active

'''instantiate new worksheet to write data into'''
wb = Workbook()
ws = wb.active

'''headers dictionary'''
dict_headers = {'A': 'PROBILL',
                'B': 'TRIP',
                'C': 'ORIGIN',
                'D': 'PICKUP BY',
                'E': 'DESTINATION',
                'F': 'DELIVER BY',
                'G': 'P/U',
                'H': 'TRAILER',
                'I': 'STATUS'}

old_dict_headers = {'J': 'TRIP',  # 9
                    'K': 'PICKUP BY',  # 10
                    'L': 'ORIGIN',  # 11
                    'Q': 'DELIVER BY',  # 16
                    'T': 'DESTINATION',  # 19
                    'U': 'PROBILL',  # 20
                    'Y': 'P/U',  # 24
                    'Z': 'TRAILER',  # 25
                    'AC': 'STATUS'}  # 28

'''setting default width for columns'''
ws.column_dimensions['A'].width = 14
ws.column_dimensions['B'].width = 13
ws.column_dimensions['C'].width = 25
ws.column_dimensions['D'].width = 30
ws.column_dimensions['E'].width = 25
ws.column_dimensions['F'].width = 30
ws.column_dimensions['G'].width = 10
ws.column_dimensions['H'].width = 10
ws.column_dimensions['I'].width = 10

'''create font objects for different text fields'''
ft = Font(bold=True, size=15)
ft_small = Font(bold=True)
title_ft = Font(bold=True, size=25)
title_2ft = Font(bold=True, size=20)


def write_headers(row_num, fontyy=ft):
    """write headers to a given row in new sheet"""
    for coord in dict_headers:
        ws[coord + str(row_num)] = dict_headers.get(coord)
        ws[coord + str(row_num)].font = fontyy
        ws[coord + str(row_num)].fill = lightblueFill
        ws[coord + str(row_num)].alignment = Alignment(horizontal='center')
    style_range(ws, f'A{row_num}:I{row_num}', border=thin_allborder)

write_headers(3)

"""Open xlsx converted original sheet"""
og_wb = openpyxl.load_workbook(f'{wd}\\xlsx_{og_filename}.xlsx')
og_sheet = og_wb.active

'''merge and format cells'''
ws.merge_cells('A1:I1')
ws['A1'].value = og_sheet['A1'].value
ws['A1'].font = title_ft
ws['A1'].alignment = Alignment(horizontal='center')
ws.merge_cells('A2:I2')
ws['A2'].value = og_sheet['O1'].value
ws['A2'].font = title_2ft
ws['A2'].alignment = Alignment(horizontal='center')
style_range(ws, 'A2:I2', border=thin_border)
style_range(ws, 'A3:I3', fill=blueFill)
style_range(ws, 'A4:I4', fill=blueFill)

# make a regex pattern which searchs if data is present in column number 10
pattern1 = re.compile(r'^\d{3,10}')
rows_list = []
'''collect the rows which contain data in rows_list'''
for row in og_sheet.rows:
    value1 = str(row[9].value)
    if re.search(pattern1, value1):
        '''remove probills start with 'V' if applicable'''
        if is_import_v:
            rows_list.append(row)
        else:
            if not str(row[20].value).lower().startswith('v'):
                rows_list.append(row)

old_row_list = []
'''get data from old sheet'''
for row in old_ws:
    value1 = str(row[1].value)
    if re.search(pattern1, value1):
        old_row_list.append(row)

    # '''check for M at the end of trips (manually added) and add it to the new records rows_list'''
    # if row[1].value[-1].lower() == 'm':
    #     rows_list.append(row)

'''fix microseconds bug in excel sheet for stringformattime'''
colQ = og_sheet['Q']
fix_pat = re.compile(r'\d\d:\d\d:\d\d$')
for cell in colQ:
    if re.search(fix_pat, str(og_sheet[cell.coordinate].value)):
        og_sheet[cell.coordinate] = str(og_sheet[cell.coordinate].value) + '.000000'
colK = og_sheet['K']
fix_pat = re.compile(r'\d\d:\d\d:\d\d$')

for cell in colK:
    if re.search(fix_pat, str(og_sheet[cell.coordinate].value)):
        og_sheet[cell.coordinate] = str(og_sheet[cell.coordinate].value) + '.000000'

pickupdate_index = 10
deliverydate_index = 16

dates = []

'''parse unique dates from Delivery dates'''


def get_dates():
    global dates
    pt3 = re.compile(r'^\d{4}-\d\d-\d\d')
    for row in rows_list:
        z = re.match(pt3, str(row[deliverydate_index].value))
        date_extracted = z.group(0)
        if date_extracted not in dates:
            dates.append(date_extracted)


get_dates()

#    queue_list.sort(key=lambda x: datetime.datetime.strptime(str(x[pickupdate_index].value), '%A, %B %d, %Y @ %H:%M'))

dates.sort(key=lambda x: datetime.datetime.strptime(str(x), '%Y-%m-%d'))

current_row = 5

'''date necessary format'''
for row in rows_list:
    dt3 = datetime.datetime.strptime(str(row[pickupdate_index].value), '%Y-%m-%d  %H:%M:%S.%f')
    row[pickupdate_index].value = dt3.strftime('%A, %B %d, %Y @ %H:%M')
    dt3 = datetime.datetime.strptime(str(row[deliverydate_index].value), '%Y-%m-%d  %H:%M:%S.%f')
    row[deliverydate_index].value = dt3.strftime('%A, %B %d, %Y @ %H:%M')


def each_date(date):
    global lol
    queue_list = []
    '''Add all rows in new rowlist with this date'''
    for row_q1 in rows_list:
        dt6 = datetime.datetime.strptime(date, '%Y-%m-%d')
        match_date = dt6.strftime('%A, %B %d, %Y')
        if match_date in str(row_q1[deliverydate_index].value):
            queue_list.append(row_q1)

    if old_filename:
        '''------------old------------'''
        old_queue_list = []
        '''Add all rows in new rowlist with this date'''
        for row_q1 in old_row_list:
            dt6 = datetime.datetime.strptime(date, '%Y-%m-%d')
            match_date = dt6.strftime('%A, %B %d, %Y')
            if match_date in str(row_q1[5].value):
                old_queue_list.append(row_q1)

        '''sort new rowlist by pickup'''
        old_queue_list.sort(key=lambda x: datetime.datetime.strptime(str(x[3].value), '%A, %B %d, %Y @ %H:%M'))

        '''manipulation op'''
        for new_row in queue_list:
            new_trip = new_row[9].value
            for old_row in old_queue_list:
                if old_row[1].value == new_trip:
                    n_pu = new_row[24].value
                    n_trailer = new_row[25].value
                    n_status = new_row[28].value
                    o_pu = old_row[6].value
                    o_trailer = old_row[7].value
                    o_status = old_row[8].value

                    if not n_pu:
                        n_pu = o_pu
                        new_row[24].value = n_pu
                    if not o_trailer:
                        n_trailer = o_trailer
                        new_row[25].value = n_trailer

                    status_priority = {'ASSGN': 1,
                                       'DISP': 2,
                                       'ATPICK': 3,
                                       'PICKD': 4,
                                       'BKRPICKD': 4,
                                       'BORDER': 5,
                                       'ENRTE': 6,
                                       'ATCONS': 7,
                                       'SP4DEL': 8,
                                       'SP4OB': 9,
                                       'SPTLD': 10,
                                       'DELVD': 11, }

                    if o_status == 'BROKER':
                        n_status = o_status
                        new_row[28].value = n_status
                    elif status_priority[o_status] > status_priority[n_status]:
                        n_status = o_status
                        new_row[28].value = n_status
    if old_filename:
        for old_row in old_queue_list:
            lol = []
            '''manual update'''
            if str(old_row[1].value)[-1].lower() == 'm':

                for i in range(29):
                    cell = og_sheet['AH8']
                    lol.append(cell)

                lol[9] = old_row[1]
                lol[10] = old_row[3]
                lol[11] = old_row[2]
                lol[16] = old_row[5]
                lol[19] = old_row[4]
                lol[20] = old_row[0]
                lol[24] = old_row[6]
                lol[25] = old_row[7]
                lol[28] = old_row[8]
                queue_list.append(lol)

    '''sort new rowlist by  pickup'''

    queue_list.sort(key=lambda x: datetime.datetime.strptime(str(x[pickupdate_index].value), '%A, %B %d, %Y @ %H:%M'))

    global current_row

    '''add heading each day'''
    dt = datetime.datetime.strptime(date, '%Y-%m-%d')
    ws[f'A{current_row}'] = dt.strftime('%A, %B %d, %Y')
    ws.merge_cells(f'A{current_row}:I{current_row}')
    style_range(ws, f'A{current_row}:I{current_row}', border=thin_border)
    ws[f'A{current_row}'].font = ft
    current_row += 1
    write_headers(current_row, ft_small)

    '''write sorted list of rows to excel'''
    print('PICKUP DATES...')

    style_range(ws, f'A{current_row}:I{current_row + len(queue_list)}', border=thin_allborder)


    for row_q in queue_list:
        current_row += 1
        ws[f'D{current_row}'] = row_q[10].value
        ws[f'C{current_row}'] = row_q[11].value
        ws[f'E{current_row}'] = row_q[19].value
        ws[f'F{current_row}'] = row_q[16].value
        ws[f'B{current_row}'] = row_q[9].value
        ws[f'A{current_row}'] = row_q[20].value

        ws[f'G{current_row}'] = row_q[24].value
        if row_q[24].value is None:
            ws[f'G{current_row}'].fill = yellowFill

        ws[f'H{current_row}'] = row_q[25].value
        if row_q[25].value is None:
            ws[f'H{current_row}'].fill = yellowFill

        ws[f'I{current_row}'] = row_q[28].value

        status = row_q[28].value

        '''color fill'''
        if status == 'ASSGN':
            ws[f'I{current_row}'].fill = ASSGNFill
        elif status == 'DISP':
            ws[f'I{current_row}'].fill = DISPFill
        elif status == 'ATPICK':
            ws[f'I{current_row}'].fill = ATPICKFill
        elif status == 'PICKD':
            ws[f'I{current_row}'].fill = PICKDFill
        elif status == 'BKRPICKD':
            ws[f'I{current_row}'].fill = BKRPICKDFill
        elif status == 'BORDER':
            ws[f'I{current_row}'].fill = BORDERFill

        elif status == 'ENRTE':
            # ws[f'I{current_row}'].fill = ENRTEFill
            style_range(ws, f'A{current_row}:I{current_row}', fill=ENRTEFill)

        elif status == 'ATCONS':
            ws[f'I{current_row}'].fill = ATCONSFill
        elif status == 'SP4DEL':
            ws[f'I{current_row}'].fill = SP4DELFill
        elif status == 'SP4OB':
            ws[f'I{current_row}'].fill = SP4OBFill
        elif status == 'SPTLD':
            ws[f'I{current_row}'].fill = SPTLDFill

        elif status == 'BROKER':
            style_range(ws, f'A{current_row}:I{current_row}', fill=ENRTEFill)



        '''center align'''
        ws[f'D{current_row}'].alignment = Alignment(horizontal='center')
        ws[f'C{current_row}'].alignment = Alignment(horizontal='center')
        ws[f'E{current_row}'].alignment = Alignment(horizontal='center')
        ws[f'F{current_row}'].alignment = Alignment(horizontal='center')
        ws[f'B{current_row}'].alignment = Alignment(horizontal='center')
        ws[f'A{current_row}'].alignment = Alignment(horizontal='center')
        ws[f'G{current_row}'].alignment = Alignment(horizontal='center')
        ws[f'H{current_row}'].alignment = Alignment(horizontal='center')
        ws[f'I{current_row}'].alignment = Alignment(horizontal='center')


        print(f'-{row_q[16].value}')

print('----------------------------------------------------')

for date in dates:
    print(f'>>Writing data for {date}')
    each_date(date)
    current_row += 2

# ws.sheet_view.showGridLines = False

name_ext = str(og_sheet['O1'].value)
# Tuesday, July 2, 2019 21:43

dt_a1 = datetime.datetime.strptime(name_ext, '%A, %B %d, %Y %H:%M')
name_ext = dt_a1.strftime('%A_%B_%d_%Y t_%H%M')
try:
    if old_filename:
        wb.save(f'{wd}\\ov_DISPATCH_PLAN {name_ext}.xlsx')
    elif not is_import_v:
        wb.save(f'{wd}\\DISPATCH_PLAN {name_ext}_noV.xlsx')
    elif is_import_v:
        wb.save(f'{wd}\\DISPATCH_PLAN {name_ext}.xlsx')


except PermissionError:
    print(
        f'>>ERROR!! You did not close the file "DISPATCH_PLAN {name_ext}.xlsx" can\'t write to it while it\'s still open')
    sys.exit()

os.system(f'del "{wd}\\xlsx_{og_filename}.xlsx"')

if old_filename:
    print(f'>>SUCCESSFULLY generated a new dispatch plan, file saved as "ov_DISPATCH_PLAN {name_ext}.xlsx"')
else:
    print(f'>>SUCCESSFULLY generated a new dispatch plan, file saved as "DISPATCH_PLAN {name_ext}.xlsx"')