from ctypes import windll
import datetime
import os
from PIL import ImageGrab
import time
import xlwings as xw
import win32com.client as win32

import paths_manager


def get_height_of_table(ws, first_row, col, row_offset, col_offset):
    h = row_offset - 1
    this_row = first_row + row_offset
    cell_val = ws.cells(this_row, col + col_offset).value
    print(f"Fila {this_row}")
    while cell_val:
        h += 1
        this_row += 1
        print(f"Fila {this_row}")
        cell_val = ws.cells(this_row, col + col_offset).value
    return h

def send_email_with_pics(email_address, pic_paths):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = email_address
    mail.Subject = f'Monitor notas para la fecha {datetime.datetime.today().strftime("%d/%m/%Y")}'

    #html_accum = '<html><head></head><body><div>'

    for i in range(len(pic_paths)):
        full_path = os.path.join(os.getcwd(), pic_paths[i])
        print(f"Adding {full_path}")
        attachment = mail.Attachments.Add(full_path)
        """
        print(f"<img src='{pic_paths[i]}' /></br></br>")
        html_accum += f"<img src='{pic_paths[i]}' /></br></br>"
        """

    #mail.HTMLBody = html_accum + '</div></body></html>'

    with open("email.html", "r", encoding='utf-8') as f:
        text = f.read()
        mail.HTMLBody = text
        print(text)
    mail.Send()

st = time.time()

"""
The following is a series of series of tuples with the shape
(
    sheet name, 
    (default range, default first row, default first column, default last row, default last column, offset first row, offset first column),
    1 if it's the range is variable 0 otherwise,
    sequence number
)
"""
tgt_range_groups = (
    (
        ('MONITOR Autocall', ('E10:AH79', 10, 5, 79, 34, 4, 0), 1, 1),
        ('Roll', ('AC31:AJ68', 31, 29, 68, 36, 2, 0), 1, 2),
        ('MONITOR Creditos', ('D10:T42', 10, 4, 42, 20, 3, 0), 1, 3),
        ('BC', ('C2:D25', 2, 3, 25, 4, -1, -1), 0, 7),
        ('%RV', ('I5:J16', 5, 9, 16, 10, -1, -1), 0, 8),
    ),
    (
        ('MONITOR Creditos', ('D45:P49', 45, 4, 49, 16, 2, 0), 1, 4),
        ('Roll', ('C12:T30', 12, 3, 30, 20, -1, -1), 0, 5),
        ('%RV', ('O1:U20', 1, 15, 20, 21, -1, -1), 0, 9),
    ),
    (
        ('Roll', ('T69:Y77', 69, 20, 77, 25, 3, 0), 1, 6),
        ('%RV', ('K28:M33', 28, 11, 33, 13, -1, -1), 0, 10),
    )
)

data = paths_manager.get_paths()
src_file = data['excel_to_img.input_file']

pic_paths = []

for tgt_range_group in tgt_range_groups:
    with xw.App(visible=False) as app:
        wb = xw.Book(src_file)
        """
        wb.app.calculate()
        time.sleep(5)
        wb.app.calculate()
        time.sleep(5)
        wb.save()
        wb = xw.Book(src_file)
        """
        for i in range(len(tgt_range_group)):
            print(tgt_range_group[i])
            ws = wb.sheets[tgt_range_group[i][0]]
            ws.calculate()
            for i in range(5):
                print(f"Sleeping, {i + 1} seconds.")
                time.sleep(1)
            wb.save()
            # Might as well close and open book here.
            if tgt_range_group[i][2]:
                table_height = get_height_of_table(ws, tgt_range_group[i][1][1], tgt_range_group[i][1][2], tgt_range_group[i][1][5], tgt_range_group[i][1][6])
            else:
                table_height = tgt_range_group[i][1][3] - tgt_range_group[i][1][1]
            rango_variable = ws.range(
                (tgt_range_group[i][1][1], tgt_range_group[i][1][2]),
                (tgt_range_group[i][1][1] + table_height, tgt_range_group[i][1][4])
            )
            # Might as well try only range
            # Like so: sheet.range("A1:B10").calculate().
            rango_variable.copy()

            img = ImageGrab.grabclipboard()
            img_name = f'test_{tgt_range_group[i][3]}.jpg'
            pic_paths.append(img_name)
            img_b = img.convert('RGB')
            img_b.save(img_name)
            if windll.user32.OpenClipboard(None):
                windll.user32.EmptyClipboard()
                windll.user32.CloseClipboard()

        wb.close()


#pic_paths = [f"test_{i}.jpg" for i in range(1, 10 + 1)]
print(pic_paths)
print("Time to send the email")
send_email_with_pics("kevinbarzola@zest.pe", pic_paths)
send_email_with_pics("joakimbaraka@zest.pe", pic_paths)

et = time.time()
print('Execution time:', et - st, 'seconds')
