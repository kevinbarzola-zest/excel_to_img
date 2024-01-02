from PIL import ImageGrab
import time
import xlwings as xw


def get_height_of_table(ws, first_row, col):
    h = 0
    this_row = first_row
    cell_val = ws.cells(this_row, col).value
    while cell_val:
        h += 1
        this_row += 1
        cell_val = ws.cells(this_row, col).value
    return h

st = time.time()
tgt_range_groups = (
    (
        ('MONITOR Autocall', ('E10:AH79', 10, 5, 79, 34), 1, 1),
        ('Roll', ('AC31:AJ68', 31, 29, 68, 36), 1, 2),
        ('MONITOR Creditos', ('D10:T42', 10, 4, 42, 20), 1, 3),
        ('Roll', ('C12:T30', 12, 3, 30, 20), 0, 5),
        ('BC', ('C2:D25', 2, 3, 25, 4), 0, 7),
        ('%RV', ('O1:U20', 1, 15, 20, 21), 0, 8),
    ),
    (
        ('MONITOR Creditos', ('D45:P49', 45, 4, 49, 16), 1, 4),
        ('Roll', ('T69:Y77', 69, 20, 77, 25), 1, 6),
        ('%RV', ('K28:M33', 28, 11, 33, 13), 0, 9),
    ),
)

for tgt_range_group in tgt_range_groups:
    with xw.App(visible=False) as app:
        wb = xw.Book("monitor_notas.xlsm")
        for i in range(len(tgt_range_group)):
            print(tgt_range_group[i])
            ws = wb.sheets[tgt_range_group[i][0]]


            ws.range(tgt_range_group[i][1]).api.CopyPicture()
            ws.api.Paste()
            pic = ws.pictures[0]
            pic.api.Copy()

            img = ImageGrab.grabclipboard()
            img.save(f'test_{tgt_range_group[i][3]}.png')

        wb.close()

et = time.time()
print('Execution time:', et - st, 'seconds')