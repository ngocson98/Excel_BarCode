import pandas as pd
import qrcode
import xlsxwriter

"""workbook = xlsxwriter.Workbook('file_data.xlsx')
worksheet = workbook.add_worksheet()
"""

try:
    from PIL import Image
except ImportError:
    import Image

xl = pd.ExcelFile('file_data.xlsx')

df = pd.read_excel(xl, 0, header=None)
# print(df.head())
df_row = df.iloc[:, 1]
# print(df_row)
max_rows = len(df.iloc[:, 1])
print(max_rows)

for i in range(max_rows):
    print("Row " + str(i + 1) + ': ' + df.at[i, 1])
    data = df.at[i, 1]
    img = qrcode.make(data)
    img.save('CreateQR/MyQRCode{}.png'.format(i + 1))
    # a = 'C' + str(i+1)
    # worksheet.insert_image(a.format(i+1), 'CreateQR/MyQRCode{}.png'.format(i+1), {'x_offset': 15, 'y_offset': 10})
    print("QR {} finish!".format(i + 1))

# worksheet.insert_image('C2', 'CreateQR/MyQRCode1.png', {'x_offset': 15, 'y_offset': 10})

# ____________________________________________________

# Widen the first column to make the text clearer.
"""worksheet.set_column('A:A', 30)
worksheet.insert_image('C1', 'CreateQR/MyQRCode1.png', {'x_scale': 0.1, 'y_scale': 0.1})

workbook.close()
"""
