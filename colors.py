import xlsxwriter
import colorsys


# Create an new Excel file and add a worksheet.
workbook = xlsxwriter.Workbook('demo-colors.xlsx')
worksheet = workbook.add_worksheet()

# Widen the first column to make the text clearer.
worksheet.set_column('B:B', 8)


def rgbfromint(rgbint):
  return '#%02x%02x%02x' % (rgbint // 256 // 256 % 256, rgbint // 256 % 256, rgbint % 256)

cantcolors = 256*256
separacion = 8

allcolors = [rgbfromint(c) for c in range(0, 0xFFFFFF+1, separacion) if c <= cantcolors]
print(len(allcolors))

col = 1
row = 1
for i, color in enumerate(allcolors):
  format = workbook.add_format()
  format.set_bg_color(color)
  worksheet.write(row, col, color, format)
  
  row += 1
  if row > 256: 
    row = 1
    col = col + 1


"""
"""

workbook.close()