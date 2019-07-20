import xlsxwriter
from PIL import Image
from pathlib import Path

workbook = xlsxwriter.Workbook('demo.xlsx')
sheet1 = workbook.add_worksheet()

H_FACTOR = 1.333
W_FACTOR = 8

image_col_index = 0
max_width = 0
max_height = 0
path_source = Path('D:/work/client/sanguoclient/resource/assets')
list_file = sorted(path_source.rglob('*.*'))
index = 0
for v in list_file:
    if v.suffix == '.png' or v.suffix == '.jpg':
        index += 1
        img = Image.open(v)
        w, h = img.size
        if w > max_width:
            max_width = w
        if h > max_height:
            max_height = h
        sheet1.insert_image(index, image_col_index, str(v))
        if h > sheet1.default_row_pixels:
            sheet1.set_row(index, h / H_FACTOR)
print('max_width={0}'.format(max_width))
print('max_height={0}'.format(max_height))

sheet1.set_column(image_col_index, image_col_index, width=max_width / W_FACTOR)
workbook.close()
print('...文档生成完毕')
