import xlsxwriter as writer

# Author : Imtiaz Adar
# Project : Excel Using Python
# Language : Python

# setup workbook and sheet
workbook = writer.Workbook('informations_.xlsx')
worksheet = workbook.add_worksheet('InformationSheet')

# design
header_format = workbook.add_format(
        {'bold': True, 'font_color': 'white', 'font_size': 40, 'top': 5, 'top_color': '#000000',
        'bottom': 5, 'bottom_color': '#000000', 'left': 5, 'left_color': '#000000',
        'right': 5, 'right_color': '#000000', 'bg_color': '#000000',
        'align': 'center', 'valign': 'center'})
cell_format = workbook.add_format(
        {'bold': True, 'font_color': '#000000', 'font_size': 26, 'top': 2, 'top_color': '#000000',
        'bottom': 2, 'bottom_color': '#000000', 'left': 2, 'left_color': '#000000',
        'right': 2, 'right_color': '#000000', 'bg_color': '#00994c',
        'align': 'center', 'valign': 'center', 'border_color': '#000000'})
serial_format = workbook.add_format(
        {'num_format': '#0', 'bold': True, 'font_color': '#000000', 'font_size': 26,
        'top': 2, 'top_color': '#000000',
        'bottom': 2, 'bottom_color': '#000000', 'left': 2, 'left_color': '#000000',
        'right': 2, 'right_color': '#000000', 'bg_color': '#00994c',
        'align': 'center', 'valign': 'center', 'border_color': '#000000'})
floating_format = workbook.add_format(
        {'num_format': '##.0000', 'bold': True, 'font_color': '#000000', 'font_size': 26,
         'top': 2, 'top_color': '#000000',
         'bottom': 2, 'bottom_color': '#000000', 'left': 2, 'left_color': '#000000',
         'right': 2, 'right_color': '#000000', 'bg_color': '00994c',
         'align': 'center', 'valign': 'center', 'border_color': '#000000'})

# informations
serials = [1, 2, 3, 4, 5, 6]
names = ['Imtiaz Adar', 'Rafsan Kabir', 'Borshon Ahmed', 'Nafi Uddin', 'Hossain Shah', 'Nazmul Ashraf']
phones = ['8801979554646', '8801839939334', '8801675748444', '8801929292934', '8801565444334', '88014535644325']
addresses = ['Dhaka', 'Chattogram', 'Bagura', 'Noakhali', 'Rajshahi', 'Rangpur']

# write into excel file
# headers
worksheet.write('A1', 'serial'.upper(), header_format)
worksheet.write('B1', 'name'.upper(), header_format)
worksheet.write('C1', 'phone'.upper(), header_format)
worksheet.write('D1', 'address'.upper(), header_format)
# values
for i in range(len(serials)):
    worksheet.write(i + 1, 0, serials[i], serial_format)
for i in range(len(names)):
    worksheet.write(i + 1, 1, names[i], cell_format)
for i in range(len(phones)):
    worksheet.write(i + 1, 2, phones[i], cell_format)
for i in range(len(addresses)):
    worksheet.write(i + 1, 3, addresses[i], cell_format)

# column sizing
worksheet.set_column('B1:B1', 50)
worksheet.set_footer('C1:C1', 50)
worksheet.set_column(0, 4, 40)

print('Written Successfully...')
workbook.close()