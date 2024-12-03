from datetime import datetime

from nagare.renderers import xls

date_formats = (
    'dd/mm/yy',
    'mm/dd/yy',
    'dd m yy',
    'd mm yy',
    'd mmm yy',
    'd mmmm yy',
    'd mmmm yyy',
    'd mmmm yyyy',
    'dd/mm/yy hh:mm',
    'dd/mm/yy hh:mm:ss',
    'dd/mm/yy hh:mm:ss.000',
    'hh:mm',
    'hh:mm:ss',
    'hh:mm:ss.000',
)

x = xls.Renderer()
with x.workbook, x.worksheet as worksheet:
    worksheet.set_column('A:B', 30)

    bold = x.format(bold=True)
    x << x.cell('A1', 'Formatted date', bold)
    x << x.cell('B1', 'Format', bold)

    date_time = datetime.strptime('2013-01-23 12:30:05.123', '%Y-%m-%d %H:%M:%S.%f')

    for row, date_format_str in enumerate(date_formats, 1):
        x << x.cell(row, 0, date_time, x.format(num_format=date_format_str, align='left'))
        x << x.cell(row, 1, date_format_str)

x.root.tofile('/tmp/example2.xlsx')
