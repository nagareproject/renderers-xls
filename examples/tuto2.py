from datetime import datetime

from nagare.renderers import xls

expenses = (
    ['Rent', '2013-01-13', 1000],
    ['Gas', '2013-01-14', 100],
    ['Food', '2013-01-16', 300],
    ['Gym', '2013-01-20', 50],
)

x = xls.Renderer()
with x.workbook:
    bold = x.format(bold=True)
    money_format = x.format(num_format='$#,##0', bold=True)
    date_format = x.format(num_format='mmmm d yyyy')

    with x.worksheet as worksheet:
        worksheet.set_column(1, 1, 15)

        x << x.cell('A1', 'Item', x.format(bold=True))
        x << x.cell('B1', 'Date', bold)
        x << x.cell('C1', 'Cost', bold)

        for row, (item, date_str, cost) in enumerate(expenses, 1):
            x << x.cell(row, 0, item)
            x << x.cell(row, 1, datetime.strptime(date_str, '%Y-%m-%d'), date_format)
            x << x.cell(row, 2, cost, money_format)

        for row, (item, date_str, cost) in enumerate(expenses):
            cell = x.ref('G2') + x.rows(row)

            x << x.cell(cell, item)
            x << x.cell(cell + 1, datetime.strptime(date_str, '%Y-%m-%d'), date_format)
            x << x.cell(cell + 2, cost, money_format)

        x << x.column('L2', [cost for item, date_str, cost in expenses], money_format)

        x << x.cell(row + 1, 0, 'Total', bold)
        x << x.cell(row + 1, 2, '=SUM(C2:C5)', money_format)

x.root.tofile('/tmp/tuto2.xlsx')
