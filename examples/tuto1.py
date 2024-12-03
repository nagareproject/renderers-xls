from nagare.renderers import xls

expenses = (
    ['Rent', 1000],
    ['Gas',   100],
    ['Food',  300],
    ['Gym',    50],
)

x = xls.Renderer()

with x.workbook:
    bold = x.format(bold=True)
    money = x.format(num_format='$#,##0')

    with x.worksheet('DATA 1'):
        x << x.cell('A1', 'Item', bold) << x.cell('B1', 'Cost', bold)

        for row, (item, cost) in enumerate(expenses, 1):
            x << x.cell(row, 0, item)
            x << x.cell(row, 1, cost, money)

        x << x.cell(row + 1, 0, 'Total', bold)
        with x.cell(row + 1, 1, '=SUM(B1:B5)'):
            x << bold

        x << x.textbox(
            'F10',
            'Hello world',
            font={'name': 'Arial', 'size': 14, 'bold': True},
            align={'vertical': 'middle', 'horizontal': 'center'},
            gradient={'colors': ['#DDEBCF', '#9CB86E', '#156B13']}
        )

    x << x.worksheet('DATA 2')

x.root.tofile('/tmp/tuto1.xlsx')
