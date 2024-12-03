from nagare.renderers import xls

data = [
    [-2, 2, 3, -1, 0],
    [30, 20, 33, 20, 15],
    [1, -1, -1, 1, -1],
]


x = xls.Renderer()
with x.workbook, x.worksheet:
    x << x.row('A1', data[0])
    x << x.row('A2', data[1])
    x << x.row('A3', data[2])

    x << x.sparkline('F1', range='Sheet1!A1:E1', markers=True)
    x << x.sparkline('F2', range='Sheet1!A2:E2', type='column', style=12)
    x << x.sparkline('F3', range='Sheet1!A3:E3', type='win_loss', negative_points=True)

x.root.tofile('/tmp/example7.xlsx')
