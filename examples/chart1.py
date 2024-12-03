from nagare.renderers import xls

data = [
    [1, 2, 3, 4, 5],
    [2, 4, 6, 8, 10],
    [3, 6, 9, 12, 15],
]

x = xls.Renderer()
with x.workbook, x.worksheet:
    x << [x.column(x.ref('A1') + i, col) for i, col in enumerate(data)]

    with x.chart_cell('A7'), x.chart(type='column'):
        x << x.series(values='=Sheet1!$A$1:$A$5')
        x << x.series(values='=Sheet1!$B$1:$B$5')
        x << x.series(values='=Sheet1!$C$1:$C$5')

x.root.tofile('/tmp/chart1.xlsx')
