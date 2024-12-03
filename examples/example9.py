from nagare.renderers import xls

x = xls.Renderer()
with x.workbook, x.worksheet:
    c = x.ref('A1')

    x << x.cell(c, 'Hello')
    x << x.cell(c + 1, 'World')

    x << x.cell(c + x.cells(2, 1), 'Here')

x.root.tofile('/tmp/example9.xlsx')
