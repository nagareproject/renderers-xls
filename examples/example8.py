from nagare.renderers import xls

x = xls.Renderer()
with x.workbook, x.worksheet:
    x << x.cell('A1', 'Hello')
    x << x.comment_cell('A1', 'This is a comment')

x.root.tofile('/tmp/example8.xlsx')
