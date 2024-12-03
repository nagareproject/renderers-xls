from nagare.renderers import xls

x = xls.Renderer()

with x.workbook:
    with x.worksheet as worksheet:
        worksheet.set_column('A:A', 20)

        x << x.cell('A1', 'Hello')
        with x.cell('A2'):
            x << 'World'
            x << x.format(bold=True)

        with x.cell(2, 0):
            x << 123

        with x.cell(3, 0):
            x << 123.456

        x << x.image('B5', 'google.png')
        x << x.image('B10', 'google.png', x_scale=0.5, y_scale=0.5)

x.root.tofile('/tmp/example1.xlsx')
