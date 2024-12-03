from nagare.renderers import xls

x = xls.Renderer()
with x.workbook, x.worksheet as worksheet:
    worksheet.set_column('A:A', 30)

    x << x.cell('A1', 'http://www.python.org/')
    x << x.url_cell('A3', 'http://www.python.org/', string='Python Home')
    x << x.url_cell('A5', 'http://www.python.org/', tip='Click here')
    x << x.url_cell(
        'A7',
        'http://www.python.org/',
        x.format(font_color='red', bold=True, underline=True, font_size=12),
    )
    x << x.url_cell('A9', 'mailto:jmcnamara@cpan.org', string='Mail me')

    x << x.str_cell('A11', 'http://www.python.org/')

x.root.tofile('/tmp/example3.xlsx')
