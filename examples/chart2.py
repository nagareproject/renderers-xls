from nagare.renderers import xls

headings = ['Number', 'Batch 1', 'Batch 2']
data = [
    [2, 3, 4, 5, 6, 7],
    [40, 40, 50, 30, 25, 50],
    [30, 25, 30, 10, 5, 10],
]


x = xls.Renderer()
with x.workbook, x.worksheet:
    bold = x.format(bold=True)

    x << x.row('A1', headings, bold)
    x << [x.column(x.ref('A2') + col, values) for col, values in enumerate(data)]

    with x.chart_cell('D2', x_offset=25, y_offset=10), x.chart(type='area') as chart:
        chart.set_title(name='Results of sample analysis')
        chart.set_x_axis(name='Test number')
        chart.set_y_axis(name='Sample length (mm)')
        chart.set_style(11)

        x << x.series(name='=Sheet1!$B$1', categories='=Sheet1!$A$2:$A$7', values='=Sheet1!$B$2:$B$7')
        x << x.series(name=['Sheet1', 0, 2], categories=['Sheet1', 1, 0, 6, 0], values=['Sheet1', 1, 2, 6, 2])

    with x.chart_cell('D18', x_offset=25, y_offset=1), x.chart(type='area', subtype='stacked') as chart:
        chart.set_title(name='Stacked Chart')
        chart.set_x_axis(name='Test number')
        chart.set_y_axis(name='Sample length (mm)')
        chart.set_style(12)

        x << x.series(name='=Sheet1!$B$1', categories='=Sheet1!$A$2:$A$7', values='=Sheet1!$B$2:$B$7')
        x << x.series(name='=Sheet1!$C$1', categories='=Sheet1!$A$2:$A$7', values='=Sheet1!$C$2:$C$7')

    with x.chart_cell('D34', x_offset=25, y_offset=10), x.chart(type='area', subtype='percent_stacked') as chart:
        chart.set_title(name='Percent Stacked Chart')
        chart.set_x_axis(name='Test number')
        chart.set_y_axis(name='Sample length (mm)')
        chart.set_style(13)

        x << x.series(name='=Sheet1!$B$1', categories='=Sheet1!$A$2:$A$7', values='=Sheet1!$B$2:$B$7')
        x << x.series(name='=Sheet1!$C$1', categories='=Sheet1!$A$2:$A$7', values='=Sheet1!$C$2:$C$7')

x.root.tofile('/tmp/chart2.xlsx')
