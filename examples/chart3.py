from nagare.renderers import xls

x = xls.Renderer()

with x.workbook:
    for chart_type in ['column', 'area', 'line', 'pie']:
        with x.worksheet(chart_type.title()) as worksheet:
            worksheet.set_zoom(30)

            style_number = 1
            for row_num in range(0, 90, 15):
                for col_num in range(0, 64, 8):
                    with x.chart_cell(row_num, col_num), x.chart(type=chart_type) as chart:
                        chart.set_title(name='Style %d' % style_number)
                        chart.set_legend(none=True)
                        chart.set_style(style_number)
                        style_number += 1

                        x << x.series(values='=Data!$A$1:$A$6')

    with x.worksheet('Data') as worksheet:
        worksheet.hide()
        x << x.column('A1', [10, 40, 50, 20, 10, 50])


x.root.tofile('/tmp/chart3.xlsx')
