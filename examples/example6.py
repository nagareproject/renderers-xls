from nagare.renderers import xls

data = [
    ['Apples', 10000, 5000, 8000, 6000],
    ['Pears', 2000, 3000, 4000, 5000],
    ['Bananas', 6000, 6000, 6500, 6000],
    ['Oranges', 500, 300, 200, 700],
]

x = xls.Renderer()
with x.workbook:
    with x.worksheet as worksheet:
        worksheet.set_column('B:G', 12)

        x << x.cell('B1', 'Default table with no data')
        x << x.table('B3:F7')

    with x.worksheet as worksheet:
        worksheet.set_column('B:G', 12)

        x << x.cell('B1', 'Default table with data')
        x << x.table(
            'B3:G7',
            columns=[
                {'header': 'Product'},
                {'header': 'Quarter 1'},
                {'header': 'Quarter 2'},
                {'header': 'Quarter 3'},
                {'header': 'Quarter 4'},
            ],
            data=data,
        )

x.root.tofile('/tmp/example6.xlsx')
