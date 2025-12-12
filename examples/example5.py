# --
# Copyright (c) 2014-2025 Net-ng.
# All rights reserved.
#
# This software is licensed under the BSD License, as described in
# the file LICENSE.txt, which you should have received as part of
# this distribution.
# --

from nagare.renderers import xls

x = xls.Renderer()

with x.workbook as workbook, x.worksheet as worksheet:
    worksheet.set_column('B:D', 12)
    worksheet.set_row(3, 30)
    worksheet.set_row(6, 30)
    worksheet.set_row(7, 30)

    merge_format = workbook.add_format(bold=True, border=1, align='center', valign='vcenter', fg_color='yellow')

    worksheet.merge_range('B4:D4', 'Merged Range', merge_format)
    worksheet.merge_range('B7:D8', 'Merged Range', merge_format)

x.root.tofile('/tmp/example5.xlsx')
