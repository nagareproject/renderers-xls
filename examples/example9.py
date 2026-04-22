# --
# Copyright (c) 2014-2026 Net-ng.
# All rights reserved.
#
# This software is licensed under the BSD License, as described in
# the file LICENSE.txt, which you should have received as part of
# this distribution.
# --

from nagare.renderers import xls

x = xls.Renderer()
with x.workbook, x.worksheet:
    c = x.ref('A1')

    x << x.cell(c, 'Hello')
    x << x.cell(c + 1, 'World')

    x << x.cell(c + x.cells(2, 1), 'Here')

x.root.tofile('/tmp/example9.xlsx')
