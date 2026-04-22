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
    x << x.cell('A1', 'Hello')
    x << x.comment_cell('A1', 'This is a comment')

x.root.tofile('/tmp/example8.xlsx')
