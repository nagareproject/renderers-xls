# --
# Copyright (c) 2008-2024 Net-ng.
# All rights reserved.
#
# This software is licensed under the BSD License, as described in
# the file LICENSE.txt, which you should have received as part of
# this distribution.
# --

import io
import random
import itertools
from collections.abc import Iterable

import xlsxwriter


class Tag:
    def __init__(self, renderer):
        self._renderer = renderer

        self._deferred = []
        self._children = []
        self.kw = {}

    def __getattr__(self, name):
        return lambda *args, **kw: self._deferred.append((name, args, kw))

    def __call__(self, *children, **attributes):
        self._children.extend(children)
        self.kw.update({name.rstrip('_'): value for name, value in attributes.items()})

        return self

    def __enter__(self):
        self._renderer.enter(self)
        return self

    def __exit__(self, exception, data, tb):
        if exception is None:
            self._renderer.exit(self)

    @xlsxwriter.worksheet.convert_cell_args
    def convert_cell_args(self, *args):
        return args

    @property
    def args(self):
        return itertools.takewhile(lambda o: not isinstance(o, Tag), self._children)

    @property
    def children(self):
        return itertools.dropwhile(lambda o: not isinstance(o, Tag), self._children)

    def before(self, parent, *args, **kw):
        return parent

    def after(self, workbook, parent, children, *args, **kw):
        pass

    def generate(self, workbook, parent):
        me = self.before(workbook, parent, *self.args, **self.kw)
        for method, args, kw in self._deferred:
            getattr(me, method)(*(args + ((kw,) if kw else ())))

        children = [child.generate(workbook, me) for child in self.children]

        self.after(workbook, parent, children, *self.args, **self.kw)

        return me


# =====================================================================================================================


class Workbook(Tag):
    def __init__(self, renderer):
        super().__init__(renderer)
        self.workbook = xlsxwriter.workbook.Workbook()

    def __getattr__(self, name):
        return lambda *args, **kw: getattr(self.workbook, name)(*args, kw)

    def __call__(self, *children, **attributes):
        if attributes:
            self.workbook = xlsxwriter.workbook.Workbook(options=attributes)
            self.workbook.write = self.write

        return super().__call__(*children)

    def tofile(self, filename):
        with self.workbook as workbook:
            workbook.filename = filename
            self.generate(workbook, workbook)

    def tostring(self):
        output = io.BytesIO()
        self.tofile(output)
        return output.getvalue()


class Format(Tag):
    @staticmethod
    def before(workbook, _, *args, **kw):
        return workbook.add_format(*args, kw)


class Worksheet(Tag):
    def __init__(self, renderer):
        super().__init__(renderer)

        self.worksheet = None
        self.columns_sizes = {}
        self.adjust_columns = {}

    def adjust_column_size(self, row, value):
        adjust = self.adjust_columns.get(row, lambda size, value: max(size, len(str(value))))
        self.columns_sizes[row] = adjust if isinstance(adjust, int) else adjust(self.columns_sizes.get(row, 0), value)

    def _autofit(self):
        for col, size in self.columns_sizes.items():
            self.worksheet.set_column(col, col, size)

    def autofit(self, cols=None):
        self.adjust_columns = cols or {}

    def before(self, workbook, _, *args, **kw):
        self.worksheet = workbook.add_worksheet(*args, **kw)
        self.worksheet.adjust_column_size = self.adjust_column_size

        return self.worksheet

    def after(self, workbook, worksheet, children, *args, **kw):
        if self.adjust_columns:
            self._autofit()


class CellRef(str):
    @xlsxwriter.worksheet.convert_cell_args
    def __new__(cls, row, col):
        self = super().__new__(cls, xlsxwriter.utility.xl_rowcol_to_cell(row, col))
        self.row = row
        self.col = col

        return self

    def row(self, row_num):
        return self.__class__(row_num, self.col)

    def col(self, col_num):
        return self.__class__(self.row, col_num)

    def __add__(self, delta):
        if isinstance(delta, int):
            delta = CellsDelta(0, delta)

        if not isinstance(delta, CellsDelta):
            raise ValueError('Invalid cell delta: {}'.format(delta))

        return self.__class__(*delta.add(self.row, self.col))

    def __sub__(self, delta):
        return self + -delta


class CellsDelta:
    def __init__(self, row_delta=0, col_delta=0):
        self.row_delta = row_delta
        self.col_delta = col_delta

    def __neg__(self):
        return CellsDelta(-self.row_delta, -self.col_delta)

    def add(self, row, col):
        return row + self.row_delta, col + self.col_delta


class RowsDelta(CellsDelta):
    def __init__(self, row_delta):
        super().__init__(row_delta, 0)


class ColumnsDelta(CellsDelta):
    def __init__(self, col_delta):
        super().__init__(0, col_delta)


class Column(Tag):
    def after(self, workbook, worksheet, children, *args, **kw):
        row, col, values, *args = self.convert_cell_args(*args, *children)
        for value in values:
            worksheet.adjust_column_size(col, value)

        return worksheet.write_column(row, col, values, *args, **kw)


class Row(Tag):
    def after(self, workbook, worksheet, children, *args, **kw):
        row, col, values, *args = self.convert_cell_args(*args, *children)
        for c, value in enumerate(values, col):
            worksheet.adjust_column_size(c, value)

        return worksheet.write_row(row, col, values, *args, **kw)


class Cell(Tag):
    def after(self, workbook, worksheet, children, *args, **kw):
        row, col, value, *args = self.convert_cell_args(*args, *children)
        worksheet.adjust_column_size(col, value)

        return worksheet.write(row, col, value, *args, **kw)


class FormulaCell(Tag):
    def after(self, workbook, worksheet, children, *args, **kw):
        row, col, value, *args = self.convert_cell_args(*args, *children)
        worksheet.adjust_column_size(col, value)

        return worksheet.write_formula(row, col, value, *args, **kw)


class StrCell(Tag):
    def after(self, workbook, worksheet, children, *args, **kw):
        row, col, value, *args = self.convert_cell_args(*args, *children)
        worksheet.adjust_column_size(col, value)

        return worksheet.write_string(row, col, value, *args, **kw)


class UrlCell(Tag):
    def after(self, workbook, worksheet, children, *args, **kw):
        row, col, value, *args = self.convert_cell_args(*args, *children)
        worksheet.adjust_column_size(col, value)

        return worksheet.write_url(row, col, value, *args, **kw)


class RichCell(Tag):
    def after(self, workbook, worksheet, children, *args, **kw):
        return worksheet.write_rich_string(*args, *children, **kw)


class CommentCell(Tag):
    def after(self, workbook, worksheet, children, *args, **kw):
        return worksheet.write_comment(*args, *children, **kw)


class Table(Tag):
    @staticmethod
    def before(workbook, worksheet, *args, **kw):
        return worksheet.add_table(*args, kw)


class Sparkline(Tag):
    @staticmethod
    def before(workbook, worksheet, *args, **kw):
        return worksheet.add_sparkline(*args, kw)


class TextBox(Tag):
    def after(self, workbook, worksheet, children, *args, **kw):
        return worksheet.insert_textbox(*args, *children, kw)


class Image(Tag):
    @staticmethod
    def before(workbook, worksheet, *args, **kw):
        return worksheet.insert_image(*args, kw)


class ChartCell(Tag):
    def after(self, workbook, worksheet, children, *args, **kw):
        return worksheet.insert_chart(*args, *children, kw)


class Chart(Tag):
    def before(self, workbook, worksheet, *args, **kw):
        return workbook.add_chart(*args, kw)


class Series(Tag):
    @staticmethod
    def before(workbook, chart, *args, **kw):
        return chart.add_series(*args, kw)


# =====================================================================================================================


class TagProp:
    def __init__(self, name, factory):
        self.factory = factory

    def __get__(self, renderer, cls):
        return self.factory(renderer)


class Renderer:
    workbook = TagProp('workbook', Workbook)
    worksheet = TagProp('worksheet', Worksheet)
    column = TagProp('column', Column)
    row = TagProp('row', Row)
    cell = TagProp('cell', Cell)
    formula_cell = TagProp('formula_cell', FormulaCell)
    str_cell = TagProp('str_cell', StrCell)
    url_cell = TagProp('url_cell', UrlCell)
    rich_cell = TagProp('rich_cell', RichCell)
    comment_cell = TagProp('comment_cell', CommentCell)
    table = TagProp('table', Table)
    sparkline = TagProp('sparkline', Sparkline)
    textbox = TagProp('textbox', TextBox)
    image = TagProp('image', Image)
    chart = TagProp('chart', Chart)
    chart_cell = TagProp('chart_cell', ChartCell)
    series = TagProp('series', Series)
    format = TagProp('format', Format)

    def __init__(self):
        self._children = [[]]

    def new(self, parent, component):
        return self.__class__()

    def start_rendering(self, view_name, args, kw):
        pass

    def end_rendering(self, rendering):
        return rendering

    @property
    def root(self):
        return self._children[0][0]

    def enter(self, current):
        self._children[-1].append(current)
        self._children.append([])

    def exit(self, current):
        current(*self._children.pop())

    def __lshift__(self, current):
        if not isinstance(current, str) and isinstance(current, Iterable):
            self._children[-1].extend(current)
        else:
            self._children[-1].append(current)

        return self

    @staticmethod
    def generate_id(prefix=''):
        return prefix + str(random.randint(10000000, 99999999))  # noqa: S311

    ref = CellRef
    cells = CellsDelta
    rows = RowsDelta
    columns = ColumnsDelta

    @staticmethod
    def col_name(column_num, column_abs=False):
        return xlsxwriter.utility.xl_col_to_name(column_num, column_abs)

    @staticmethod
    def col_number(column_name):
        return xlsxwriter.utility.xl_cell_to_rowcol(column_name + '1')[1]

    @staticmethod
    def cell_name(row_num, column_num, row_abs=False, column_abs=False):
        return xlsxwriter.utility.xl_rowcol_to_cell(row_num, column_num, row_abs, column_abs)

    @staticmethod
    def cell_number(cell_name):
        return xlsxwriter.utility.xl_cell_to_rowcol(cell_name)

    @staticmethod
    def range_name(first_row, first_column, last_row, last_column):
        return xlsxwriter.utility.xl_range(first_row, first_column, last_row, last_column)

    @staticmethod
    def range(from_cell_ref, to_cell_ref):
        return f'{from_cell_ref}:{to_cell_ref}'
