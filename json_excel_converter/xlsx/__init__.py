import functools
import json

import xlsxwriter
from io import BytesIO

from json_excel_converter import Writer as bWriter


class Formatter:
    def __init__(self, formats):
        self.workbook = None
        self.formats = formats
        self.format_cache = {}

    def merge_format(self, is_header, cell_data, rowidx, colidx, first, last):
        my_format = self.format_data(is_header, cell_data, rowidx, colidx, first, last)
        if self.previous:
            prev_format = self.previous.merge_format(is_header, cell_data,
                                                     rowidx, colidx, first, last)
            return {
                **prev_format,
                **my_format
            }
        else:
            return my_format

    def format(self, is_header, cell_data, rowidx, colidx, first, last):
        fmt = {}
        for format in self.formats:
            fmt.update(format.data_format(is_header, cell_data, rowidx, colidx, first, last))
        strfmt = json.dumps(fmt, sort_keys=True)
        if strfmt in self.format_cache:
            return self.format_cache[strfmt]
        fmt = self.workbook.add_format(fmt)
        self.format_cache[strfmt] = fmt
        return fmt

    def format_header(self, cell_data, rowidx, colidx, first, last):
        return self.format(True, cell_data, rowidx, colidx, first, last)

    def format_data(self, cell_data, rowidx, colidx, first, last):
        return self.format(False, cell_data, rowidx, colidx, first, last)


class Writer(bWriter):
    """
    XLSXWriter can not reset the sheet, so all the data are cached and written out on "finish"
    """

    def __init__(self, file=None, workbook=None, sheet=None,
                 sheet_name=None, start_row=1, start_col=0,
                 formats=()):
        super().__init__()
        self.file = file
        self.workbook = workbook
        self.sheet = sheet
        self.sheet_name = sheet_name
        self.headers = []
        self.rows = []
        self.current_row = start_row
        self.start_row = start_row
        self.start_col = start_col
        self.formatter = Formatter(formats)

    def start(self):
        self.headers = []
        self.rows = []
        self.current_row = 0

    def reset(self):
        self.headers = []
        self.rows = []

    def finish(self):
        close = not self.sheet
        if not self.sheet:
            self.workbook = xlsxwriter.Workbook(self.file)
            self.sheet = self.workbook.add_worksheet(self.sheet_name)
        self.formatter.workbook = self.workbook
        self.before_write()
        self.before_headers()
        for header_idx, h in enumerate(self.headers):
            self.output_header_row(h, header_idx)
        self.after_headers()
        self.before_rows()
        for row_idx, r in enumerate(self.rows):
            self.output_row(r, row_idx, first=not row_idx, last=row_idx == len(self.rows) - 1)
        self.after_rows()
        self.after_write()
        if close:
            self.workbook.close()

    def before_write(self):
        # might be overwritten by descendants
        pass

    def after_write(self):
        # might be overwritten by descendants
        pass

    def before_headers(self):
        # might be overwritten by descendants
        pass

    def after_headers(self):
        # might be overwritten by descendants
        pass

    def before_rows(self):
        # might be overwritten by descendants
        pass

    def after_rows(self):
        # might be overwritten by descendants
        pass

    def output_header_row(self, header, header_idx):
        col = self.start_col
        for idx, h in enumerate(header):
            if h.has_children:
                span = 1
            else:
                span = len(self.headers) - header_idx

            col = self.output_header_cell(col, h, span=span, header_idx=header_idx,
                                          first=not header_idx, last=not h.has_children)
        self.current_row += 1

    def output_header_cell(self, col, cell_data, span, header_idx, first, last):
        cell_data.span = span
        return self.output_cell(
            col, cell_data,
            self.formatter.format_header(cell_data, header_idx, col, first, last))

    def output_row(self, row, row_idx, first, last):
        col = self.start_col
        for r in row:
            col = self.output_row_cell(col, r, row_idx, first, last)
        self.current_row += 1

    def output_row_cell(self, col, cell_data, row_idx, first, last):
        return self.output_cell(col, cell_data,
                                self.formatter.format_data(cell_data, row_idx, col, first, last))

    def output_cell(self, col, cell_data, cell_format):
        if cell_data.columns > 1 or cell_data.span > 1:
            self.sheet.merge_range(self.current_row, col, self.current_row + cell_data.span - 1,
                                   col + cell_data.columns - 1, cell_data.value, cell_format)
        else:
            if cell_data.value != '':
                self.sheet.write(self.current_row, col, cell_data.value, cell_format)
        return col + cell_data.columns

    def write_header(self, header):
        self.headers.append((list(header)))

    def write_row(self, row):
        self.rows.append((list(row)))
