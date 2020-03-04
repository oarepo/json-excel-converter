import json

import xlsxwriter

from json_excel_converter import Writer as bWriter


class Formatter:
    def __init__(self, formats):
        self.workbook = None
        self.formats = formats
        self.format_cache = {}

    def format(self, cell_data, rowidx, colidx, first, last):
        fmt = {}
        for format in self.formats:
            fmt.update(format.data_format(cell_data, rowidx, colidx, first, last))
        strfmt = json.dumps(fmt, sort_keys=True)
        if strfmt in self.format_cache:
            return self.format_cache[strfmt]
        fmt = self.workbook.add_format(fmt)
        self.format_cache[strfmt] = fmt
        return fmt


class Token:
    def __eq__(self, other):
        return self is other

    def __hash__(self):
        return id(self)


DEFAULT_COLUMN_WIDTH = Token()
DEFAULT_ROW_HEIGHT = Token()


class Writer(bWriter):
    """
    XLSXWriter can not reset the sheet, so all the data are cached and written out on "finish"
    """

    def __init__(self, file=None, workbook=None, sheet=None,
                 sheet_name=None, start_row=1, start_col=0,
                 header_formats=(), data_formats=(),
                 column_widths=None, row_heights=None):
        super().__init__()
        self.file = file
        self.workbook = workbook
        self.sheet = sheet
        self.sheet_name = sheet_name
        self.headers = []
        self.rows = []
        self.raw = []
        self.current_row = start_row
        self.start_row = start_row
        self.start_col = start_col
        self.header_formatter = Formatter(header_formats)
        self.data_formatter = Formatter(data_formats)
        self.column_widths = column_widths or {}
        self.row_heights = row_heights or {}

    def start(self):
        self.headers = []
        self.rows = []
        self.raw = []
        self.current_row = 0

    def reset(self):
        self.headers = []
        self.rows = []
        self.raw = []

    def finish(self):
        close = not self.sheet
        if not self.sheet:
            self.workbook = xlsxwriter.Workbook(self.file)
            self.sheet = self.workbook.add_worksheet(self.sheet_name)
        self.header_formatter.workbook = self.workbook
        self.data_formatter.workbook = self.workbook
        self.before_write()
        self.before_headers()
        for header_idx, h in enumerate(self.headers):
            self.output_header_row(h, header_idx)
        self.after_headers()
        self.before_rows()
        for row_idx, r in enumerate(self.rows):
            self.output_row(r, row_idx, first=not row_idx,
                            last=row_idx == len(self.rows) - 1,
                            raw=self.raw[row_idx])
        self.after_rows()
        self.after_write()

        if DEFAULT_COLUMN_WIDTH in self.column_widths:
            max_col = 0
            for header in self.headers:
                col = 0
                for cell in header:
                    col += cell.columns
                if col > max_col:
                    max_col = col
            for col in range(max_col):
                self.sheet.set_column(col, col, self.column_widths[DEFAULT_COLUMN_WIDTH])

        for header in self.headers:
            col = 0
            for cell in header:
                if cell.path in self.column_widths:
                    self.sheet.set_column(col, col + cell.columns - 1,
                                          self.column_widths[cell.path])
                col += cell.columns

        if DEFAULT_ROW_HEIGHT in self.row_heights:
            for row_idx in range(self.current_row):
                self.sheet.set_row(row_idx, self.row_heights[DEFAULT_ROW_HEIGHT])

        for row_idx, height in self.row_heights.items():
            if row_idx is DEFAULT_ROW_HEIGHT:
                continue
            self.sheet.set_row(row_idx, height)

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
            self.header_formatter.format(cell_data, header_idx, col, first, last), data=None)

    def output_row(self, row, row_idx, first, last, raw):
        col = self.start_col
        for r in row:
            col = self.output_row_cell(col, r, row_idx, first, last, raw)
        self.current_row += 1

    def output_row_cell(self, col, cell_data, row_idx, first, last, raw):
        return self.output_cell(col, cell_data,
                                self.data_formatter.format(cell_data, row_idx, col, first, last),
                                raw)

    def output_cell(self, col, cell_data, cell_format, data):
        if cell_data.columns > 1 or cell_data.span > 1:
            self.write_cell_range(self.current_row, col,
                                  self.current_row + cell_data.span - 1,
                                  col + cell_data.columns - 1, cell_data, cell_format)
        else:
            if cell_data.value != '':
                self.write_cell(self.current_row, col, cell_data, cell_format, data)
        return col + cell_data.columns

    def write_cell(self, row, col, cell_data, cell_format, data):
        if cell_data.value != '':
            if cell_data.url:
                self.sheet.write_url(row, col, cell_data.url,
                                     string=cell_data.value, cell_format=cell_format)
            else:
                self.sheet.write(row, col, cell_data.value, cell_format)

    def write_cell_range(self, row, col, trow, tcol, cell_data, cell_format):
        self.sheet.merge_range(row, col, trow, tcol, cell_data.value, cell_format)

    def write_header(self, header):
        self.headers.append((list(header)))

    def write_row(self, row, data):
        self.rows.append((list(row)))
        self.raw.append(data)
