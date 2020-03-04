from json_excel_converter import Converter, Options
from json_excel_converter.xlsx import Writer, DEFAULT_COLUMN_WIDTH, DEFAULT_ROW_HEIGHT
from json_excel_converter.xlsx.formats import (
    LastUnderlined,
    ColumnBorder,
    Bold,
    Centered, Format)
from json_excel_converter.linearize import Value


def test_writer():
    w = Writer('/tmp/test.xlsx',
               header_formats=(Centered, Bold, LastUnderlined,))
    w.start()
    w.write_header([
        Value('a', 1),
        Value('b', 2, has_children=True),
        Value('c', 1)
    ])
    w.write_header([
        Value(None, 1),
        Value('b1', 1),
        Value('b2', 1),
        Value(None, 1)
    ])
    w.write_row([Value(1), Value(2), Value(3), Value(4)], data=None)
    w.write_row([Value('1'), Value('2'), Value('3'), Value('4')], data=None)
    w.finish()


def test_writer_2():
    w = Writer('/tmp/test2.xlsx',
               header_formats=(Centered, Bold,
                               LastUnderlined, ColumnBorder,),
               data_formats=(ColumnBorder,))
    w.start()
    w.write_header([
        Value('a', 1),
        Value('b', 2, has_children=True),
        Value('c', 1)
    ])
    w.write_header([
        Value(None, 1),
        Value('b1', 1),
        Value('b2', 1),
        Value(None, 1)
    ])
    w.write_row([Value(1), Value(2), Value(3), Value(4)], data=None)
    w.write_row([Value('1'), Value('2'), Value('3'), Value('4')], data=None)
    w.finish()


def test_red_header_url_link():
    data = [
        {'a': 'Hello', 'b': 'aaa'},
        {'a': 'World', 'b': 'bbb'}
    ]

    class UrlWriter(Writer):
        def write_cell(self, row, col, cell_data, cell_format, data):
            if cell_data.path == 'a' and data:
                self.sheet.write_url(row, col,
                                     'https://test.org/' + data['b'],
                                     string=cell_data.value)
            else:
                super().write_cell(row, col, cell_data, cell_format, data)

    w = UrlWriter('/tmp/test3.xlsx',
                  header_formats=(
                      Centered, Bold, LastUnderlined,
                      Format({
                          'font_color': 'red'
                      })
                  ),
                  data_formats=(
                      Format({
                          'font_color': 'green'
                      }),
                  ),
                  column_widths={
                      'a': 30,
                      DEFAULT_COLUMN_WIDTH: 20
                  },
                  row_heights={
                      DEFAULT_ROW_HEIGHT: 20,
                      0: 10,
                      1: 30
                  })

    conv = Converter()
    conv.convert(data, w)


def test_url():
    data = [
        {'a': 'https://google.com'},
    ]

    w = Writer('/tmp/test4.xlsx',
               header_formats=(
                   Centered, Bold, LastUnderlined,
               )
               )
    options = Options()
    options['a'].url = lambda data: data['a']
    conv = Converter(options)
    conv.convert(data, w)
