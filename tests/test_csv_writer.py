from json_excel_converter.csv import Writer
from json_excel_converter.linearize import Value


def test_writer():
    w = Writer()
    w.start()
    w.write_header([
        Value('a', 1),
        Value('b', 2),
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
    assert w.file.getvalue() == (
            'a,b,,c\r\n' +
            ',b1,b2,\r\n' +
            '1,2,3,4\r\n' +
            '1,2,3,4\r\n'
    )
