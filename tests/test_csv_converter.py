from json_excel_converter import Converter
from json_excel_converter.csv import Writer


def test_csv_converter():
    data = [
        {
            'a': [
                {
                    'b': '1'
                },
                {
                    'c': '2'
                }
            ],
            'd': 'abc'
        }
    ]
    data += data

    conv = Converter()
    w = Writer()
    conv.convert(data, w)
    assert w.file.getvalue().strip().replace('\r\n', '\n') == """
a,,a,,d
b,c,b,c,
1,,,2,abc
1,,,2,abc
    """.strip().replace('\r\n', '\n')


def test_array_grow():
    data = [
        {'a': ['1']},
        {'a': ['1', '2']},
        {'a': ['1', '2', '3']},
    ]
    conv = Converter()
    w = Writer()
    conv.convert(data, w)
    assert w.file.getvalue().strip().replace('\r\n', '\n') == """
a,a,a
1,,
1,2,
1,2,3
    """.strip().replace('\r\n', '\n')
