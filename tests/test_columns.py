from json_excel_converter.linearize import Columns, Value
from json_excel_converter.options import Options


def test_empty():
    cols = Columns()
    cols.check({})
    assert [] == list(cols.output({}))
    assert len(cols.columns) == 0
    assert cols.columns_taken == 0


def test_simple_data():
    cols = Columns()
    val = {'a': 'aa'}
    cols.check(val)
    assert list(cols.output(val)) == [Value('aa', 1, path='a')]
    assert len(cols.columns) == 1
    assert cols.columns_taken == 1


def test_simple_data2():
    cols = Columns()
    val = {'a': 'aa', 'b': 'bb'}
    cols.check(val)
    assert list(cols.output(val)) == [Value('aa', 1, path='a'), Value('bb', 1, path='b')]
    assert len(cols.columns) == 2
    assert cols.columns_taken == 2


def test_array():
    cols = Columns()
    vals = {'a': ['a', 'aa'], 'b': 'bb'}
    cols.check(vals)
    assert list(cols.output(vals)) == [
        Value('a', 1, path='a'), Value('aa', 1, path='a'), Value('bb', 1, path='b')
    ]
    assert len(cols.columns) == 2
    assert cols.columns_taken == 3
    assert cols.depth == 1
    assert list(cols.get_header_row(0)) == [
        Value('a', 1, path='a'), Value('a', 1, path='a'), Value('b', 1, path='b')
    ]


def test_nested():
    cols = Columns()
    vals = {'a': [{'c': 'c1'}, {'c': 'c2'}], 'b': 'bb'}
    cols.check(vals)

    assert cols.depth == 2
    assert list(cols.get_header_row(0)) == [
        Value('a', 1, path='a', has_children=True),
        Value('a', 1, path='a', has_children=True),
        Value('b', 1, path='b')
    ]
    assert list(cols.get_header_row(1)) == [
        Value('c', 1, path='a.c'), Value('c', 1, path='a.c'), Value('', 1, path='b')
    ]

    assert list(cols.output(vals)) == [
        Value('c1', 1, path='a.c'), Value('c2', 1, path='a.c'), Value('bb', 1, path='b')
    ]
    assert len(cols.columns) == 2
    assert cols.columns_taken == 3


def test_fields():
    cols = Columns(Options(fields=['a']))
    val = {'a': 'aa', 'b': 'bb'}
    cols.check(val)
    assert list(cols.output(val)) == [Value('aa', 1, path='a')]
    assert len(cols.columns) == 1
    assert cols.columns_taken == 1


def test_excludes():
    cols = Columns(Options(excludes=['a']))
    val = {'a': 'aa', 'b': 'bb'}
    cols.check(val)
    assert list(cols.output(val)) == [Value('bb', 1, path='b')]
    assert len(cols.columns) == 1
    assert cols.columns_taken == 1


def test_ordering():
    cols = Columns(Options(ordering=['b', 'a']))
    val = {'a': 'aa', 'b': 'bb'}
    cols.check(val)
    assert list(cols.output(val)) == [Value('bb', 1, path='b'), Value('aa', 1, path='a')]
    assert len(cols.columns) == 2
    assert cols.columns_taken == 2


def test_array_header_translation():
    def translator(header, idx, cardinality):
        if cardinality > 1:
            return '%s %s' % (header, idx + 1)
        else:
            return header

    cols = Columns(options=Options(header_translator=translator))
    vals = {'a': ['a', 'aa'], 'b': 'bb'}
    cols.check(vals)
    assert list(cols.get_header_row(0)) == [
        Value('a 1', 1, path='a'), Value('a 2', 1, path='a'), Value('b', 1, path='b')
    ]
