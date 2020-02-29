# JSON to excel converter

![](https://img.shields.io/github/license/oarepo/json-excel-converter.svg)
![](https://img.shields.io/travis/oarepo/json-excel-converter.svg)
![](https://img.shields.io/coveralls/oarepo/json-excel-converter.svg)
![](https://img.shields.io/pypi/v/json-excel-converter.svg)

A package that converts json to CSV, excel or other table formats

## Sample output

### Simple json

```json
[
  {
    "col1": "val1",
    "col2": "val2" 
  }
]
```

the generated CSV/excel is:

```
col1          col2
==================
val1          val2
```

### Nested json

```json
[
  {
    "col1": "val1",
    "col2": {
      "col21": "val21",
      "col22": "val22"
    }
  }
]
```

the generated CSV/excel is (in excel, col2 spans two cells horizontally):

```
col1          col2
              col21         col22
=================================
val1          val21         val22
```

### json with array property

```json
[
  {
    "col1": "val1",
    "col2": [
      {
        "col21": "val21"
      },
      {
        "col21": "val22"
      }
    ]
  }
]
```

the generated CSV/excel is (in excel, col2 spans two cells horizontally):

```
col1          col2         
              col21         col21
=================================
val1          val21         val22
```


## Installation

```bash
pip install json-excel-converter[extra]
```

where extra is:

 * ``xlsxwriter`` to use the xlsxwriter

## Usage

### Simple usage

```python

from json_excel_converter import Converter 
from json_excel_converter.xlsx import Writer

data = [
    {'a': [1], 'b': 'hello'},
    {'a': [1, 2, 3], 'b': 'world'}
]

conv = Converter()
conv.convert(data, Writer(file='/tmp/test.xlsx'))
```

### Streaming usage with restarts

```python

from json_excel_converter import Converter, LinearizationError 
from json_excel_converter.csv import Writer

conv = Converter()
writer = Writer(file='/tmp/test.csv')
while True:
    try:
        data = get_streaming_data()     # custom function to get iterator of data
        conv.convert_streaming(data, writer)
        break
    except LinearizationError:
        pass
```

### Arrays

When first row is passed, the library creates the layout of columns. In case of arrays,
a column (or more columns if the array contains json objects) is created for each
of the items in the array.

Then on second row it might happen that the array contains more items. The library reacts
by adjusting the number of columns in the layout and raising ``LinearizationError``.

``Converter.convert`` captures this error and restarts the processing. In case of CSV
this means truncating the output file to 0 bytes and processing the data again.

``Converter.convert_streaming`` just raises this exception as data stream does not 
generally supports rewinding.

If you know the size of the array in advance, you might pass it in options. Then no
processing restarts are required and ``LinearizationError`` is not raised.

 ```python

from json_excel_converter import Converter, Options
from json_excel_converter.xlsx import Writer

data = [
    {'a': [1]},
    {'a': [1, 2, 3]}
]
options = Options()
options['a'].cardinality = 3

conv = Converter(options=options)
conv.convert(data, Writer(file='/tmp/test.xlsx'))
```

### XLSX Formatting

For xlsx, the library can format the output as well. To do so, pass the writer and array
of formats from the ``json_excel_converter.xlsx.formats`` package or create your own.

```python
from json_excel_converter import Converter

from json_excel_converter.xlsx import Writer
from json_excel_converter.xlsx.formats import UnderlinedHeaderFormat, BoldHeaderFormat, \
    CenteredHeaderFormat, HeaderFormat

data = [
    {'a': 'Hello'},
    {'a': 'World'}
]

w = Writer('/tmp/test3.xlsx',
           formats=(CenteredHeaderFormat, BoldHeaderFormat, UnderlinedHeaderFormat, 
                    HeaderFormat({
                        'font_color': 'red'
                    })))

conv = Converter()
conv.convert(data, w)
```

See https://xlsxwriter.readthedocs.io/format.html for details on formats in xlsxwriter