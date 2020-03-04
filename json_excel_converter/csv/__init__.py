import csv
from io import StringIO

from json_excel_converter import Writer as bWriter


class Writer(bWriter):
    def __init__(self, file=None):
        super().__init__()
        if file is None:
            file = StringIO()
        self.file = file
        self.fd = None
        self.csv = None

    def start(self):
        if isinstance(self.file, str):
            self.fd = open(self.file, 'w')
        else:
            self.fd = self.file
        self.csv = csv.writer(self.fd)

    def reset(self):
        self.fd.seek(0)
        self.fd.truncate()

    def finish(self):
        self.fd.flush()
        if isinstance(self.file, str):
            self.fd.close()
        self.fd = None

    def write_header(self, header):
        self.write_row(header, None)

    def write_row(self, row, data):
        out = []
        for h in row:
            v = h.value
            if v is None:
                v = ''
            out.append(v)
            if h.columns > 1:
                for _ in range(h.columns - 1):
                    out.append('')
        self.csv.writerow(out)
