from .linearize import Columns, LinearizationError
from .options import Options, EMPTY_OPTIONS


class Converter:
    def __init__(self, options=EMPTY_OPTIONS):
        self.options = options
        self.conv = None

    def convert(self, data, writer):
        self.conv = None
        while True:
            try:
                self.convert_streaming(data, writer)
                break
            except LinearizationError:
                writer.reset()

    def convert_streaming(self, data, writer):
        writer.start()
        if not self.conv:
            self.conv = Columns(options=self.options)
        for idx, d in enumerate(data):
            errors = self.conv.check(d)
            if not idx:
                self._write_header(writer, self.conv)
            elif errors:
                raise LinearizationError(errors)
            row = self.conv.output(d)
            writer.write_row(row)
        writer.finish()
        self.conv = None

    def reset(self):
        self.conv = None

    @staticmethod
    def _write_header(writer, columns):
        depth = columns.depth
        for d in range(depth):
            writer.write_header(columns.get_header_row(d))
