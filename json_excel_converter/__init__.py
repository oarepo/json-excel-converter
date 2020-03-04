from .linearize import LinearizationError, Value
from .converter import Converter
from .options import Options


class Writer:
    def __init__(self):
        pass

    def start(self):
        """
        Called before any headers/rows are output
        """
        pass  # pragma: no cover

    def reset(self):
        """
        Called when the process has not finished successfully. Should be used to clean up any
        resources
        """
        pass  # pragma: no cover

    def write_header(self, header):
        """
        Writes a header. Can be called multiple times to write subheaders
        :param header: a list of (headerName, colspan) if there is a header
               or (None, colspan) if the space should be left empty
        """
        raise NotImplemented()  # pragma: no cover

    def write_row(self, row, data):
        """
        Writes a row
        :param row: a list of values, a value can be None to leave its space empty
        """
        raise NotImplemented()  # pragma: no cover

    def finish(self):
        """
        Called when all the rows have been written successfully
        """
        pass  # pragma: no cover
