class Format:
    @classmethod
    def data_format(self, is_header, cell_data, rowidx, colidx, first, last):
        """
        Override this method to get custom formatting
        """
        return {}


class HeaderFormat(Format):
    def __init__(self, fmt):
        self.fmt = fmt

    def data_format(self, is_header, cell_data, rowidx, colidx, first, last):
        if is_header:
            return self.fmt
        return {}


class BoldHeaderFormat(Format):
    @classmethod
    def data_format(self, is_header, cell_data, rowidx, colidx, first, last):
        if is_header:
            return {
                'bold': True
            }
        return {}


class CenteredHeaderFormat(Format):
    @classmethod
    def data_format(self, is_header, cell_data, rowidx, colidx, first, last):
        if is_header:
            return {
                'align': 'center',
                'valign': 'vcenter'
            }
        return {}


class UnderlinedHeaderFormat(Format):
    @classmethod
    def data_format(self, is_header, cell_data, rowidx, colidx, first, last):
        if is_header and last:
            return {
                'bottom': 1
            }
        return {}


class ColumnBorderFormat(Format):
    @classmethod
    def data_format(self, is_header, cell_data, rowidx, colidx, first, last):
        return {
            'left': 1,
            'right': 1
        }


class DataFormat(Format):
    def __init__(self, fmt):
        self.fmt = fmt

    def data_format(self, is_header, cell_data, rowidx, colidx, first, last):
        if not is_header:
            return self.fmt
        return {}
