class Format:
    def __init__(self, fmt=None):
        self.fmt = fmt or {}

    def data_format(self, cell_data, rowidx, colidx, first, last):
        return self.fmt


class Bold(Format):
    @classmethod
    def data_format(cls, cell_data, rowidx, colidx, first, last):
        return {
            'bold': True
        }


class Centered(Format):
    @classmethod
    def data_format(cls, cell_data, rowidx, colidx, first, last):
        return {
            'align': 'center',
            'valign': 'vcenter'
        }


class LastUnderlined(Format):
    @classmethod
    def data_format(cls, cell_data, rowidx, colidx, first, last):
        if last:
            return {
                'bottom': 1
            }
        return {}


class ColumnBorder(Format):
    @classmethod
    def data_format(cls, cell_data, rowidx, colidx, first, last):
        return {
            'left': 1,
            'right': 1
        }
