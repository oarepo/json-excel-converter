from collections import OrderedDict

from .options import Options


class Value:
    def __init__(self, value, columns=1, span=1, path=None, has_children=False):
        self.value = value
        self.columns = columns
        self.span = span
        self.path = path
        self.has_children = has_children

    def __eq__(self, other):
        return isinstance(other, Value) and self.value == other.value and \
               self.columns == other.columns and self.span == other.span and \
               self.has_children == other.has_children and self.path == other.path

    def __str__(self):
        return '%s/%s' % (repr(self.value), self.columns)

    def __repr__(self):
        return str(self)


class LinearizationError(Exception):
    """
    An error raised when column is an array that is bigger than the space allocated
    """

    def __init__(self, errors):
        super().__init__()
        self.errors = errors


class Column:
    def __init__(self, parent, name, path, cardinality, options):
        """
        Represents a column in excel

        :param parent:          parent columns
        :param name:            name of the column
        :param cardinality:     cardinality of the column - if the data are array,
                                the maximum length of the array
        :param options:         an instance of Options class
        """
        self.parent = parent
        self.name = name
        self.options = options
        self.cardinality = cardinality
        self.children = None  # an instance of Columns class
        self.path = path

    def check(self, value):
        """
        Returns a list of errors
        :param value: a value to be checked against this column
        """
        errors = []
        if isinstance(value, (list, tuple)):
            if len(value) > self.cardinality:
                self.cardinality = len(value)
                errors.append(('cardinality', self.path, self.cardinality, len(value)))
            for array_value in value:
                errors.extend(self.check(array_value))
        elif isinstance(value, dict):
            if not self.children:
                self.children = Columns(path=self.path, options=self.options)
                errors.append(('nochildren', self.path, value))
            errors.extend(self.children.check(value))
        elif value is not None:
            if self.children:
                raise ValueError(
                    'Inconsistent JSON: %s: sometimes a primitive is used, sometimes a dict. '
                    'Value %s, previous dict %s' % (self.path, value, self.children.columns))
        return errors

    def empty(self, already_output=0):
        """
        return empty cells that this column (with subcolumns) take
        :param already_output:  the number of columns (for cardinality>1)
                                that has already been output
        :return:    generator of columns
        """
        if self.cardinality > already_output:
            return Value('', (self.cardinality - already_output) * self.columns_taken)

    def output(self, value, index=0):
        """
        Output the value
        :param value:
        :return:
        """
        if isinstance(value, (list, tuple)):
            # output the values from the array
            for idx, v in enumerate(value):
                yield from self.output(v, idx)
            # output any extra empty space if more items are allocated
            if len(value) < self.cardinality:
                yield self.empty(already_output=len(value))
        elif isinstance(value, dict):
            yield from self.children.output(value)
        else:
            # otherwise it is a primitive value, so just return it
            yield Value(self.options.value_translator(value, self.path, index, self.cardinality),
                        path=self.path, columns=self.columns_taken)

    def get_header_row(self, level):
        """
        Returns the header row

        :param level:   which header row to return, 0 == main header, 1 == subheader, ...
        :return: a generator of Value instances
        """
        if not level:
            for c in range(self.cardinality):
                header_value = self.options.header_translator(
                    self.name, self.path, c, self.cardinality)
                if self.children:
                    yield Value(header_value, self.columns_taken,
                                path=self.path, has_children=True)
                else:
                    yield Value(header_value, path=self.path)
        elif not self.children:
            yield Value('', self.columns_taken * self.cardinality, path=self.path)
        else:
            for idx in range(self.cardinality):
                yield from self.children.get_header_row(level - 1)

    @property
    def columns_taken(self):
        """
        Returns a number of columns that a single instance (i.e. cardinality=1) takes
        """
        if not self.children:
            return 1
        return self.children.columns_taken

    @property
    def depth(self):
        if self.children:
            return 1 + self.children.depth
        return 1

    def __repr__(self):
        return self.name


class Columns:
    def __init__(self, options=None, path=None):
        """
        Creates a new instance, options is an instance of Options class
        :param options:
        """
        if path:
            self.path = path + '.'
        else:
            self.path = ''
        self.columns = OrderedDict()
        self.options = options or Options()

    def check(self, value):
        """
        Checks that the value is consistent with columns

        :param value:
        :return:   list of errors
        """
        errors = []
        pairs = []
        if self.options.fields:
            pairs = [
                (k, value[k])
                for k in self.options.fields
                if k in value
            ]
        elif self.options.excludes:
            pairs = [
                (k, value[k])
                for k in sorted(value.keys(), key=self.options.sort_key)
                if k not in self.options.excludes
            ]
        else:
            pairs = [
                (k, value[k])
                for k in sorted(value.keys(), key=self.options.sort_key)
            ]
        for k, v in pairs:
            if k in self.columns:
                errors.extend(self.columns[k].check(v))
            else:
                errors.append(('nochild', k, v))
                child_cardinality = self.options[k].cardinality
                column = Column(parent=self,
                                path=self.path + k,
                                name=k,
                                cardinality=child_cardinality,
                                options=self.options[k])
                self.columns[k] = column
                column.check(v)  # do not add to errors as this subtree is a new one
        return errors

    def output(self, json):
        for k, column in self.columns.items():
            if k not in json:
                yield column.empty()
            else:
                yield from column.output(json[k])

    @property
    def depth(self):
        if self.columns:
            return max(c.depth for c in self.columns.values())
        return 0

    def get_header_row(self, level):
        for c in self.columns.values():
            yield from c.get_header_row(level)

    @property
    def columns_taken(self):
        return sum(c.columns_taken * c.cardinality for c in self.columns.values())
