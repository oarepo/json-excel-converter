class Options:

    def __init__(self, cardinality=1,
                 ordering=None, fields=None, excludes=None,
                 sort_key=None,
                 header_translator=None,
                 value_translator=None,
                 parent=None):
        self.children = {}
        self.cardinality = cardinality
        self.ordering = ordering
        self.sort_key = sort_key
        self.fields = fields
        self.excludes = excludes or set()
        self.header_translator = header_translator or (
            parent.header_translator if parent else
            (lambda header, path, index, cardinality: header)
        )
        self.value_translator = value_translator or (
            parent.value_translator if parent else
            (lambda value, path, index, cardinality: value)
        )

    def __getitem__(self, item):
        if item in self.children:
            return self.children[item]
        else:
            opts = self.children[item] = Options(parent=self)
            return opts

    def __setitem__(self, key, value):
        self.children[key] = value

    @property
    def ordering(self):
        return self._ordering

    @ordering.setter
    def ordering(self, value):
        if value:
            self._ordering = {
                k: idx for idx, k in enumerate(value)
            }
        else:
            self._ordering = {}

    @property
    def sort_key(self):
        if self._sort_key:
            return self._sort_key
        return lambda x: (self.ordering.get(x, len(self.ordering)), x)

    @sort_key.setter
    def sort_key(self, func):
        self._sort_key = func


EMPTY_OPTIONS = Options()
