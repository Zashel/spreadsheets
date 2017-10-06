from functools import reduce

class CoordinatesError(Exception): pass

def sum_slices(*args):
    assert len(args) >= 1
    assert all([isinstance(arg, slice) for arg in args])
    return reduce(lambda x, y: slice(x.start+y.start, x.stop+y.stop), args)

def sub_slices(*args):
    assert len(args) >= 1
    assert all([isinstance(arg, slice) for arg in args])
    return reduce(lambda x, y: slice(x.start-y.start, x.stop-y.stop), args)

class _RelativeCell:
    def __init__(self, coordinates):
        self.coordinates = coordinates

    def __call__(self, cell):
        assert isinstance(cell, Cell)
        input(cell.coordinates)
        return cell.spreadsheet.__getitem__(sum_slices(cell.coordinates, self.coordinates))

class _RelativeCells:
    def __init__(self, coordinates):
        self.coordinates = coordinates

    def __call__(self, cell):
        assert isinstance(cell, Cell)
        start = self.coordinates.start(cell).coordinates
        stop = self.coordinates.stop(cell).coordinates
        return cell.spreadsheet.subslice(start, stop)

class _Relative:
    class Cell:
        def __getitem__(self, item):
            assert isinstance(item, slice)
            return _RelativeCell(item)
    class Cells:
        def __getitem__(self, item):
            assert isinstance(item, slice)
            assert all([isinstance(i, _RelativeCell) for i in (item.start, item.stop)])
            return _RelativeCells(item)

RelativeCell = _Relative.Cell()
RelativeCells = _Relative.Cells()

class Spreadsheet(list):
    """
    Spreadsheetclass to form Excel spreadsheets 12 in Sylk format
    """
    def __init__(self, data=None):
        list.__init__(self)
        if data is not None:
            self.extend(data)

    def __getitem__(self, item):
        """
        Gets either the row, either the location by a slice
        :param item: slice or int
        :return: Row or Item
        """
        if isinstance(item, slice):
            return list.__getitem__(self, item.stop)[item.start]
        return list.__getitem__(self, item)

    def __setitem__(self, key, item):
        """
        Sets items
        :param: key: index of item
        :param item: New item to set
        :return: None
        """
        if isinstance(key, slice):
            self[key.stop][key.start] = item

    def __sylk__(self):
        return b""

    def __csv__(self):
        pass

    @property
    def Columns(self):
        """
        Gests given columns
        :param coordinates: Slicefrom zashel.
        :return:
        """
        class ColumnsGenerator:
            def __init__(self, spreadsheet):
                self.spreadsheet = spreadsheet

            def __getitem__(self, coordinates):
                if isinstance(coordinates, slice):
                    final = list()
                    for x in range(coordinates.start, coordinates.stop+1):
                        this = list()
                        for row in self.spreadsheet:
                            if x >= len(row):
                                row[x] = None
                            this.append(row[x])
                        final.append(this)
                    return Columns(self.spreadsheet, final)
                else:
                    raise CoordinatesError("Coordinates may be slices with the form [first column:last column]")
        return ColumnsGenerator(self)

    @property
    def Rows(self):
        class RowsGenerator:
            def __init__(self, spreadsheet):
                self.spreadsheet = spreadsheet

            def __getitem__(self, coordinates):
                final = list()
                if isinstance(coordinates, slice):
                    for x in range(coordinates.start, coordinates.stop+1):
                        final.append(self.spreadsheet[x])
                    return Rows(self.spreadsheet, final)
                else:
                    raise CoordinatesError("Coordinates may be slices with the form [first row:last row]")

        return RowsGenerator(self)

    def append(self, item):
        """
        Appends only Rows to Spreadsheet
        :param item: Item to append to Spreadsheet
        :return: None
        """
        if any([isinstance(item, typo) for typo in (list, tuple)]):
            list.append(self, Rows(self, item))
        else:
            list.append(self, Rows(self, [item]))

    def extend(self, items):
        """
        Extends only Rows to Spreadsheet
        :param items: Items to extend spreadsheet with
        :return: None
        """
        for item in items:
            self.append(item)

    def subslice(self, start, stop):
        print(start, stop)
        final = list()
        for row_index in range(start.stop, stop.stop+1):
            row = list()
            for col_index in range(start.start, stop.start+1):
                row.append(self[col_index:row_index])
            final.append(row)
        return Spreadsheet(final)

    def to_sylk(self):
        pass

class Columns(list):
    """
    Column Class to Spreadsheet
    """
    def __init__(self, spreadsheet, *args, **kwargs):
        self._spreadsheet = spreadsheet
        list.__init__(self, *args, **kwargs)

    @property
    def spreadsheet(self):
        return self._spreadsheet

    def __sylk__(self):
        return b""

    def __csv__(self):
        pass

class Rows(list):
    """
    Row class to Spreadsheet
    """
    def __init__(self, spreadsheet, iterable=None):
        self._spreadsheet = spreadsheet
        if iterable is None:
            iterable = list()
        list.__init__(self, [isinstance(arg, Cell) and arg or Cell(spreadsheet, arg) for arg in iterable])

    def __setitem__(self, key, value):
        if isinstance(key, int):
            if key >= len(self):
                for x in range(len(self), key+1):
                    self.append(None)
                self[key] = value
            else:
                if isinstance(value, Cell):
                    self[key] = None
                    start = value.coordinates
                    self[key].value = RelativeCell.__getitem__(sub_slices(start, self[key].coordinates))
                    print("Vamos por aqui ", object.__getattribute__(self[key], "_value"))
                else:
                    self[key].value = value
        elif isinstance(key, slice):
            if any([isinstance(value, typo) for typo in (list, tuple)]):
                if abs(key.stop - key.start) == len(value)-1:
                    for index, x in enumerate(range(key.start, key.stop+1)):
                        self[x] = value[index]
                else:
                    raise TypeError("you can only assign iterables with the same length of slice")
            else:
                raise TypeError("you can only assign an iterable")

    def __sylk__(self):
        pass

    def __csv__(self):
        pass

    @property
    def spreadsheet(self):
        return self._spreadsheet

    def append(self, item):
        if not any([isinstance(item, typo) for typo in (list, tuple)]):
            item = [item]
        [list.append(self, isinstance(i, Cell) and i or Cell(self.spreadsheet, i)) for i in item]

    def extend(self, items):
        if any([isinstance(items, typo) for typo in (list, tuple)]):
            [self.append(item) for item in items]
        else:
            TypeError("values may be a list or a tuple")


class Cell:
    """
    Cell Class to Spreadsheet
    """
    def __init__(self, spreadsheet, value):
        object.__setattr__(self, "_spreadsheet", spreadsheet)
        object.__setattr__(self, "_value", value)

    def __getattribute__(self, item):
        if item in ("coordinates", "spreadsheet", "value"):
            return object.__getattribute__(self, item)
        else:
            return object.__getattribute__(self, "_value").__getattribute__(item)

    def __str__(self):
        return str(eval(self.__repr__()))

    def __repr__(self):
        return str(self.value.__repr__())

    @property
    def coordinates(self):
        for index, row in enumerate(self.spreadsheet):
            try:
                column = row.index(self)
            except ValueError:
                pass
            else:
                return slice(column, index)
        else:
            raise IndexError("Corrdinates not found")

    @property
    def spreadsheet(self):
        return object.__getattribute__(self, "_spreadsheet")

    @property
    def value(self):
        to_return = object.__getattribute__(self, "_value")
        if isinstance(to_return, _RelativeCell):
            return object.__getattribute__(to_return(self), "_value")
        else:
            return to_return

    @value.setter
    def value(self, value):
        object.__setattr__(self, "_value", value) #TODO -> Check what Instantiate (Item or Function)

class Function:
    """
    Function Class for Cells
    """
    def __init__(self, cell, value):
        self._cell = cell
        self._value = value

    def __sylk__(self):
        pass

    def __csv__(self):
        pass
