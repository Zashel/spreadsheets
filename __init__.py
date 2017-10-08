import datetime
import re
from functools import reduce

class CoordinatesError(Exception): pass

def sylk(item):
    """
    Gives the sylk representation of an object
    :param item: item to get the sylk representation of
    :return: sylk representation of item
    """
    if hasattr(item, "__sylk__"):
        return item.__sylk__()
    else:
        return item.__repr__()

def sum_slices(*args):
    assert len(args) >= 1
    assert all([isinstance(arg, slice) for arg in args])
    return reduce(lambda x, y: slice(x.start+y.start, x.stop+y.stop), args)

def sub_slices(*args):
    assert len(args) >= 1
    assert all([isinstance(arg, slice) for arg in args])
    return reduce(lambda x, y: slice(x.start-y.start, x.stop-y.stop), args)

def get_coordinates_by_name(name):
    name = name.upper()
    columns = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    col, row = re.findall(r"([A-Z]+)([0-9]+)", name)[0]
    final_col = int()
    for index, iter in enumerate(range(len(col)-1, -1, -1)):
        final_col += (pow((len(columns)), index))*(columns.index(col[iter])+1)
    return slice(final_col-1, int(row)-1)

class _Relatives:
    pass

class _RelativeCell(_Relatives):
    def __init__(self, coordinates):
        _Relatives.__init__(self)
        self.coordinates = coordinates

    def __call__(self, cell):
        assert isinstance(cell, Cell)
        return cell.spreadsheet.__getitem__(sum_slices(cell.coordinates, self.coordinates))

class _RelativeCells(_Relatives):
    def __init__(self, coordinates):
        _Relatives.__init__(self)
        self.coordinates = coordinates

    def __call__(self, cell):
        assert isinstance(cell, Cell)
        start = self.coordinates.start(cell).coordinates
        stop = self.coordinates.stop(cell).coordinates
        return cell.spreadsheet.subslice(start, stop)

class _RelativeColumns(_Relatives):
    def __init__(self, coordinates):
        _Relatives.__init__(self)
        assert isinstance(coordinates, slice)
        self.coordinates = coordinates

    def __call__(self, cell):
        assert isinstance(cell, Cell)
        start = cell.coordinates.start + self.coordinates.start
        stop = cell.coordinates.start + self.coordinates.stop
        return cell.spreadsheet.Columns[start, stop]

class _RelativeRows(_Relatives):
    def __init__(self, coordinates):
        _Relatives.__init__(self)
        assert isinstance(coordinates, slice)
        self.coordinates = coordinates

    def __call__(self, cell):
        assert isinstance(cell, Cell)
        start = cell.coordinates.stop + self.coordinates.start
        stop = cell.coordinates.stop + self.coordinates.stop
        return cell.spreadsheet.Rows[start, stop]

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
    class Columns:
        def __getitem__(self, item):
            assert isinstance(item, slice)
            return _RelativeColumns(item)
    class Rows:
        def __getitem__(self, item):
            assert isinstance(item, slice)
            return _RelativeRows(item)

RelativeCell = _Relative.Cell()
RelativeCells = _Relative.Cells()
Relativecolumns = _Relative.Columns()
RelativeRows = _Relative.Rows()

class Spreadsheet(list):
    """
    Spreadsheetclass to form Excel spreadsheets 12 in Sylk format
    """
    functions = dict()
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
        elif isinstance(item, str):
            coord = get_coordinates_by_name(item)
            return list.__getitem__(self, coord.stop)[coord.start]
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
        elif isinstance(key, str):
            coord = get_coordinates_by_name(key)
            self[coord.stop][coord.start] = item

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
                    return Spreadsheet(final)
                else:
                    raise CoordinatesError("Coordinates may be slices with the form [first row:last row]")
        return RowsGenerator(self)

    def append(self, item):
        """
        Appends only Rows to Spreadsheet
        :param item: Item to append to Spreadsheet
        :return: None
        """
        row = Rows(self)
        list.append(self, row)
        if not any([isinstance(item, typo) for typo in (list, tuple)]):
            item = [item]
        row.extend(item)

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

class Columns(Spreadsheet):
    """
    Column Class to Spreadsheet
    """
    def __init__(self, spreadsheet, *args, **kwargs):
        self._spreadsheet = spreadsheet
        Spreadsheet.__init__(self, *args, **kwargs)

    @property
    def spreadsheet(self):
        return self._spreadsheet

    def __sylk__(self):
        return

    def __csv__(self):
        pass


def verify(value, coords):
    if isinstance(value, Cell):
        start = value.coordinates
        return RelativeCell.__getitem__(sub_slices(start, coords))
    elif isinstance(value, str) and value.startswith("="):
        value = value[1:].lower()
        functions = re.findall(r"([a-z_\.]+)\(", value)
        if len(functions) == 0:
            coord = get_coordinates_by_name(value)
            return RelativeCell.__getitem__(sub_slices(coord, coords))
        else:
            for f in functions:
                if f not in Spreadsheet.functions:
                    Spreadsheet.functions[f] = Function(f)
                value = re.sub(r"([a-z_\.]+)\(", lambda x: "Spreadsheet.functions[\"{}\"](".format(x.group(0).strip("(")), value)
                new_value = dict()
                new_value["value"] = value
                repr = re.sub(r"Spreadsheet.functions\[[\w\W]+]\]\([\w\W]+\)", lambda x: x.group(0)+".__repr__()", value)
                new_value["repr"] = repr
                new_value["sylk"] = re.sub(r"Spreadsheet.functions\[[\w\W]+]\]\([\w\W]+\)", lambda x: "sylk("+x.group(0)+")", value)
                new_value["eval"] = str(eval(repr))
                return new_value
    else:
        return value


class Rows(list):
    """
    Row class to Spreadsheet
    """
    def __init__(self, spreadsheet, iterable=None):
        self._spreadsheet = spreadsheet
        if iterable is None:
            iterable = list()
        list.__init__(self)
        self.extend(iterable)

    def __setitem__(self, key, value):
        if isinstance(key, int):
            if key >= len(self):
                for x in range(len(self), key+1):
                    list.append(self, Cell(self.spreadsheet, None))
            self[key].value = verify(value, self[key].coordinates)
        elif isinstance(key, slice):
            if any([isinstance(value, typo) for typo in (list, tuple)]):
                if abs(key.stop - key.start) == len(value)-1:
                    for index, x in enumerate(range(key.start, key.stop+1)):
                        self[x] = None
                        self[x].value = self.verify(value[index], self[x].coordinates)
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
        for i in item:
            list.append(self, Cell(self.spreadsheet, None))
            input(self.spreadsheet)
            self[-1].value = verify(i, self[-1].coordinates)

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
        object.__setattr__(self, "_coordinates", {"time": datetime.datetime.now() - \
                                                          datetime.timedelta(seconds=10),
                                                  "slice": None})

    def __getattribute__(self, item):
        if item in ("coordinates", "spreadsheet", "value"):
            return object.__getattribute__(self, item)
        else:
            return object.__getattribute__(self, "_value").__getattribute__(item)

    def __str__(self):
        return str(eval(self.__repr__()))

    def __repr__(self):
        return str(self.value.__repr__())

    def __sylk__(self):
        value = object.__getattribute__(self, "_value")
        if isinstance(value, dict) and "sylk" in value:
            return value["sylk"]

    @property
    def coordinates(self):
        coordinates = object.__getattribute__(self, "_coordinates")
        if coordinates["time"] < datetime.datetime.now():
            for index, row in enumerate(self.spreadsheet):
                try:
                    column = row.index(self)
                except ValueError:
                    pass
                else:
                    object.__getattribute__(self, "_coordinates")["slice"] = slice(column, index)
                    object.__getattribute__(self, "_coordinates")["time"] = datetime.datetime.now() + \
                                                                           datetime.timedelta(seconds=0.001)
                    return slice(column, index)
            else:
                raise IndexError("Coordinates not found")
        else:
            return object.__getattribute__(self, "_coordinates")["slice"]

    @property
    def spreadsheet(self):
        return object.__getattribute__(self, "_spreadsheet")

    @property
    def value(self):
        to_return = object.__getattribute__(self, "_value")
        if isinstance(to_return, _RelativeCell):
            return object.__getattribute__(to_return(self), "_value")
        elif isinstance(to_return, dict) and "eval" in to_return:
            print(to_return["eval"])
            return eval(to_return["eval"])
        else:
            return to_return

    @value.setter
    def value(self, value):
        object.__setattr__(self, "_value", value) #TODO -> Check what Instantiate (Item or Function)


class Function:
    """
    Function Class for Cells
    """
    def __init__(self, function):
        self._function = function

    @property
    def function(self):
       return self._function

    def __call__(self, *args, **kwargs):
        class callable:
            def __init__(self, function, *args, **kwrags):
                self.function = function
                self.args = args
                self.kwargs = kwargs

            def __repr__(self):
                return "Functions.{function}({args})".format(function = self.function,
                                                             args =  ", ".join([arg.__repr__() for arg in args]+
                                                                               [key+"="+self.kwargs[key].__repr__() for key in self.kwargs]))
            def __sylk__(self):
                return "{function}({args})".format(function = self.function, #TODO Verify semycolon in args
                                                   args = "; ".join([sylk(arg) for arg in args]+
                                                                    [key+"="+sylk(self.kwargs[key]) for key in self.kwargs]))
        return callable(self.function, *args, **kwargs)

    def __sylk__(self):
        pass

    def __csv__(self):
        pass


class Functions:
    @staticmethod
    def sum(*args):
        return sum(args)
