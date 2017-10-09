import datetime
import re
import weakref
from .functions import Functions
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

def get_column_by_name(name):
    name = name.upper()
    columns = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    final_col = int()
    for index, iter in enumerate(range(len(name) - 1, -1, -1)):
        final_col += (pow((len(columns)), index)) * (columns.index(name[iter]) + 1)
    return final_col-1

def get_coordinates_by_name(name):
    name = name.upper()
    col, row = re.findall(r"([A-Z]+)([0-9]+)", name)[0]
    return slice(get_column_by_name(col), int(row)-1)

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
    sheets = weakref.WeakValueDictionary()
    def __init__(self, data=None, *, name=None):
        list.__init__(self)
        if name is None:
            self._name = str(len(Spreadsheet.sheets))
        else:
            self._name = name
        Spreadsheet.sheets[self.name] = self
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
            return self.range(item)
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

    @property
    def name(self):
        return self._name

    def append(self, item):
        """
        Appends only Rows to Spreadsheet
        :param item: Item to append to Spreadsheet
        :return: None
        """
        row = Rows(self)
        list.append(self, row)
        if not isinstance(item, (list, tuple)):
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
        final = list()
        for row_index in range(start.stop, stop.stop+1):
            row = list()
            for col_index in range(start.start, stop.start+1):
                row.append(self[row_index][col_index])
            final.append(row)
        if len(final) == 1:
            final = final[0]
        return Spreadsheet(final)

    def range(self, range):
        range = range.lower()
        data = re.findall(r"^([a-z]+[0-9]+)(:[a-z]+[0-9]+)?$", range) #Range by cells
        if len(data) == 1:
            init, end = data[0]
            end = end.strip(":")
            if end != "":
                return self.subslice(get_coordinates_by_name(init), get_coordinates_by_name(end))
            else:
                coords = get_coordinates_by_name(init)
                return self[coords.stop][coords.start]
        elif len(data) == 0:
            data = re.findall(r"^([a-z]+)(:[a-z]+)?$", range) #Range by columns
            if len(data) == 1:
                init, end = data[0]
                end = end.strip(":")
                return self.Columns[get_column_by_name(init):get_column_by_name(end)]
            elif len(data) == 0:
                data = re.findall(r"^([0-9]+)(:[0-9]+)?$", range) #Range by Rows
                if len(data) == 1:
                    init, end = data[0]
                    end = end.strip(":")
                    return self.Rows[int(init)-1:int(end)-1]
        raise TypeError


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
            data = verify(value, self[key])
            if isinstance(data, Cell):
                list.__setitem__(self, x, data)
            else:
                self[key].value = verify(value, self[key])
        elif isinstance(key, slice):
            if isinstance(value, (list, tuple)):
                if abs(key.stop - key.start) == len(value)-1:
                    for index, x in enumerate(range(key.start, key.stop+1)):
                        self[x] = None
                        data = verify(value[index], self[x])
                        if isinstance(data, Cell):
                            list.__setitem__(self, x, data)
                        else:
                            self[x].value = verify(value[index], self[x])
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
        if not isinstance(item, (list, tuple)):
            item = [item]
        for i in item:
            list.append(self, Cell(self.spreadsheet, None))
            data = verify(i, self[-1])
            if isinstance(data, Cell):
                list.__setitem__(self, -1, data)
            else:
                self[-1].value = verify(i, self[-1])

    def extend(self, items):
        if isinstance(items, (list, tuple)):
            [self.append(item) for item in items]
        else:
            TypeError("values may be a list or a tuple")

    def index(self, value):
        for index, item in enumerate(self):
            if item == value:
                return index
        else:
            raise ValueError

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

    def __eq__(self, other):
        if isinstance(other, Cell):
            return id(self) == id(other)
        else:
            self.value == other

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
                    return object.__getattribute__(self, "_coordinates")["slice"]
            else:
                raise IndexError("Coordinates not found")
        return object.__getattribute__(self, "_coordinates")["slice"]

    @property
    def spreadsheet(self):
        return object.__getattribute__(self, "_spreadsheet")

    @property
    def value(self):
        to_return = object.__getattribute__(self, "_value")
        if isinstance(to_return, _RelativeCell):
            return object.__getattribute__(to_return(self), "_value")
        elif all([isinstance(to_return, typo) for typo in (Cell, dict)]):
            return eval(object.__getattribute__(to_return, "_value")["eval"])
        elif isinstance(to_return, dict) and "eval" in to_return:
            return eval(str(eval(to_return["repr"])))
        else:
            return to_return

    @value.setter
    def value(self, value):
        object.__setattr__(self, "_value", value)


class Function:
    """
    Function Class for Cells
    """
    def __init__(self, function):
        self._function = function

    @property
    def function(self):
       return self._function

    def __call__(self, args, _sheetname, **kwargs):
        class callable:
            def __init__(self, function, args, *, _sheetname, **kwargs):
                self.function = function
                self._args = args.split(";") #arguments in function may be separated by semicolons
                self.kwargs = kwargs
                self.spreadsheet = Spreadsheet.sheets[str(_sheetname)]
                self.sheetname = _sheetname

            @property
            def args(self):
                final = list()
                for arg in self._args:
                    arg = arg.lower()
                    data = re.findall(r"^([\W\w]+!)?([a-z]+[0-9]+(?::[a-z]+[0-9]+)?)$", arg)
                    if len(data) == 1:
                        sheetname, cell = data[0]
                        if cell != "":
                            sheetname = sheetname != "" and sheetname or self.sheetname
                        final.append(Spreadsheet.sheets[str(sheetname)].range(cell))
                    else:
                        data = re.findall(r"^([\W\w]+!)?([a-z]+(?::[a-z]+)?)$", arg)
                        if len(data) == 1:
                            sheetname, cell = data[0]
                            if cell != "":
                                sheetname = sheetname != "" and sheetname or self.sheetname
                            final.append(Spreadsheet.sheets[str(sheetname)].range(cell))
                        else:
                            data = re.findall(r"^([\W\w]+!)?([0-9]+(?::[0-9]+)?)$", arg)
                            if len(data) == 1:
                                sheetname, cell = data[0]
                                if cell != "":
                                    sheetname = sheetname != "" and sheetname or self.sheetname
                                final.append(Spreadsheet.sheets[str(sheetname)].range(cell))
                            else:
                                final.append(arg)
                return final

            def __repr__(self):
                return "Functions.{function}({args})".format(function = self.function,
                                                             args =  ", ".join([arg.__repr__() for arg in self.args]+
                                                                               [key+"="+self.kwargs[key].__repr__()
                                                                                for key in self.kwargs]))

            def __sylk__(self):
                return "{function}({args})".format(function = self.function, #TODO Verify semicolon in args
                                                   args = "; ".join([sylk(arg) for arg in self.args]+
                                                                    [key+"="+sylk(self.kwargs[key])
                                                                     for key in self.kwargs]))
        return callable(self.function, args, _sheetname=_sheetname, **kwargs)

    def __sylk__(self):
        pass

    def __csv__(self):
        pass


def verify(value, cell):
    coords = cell.coordinates
    spreadsheet = cell.spreadsheet
    sheetname = spreadsheet.name
    if isinstance(value, Cell):
        start = value.coordinates
        if value.spreadsheet.name == sheetname:
            return RelativeCell.__getitem__(sub_slices(start, coords))
        else:
            return value
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
                value = re.sub(r"([a-z_\.]+)\(",
                               lambda x: "Spreadsheet.functions[\"{}\"](".format(x.group(0).strip("(")),
                               value)
                new_value = dict()
                new_value["value"] = value
                #repr = re.sub(r"Spreadsheet.functions\[[\w\W]+]\]\([\w\W]+\)",
                #              lambda x: "{}_sheetname={}).__repr__()".format(x.group(0)[-1], sheetname),
                #              value)
                #print(repr)
                repr = value.replace("(", "(\"\"\"").replace(")", "\"\"\", _sheetname={})".format(sheetname))
                new_value["repr"] = repr
                new_value["sylk"] = re.sub(r"Spreadsheet.functions\[[\w\W]+]\]\([\w\W]+\)",
                                           lambda x: "sylk("+x.group(0)+")",
                                           value)
                try:
                    new_value["eval"] = str(eval(repr))
                except SyntaxError:
                    raise SyntaxError(value)
                return new_value
    else:
        return value
