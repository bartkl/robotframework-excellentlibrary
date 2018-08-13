import openpyxl
import os
import errno
import json
from robot.api import logger
from openpyxl.utils import column_index_from_string, range_boundaries


# Base exception.
class ExcellentLibraryException(Exception):
    def __str__(self):
        return self.message


class AliasAlreadyInUseException(ExcellentLibraryException):
    def __init__(self, alias):
        self.message = ("The alias `{}' is already in use by another "
                        "workbook.".format(alias))


class ExcelFileNotFoundException(ExcellentLibraryException):
    def __init__(self, filepath):
        self.message = "file `{}' does not exist.".format(filepath)


class FileAlreadyExists(ExcellentLibraryException):
    def __init__(self, filepath):
        self.message = "The file `{0}' already exists.".format(filepath)


class FileAlreadyOpened(ExcellentLibraryException):
    def __init__(self, filepath, alias):
        self.message = ("The workbook with filepath `{0}' is already opened "
                       "with alias `{1}'.")\
                       .format(filepath, alias)


class InvalidCellCoordinatesException(ExcellentLibraryException):
    def __init__(self):
        self.message = ("Please supply sufficient coordinates for "
                        "identifying a cell.")


class NoAliasSuppliedException(ExcellentLibraryException):
    def __init__(self):
        self.message = ("Please supply an alias in order to "
                        "identify the workbook.")


class SheetExistsAlreadyException(ExcellentLibraryException):
    def __init__(self, name):
        self.message = "sheet `{}' already exists.".format(name)


class SheetNotFoundException(ExcellentLibraryException):
    def __init__(self, name):
        self.message = "Could not find sheet `{0}'.".format(name)


class TooFewColumnNamesSuppliedException(ExcellentLibraryException):
    def __init__(self,):
        self.message = ("The amount of column names supplied is smaller than "
                        "the amount of columns.")


class UnknownWorkbookException(ExcellentLibraryException):
    def __init__(self, alias):
        self.message = "No opened workbook found with alias `{}'."\
                       .format(alias)


class UnopenedWorkbookException(ExcellentLibraryException):
    def __init__(self, alias):
        self.message = \
            ("workbook with alias `{}' is unknown, please open it before "
             "using it.".format(alias))


class ExcellentLibrary:
    """This library is built on top of _OpenPyXL_ in order to bring its
    functionality to _Robot Framework_. The major motivation for this was
    to add support for _Excel 2010_ (XSLX) files, which _ExcelLibrary_ does
    not support.

    = Usage =

    == Workbooks ==
    === Opening and switching ===
    To open an Excel file (workbook), use `Open workbook` . You can open
    several workbooks simultaneously, between which you can switch with
    `Switch workbook`. For example:

    | ${some workbook} =   |  Some file.xlsx   |
    | Open workbook        |  ${some workbook} |
    | Switch workbook      |  ${some workbook} |

    After opening a workbook, it will also be made the active workbook. So if
    for some reason you wish to open a workbook but not switch to it, you have
    to manually switch to the workbook you wish to be working on.

    Note that whenever you try to open a workbook that's already opened, you
    will get a warning pointing that out to you. It won't be reopened, but it
    will be made the active workbook.

    === Creating ===
    To create a new Excel file (workbook), you can use the `Create workbook`
    keyword. An example should clarify its use:

    | ${new workbook} =   | New file.xlsx   |
    | Create workbook     | ${new workbook} |  # It saves to file too! |

    === Saving ===
    Whenever you've made it any change to the workbook, you yourself must
    call `Save` to save the changes to the file like so:

    | Save |  # Saves changes to the current workbook to disk. |

    An exception to this rule is the `Create workbook` keyword, which performs
    the save itself.

    *Caution*: There may be other exceptions depending on the implementation
    in _OpenPyXL_.

    == Sheets ==

    === Switching between them ===
    You can switch between sheets by identifying them by their name as following:

    | Switch sheet      |  Orders  |  # Acts on active workbook. |

    Note that if each workbook has its own active sheet, so whenever you switch
    between workbooks they all keep track of their own active sheet:

    | Open workbook     |  Debit.xlsx    |
    | Switch sheet      |  Orders        |
    | Open workbook     |  Credit.xlsx   |  # Switches to this workbook too!  |
    | Switch sheet      |  People        |
    | Switch workbook   |  Debit1.xlsx   |  # Active sheet is still *Orders*. |


    === Creating sheets ===
    To create a new sheet in the active workbook, simply use
    `Create sheet`. For example:

    | Create sheet    |  New sheet                |
    |  Save           |  # Don't forget to save!  |

    Don't forget to save your changes to the workbook as soon as you're done.

    Whenever a sheet with the given name already exists an
    ``SheetExistsAlreadyException`` is raised.

    == Data ==
    === Reading and identifying a cell ===
    Several keywords, including `Write To Cell` and `Read From Cell` require
    you to identify the cell with which you wish to interact. Basically there
    are two ways to choose from:
    - _A1 notation_, provided through the ``cell`` parameter.\
    This is the well-known shorthand notation which numbers\
    the columns _A_, _B_, _C_, ... and the rows 1, 2, 3, ...\
    For example, _B4_ will refer to row 4, column 2.
    - _Row/column coordinates__, provided through the ``row_nr`` and\
    ``col_nr`` parameters.\
    This is exactly what you'd expect: the row and column numbers\
    (starting from 1) of the cell you want to interact with.

    _NOTE_: Since ``cell`` is the first named argument, you can simply pass in
    the value without having to mention the parameter name.

    *Examples*:

    | ${no what}    | Read From Cell |          |          |          | # Bad |
    | Write To Cell | Hi.            | B1       |          |          | # OK  |
    | Write To Cell | Hi again.      | cell=D1  |          |          | # OK  |
    | ${what}       | Read From Cell | row_nr=1 | col_nr=2 |          | # OK  |
    | ${no what}    | Read From Cell | row_nr=1 |          |          | # Bad |
    | Write To Cell | Hello          | cell=D1  | row_nr=1 | col_nr=4 | # Bad |

    If desired one can trim the surrounding whitespace of a cell value by
    passing ``trim=${TRUE}``. By default no trimming is applied.

    === Writing data to sheets ===
    To write plain text data to a cell, the following straight-forward use of
    `Write To Cell` keyword will do:

    | Write To Cell | Hello   | B1       | # OK |

    See *Identiying a cell* for more information on cell identification. Here
    I will stick with the A1 notation.

    It is possible to format the cell using the ``number_format`` parameter.
    In order for this to work properly with the data you're writing, you must
    make sure that the data type of the latter is compatible with what the
    number formatting expects. For example, to format a cell as a number
    that's rounded to two decimals, one should write data of a type number. To
    format a cell to hold a datetime value, a Python datetime object should be
    passed in for it to function.

    Some examples:
    | Write To Cell | Hello      | B1 |                           | # OK  |
    | Write To Cell | ${2}       | B1 |                           | # OK  |
    | Write to cell | 1.233      | A1 | number_format=#.#         | # Bad |
    | Write to cell | ${1.233}   | A1 | number_format=#.#         | # OK  |
    | Write to cell | 2018-04-01 | C1 | number_format=yyyy-dd-mm  | # Bad |
    | ${now}        | DateTime.Get current date | |               |       |
    | Write to cell | ${now}     | D1 | number_format=yyyy-dd-mm  | # OK  |
    | Write to cell | ${now}     | D1 | number_format=jjjj-dd-mm  | # Bad |

    _NOTE_: The ``numer_format`` parameter seems to assume the US locale, so
    make sure to delimit numbers with dots ("."), and format your dates using
    ``yyyy`` for example rather than ``jjjj`` (Dutch). Excel will honor your
    own locale settings anyways, so don't worry about it.

    The OpenPyXL documentation is quite immature, so if you really need to
    understand the implementation better you are forced to experiment or
    read the source code.

    ===  Reading cells ===
    Reading cells is easy:

    | ${value} | Read From Cell   | B1       | # OK |

    See *Identiying a cell* for more information on cell identification. Here
    I will stick with the A1 notation.

    _NOTE_: This respects the number format of the cell, so reading a cell
    with, for example, number format ``#.#`` will yield a value of type number
    in Robot Framework.

    _@TODO_: The ``Workbook`` object has a ``guess_types`` boolean which can
    be used to manipulate this data type inferring behavior when reading
    cells. This should be looked into.

    === Reading the entire sheet ===
    When you're dealing with reasonable amounts of data, it can be useful to
    read all the data in a sheet to a list and work with that object in Robot
    Framework. For this use the ``Read sheet data`` keyword.

    Its parameters are documented extensively in the keyword documentation, so
    make sure to read that. Here an example will be shown:

    | Open Workbook | ${CURDIR}${/}..${/}Data${/}Orders.xlsx          |
    | @{sheet}      | Read Sheet Data                                 |
    | ...           |     get_column_names_from_header_row=${TRUE}    |
    | ...           |     skip_first_row=${TRUE}                      |
    | :FOR | ${row}  | IN  | @{sheet}                                 |
    | \  |  ${order id}        | Set variable  &{row}[OrderID]        |
    | \  |  ${price}           | Set variable  &{row}[Price]          |

    Here the keys of the ``&{row}`` dictionary correspond to the column names
    as fetched from header row, which was instructed by the
    ``get_column_names_from_header_row`` parameter.

    Here's an example without column names:

    | @{sheet}                 | Read Sheet Data                       |
    | :FOR | ${row}            | IN  | @{sheet}                        |
    | \    |  Log list  ${row} |                                       |

    === Using the row iterator ===
    Use the `Get Row Iterator` keyword to obtain an iterator object which can
    be used to iterate over all the rows. Only use this if the use case is
    advanced and truly requires it. It's technically somewhat harder and
    to use and generally leads to less readable code.

    = Exceptions =
    - ``ExcellentLibraryException``: Base exception; non-functional.
    - ``ExcelFileNotFoundException``: The supplied file could not be found.
    - ``InvalidCellCoordinatesException``: The provided coordinates were\
    incomplete or invalid. See *Identifying a cell*.
    - ``UnopenedWorkbookException``: The workbook you were trying to use is\
    nog among the opened ones. Please make sure to open a workbook before\
    trying to use it.
    - ``SheetExistsAlreadyException``: A new sheet is attempted to be created,\
    but a sheet with the supplied title already exists.
    """


    __version__ = '0.8.4'
    ROBOT_LIBRARY_SCOPE = 'GLOBAL'

    def __init__(self):
        self.workbooks = {}
        self.active_workbook_alias = None
        self.active_workbook = None

    def _add_to_workbooks(self, filepath, workbook, alias=None):
        """Adds the specified workbook to the opened workbooks dictionary. The
        supplied alias will be used as the key for the dictionary entry. This 
        defaults to the filepath in case no alias is given. The values in this
        dictionary are dictionaries themselves, holding the filepath and the
        OpenPyXL Workbook object.
        """
        if not alias:
            alias = filepath  # Setting the default.

        if alias in self.workbooks.keys():
            raise AliasAlreadyInUseException(alias)

        for workbook_entry in self.workbooks.values()   :
            if filepath == workbook_entry["filepath"]:
                existing_alias = self._get_alias_of_workbook_by_filepath(filepath)
                raise FileAlreadyOpened(filepath, existing_alias)

        self.workbooks[alias] = {"filepath": filepath,
                                 "workbook": workbook}
        self._set_new_active_workbook(alias)

    def _get_alias_of_workbook_by_filepath(self, filepath):
        """Gets the alias of supplied workbook. Only supports opened
        workbooks.
        """
        for alias, workbook_entry in self.workbooks.iteritems():
            if workbook_entry["filepath"] == filepath:
                return alias

    def _get_column_names_from_header_row(self, sheet):
        """Gets values from header row and returns them as a list.
        """
        column_names = []
        header_row = sheet[1]
        for cell in header_row:
            column_names.append(cell.value)
        return column_names

    def _remove_from_workbooks(self, alias):
        """Removes the workbook provided, identified by its file path.
        """
        try:
            del self.workbooks[alias]
        except AttributeError:
            pass

    def _resolve_cell_coordinates(self, locator):
        """Resolves the cell coordinates based on several possible forms
        of `locator` parameter.

        See the section **Identifying cell coordinates** for more information.
        """

        locator = locator.strip()
        if "," in locator:  # Coordinates.
            if locator.startswith("coords:"):
                locator = locator[7:]
            locator = locator.lstrip('(').rstrip(')')
            coords_parts = locator\
                           .lstrip('(')\
                           .rstrip(')')\
                           .split(',')
            row_nr, col_nr = coords_parts[1].strip(),\
                             coords_parts[0].strip()

        else:  # Assume A1 notation.
            if locator.startswith("a1:"):
                locator = locator[3:]

            a1_col = ""
            a1_row = ""
            for char in locator:
                if char.isdigit():
                    a1_row += char
                else:
                    a1_col += char

            col_nr = column_index_from_string(a1_col)
            row_nr = a1_row

        return int(row_nr), int(col_nr)

    def _set_new_active_workbook(self, alias):
        if not alias in self.workbooks.keys():
            raise UnknownWorkbookException(alias)
        self.active_workbook_alias = alias
        self.active_workbook = self.workbooks[alias]["workbook"]

    def close_all_workbooks(self):
        """Closes all opened workbooks.
        """
        for alias in self.workbooks.keys():
            self.close_workbook(alias)

    def close_workbook(self, alias=None):
        """Closes an Excel workbook.

        Changes made to a file won't be saved automatically.
        Use `Save` to save the changes to the file.

        If the file specified is the active workbook, then a new workbook becomes active.
        """
        if not alias:
            alias = self.active_workbook_alias
        try:
            if alias == self.active_workbook_alias:
                set_new_active_workbook = True
            else:
                set_new_active_workbook = False

            self.workbooks[alias]["workbook"].close()
            self._remove_from_workbooks(alias)

            if set_new_active_workbook and len(self.workbooks) > 0:
                new_alias = self.workbooks.keys()[0]
                self._set_new_active_workbook(new_alias)

        except KeyError:
            logger.warning("Cannot close workbook with alias `{}': workbook "
                            "not opened.".format(alias))

    def create_sheet(self, name):
        """Creates a sheet in the active workbook.

        The ``name`` parameter must be used to supply the name of the sheet.
        If the sheet already exists, a ``SheetAlreadyExistsException`` will be
        raised.
        """
        if not name in self.active_workbook.sheetnames:
            self.active_workbook.create_sheet(title=name)
        else:
            raise SheetExistsAlreadyException(name)

    def create_workbook(self, filepath, overwrite_file_if_exists=False, alias=None):
        """Creates a new workbook and saves it to disk.
        It will then also be considered opened, i.e. it will be added to the
        internal dictionary of opened workbooks.

        The ``filepath`` must be supplied.
        _NOTE_: It is advised to supply an absolute path to avoid confusion
        regarding what the current working directory is.
        """
        workbook = openpyxl.Workbook()
        if os.path.isfile(filepath) and not overwrite_file_if_exists:
            raise FileAlreadyExists(filepath)
        else:
            workbook.save(filepath)

        # Add it to the opened workbooks dictionary.
        self._add_to_workbooks(filepath, workbook, alias=alias)

    def get_column_count(self):
        """Returns the number of non-empty columns in the active sheet.

        Technically this looks up the maximum column number for which the column is
        non-empty.
        """
        sheet = self.active_workbook.active
        return sheet.max_column

    def get_row_count(self):
        """Returns the number of non-empty rows in the active sheet.

        Technically this looks up the maximum row number for which the row is
        non-empty.
        """
        sheet = self.active_workbook.active
        return sheet.max_row

    def get_row_iterator(self):
        """Returns an iterator for looping over the rows in the active sheet.

        This won't be needed often and it is advised to avoid this as much as
        possible, since it unfriendly to read and hacky in its use with
        respect to Robot Framework.
        """
        sheet = self.active_workbook.active
        return sheet.iter_rows()

    def log_opened_workbooks(self):
        """Logs the dictionary in which the opened workbooks are kept.
        """
        logger.info(self.workbooks)

    def open_workbook(self, filepath, alias=None):
        """Opens an Excel workbook.

        Once a workbook is opened, one can work with it, i.e. manipulate data
        in its sheets, create new sheets, etc.

        Also, once opened, a workbook (technically, the file handle) is added
        to an internal dictionary. This way you can have several workbooks
        open simultaneously, swtiching between them when desired.

        The ``filepath`` parameter should point to the location of the file on
        the filesystem. It is advisable to make this an absolute path to avoid
        confusion.

        *Warning*: make sure to explicitly switch to the sheet you want to
        work with by using the `Switch sheet` keyword. Contrary to what you
        might expect, the active sheet by default is not necessarily the first
        one in tab-order!

        """
        try:
            workbook = openpyxl.load_workbook(filepath)
        except IOError as e:
            if e.errno == errno.ENOENT:
                raise ExcelFileNotFoundException(filepath)
            else:
                raise e
        self._add_to_workbooks(filepath, workbook, alias=alias)

    def read_from_cell(self,
                       cell,
                       cell_obj=None,
                       trim=False):
        """Reads the data from the given cell.

        For an explanation of how to identify a cell, please see the secion
        *Identifying a cell* at the top.

        The ``cell_obj`` argument can be used to pass a OpenPyXL Cell object
        to read from.
        """
        sheet = self.active_workbook.active

        if cell_obj:
            cell = cell_obj
        else:
            row_nr, col_nr = self._resolve_cell_coordinates(cell)
            cell = sheet.cell(row_nr, col_nr)

        if trim and cell.value is not None:
            return cell.value.strip()
        else:
            return cell.value

    def read_sheet_data(self,
                        column_names=None,
                        get_column_names_from_header_row=False,
                        cell_range=None,
                        trim=False):
        """Reads all the data from the active sheet.

        This keyword can output the sheet data in two formats:
            - _As a list of dictionaries_. In the case column names are
            supplied or obtained (see relevant parameters decribed below),
            the rows will be represented through dictionaries, of which the
            keys will be the column names.
            - _As a list of lists_. If no column names are provided or
            obtained, each row will be read from the sheet as a list, and
            the returned data will therefore be a list of all such lists.

        To use column names the following two parameters can be used.

        If ``column_names`` is provided it is expected to be a list which will
        be used to name the columns  in the supplied order.

        If ``get_column_names_from_header_row`` is truthy, the column names
        will be read from the first row in the sheet.

        _NOTE_: If both parameters are supplied, the ``@{column_names}`` list
        will have precedence. You will get a warning in your log when this
        situation occurs though.

        Use ``cell_range`` if you want to get data from only that range in the
        sheet, rather than all of the data in it. _TODO_: This should be moved
        to a separate keyword ``Read Cells In Range``.

        If ``trim`` is truthy, all cell values are trimmed, i.e. the
        surrounding whitespace is removed.
        """
        sheet = self.active_workbook.active
        skip_first_row = False

        if get_column_names_from_header_row:
            if column_names:
                logger.warning("Both the `column_names' and "
                                "`get_column_names_from_header_row' "
                                "parameters were supplied. Using "
                                "`column_names' and ignoring the other.")
            else:
                skip_first_row = True
                column_names = self._get_column_names_from_header_row(sheet)

        if cell_range:
            min_col, min_row, max_col, max_row = range_boundaries(cell_range)
            row_iterator = sheet.iter_rows(min_col=min_col, min_row=min_row,
                                           max_col=max_col, max_row=max_row)
            if column_names:
                column_names = column_names[min_col-1:max_col]
        else:
            row_iterator = sheet.iter_rows()

        if skip_first_row:
            next(row_iterator)  # Skip first row in the case of a header

        if column_names:
            sheet_data = []
            for row in row_iterator:
                row_data = {}
                for i, cell in enumerate(row):
                    try:
                        row_data[column_names[i]] =\
                        self.read_from_cell(None, cell_obj=cell, trim=trim)
                    except IndexError:
                        raise TooFewColumnNamesSuppliedException

                if not all(value is None for value in row_data.itervalues()):
                    sheet_data.append(row_data)
        else:
            sheet_data = []
            for row in row_iterator:
                row_data = [
                    self.read_from_cell(None, cell_obj=cell, trim=trim)
                    for cell in row
                ]
                if not all(value is None for value in row_data):
                    sheet_data.append(row_data)

        return sheet_data

    def remove_sheet(self, name):
        """Removes the given sheet from the active workbook.

        The ``name`` parameter must be used to supply the name of the sheet.
        If the sheet does not exist, a ``SheetNotFoundException`` will be
        raised.
        """
        try: 
            del self.active_workbook[name]
        except KeyError:
            raise SheetNotFoundException(name)        

    def save(self):
        """Saves the changes to the currently active workbook.

        _NOTE_: When manipulating sheets/cells, you are working with
        object representations in memory, not the factual data on disk.
        Only when you choose to make the changes persistent by calling this
        keyword, those changes will be written to disk.
        """
        filepath = self.workbooks[self.active_workbook_alias]["filepath"]
        self.active_workbook.save(filepath)

    def switch_sheet(self, sheet_name):
        """Switches to the sheet with the supplied name within the active
        workbook.

        Please supply the ``sheet_name`` parameter to identify which sheet you
        want to switch to.
        """
        sheet = self.active_workbook[sheet_name]
        index = self.active_workbook.index(sheet)
        self.active_workbook.active = index

    def switch_workbook(self, alias):
        """Switches between opened workbooks.

        Switches to the workbook identified by ``alias``, i.e. make
        that the active workbook.

        _NOTE_: You can only switch to workbooks which are opened. This
        keyword won't do that for you, so make sure you've opened the
        workbook you want to switch to using `Open workbook`.
        """
        try:
            self._set_new_active_workbook(alias)
        except KeyError:
            raise UnopenedWorkbookException(alias)

    def write_to_cell(self,
                      cell,
                      value,
                      number_format=None):
        """Writes value to the supplied cell.

        For an explanation of how to identify a cell, please see the secion
        *Identifying a cell* at the top.

        For the use of ``number_format``, please read the secion *Writing data to sheets*.
        """
        sheet = self.active_workbook.active
        row_nr, col_nr = self._resolve_cell_coordinates(cell)

        cell = sheet.cell(row_nr, col_nr)

        if number_format:
            cell.number_format = number_format

        cell.value = value


