"""Read saved workbooks or take a temporary snapshot of desktop Excel."""

import datetime
import io
import sys
from contextlib import contextmanager
from pathlib import Path
from tempfile import TemporaryDirectory

import openpyxl
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE


SNAPSHOT_SHEET = 'ConverterInput'

# Bound the Mac value read, and the workbook built from it, for a sheet whose used
# range was blown up by whole-sheet formatting: one huge Apple Event times out
# (-1712), and openpyxl needs seconds and hundreds of MB for a million cells. Real
# order sheets run to a few thousand rows and some 30 columns, well inside both.
MAX_SNAPSHOT_ROW = 20000
MAX_SNAPSHOT_COLUMN = 64

# Everything openpyxl can store. Anything else (an appscript keyword for an error
# cell, for instance) becomes None, which the readers already treat as a non-number.
STORABLE = (str, bool, int, float, datetime.datetime, datetime.date, datetime.time)


def _grid(values, rows):
    """Normalise an Apple Event result into a list of row lists."""
    if not isinstance(values, list):
        return [[values]]
    if not values or not isinstance(values[0], list):
        # Excel flattens a single row, and a single column, into one list.
        return [list(values)] if rows == 1 else [[value] for value in values]
    return [list(row) for row in values]


def _at(grid, index, column):
    if index >= len(grid) or column >= len(grid[index]):
        return None
    return grid[index][column]


def _storable(value, precise):
    # .value rounds floats to 4 decimals, while value2 keeps every digit but turns
    # dates into serial numbers, so take the precise number only for a number.
    if isinstance(value, float) and isinstance(precise, (int, float)) and not isinstance(precise, bool):
        value = precise
    if isinstance(value, str):
        return ILLEGAL_CHARACTERS_RE.sub('', value)
    if isinstance(value, (datetime.datetime, datetime.time)) and value.tzinfo is not None:
        # openpyxl refuses a timezone, and a sheet's clock time carries none anyway.
        return value.replace(tzinfo=None)
    return value if isinstance(value, STORABLE) else None


def _trim(grid):
    """Drop the formatted-but-empty tail, which openpyxl would pay rows and memory for."""
    while grid and all(value is None for value in grid[-1]):
        grid.pop()
    width = max((len(row) for row in grid), default=0)
    while width and all(_at(grid, index, width - 1) is None for index in range(len(grid))):
        width -= 1
    # Leading and interior blanks stay: every value keeps the address it has in Excel.
    return [row[:width] for row in grid] if width else []


def _mac_sheet_values(sheet):
    """Read the active sheet's values over Apple Events, without writing a file."""
    last = sheet.used_range.last_cell
    rows = min(last.row, MAX_SNAPSHOT_ROW)
    columns = min(last.column, MAX_SNAPSHOT_COLUMN)
    # Always read from A1 so every value keeps the position it has in the file.
    cells = sheet.range((1, 1), (rows, columns))
    values = _grid(cells.options(ndim=2).value, rows)
    precise = _grid(cells.api.value2.get(), rows)
    # value2 is a second Apple Event and may come back shorter, or as a bare scalar,
    # so a missing cell falls back to the rounded .value rather than raising.
    return _trim([
        [_storable(value, _at(precise, index, column)) for column, value in enumerate(row)]
        for index, row in enumerate(values)
    ])


@contextmanager
def _snapshot_of(values):
    source = openpyxl.Workbook()
    buffer = io.BytesIO()
    try:
        sheet = source.active
        sheet.title = SNAPSHOT_SHEET
        for index, row in enumerate(values, start=1):
            for column, value in enumerate(row, start=1):
                if value is None:
                    continue
                cell = sheet.cell(row=index, column=column, value=value)
                if isinstance(value, str):
                    # Excel gave us text, so keep "=SUM(A1)" and "#VALUE!" as text.
                    # Stored as a formula or an error they read back as None.
                    cell.data_type = 's'
        source.save(buffer)
    finally:
        source.close()
    with buffer:
        workbook = openpyxl.load_workbook(buffer, read_only=True, data_only=True)
        try:
            yield workbook[SNAPSHOT_SHEET]
        finally:
            workbook.close()


def sheet_names(path):
    workbook = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        return workbook.sheetnames
    finally:
        workbook.close()


@contextmanager
def open_sheet(path, name):
    if not path or not Path(path).is_file():
        raise ValueError('변환할 Excel 파일을 먼저 선택해주세요.')
    if not name:
        raise ValueError('변환할 시트를 선택해주세요.')
    workbook = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        yield workbook[name]
    finally:
        workbook.close()


def active_book():
    # File-based conversion must not initialize Excel automation (or its exit hooks).
    import xlwings

    app = xlwings.apps.active
    book = app.books.active if app is not None else None
    if book is None:
        raise ValueError('Excel을 실행하고 변환할 문서를 열어주세요.')
    return book


@contextmanager
def active_sheet():
    book = active_book()
    sheet = book.sheets.active
    if sys.platform == 'darwin':
        # Sandboxed Excel cannot save the snapshot without the "파일 접근 권한 부여"
        # window, and returns OSERROR -50 when it is cancelled, so never ask it to
        # write a file on Mac. Read the values instead and build the copy in memory.
        with _snapshot_of(_mac_sheet_values(sheet)) as snapshot_sheet:
            yield snapshot_sheet
        return
    with TemporaryDirectory(prefix='spdocuments-') as directory:
        path = Path(directory) / 'input.xlsx'
        snapshot = None
        try:
            # Use the source book's Excel instance, including on Windows with
            # multiple instances. Always save as xlsx, even for unsaved/xls books.
            snapshot = book.app.books.add()
            sheet.copy(before=snapshot.sheets[0], name=SNAPSHOT_SHEET)
            # Only the disposable copy loses macros when saved as xlsx.
            with book.app.properties(display_alerts=False):
                snapshot.save(str(path))
        finally:
            try:
                if snapshot is not None:
                    snapshot.close()
            finally:
                book.activate()
        # Read the disposable snapshot through memory so a partially consumed
        # read-only iterator cannot retain a Windows handle to input.xlsx.
        with io.BytesIO(path.read_bytes()) as buffer:
            workbook = openpyxl.load_workbook(buffer, read_only=True, data_only=True)
            try:
                yield workbook[SNAPSHOT_SHEET]
            finally:
                workbook.close()


def selected_range():
    book = active_book()
    selection = book.app.selection
    if selection is None:
        raise ValueError('Excel에서 변환할 셀 영역을 선택해주세요.')
    data = [book.sheets.active.range('TITLES').value]
    data.extend(selection.value)
    return data
