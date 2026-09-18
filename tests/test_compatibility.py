import contextlib
import datetime
import hashlib
import importlib
import io
import json
from pathlib import Path
import sys
from tempfile import TemporaryDirectory
from types import SimpleNamespace
import unittest
from unittest.mock import MagicMock, patch

import openpyxl

from src import excel_access, platform_utils
from src.gui.SpExGui import SpExGui
from src.gui.ToggleGui import ToggleGui
from src.write.ecount import EcountWriter as ecount_writer_module


ROOT = Path(__file__).resolve().parents[1]


class FileInputTests(unittest.TestCase):
    def test_both_file_choosers_accept_native_unicode_paths(self):
        with TemporaryDirectory() as directory:
            path = Path(directory) / '출고 파일.xlsx'
            workbook = openpyxl.Workbook()
            workbook.active.title = '주문'
            workbook.save(path)
            workbook.close()
            for gui_class in (SpExGui, ToggleGui):
                with self.subTest(gui=gui_class.__name__):
                    gui = SimpleNamespace(fileNameLabel=MagicMock(), sheetCombobox=MagicMock())
                    gui_class.setSpExFileName(gui, str(path))
                    gui.sheetCombobox.configure.assert_called_once_with(values=['주문'])
            # On Windows this also detects unclosed input handles.
            path.unlink()

    def test_cancel_preserves_current_file(self):
        for gui_class in (SpExGui, ToggleGui):
            with self.subTest(gui=gui_class.__name__), patch('tkinter.filedialog.askopenfilename', return_value=''):
                gui = SimpleNamespace(setSpExFileName=MagicMock(), fileNameLabel=MagicMock(), spExFile='기존 파일')
                gui_class.findFile(gui)
                gui.setSpExFileName.assert_not_called()

    def test_file_mode_does_not_import_excel_automation(self):
        self.assertNotIn('xlwings', sys.modules)
        self.assertTrue(excel_access.sheet_names(ROOT / 'example.xlsx'))
        self.assertNotIn('xlwings', sys.modules)


class DesktopOperationTests(unittest.TestCase):
    def test_document_opener_on_both_platforms(self):
        path = str((ROOT / '한글 파일 & $(example).xlsx').resolve())
        with patch.object(platform_utils.sys, 'platform', 'darwin'), patch.object(platform_utils.subprocess, 'run') as run:
            platform_utils.open_file(path)
            run.assert_called_once_with(['/usr/bin/open', path], check=True)
        with patch.object(platform_utils.sys, 'platform', 'win32'), patch.object(platform_utils.os, 'startfile', create=True) as startfile:
            platform_utils.open_file(path)
            startfile.assert_called_once_with(path)

    def test_missing_excel_has_an_actionable_error(self):
        fake_xlwings = SimpleNamespace(apps=SimpleNamespace(active=None))
        with patch.dict(sys.modules, {'xlwings': fake_xlwings}):
            with self.assertRaisesRegex(ValueError, 'Excel을 실행'):
                excel_access.active_book()

    @staticmethod
    def _snapshot_book():
        book = MagicMock()
        snapshot = book.app.books.add.return_value

        def save(path):
            workbook = openpyxl.Workbook()
            workbook.active.title = excel_access.SNAPSHOT_SHEET
            workbook.active.append(['한글', 42])
            workbook.save(path)
            workbook.close()

        snapshot.save.side_effect = save
        return book, snapshot

    @staticmethod
    def _on_windows():
        # The file-based snapshot is the Windows path; Mac never asks Excel to save.
        return patch.object(excel_access.sys, 'platform', 'win32')

    def test_snapshot_uses_source_instance_and_releases_temporary_files(self):
        book, snapshot = self._snapshot_book()
        with self._on_windows(), patch.object(excel_access, 'active_book', return_value=book):
            with excel_access.active_sheet() as sheet:
                self.assertEqual(next(sheet.values), ('한글', 42))
                path = Path(snapshot.save.call_args.args[0])
                self.assertEqual(path.suffix, '.xlsx')
                snapshot.close.assert_called_once()
                self.assertTrue(path.is_file())
            self.assertFalse(path.parent.exists())
        book.sheets.active.copy.assert_called_once_with(
            before=snapshot.sheets[0], name=excel_access.SNAPSHOT_SHEET)
        book.activate.assert_called_once()
        book.close.assert_not_called()
        book.save.assert_not_called()

    @staticmethod
    def _mac_book(values, precise, last_row=None, last_column=None):
        book = MagicMock()
        sheet = book.sheets.active
        sheet.used_range.last_cell.row = len(values) if last_row is None else last_row
        sheet.used_range.last_cell.column = len(values[0]) if last_column is None else last_column
        cells = sheet.range.return_value
        cells.options.return_value.value = values
        cells.api.value2.get.return_value = precise
        return book, sheet

    @contextlib.contextmanager
    def _mac_sheet(self, book):
        with patch.object(excel_access.sys, 'platform', 'darwin'), \
                patch.object(excel_access, 'TemporaryDirectory') as temporary, \
                patch.object(excel_access, 'active_book', return_value=book):
            with excel_access.active_sheet() as sheet:
                yield sheet
            # A 1-D result would silently transpose the sheet, so pin the option.
            book.sheets.active.range.return_value.options.assert_called_once_with(ndim=2)
            # Mac must never ask sandboxed Excel to write anything, anywhere.
            temporary.assert_not_called()
            book.app.books.add.assert_not_called()
            book.activate.assert_not_called()
            book.save.assert_not_called()

    def test_mac_snapshot_reads_values_instead_of_saving_a_file(self):
        stamp = datetime.datetime(2025, 12, 24, 9, 30)
        error = object()
        book, _ = self._mac_book(
            [['한글', 2.0, None, error],
             [stamp, 3291.6667, True, '줄\x07바꿈']],
            [['한글', 2.0, '', error],
             [46000.0, 3291.6666666666665, True, '줄\x07바꿈']])
        with self._mac_sheet(book) as sheet:
            rows = list(sheet.values)
        # 2.0 comes back as int, the error cell as None, and the date as a date.
        self.assertEqual(rows[0], ('한글', 2, None, None))
        self.assertEqual((rows[1][0], rows[1][2], rows[1][3]), (stamp, True, '줄바꿈'))
        # value2 keeps the digits .value rounds off; openpyxl then stores 16 of them.
        self.assertNotEqual(rows[1][1], 3291.6667)
        self.assertAlmostEqual(rows[1][1], 3291.6666666666665, places=9)

    def test_mac_snapshot_of_an_empty_sheet_has_no_rows(self):
        book, _ = self._mac_book([[None]], '')
        with self._mac_sheet(book) as sheet:
            self.assertEqual(list(sheet.values), [])

    def test_mac_snapshot_reads_a_single_cell_sheet(self):
        book, _ = self._mac_book([['한 칸']], '한 칸')
        with self._mac_sheet(book) as sheet:
            self.assertEqual(list(sheet.values), [('한 칸',)])

    def test_mac_snapshot_keeps_text_that_looks_like_a_formula_or_an_error(self):
        row = ['=SUM(A1)', '#VALUE!', '한글']
        book, _ = self._mac_book([row], [row])
        with self._mac_sheet(book) as sheet:
            self.assertEqual(list(sheet.values), [tuple(row)])

    def test_mac_snapshot_reads_a_single_row_without_transposing_it(self):
        book, _ = self._mac_book([[1.0, 2.0, 3.0]], [1.5, 2.5, 3.5])
        with self._mac_sheet(book) as sheet:
            self.assertEqual(list(sheet.values), [(1.5, 2.5, 3.5)])

    def test_mac_snapshot_reads_a_single_column_without_transposing_it(self):
        book, _ = self._mac_book([[1.0], [2.0], [3.0]], [1.5, 2.5, 3.5])
        with self._mac_sheet(book) as sheet:
            self.assertEqual(list(sheet.values), [(1.5,), (2.5,), (3.5,)])

    def test_mac_snapshot_survives_a_short_or_missing_value2_result(self):
        values = [['한글', 2.5], [3.5, '나라']]
        for precise in (None, [['한글', 2.25]]):
            with self.subTest(precise=precise):
                book, _ = self._mac_book(values, precise)
                with self._mac_sheet(book) as sheet:
                    rows = list(sheet.values)
                # Whatever value2 omits keeps the rounded .value instead of raising.
                self.assertEqual(rows[1], (3.5, '나라'))
                self.assertEqual(rows[0][0], '한글')
                self.assertIn(rows[0][1], (2.5, 2.25))

    def test_mac_snapshot_reads_from_a1_whatever_the_used_range_starts_at(self):
        # A used range of B3:D7 still has to keep every value at its own address.
        values = [[None] * 4 for _ in range(7)]
        values[2][1] = '가'
        book, sheet = self._mac_book(values, [[''] * 4] * 7, last_row=7, last_column=4)
        with self._mac_sheet(book) as snapshot:
            self.assertEqual(list(snapshot.values)[2][1], '가')
        sheet.range.assert_called_once_with((1, 1), (7, 4))

    def test_mac_snapshot_stops_at_the_row_and_column_caps(self):
        book, sheet = self._mac_book([['한글']], [['한글']], last_row=1048576, last_column=16384)
        with self._mac_sheet(book):
            pass
        sheet.range.assert_called_once_with(
            (1, 1), (excel_access.MAX_SNAPSHOT_ROW, excel_access.MAX_SNAPSHOT_COLUMN))
        self.assertLessEqual(
            excel_access.MAX_SNAPSHOT_ROW * excel_access.MAX_SNAPSHOT_COLUMN, 2_000_000)

    def test_mac_snapshot_drops_a_formatted_but_empty_tail(self):
        values = [['품목', '수량', '금액', None, None]]
        values.append(['사과', 2.0, 3000.0, None, None])
        values.append(['배', 1.0, 2500.0, None, None])
        values.extend([[None] * 5 for _ in range(400)])
        book, _ = self._mac_book(values, [[''] * 5 for _ in values])
        with self._mac_sheet(book) as sheet:
            rows = list(sheet.values)
        # Only the two empty trailing columns and the 400 empty rows go away.
        self.assertEqual(rows, [('품목', '수량', '금액'), ('사과', 2, 3000), ('배', 1, 2500)])

    def test_mac_snapshot_keeps_blank_rows_and_cells_inside_the_data(self):
        values = [['가', None, '나'], [None, None, None], [None, '다', None]]
        book, _ = self._mac_book(values, [[''] * 3 for _ in values])
        with self._mac_sheet(book) as sheet:
            self.assertEqual(list(sheet.values), [('가', None, '나'), (None, None, None), (None, '다', None)])

    def test_mac_snapshot_stores_an_aware_datetime_without_its_timezone(self):
        aware = datetime.datetime(2025, 12, 24, 9, 30, tzinfo=datetime.timezone.utc)
        book, _ = self._mac_book([[aware]], [[46015.0]])
        with self._mac_sheet(book) as sheet:
            self.assertEqual(list(sheet.values), [(aware.replace(tzinfo=None),)])

    def test_failed_snapshot_closes_only_the_copy(self):
        book = MagicMock()
        snapshot = book.app.books.add.return_value
        snapshot.save.side_effect = OSError('cannot save')
        with self._on_windows(), patch.object(excel_access, 'active_book', return_value=book):
            with self.assertRaisesRegex(OSError, 'cannot save'):
                with excel_access.active_sheet():
                    self.fail('A failed snapshot must not be exposed')
        snapshot.close.assert_called_once()
        book.activate.assert_called_once()
        book.close.assert_not_called()
        self.assertFalse(Path(snapshot.save.call_args.args[0]).parent.exists())

    def test_output_open_error_survives_cleanup_with_partial_snapshot_reader(self):
        book = MagicMock()
        snapshot = book.app.books.add.return_value
        snapshot_path = None

        def save(path):
            nonlocal snapshot_path
            snapshot_path = Path(path)
            workbook = openpyxl.Workbook()
            sheet = workbook.active
            sheet.title = excel_access.SNAPSHOT_SHEET
            sheet.append(['시간', '챙길것', '품목', '수량', '보험사', '지점',
                          '주문자이름', '주소', '전화번호', '단가', '금액', '수금'])
            sheet.append(['#끝'])
            for row in range(1000):
                sheet.append([f'unread trailing row {row} ' + ('x' * 200)])
            workbook.save(path)
            workbook.close()

        snapshot.save.side_effect = save
        opener_error = RuntimeError('output opener failed')
        caught = None
        writer = None
        with TemporaryDirectory() as output_directory, \
                self._on_windows(), \
                patch.object(excel_access, 'active_book', return_value=book), \
                patch.object(ecount_writer_module, 'getTempDir', return_value=output_directory), \
                patch.object(ecount_writer_module, 'open_file', side_effect=opener_error):
            try:
                with excel_access.active_sheet() as sheet:
                    writer = ecount_writer_module.EcountWriter.fromSheet(sheet)
                    writer.getDocsFromSpEx()
            except Exception as error:
                caught = error

        self.assertIs(caught, opener_error)
        self.assertIsNotNone(caught.__traceback__)
        self.assertIsNotNone(writer.spExReader.it.gi_frame)
        self.assertIsNotNone(snapshot_path)
        self.assertFalse(snapshot_path.parent.exists())

    def test_selected_range_preserves_headers_and_rows(self):
        book = MagicMock()
        book.sheets.active.range.return_value.value = ['품목', '수량']
        book.app.selection.value = [['상품1', 2], ['상품2', 3]]
        with patch.object(excel_access, 'active_book', return_value=book):
            self.assertEqual(excel_access.selected_range(), [['품목', '수량'], ['상품1', 2], ['상품2', 3]])


class ConversionRegressionTests(unittest.TestCase):
    def test_sample_outputs_match_before_the_port(self):
        # Digests of every saved cell, captured from master (44c85da).
        # Ignore ZIP timestamps and XML metadata; preserve actual document content.
        cases = [
            ('src.write.ecount.EcountWriter', 'EcountWriter',
             'aef7bd4e5a4f65fd0a37496bfae8bdb6503e279c3f09a924a08c6c51dbcedafc'),
            ('src.write.wehago.WehagoWriter', 'WehagoWriter',
             '1a333566a5ba16384392807e747adf87c8296993500446f3df3ce5d7fd87ab37'),
        ]
        for module_name, class_name, expected in cases:
            with self.subTest(writer=class_name), TemporaryDirectory() as directory:
                module = importlib.import_module(module_name)
                workbook = openpyxl.load_workbook(ROOT / 'example.xlsx', read_only=True, data_only=True)
                try:
                    with patch.object(module, 'getTempDir', return_value=directory), \
                            patch.object(module, 'open_file') as opened, \
                            contextlib.redirect_stdout(io.StringIO()):
                        getattr(module, class_name).fromSheet(workbook.active).getDocsFromSpEx()
                finally:
                    workbook.close()
                paths = list(Path(directory).glob('*.xlsx'))
                self.assertEqual(len(paths), 1)
                opened.assert_called_once()
                self.assertEqual(Path(opened.call_args.args[0]), paths[0])
                output = openpyxl.load_workbook(paths[0], read_only=True, data_only=True)
                try:
                    values = {sheet.title: list(sheet.values) for sheet in output}
                finally:
                    output.close()
                actual = hashlib.sha256(json.dumps(values, ensure_ascii=False, default=str).encode()).hexdigest()
                self.assertEqual(actual, expected)


if __name__ == '__main__':
    unittest.main()
