import unittest
from unittest import mock


class TestLoadHwResultsClosesWorkbook(unittest.TestCase):
    def test_closes_workbook_on_success(self) -> None:
        from app import data_loader

        class _Ws:
            def iter_rows(self, values_only=True):
                yield ("ФИО", "Результат")
                yield ("Иван Иванов", "зач")

        class _Wb:
            def __init__(self):
                self.active = _Ws()
                self.closed = False

            def close(self):
                self.closed = True

        wb = _Wb()
        with mock.patch.object(data_loader.openpyxl, "load_workbook", return_value=wb):
            hw = data_loader.HwFile(idx=1, label="HW01", path=mock.Mock(), date=None, mtime=0.0)
            out = data_loader.load_hw_results(hw)

        self.assertTrue(wb.closed)
        self.assertIn("Иван Иванов", out)

    def test_closes_workbook_on_exception(self) -> None:
        from app import data_loader

        class _Ws:
            def iter_rows(self, values_only=True):
                raise RuntimeError("boom")

        class _Wb:
            def __init__(self):
                self.active = _Ws()
                self.closed = False

            def close(self):
                self.closed = True

        wb = _Wb()
        with mock.patch.object(data_loader.openpyxl, "load_workbook", return_value=wb):
            hw = data_loader.HwFile(idx=1, label="HW01", path=mock.Mock(), date=None, mtime=0.0)
            with self.assertRaisesRegex(RuntimeError, "boom"):
                data_loader.load_hw_results(hw)

        self.assertTrue(wb.closed)

    def test_raises_when_header_not_found_and_closes(self) -> None:
        from app import data_loader

        class _Ws:
            def iter_rows(self, values_only=True):
                yield ("abc", "def")
                yield ("Иван Иванов", "зач")

        class _Wb:
            def __init__(self):
                self.active = _Ws()
                self.closed = False

            def close(self):
                self.closed = True

        wb = _Wb()
        with mock.patch.object(data_loader.openpyxl, "load_workbook", return_value=wb):
            hw = data_loader.HwFile(idx=1, label="HW01", path=mock.Mock(), date=None, mtime=0.0)
            with self.assertRaises(ValueError):
                data_loader.load_hw_results(hw)

        self.assertTrue(wb.closed)


class TestWorkbookSheetSupport(unittest.TestCase):
    def test_uses_explicit_sheet_name_when_present(self) -> None:
        from app import data_loader

        class _Ws:
            def __init__(self, rows):
                self._rows = rows

            def iter_rows(self, values_only=True):
                yield from self._rows

        class _Wb:
            def __init__(self):
                self.active = _Ws([("ФИО", "Результат"), ("Не тот студент", "не зач")])
                self._sheets = {
                    "ДЗ 02": _Ws([("ФИО", "Результат"), ("Иван Иванов", "зач")])
                }

            def __getitem__(self, key):
                return self._sheets[key]

            def close(self):
                pass

        with mock.patch.object(data_loader.openpyxl, "load_workbook", return_value=_Wb()):
            hw = data_loader.HwFile(
                idx=2,
                label="HW02",
                path=mock.Mock(),
                sheet_name="ДЗ 02",
                date=None,
                mtime=0.0,
            )
            out = data_loader.load_hw_results(hw)

        self.assertIn("Иван Иванов", out)
        self.assertNotIn("Не тот студент", out)

    def test_discovers_hw_sheets_in_generic_workbook_name(self) -> None:
        from app import data_loader

        path = data_loader.Path("results.xlsx")

        class _Wb:
            sheetnames = ["Свод", "ДЗ 01", "ДЗ 02", "Прочее"]

            def close(self):
                pass

        with mock.patch.object(data_loader.Path, "glob", return_value=[path]):
            with mock.patch.object(data_loader.Path, "stat", return_value=mock.Mock(st_mtime=123.0)):
                with mock.patch.object(data_loader, "_extract_date_from_path", return_value=None):
                    with mock.patch.object(
                        data_loader.openpyxl, "load_workbook", return_value=_Wb()
                    ):
                        hw_files = data_loader.discover_hw_files(data_loader.Path("."))

        self.assertEqual([hw.idx for hw in hw_files], [1, 2])
        self.assertEqual([hw.sheet_name for hw in hw_files], ["ДЗ 01", "ДЗ 02"])
        self.assertTrue(all(hw.path == path for hw in hw_files))


class TestBuildDataWarnings(unittest.TestCase):
    def test_includes_warnings_when_hw_load_fails(self) -> None:
        from app import data_loader

        hw = data_loader.HwFile(
            idx=1,
            label="HW01",
            path=data_loader.Path("HW01.xlsx"),
            date=None,
            mtime=0.0,
        )

        with mock.patch.object(data_loader, "load_hw_results", side_effect=ValueError("no header")):
            data = data_loader.build_data(data_loader.Path("."), hw_files=[hw])

        self.assertIn("warnings", data["meta"])
        self.assertTrue(data["meta"]["warnings"])


if __name__ == "__main__":
    unittest.main()
