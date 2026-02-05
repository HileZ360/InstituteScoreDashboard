import unittest
from unittest import mock


class TestLoadHwResultsClosesWorkbook(unittest.TestCase):
    def test_closes_workbook_on_success(self) -> None:
        from app import data_loader

        class _Ws:
            def iter_rows(self, values_only=True):
                # Header row -> detect indices and continue
                yield ("ФИО", "Результат")
                yield ("Ivan Ivanov", "зач")

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
        self.assertIn("Ivan Ivanov", out)

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


if __name__ == "__main__":
    unittest.main()

