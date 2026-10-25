##
# @file test_offline.py
# @brief Offline mode: Excel upload, input checks, and a full offline run.
##

import io
import unittest
from unittest import mock

from openpyxl import Workbook

from support import GuiTestCase, EXCEL, gui

SHEET = "I4916QGLLGR01A30M04A_ONT_V01.xlsx"


class OfflineInputTest(GuiTestCase):

    def setUp(self):
        super().setUp()
        self.fill_inputs()

    def files(self):
        return sorted(path.name for path in self.app_dir.iterdir())

    def upload(self, path):
        with mock.patch.object(gui.filedialog, "askopenfilename", lambda **kwargs: str(path)):
            self.app.upload_excel()
        self.app.update()

    def source_sheet(self, name=SHEET, content=EXCEL):
        folder = self.app_dir / "source"
        folder.mkdir(exist_ok=True)
        path = folder / name
        path.write_bytes(content)
        return path

    def test_no_excel_blocks_start(self):
        self.assertEqual(self.chip("Excel"), "✗ Excel")
        self.assertFalse(self.start_enabled())
        self.assertIn("No Excel sheet found", self.app.detail_label.cget("text"))

    def test_upload_sheet(self):
        self.upload(self.source_sheet())
        self.assertEqual(self.files(), [SHEET, "source"])
        self.assertEqual(self.chip("Excel"), "✓ Excel")
        self.assertEqual(self.app.excel_file_label.cget("text"), SHEET)
        self.assertTrue(self.start_enabled())

    def test_upload_replaces_previous_sheet(self):
        self.place_excel("old.xlsx")
        self.upload(self.source_sheet())
        self.assertEqual(self.popup_titles("ask"), ["Replace Excel sheet"])
        self.assertEqual(self.files(), [SHEET, "source"])

    def test_upload_cancelled_keeps_previous_sheet(self):
        self.place_excel("old.xlsx")
        with mock.patch.object(gui.messagebox, "askokcancel", lambda *args, **kwargs: False):
            self.upload(self.source_sheet())
        self.assertEqual(self.files(), ["old.xlsx", "source"])

    def test_upload_invalid_sheet(self):
        self.place_excel("old.xlsx")
        book = Workbook()
        book.active.append(["Name", "Data"])        # No "Jenkins Parameter" / "Value" columns
        data = io.BytesIO()
        book.save(data)
        self.upload(self.source_sheet("notes.xlsx", data.getvalue()))
        self.assertIn("Invalid Excel sheet", self.popup_titles("error"))
        self.assertEqual(self.files(), ["old.xlsx", "source"])      # Nothing changed

    def test_sheet_open_in_excel(self):
        (self.app_dir / "~$sheet.xlsx").write_bytes(b"lock")
        self.app.reload_excel()
        self.assertEqual(self.chip("Excel"), "✗ Excel")
        self.assertIn("close it first", self.app.excel.error)

    def test_sim_lengths(self):
        self.place_excel()
        self.assertTrue(self.start_enabled())
        cases = (("SIM CCID", self.app.ccid_var, "1" * 19, "19/20", "CCID"),
                 ("SIM IMEI", self.app.imei_var, "2" * 14 + "A", "15/15", "IMEI"),     # Digits only
                 ("SIM IMSI", self.app.imsi_var, "3" * 16, "16/15", "IMSI"))
        for label, var, value, counter, chip in cases:
            with self.subTest(label):
                good = var.get()
                var.set(value)
                self.app.update()
                self.assertEqual(self.app.length_labels[label][0].cget("text"), counter)
                self.assertEqual(self.chip(chip), f"✗ {chip}")
                self.assertFalse(self.start_enabled())
                var.set(good)
        self.app.ccid_var.set("8960" + "1" * 15 + "F")      # A CCID may end with a hex digit
        self.assertEqual(self.chip("CCID"), "✓ CCID")
        self.assertTrue(self.start_enabled())

    def test_jira_id_selects_module(self):
        for jira_id, text in (("SW9X07-10", "✓ MDM9X07 (LTE)"), ("qcx216-5", "✓ QCX216 (LTE)"),
                              ("SDX35-112", "✓ SDX35 (NR)"), ("ABC-1", "✗ Must contain QCX216, SW9X07, SDX35"),
                              ("", "e.g. SW9X07-10")):
            with self.subTest(jira_id):
                self.app.jira_var.set(jira_id)
                self.app.update()
                self.assertEqual(self.app.jira_desc.cget("text"), text)

    def test_missing_port_blocks_start(self):
        self.place_excel()
        self.app.port_var.set("")
        self.app.update()
        self.assertEqual(self.chip("AT port"), "✗ AT port")
        self.assertFalse(self.start_enabled())

    def test_not_an_at_port(self):
        self.place_excel()
        with mock.patch.object(gui, "probe_at_port", lambda port: (False, f"{port} did not answer.")):
            self.press_start()
        self.assertEqual(self.launched, [])
        self.assertIn("COM port is not an AT port", self.popup_titles("error"))
        self.assertEqual(self.chip("AT port"), "✗ AT port")


class OfflineSessionTest(GuiTestCase):

    def test_new_session_starts_clean(self):
        # Sheet and logs of the previous session are deleted when the tool opens
        self.close_app(self.app)
        (self.app_dir / SHEET).write_bytes(EXCEL)
        (self.app_dir / "Eagle_Wireless_SW_AT_Pretest_Log_20261021_090000.txt").write_text("old")
        (self.app_dir / "keep.txt").write_text("not a log")
        self.app = self.open_app()
        self.assertEqual(sorted(path.name for path in self.app_dir.iterdir()), ["keep.txt"])

    def test_full_offline_run(self):
        self.fill_inputs()
        self.place_excel()
        self.use_fake_worker("pass")
        self.press_start()
        self.assertTrue(self.app.is_running)
        self.wait_finished()
        self.assertEqual(self.app.status_label.cget("text"), "Completed  ·  ALL PASS")
        self.assertEqual(self.app.run_info["mode"], "Offline")
        self.assertEqual(self.app.run_info["firmware"], "I4916QGLLGR01A30M04A_ONT_V01")
        self.assertEqual(self.app.run_info["firmware_version"], "V01")
        self.assertEqual(str(self.app.download_button.cget("state")), "normal")
        self.assertEqual(self.app.log_path_label.cget("text"), "Eagle_Wireless_SW_AT_Pretest_Log_20261022_140000.txt")
        self.pump(0.5)
        self.assertFalse([request for request in self.jira.requests if request[0] == "POST"])   # Not posted


if __name__ == "__main__":
    unittest.main()
