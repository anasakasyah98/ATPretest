##
# @file test_download.py
# @brief Online mode Start: download the newest Excel sheet from the JIRA
#        ticket into the tool folder, then run the pre-test.
##

import unittest

from support import GuiTestCase, EXCEL, gui

NEWEST_SHEET = "I4916QGLLGR01A30M04A_ONT_V01.xlsx"      # Saved without the ID JIRA adds


class DownloadTest(GuiTestCase):

    def setUp(self):
        super().setUp()
        self.fill_inputs()
        self.go_online()
        self.enter_token()
        self.wait_topic()

    def files(self):
        return sorted(path.name for path in self.app_dir.iterdir())

    def place_old_sheets(self):
        for name in ("old_A.xlsx", "old_B.xlsx"):
            (self.app_dir / name).write_bytes(EXCEL)

    def test_downloads_newest_sheet_and_starts(self):
        self.place_old_sheets()
        self.assertTrue(self.start_enabled())
        self.press_start()
        self.assertEqual(self.launched, ["SDX35-112"])
        self.assertEqual(self.files(), [NEWEST_SHEET])
        self.assertEqual(self.popups, [])

    def test_sheet_open_in_excel(self):
        self.place_old_sheets()
        with open(self.app_dir / "old_A.xlsx", "rb"):       # Locked, like a sheet open in Excel
            self.press_start()
        self.assertEqual(self.launched, [])
        self.assertIn("Excel download failed", self.popup_titles("error"))
        self.assertIn("old_A.xlsx", self.popups[-1][2])
        self.popups.clear()
        self.press_start()          # Closed: works now
        self.assertEqual(self.launched, ["SDX35-112"])
        self.assertEqual(self.files(), [NEWEST_SHEET])

    def test_ticket_without_excel(self):
        self.set_jira("SDX35-200")
        self.press_start()
        self.assertEqual(self.launched, [])
        self.assertIn("SDX35-200 has no Excel (.xlsx) attachment.", self.popups[-1][2])
        self.assertTrue(self.help_shown())

    def test_token_revoked_before_start(self):
        self.jira.good_auth = "Basic revoked"
        self.press_start()
        self.assertEqual(self.launched, [])
        self.assertIn("JIRA login failed", self.popups[-1][2])


class AttachmentNameTest(unittest.TestCase):

    def test_jira_id_removed(self):
        name = "I4916QGLLGR01A30M04A_ONT_V01 (16c488dd-2a91-4bd5-8559-efc4a6361bf3)"
        self.assertEqual(gui.RE_ATTACHMENT_ID.sub("", name), "I4916QGLLGR01A30M04A_ONT_V01")
        self.assertEqual(gui.RE_ATTACHMENT_ID.sub("", "Build (V2)"), "Build (V2)")


if __name__ == "__main__":
    unittest.main()
