##
# @file test_internet.py
# @brief Online / Offline button and the internet check.
##

import unittest
from unittest import mock

from support import GuiTestCase, gui


class InternetCheckTest(GuiTestCase):

    def test_starts_offline(self):
        self.assertEqual(self.app.connection_button.cget("text"), "Offline")
        self.assertEqual(self.app.connection_button.cget("bg"), gui.COLORS["offline"])
        self.assertFalse(self.app.internet_button.winfo_manager())     # ⟳ only in online mode

    def test_online_is_grey_while_checking_then_green(self):
        self.app._toggle_connection()
        self.app.update()
        self.assertEqual(self.app.connection_button.cget("bg"), gui.COLORS["muted"])
        self.assertTrue(self.app.is_internet_checking)
        self.wait_until(lambda: not self.app.is_internet_checking, what="the internet check")
        self.assertTrue(self.app.is_internet_ok)
        self.assertEqual(self.app.connection_button.cget("bg"), gui.COLORS["online"])
        self.assertEqual(self.app.internet_button.cget("text"), "⟳")
        self.assertFalse(self.help_shown())
        # Order left to right: ⟳, Online
        self.assertEqual([w.cget("text") for w in self.app.connection_button.master.pack_slaves()],
                         ["⟳", "Online"])

    def test_no_internet_stays_grey_with_help(self):
        with mock.patch.object(gui, "JIRA_BASE_URL", "http://127.0.0.1:1"), \
                mock.patch.object(gui, "INTERNET_CHECK_SEC", 2), mock.patch.object(gui, "INTERNET_RETRY_MS", 500):
            self.go_online()
        self.assertFalse(self.app.is_internet_ok)
        self.assertEqual(self.app.connection_button.cget("bg"), gui.COLORS["muted"])
        self.assertIn("No internet connection", self.popup_titles("error"))
        self.assertTrue(self.help_shown())
        self.assertEqual([w.cget("text") for w in self.app.connection_button.master.pack_slaves()],
                         ["?", "⟳", "Online"])
        self.assertTrue(self.chip("Excel").startswith("✗"))
        self.assertFalse(self.start_enabled())

    def test_switching_offline_cancels_the_check(self):
        with mock.patch.object(gui, "JIRA_BASE_URL", "http://127.0.0.1:1"):
            self.app._toggle_connection()
            self.app.update()
            self.app._toggle_connection()
            self.pump(0.5)
        self.assertFalse(self.app.is_internet_checking)
        self.assertEqual(self.app.connection_button.cget("text"), "Offline")
        self.assertEqual(self.app.connection_button.cget("bg"), gui.COLORS["offline"])
        self.assertEqual(self.popups, [])

    def test_help_box_tips(self):
        self.app._show_connection_help()
        kind, title, text = self.popups[-1]
        for tip in ("Turn ON the internet connection", "JIRA login username", "API token that is still alive",
                    "correct JIRA ID", f"Create API token here : {gui.API_TOKEN_URL}"):
            self.assertIn(tip, text)


if __name__ == "__main__":
    unittest.main()
