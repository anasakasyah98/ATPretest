##
# @file test_token.py
# @brief API token check after the internet check, remembered token, and the
#        token box that cannot be copied.
##

import os
import ctypes
import unittest

from support import GuiTestCase, gui, USER, TOKEN


class TokenCheckTest(GuiTestCase):

    def desc(self):
        return self.app.token_desc.cget("text")

    def test_nothing_remembered(self):
        self.go_online()
        self.assertEqual(self.desc(), "No token found. Please input API token.")
        self.assertEqual(self.chip("Excel"), "✗ API")

    def test_typed_bad_token(self):
        self.go_online()
        self.enter_token("bad-token")
        self.assertEqual(self.desc(), "✗ Token is invalid or expired. Please input API token.")
        self.assertEqual(self.chip("Excel"), "✗ API")
        self.assertTrue(self.help_shown())
        self.assertEqual(self.tokens, {})       # Not remembered

    def test_typed_good_token_is_remembered(self):
        self.go_online()
        self.enter_token()
        self.assertEqual(self.desc(), "✓ Token valid")
        self.assertEqual(self.chip("Excel"), "✓ API")
        self.assertEqual(self.tokens, {USER: TOKEN})

    def test_remembered_token_is_filled_in(self):
        self.tokens[USER] = TOKEN
        self.go_online()
        self.wait_until(lambda: self.app.token_state[1] == "ok", what="the token check")
        self.assertEqual(self.app.token_var.get(), TOKEN)
        self.assertEqual(self.desc(), "✓ Token valid")
        self.assertEqual(self.chip("Excel"), "✓ API")

    def test_remembered_expired_token(self):
        self.tokens[USER] = "expired-token"
        self.go_online()
        self.wait_until(lambda: self.desc() != "Checking the API token…", what="the token check")
        self.assertEqual(self.desc(), "Token expired. Please input API token.")
        self.assertEqual(self.app.token_var.get(), "")
        self.assertTrue(self.help_shown())

    def test_token_cannot_be_copied(self):
        self.tokens[USER] = TOKEN
        self.go_online()
        self.wait_until(lambda: self.app.token_var.get() == TOKEN, what="the token to be filled in")
        self.app.clipboard_clear()
        self.app.clipboard_append("unchanged")
        self.app.token_entry.selection_range(0, "end")
        for sequence in ("<<Copy>>", "<<Cut>>"):
            self.app.token_entry.event_generate(sequence)
        self.app.update()
        self.assertEqual(self.app.clipboard_get(), "unchanged")
        self.assertEqual(self.app.token_var.get(), TOKEN)

    def test_offline_shows_default_text(self):
        self.go_online()
        self.enter_token()
        self.app._toggle_connection()
        self.app.update()
        self.assertEqual(self.desc(), f"Used with {USER}")


@unittest.skipUnless(os.name == "nt", "Windows only")
class TokenStoreTest(unittest.TestCase):
    # Uses the real token store with a throwaway username, deleted afterwards
    username = "eaglepretest.selftest@eaglewireless.com"

    def tearDown(self):
        ctypes.WinDLL("advapi32").CredDeleteW(gui.CREDENTIAL_TARGET.format(username=self.username), 1, 0)

    def test_save_and_load(self):
        self.assertEqual(gui.load_token(self.username), "")
        gui.save_token(self.username, "token-1")
        self.assertEqual(gui.load_token(self.username), "token-1")
        gui.save_token(self.username, "token-2")        # Replaced, not added
        self.assertEqual(gui.load_token(self.username), "token-2")


if __name__ == "__main__":
    unittest.main()
