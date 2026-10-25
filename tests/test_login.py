##
# @file test_login.py
# @brief Login page: username rule and the hashed team password.
##

import hashlib
import unittest
from unittest import mock

from support import gui

TEST_PASSWORD = "test-password-123"
TEST_SALT = "00112233445566778899aabbccddeeff"
TEST_ITERATIONS = 1000      # Fast for the test, the tool uses LOGIN_PASSWORD_ITERATIONS


class LoginTest(unittest.TestCase):

    def setUp(self):
        test_hash = hashlib.pbkdf2_hmac("sha256", TEST_PASSWORD.encode(), bytes.fromhex(TEST_SALT),
                                        TEST_ITERATIONS).hex()
        for name, value in (("LOGIN_PASSWORD_SALT", TEST_SALT), ("LOGIN_PASSWORD_ITERATIONS", TEST_ITERATIONS),
                            ("LOGIN_PASSWORD_HASH", test_hash)):
            patcher = mock.patch.object(gui, name, value)
            patcher.start()
            self.addCleanup(patcher.stop)
        self.window = gui.LoginWindow()
        self.addCleanup(self.close)

    def close(self):
        try:
            for job in self.window.tk.call("after", "info"):
                self.window.after_cancel(job)
            self.window.destroy()
        except Exception:
            pass    # Already closed by a successful login

    def log_in(self, username, password):
        self.window.username_var.set(username)
        self.window.password_var.set(password)
        self.window.login()

    def test_correct_password(self):
        self.log_in("  Name.Test@EagleWireless.com ", TEST_PASSWORD)
        self.assertTrue(self.window.is_authenticated)
        self.assertEqual(self.window.username, "name.test@eaglewireless.com")

    def test_wrong_password(self):
        self.log_in("name.test@eaglewireless.com", "wrong")
        self.assertFalse(self.window.is_authenticated)
        self.assertEqual(self.window.error_label.cget("text"), "✗ Invalid password")
        self.assertEqual(self.window.password_var.get(), "")

    def test_wrong_username(self):
        for username in ("name.test@gmail.com", "@eaglewireless.com"):
            with self.subTest(username):
                self.log_in(username, TEST_PASSWORD)
                self.assertFalse(self.window.is_authenticated)
                self.assertIn("Username must end with", self.window.error_label.cget("text"))


class PasswordStoreTest(unittest.TestCase):

    def test_only_the_hash_is_stored(self):
        self.assertFalse(hasattr(gui, "LOGIN_PASSWORD"))
        self.assertEqual(len(gui.LOGIN_PASSWORD_HASH), 64)
        self.assertEqual(len(bytes.fromhex(gui.LOGIN_PASSWORD_SALT)), 16)
        self.assertGreaterEqual(gui.LOGIN_PASSWORD_ITERATIONS, 600000)
        self.assertNotEqual(gui.login_password_hash(""), gui.LOGIN_PASSWORD_HASH)

    def test_version_from_constants(self):
        self.assertEqual(gui.APP_VERSION, gui.const.SOFTWARE_VERSION)
        self.assertIn(f"Version  : {gui.const.SOFTWARE_VERSION}\n", gui.const.STARTUP_MESSAGE_STRING)


if __name__ == "__main__":
    unittest.main()
