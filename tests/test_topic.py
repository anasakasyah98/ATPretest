##
# @file test_topic.py
# @brief Online mode only accepts a Pretest JIRA (the JIRA topic is checked).
##

import unittest

from support import GuiTestCase, gui


class TopicCheckTest(GuiTestCase):

    def setUp(self):
        super().setUp()
        self.fill_inputs()
        self.go_online()
        self.enter_token()
        self.wait_topic()

    def desc(self):
        return self.app.jira_desc.cget("text")

    def test_pretest_jira_spellings(self):
        for jira_id in ("SDX35-112", "SDX35-113", "SDX35-114"):    # Pretest, pre-test, PRE TEST
            with self.subTest(jira_id):
                self.set_jira(jira_id)
                self.assertEqual(self.desc(), "✓ SDX35 (NR) - Pretest JIRA")
                self.assertEqual(self.chip("JIRA"), "✓ JIRA")
                self.assertTrue(self.start_enabled())
                self.assertFalse(self.help_shown())

    def test_not_a_pretest_jira(self):
        self.set_jira("SDX35-300")
        self.assertEqual(self.desc(), "✗ Not a Pretest JIRA")
        self.assertEqual(self.chip("JIRA"), "✗ JIRA")
        self.assertFalse(self.start_enabled())
        self.assertTrue(self.help_shown())

    def test_jira_not_found(self):
        self.set_jira("SDX35-999")
        self.assertEqual(self.desc(), "✗ JIRA not found")
        self.assertFalse(self.start_enabled())
        self.app.start_test()
        self.assertIn("Online mode: JIRA ID - JIRA not found.", self.popups[-1][2])
        self.assertEqual(self.launched, [])

    def test_checking_while_typed(self):
        self.app.jira_var.set("SDX35-113")
        self.app.update()
        self.assertEqual(self.desc(), "Checking the JIRA topic…")
        self.assertEqual(self.chip("JIRA"), "… JIRA")
        self.assertFalse(self.start_enabled())

    def test_fast_typing_checks_only_the_last_id(self):
        self.jira.requests.clear()
        for text in ("S", "SD", "SDX35-1", "SDX35-11", "SDX35-113"):
            self.app.jira_var.set(text)
            self.app.update()
        self.wait_topic()
        self.assertEqual(self.jira.summary_requests(), ["SDX35-113"])

    def test_offline_does_not_check(self):
        self.app._toggle_connection()
        self.jira.requests.clear()
        self.app.jira_var.set("SDX35-300")
        self.pump(0.5)
        self.assertEqual(self.desc(), "✓ SDX35 (NR)")
        self.assertEqual(self.jira.requests, [])

    def test_rechecked_at_start(self):
        # The topic is renamed after the check: the download at Start still blocks it
        self.jira.tickets["SDX35-112"]["summary"] = "SDX35 release"
        self.press_start()
        self.assertEqual(self.launched, [])
        self.assertIn("SDX35-112 is not a Pretest JIRA.", self.popups[-1][2])


class PretestPatternTest(unittest.TestCase):

    def test_pattern(self):
        for topic in ("SDX35 Pretest", "pretest", "PRE-TEST run", "Pre test", "pre_test"):
            self.assertTrue(gui.RE_PRETEST.search(topic), topic)
        for topic in ("SDX35 release", "Pre-release test", "Prestest"):
            self.assertFalse(gui.RE_PRETEST.search(topic), topic)


if __name__ == "__main__":
    unittest.main()
