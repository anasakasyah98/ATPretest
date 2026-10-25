##
# @file test_parser.py
# @brief Output parser: the console output of the worker process is read into
#        the Test Cases table, the progress, and the final status.
##

import unittest

from support import GuiTestCase, gui
from test_case import at_cmd_test_case_list

TOTAL = len(at_cmd_test_case_list)


class ParserTest(GuiTestCase):

    def run_fake(self, scenario, jira_id="SDX35-112"):
        self.fill_inputs(jira_id)
        self.place_excel()
        self.use_fake_worker(scenario)
        self.press_start()
        self.assertTrue(self.app.is_running, self.popups)

    def unsupported(self, key):
        profile = gui.MODULE_PROFILES[key]
        return {number for number, test_case in enumerate(at_cmd_test_case_list, start=1)
                if profile.is_unsupported(test_case)}

    def table_result(self, number):
        return self.app.tree.item(str(number), "values")[5]

    def test_all_pass(self):
        self.run_fake("pass")
        self.wait_finished()
        skipped = self.unsupported("SDX35")
        expected = {number: "UNSUPPORTED" if number in skipped else "PASS" for number in range(1, TOTAL + 1)}
        self.assertEqual(self.results(), expected)
        self.assertEqual(self.app.counts, {"PASS": TOTAL - len(skipped), "FAIL": 0, "UNSUPPORTED": len(skipped)})
        self.assertTrue(self.app.is_summary_seen)
        self.assertEqual(self.app.status_label.cget("text"), "Completed  ·  ALL PASS")
        self.assertEqual(self.app.detail_label.cget("text"),
                         f"{TOTAL - len(skipped)} passed, {len(skipped)} skipped")
        self.assertEqual(self.table_result(1), "✓ PASS")
        self.assertEqual(self.table_result(min(skipped)), "– SKIP")
        self.assertEqual(self.app.percent_label.cget("text"), "100 %")

    def test_failures(self):
        self.run_fake("fail:3,5")
        self.wait_finished()
        self.assertEqual(self.results()[3], "FAIL")
        self.assertEqual(self.results()[5], "FAIL")
        self.assertEqual(self.app.counts["FAIL"], 2)
        self.assertEqual(self.app.status_label.cget("text"), "Completed  ·  2 FAILED")
        self.assertEqual(self.table_result(3), "✗ FAIL")
        self.assertEqual(self.app.tree.selection(), ("3",))       # First failure selected

    def test_command_response_and_expected(self):
        self.run_fake("pass")
        self.wait_finished()
        record = self.app.records[3]
        test_case = at_cmd_test_case_list[2]
        self.assertEqual(record.send, test_case.at_command)
        self.assertEqual(record.receive, f"{test_case.at_command}\n+RESP: 3\n\nOK\n\n")     # Every response line kept
        self.assertEqual(record.expected, test_case.expected_response)
        self.assertEqual(self.app.tree.item("3", "values")[4], test_case.expected_response)

    def test_run_that_did_not_complete(self):
        self.run_fake("crash")
        self.wait_finished()
        self.assertFalse(self.app.is_summary_seen)
        self.assertEqual(self.app.status_label.cget("text"), "Pre-test did not complete")
        self.assertEqual(self.app.detail_label.cget("text"),
                         "Cannot open COM99: could not open port 'COM99': FileNotFoundError")
        self.assertEqual(self.app.counts, {"PASS": 0, "FAIL": 0, "UNSUPPORTED": 0})

    def test_progress_while_running(self):
        self.run_fake("stop:10")
        self.wait_until(lambda: self.app.records[10].result, what="10 results")
        self.assertTrue(self.app.is_running)
        self.assertEqual(self.app.percent_label.cget("text"), f"{int(10 * 100 / TOTAL)} %")
        self.assertEqual(self.app.status_label.cget("text"), f"Running  ·  10/{TOTAL} complete")
        self.assertIn(at_cmd_test_case_list[10].at_command, self.app.detail_label.cget("text"))   # Next command
        self.assertEqual(self.app.tree.item("11", "tags"), ("RUNNING",))

    def test_stopped_by_user(self):
        self.run_fake("stop:4")
        self.wait_until(lambda: self.app.records[4].result, what="4 results")
        self.app.stop_test()
        self.wait_finished()
        self.assertTrue(self.app.is_aborted)
        self.assertEqual(self.app.status_label.cget("text"), "Stopped by user")
        self.assertEqual(self.app.counts["PASS"], 4)

    def test_modem_reset_wait(self):
        # MDM9X07 waits for the modem after AT+CFUN=1,1 (command 18)
        self.assertEqual(at_cmd_test_case_list[17].at_command, gui.const.RESET_AT_COMMAND)
        self.run_fake("stop:18", jira_id="SW9X07-10")
        self.wait_until(lambda: self.app.records[18].result, what="18 results")
        self.assertTrue(self.app.status_label.cget("text").startswith("Modem resetting"))
        self.assertIsNotNone(self.app.reset_deadline)

    def test_log_file(self):
        self.run_fake("pass")
        self.wait_finished()
        self.assertEqual(self.app.log_file_path.name, "Eagle_Wireless_SW_AT_Pretest_Log_20261022_140000.txt")
        self.assertEqual(str(self.app.download_button.cget("state")), "normal")


class ParserPatternTest(unittest.TestCase):

    def test_patterns(self):
        self.assertEqual(gui.RE_INDEX.match("[3/44]").groups(), ("3", "44"))
        self.assertIsNone(gui.RE_INDEX.match("[3/44] extra"))
        self.assertEqual(gui.RE_RESULT.match("Result    : Pretest_LTE_CD-01-001 [1] : PASS").group(1), "PASS")
        self.assertEqual(gui.RE_RESULT.match("Result    : X [2] : UNSUPPORTED").group(1), "UNSUPPORTED")
        self.assertIsNone(gui.RE_RESULT.match("Result    : X [2] : MAYBE"))
        self.assertEqual(gui.RE_SEND.match("Send      : AT+QSVN=\"13\"").group(1), "AT+QSVN=\"13\"")
        self.assertEqual(gui.RE_TOTAL.match("Total : 44 | Pass : 40 | Fail : 2").groups(), ("44", "40", "2"))


if __name__ == "__main__":
    unittest.main()
