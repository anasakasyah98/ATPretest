##
# @file test_result.py
# @brief PDF report, result file names, and the result comment posted to the
#        JIRA ticket after an online pre-test.
##

import re
import time
import unittest

from support import GuiTestCase, TOKEN


class ResultTest(GuiTestCase):

    ##
    # @brief Fake a finished SDX35 run (no modem): every supported command gets a result.
    ##
    def finish_run(self, token=TOKEN, fails=0, aborted=False):
        app = self.app
        app.select_module("SDX35")
        app._populate_table()
        app.run_started_at = time.mktime((2026, 10, 22, 14, 0, 0, 0, 0, -1))
        app.run_info = {"module": "SDX35 (NR)", "jira": "SDX35-112", "ccid": "1" * 20, "imei": "2" * 15,
                        "imsi": "3" * 15, "firmware": "I4916QGLLGR01A30M04A_ONT_V01", "firmware_version": "V01",
                        "mode": "Online" if token else "Offline"}
        app.run_token = token
        count = 0
        for number, record in app.records.items():
            if "NA" not in app.tree.item(str(number), "tags"):
                record.receive = f"{record.send}\n+RESP: \"value <&>\"\n\nOK\n"
                app._on_result(record, "FAIL" if count < fails else "PASS")
                count += 1
        log = self.app_dir / "Eagle_Wireless_SW_AT_Pretest_Log_20261022_140000.txt"
        log.write_text("log line\n")
        app.log_file_path = log
        app.is_running, app.is_aborted, app.is_summary_seen = True, aborted, True
        app._on_finished()

    def wait_post(self):
        self.wait_until(lambda: "Result posted" in self.app.detail_label.cget("text") or self.popups,
                        what="the result to be posted")

    def comment(self):
        return [item[2] for item in self.jira.posted if item[0] == "comment"][0]

    def test_all_pass_comment(self):
        self.finish_run()
        self.wait_post()
        attached = [(item[2], item[3][:4]) for item in self.jira.posted if item[0] == "attach"]
        self.assertEqual(attached, [("EaglePretest_Report_20261022_140000.pdf", b"%PDF"),
                                    ("EaglePretest_Report_20261022_140000.txt", b"log ")])
        self.assertEqual(self.comment().split("\n\n"), [
            "*EaglePretest*",
            "AT command pre-test has been completed. Here is the result:",
            "Result → All pass, no fail",
            "Report → [^EaglePretest_Report_20261022_140000.pdf]",
            "Log → [^EaglePretest_Report_20261022_140000.txt]"])
        self.assertIn("Result posted to SDX35-112", self.app.detail_label.cget("text"))

    def test_fail_comment(self):
        self.finish_run(fails=2)
        self.wait_post()
        self.assertIn("Result → Not all pass, 2 fail", self.comment())

    def test_post_rejected(self):
        self.go_online()        # The ? box is only shown in online mode
        self.finish_run(token="bad-token")
        self.wait_post()
        self.assertIn("Cannot post to JIRA", self.popup_titles("error"))
        self.assertTrue(self.help_shown())

    def test_offline_or_stopped_run_is_not_posted(self):
        for token, aborted in (("", False), (TOKEN, True)):
            with self.subTest(token=token, aborted=aborted):
                self.jira.requests.clear()
                self.finish_run(token=token, aborted=aborted)
                self.pump(1)
                self.assertFalse([request for request in self.jira.requests if request[0] == "POST"])

    def test_pdf_report(self):
        self.finish_run(token="")
        target = self.app_dir / self.app._result_file_name(".pdf")
        self.app._write_report(target)
        data = target.read_bytes()
        self.assertTrue(data.startswith(b"%PDF"))
        self.assertLess(len(data), 200_000)     # Logo is downscaled
        self.assertTrue(re.fullmatch(r"EaglePretest_Report_\d{8}_\d{6}\.pdf", target.name))


if __name__ == "__main__":
    unittest.main()
