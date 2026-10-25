##
# @file support.py
# @brief Shared helpers for the GUI tests: a mock JIRA server on this PC and a
#        test case base class that opens the GUI against it.
#
# Nothing touches the real JIRA, the real tool folder, or a modem:
# - JIRA calls go to MockJira (127.0.0.1)
# - the tool folder is a temporary folder
# - the API token store is kept in memory
# - the COM port check always passes, and the pre-test itself is not run
# - popups are recorded instead of shown
#
# @author Anas Akasyah Bin Abdul Patas
# @copyright Eagle Wireless
##

import io
import sys
import json
import time
import email
import base64
import tempfile
import threading
import unittest
import warnings
from unittest import mock
from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from openpyxl import Workbook

import gui

USER = "eaglepretest.test@eaglewireless.com"
TOKEN = "good-token"
PORT = "COM99"


def build_excel(rows=(("platformName", "SDX35"), ("subversion", "V01"))):
    book = Workbook()
    sheet = book.active
    sheet.append(["Jenkins Parameter", "Value"])
    for row in rows:
        sheet.append(row)
    data = io.BytesIO()
    book.save(data)
    return data.getvalue()


EXCEL = build_excel()


##
# @brief Local stand-in for the JIRA REST API used by the tool.
##
class MockJira:

    def __init__(self):
        self.good_auth = "Basic " + base64.b64encode(f"{USER}:{TOKEN}".encode()).decode()
        # JIRA ID -> {"summary": topic, "attachments": [(file name, created, content)]}
        self.tickets = {
            "SDX35-112": {"summary": "SDX35 V01 Pretest", "attachments": [
                ("notes.txt", "2026-10-21T10:00:00.000+0800", b"notes"),
                ("old_params.xlsx", "2026-10-01T10:00:00.000+0800", EXCEL),
                ("I4916QGLLGR01A30M04A_ONT_V01 (16c488dd-2a91-4bd5-8559-efc4a6361bf3).xlsx",
                 "2026-10-20T10:00:00.000+0800", EXCEL)]},
            "SDX35-113": {"summary": "sdx35 pre-test build", "attachments": []},
            "SDX35-114": {"summary": "SDX35 PRE TEST", "attachments": []},
            "SDX35-200": {"summary": "SDX35 Pretest without sheet", "attachments": []},
            "SDX35-300": {"summary": "SDX35 release build", "attachments": []},
        }
        self.requests = []      # (method, path)
        self.posted = []        # ("attach", JIRA ID, file name, content) / ("comment", JIRA ID, text)
        self.server = HTTPServer(("127.0.0.1", 0), self._handler())
        threading.Thread(target=self.server.serve_forever, daemon=True).start()
        self.url = f"http://127.0.0.1:{self.server.server_port}"

    def close(self):
        self.server.shutdown()
        self.server.server_close()

    def summary_requests(self):
        return [path.split("/")[5].split("?")[0] for method, path in self.requests
                if method == "GET" and path.endswith("?fields=summary")]

    def _handler(self):
        jira = self

        class Handler(BaseHTTPRequestHandler):
            def log_message(self, *args):
                pass

            def send(self, code, body=b"{}", kind="application/json"):
                self.send_response(code)
                self.send_header("Content-Type", kind)
                self.end_headers()
                self.wfile.write(body)

            def do_GET(self):
                jira.requests.append(("GET", self.path))
                if self.path == "/":
                    return self.send(404)       # The site answers (internet check)
                if self.headers.get("Authorization") != jira.good_auth:
                    return self.send(401)
                if self.path == "/rest/api/3/myself":
                    return self.send(200)
                if self.path.startswith("/rest/api/3/issue/"):
                    ticket = jira.tickets.get(self.path.split("/")[5].split("?")[0])
                    if ticket is None:
                        return self.send(404)
                    attachments = [{"filename": name, "created": created,
                                    "content": f"{jira.url}/attachment/{id(content)}/{index}"}
                                   for index, (name, created, content) in enumerate(ticket["attachments"])]
                    fields = {"summary": ticket["summary"], "attachment": attachments}
                    return self.send(200, json.dumps({"fields": fields}).encode())
                if self.path.startswith("/attachment/"):
                    index = int(self.path.rsplit("/", 1)[1])
                    for ticket in jira.tickets.values():
                        for position, (_, _, content) in enumerate(ticket["attachments"]):
                            if position == index and str(id(content)) == self.path.split("/")[2]:
                                return self.send(200, content, "application/octet-stream")
                return self.send(404)

            def do_POST(self):
                jira.requests.append(("POST", self.path))
                body = self.rfile.read(int(self.headers["Content-Length"]))
                if self.headers.get("Authorization") != jira.good_auth:
                    return self.send(401)
                parts = self.path.split("/")
                jira_id = parts[5] if len(parts) > 6 else ""
                if jira_id not in jira.tickets:
                    return self.send(404)
                if parts[6] == "attachments":
                    if self.headers.get("X-Atlassian-Token") != "no-check":
                        return self.send(403)
                    message = email.message_from_bytes(b"Content-Type: " + self.headers["Content-Type"].encode()
                                                       + b"\r\n\r\n" + body)
                    part = message.get_payload()[0]
                    jira.posted.append(("attach", jira_id, part.get_filename(), part.get_payload(decode=True)))
                    return self.send(200, b"[]")
                if parts[6] == "comment":
                    jira.posted.append(("comment", jira_id, json.loads(body)["body"]))
                    return self.send(201)
                return self.send(404)

        return Handler


##
# @brief Opens the GUI against MockJira for every test, and closes it after.
##
class GuiTestCase(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        gui.enable_dpi_awareness()

    def setUp(self):
        warnings.simplefilter("ignore", ResourceWarning)     # Closed by the garbage collector, not a failure
        self.jira = MockJira()
        self.addCleanup(self.jira.close)
        folder = tempfile.TemporaryDirectory()
        self.addCleanup(folder.cleanup)
        self.app_dir = Path(folder.name)
        self.tokens = {}            # Stand-in for the API token store
        self.popups = []            # (kind, title, text)
        self.launched = []          # JIRA IDs the pre-test was started with

        def popup(kind, answer=None):
            def show(title, text, **kwargs):
                self.popups.append((kind, title, text))
                return answer
            return show

        for target, name, value in (
                (gui, "JIRA_BASE_URL", self.jira.url),
                (gui, "JIRA_BROWSE_URL", self.jira.url + "/browse/"),
                (gui, "TOKEN_CHECK_DELAY_MS", 200),
                (gui, "app_directory", lambda: self.app_dir),
                (gui, "load_token", lambda username: self.tokens.get(username, "")),
                (gui, "save_token", lambda username, token: self.tokens.__setitem__(username, token)),
                (gui, "probe_at_port", lambda port: (True, "")),
                (gui.messagebox, "showerror", popup("error")),
                (gui.messagebox, "showwarning", popup("warning")),
                (gui.messagebox, "showinfo", popup("info")),
                (gui.messagebox, "askokcancel", popup("ask", True)),
                (gui.messagebox, "askyesno", popup("ask", True))):
            patcher = mock.patch.object(target, name, value)
            patcher.start()
            self.addCleanup(patcher.stop)

        self.app = self.open_app()

    def open_app(self):
        app = gui.PretestApp(USER)
        app._launch_test = lambda: self.launched.append(app.jira_var.get().strip().upper())
        self.addCleanup(self.close_app, app)
        app.update()
        return app

    @staticmethod
    def close_app(app):
        try:
            for job in app.tk.call("after", "info"):
                app.after_cancel(job)
            app.destroy()
        except Exception:
            pass    # Already closed by the test

    # Waiting

    def pump(self, seconds=0.3):
        end = time.monotonic() + seconds
        while time.monotonic() < end:
            self.app.update()
            time.sleep(0.02)

    def wait_until(self, condition, timeout=10, what="condition"):
        end = time.monotonic() + timeout
        while time.monotonic() < end:
            self.app.update()
            if condition():
                self.app.update()
                return
            time.sleep(0.02)
        self.fail(f"Timed out waiting for {what}")

    # Steps

    def fill_inputs(self, jira_id="SDX35-112"):
        app = self.app
        app.ports[PORT] = "Eagle USB AT Port"
        app.port_var.set(PORT)
        app._set_port_state(PORT, "ok", "")
        app.ccid_var.set("1" * 20)
        app.imei_var.set("2" * 15)
        app.imsi_var.set("3" * 15)
        app.jira_var.set(jira_id)

    def go_online(self):
        self.app._toggle_connection()
        self.wait_until(lambda: not self.app.is_internet_checking, what="the internet check")

    def enter_token(self, token=TOKEN):
        self.app.token_var.set(token)
        self.wait_until(lambda: self.app.token_state[1] != "checking", what="the token check")

    def wait_topic(self):
        jira_id = self.app.jira_var.get().strip().upper()
        self.wait_until(lambda: self.app.topic_state[0] == jira_id
                        and self.app.topic_state[1] not in ("none", "checking"), what="the JIRA topic check")

    def set_jira(self, jira_id):
        self.app.jira_var.set(jira_id)
        self.wait_topic()

    ##
    # @brief Run the real pre-test start, with fake_worker.py in place of the
    #        worker process (no modem).
    ##
    def use_fake_worker(self, scenario):
        real_popen = gui.subprocess.Popen
        worker = str(Path(__file__).with_name("fake_worker.py"))

        def popen(cmd, **kwargs):
            module_key = cmd[cmd.index(gui.WORKER_FLAG) + 1]
            return real_popen([sys.executable, "-u", worker, module_key, scenario], **kwargs)

        patcher = mock.patch.object(gui.subprocess, "Popen", popen)
        patcher.start()
        self.addCleanup(patcher.stop)
        del self.app._launch_test       # Back to the real start
        self.addCleanup(self.stop_worker)

    def stop_worker(self):
        process = self.app.process
        if process is not None and process.poll() is None:
            process.kill()
            process.wait()

    def place_excel(self, name="I4916QGLLGR01A30M04A_ONT_V01.xlsx", content=EXCEL):
        (self.app_dir / name).write_bytes(content)
        self.app.reload_excel()

    def wait_finished(self, timeout=30):
        self.wait_until(lambda: not self.app.is_running, timeout, what="the pre-test to finish")

    def results(self):
        return {number: record.result for number, record in self.app.records.items()}

    def press_start(self):
        self.app.start_test()
        self.app.update()
        self.wait_until(lambda: not self.app.is_checking_port and not self.app.is_downloading, what="Start")

    # Reading the UI

    def chip(self, name):
        return self.app.precheck_chips[name].cget("text")

    def start_enabled(self):
        return str(self.app.start_button.cget("state")) == "normal"

    def help_shown(self):
        return bool(self.app.help_button.winfo_manager())

    def popup_titles(self, kind=None):
        return [title for popup_kind, title, _ in self.popups if kind in (None, popup_kind)]
