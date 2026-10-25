##
# @file gui.py
# @brief Graphical front-end for the Eagle Wireless AT command pre-test tool.
#
# Uses the existing pre-test modules (mdm9x07, qcx216, sdx35) as services
# without modifying them. Every run is executed in a child process (this same
# script/executable started with --worker) so that:
# - the console output of the module is streamed live into the UI,
# - module state and the log file are fresh for every run,
# - a running test can be stopped safely (the COM port is released).
#
# Features:
# - Login page (Eagle Wireless username and team password)
# - Module selected from the JIRA ID (QCX216, SW9X07 -> MDM9X07, SDX35)
# - Offline mode: Jenkins Build Parameter Excel Sheet uploaded by the user
# - Online mode: Excel sheet downloaded from the JIRA ticket, result posted back as a comment
# - Live log console with PASS / FAIL / UNSUPPORTED highlighting
# - Progress bar with modem reset (AT+CFUN=1,1) countdown
# - Test case table with per-command details
# - Result download: log (.txt) or report (.pdf)
#
# @author Anas Akasyah Bin Abdul Patas
# @version 1.0.0
# @date 2026-10-26
# @copyright Eagle Wireless
##

# Libraries
import io
import os
import re
import sys
import hmac
import json
import time
import queue
import base64
import ctypes
import shutil
import hashlib
import logging
import tempfile
import itertools
import threading
import contextlib
import subprocess
import urllib.error
import urllib.parse
import urllib.request
from ctypes import wintypes
from dataclasses import dataclass
from pathlib import Path
from typing import Callable
from typing import Optional
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog
from tkinter.scrolledtext import ScrolledText
import serial.tools.list_ports
import constants as const
import main as pretest_main
import mdm9x07
import qcx216
import sdx35
from test_case import at_cmd_test_case_list
from test_case import at_cmd_general_info_list

WORKER_FLAG                         = "--worker"
MARKER_LOG_FILE                     = "@@GUI_LOG_FILE@@ "
MARKER_EXIT                         = "@@GUI_WORKER_DONE@@"

APP_TITLE                           = "EaglePretest"
APP_VERSION                         = const.SOFTWARE_VERSION

LOGIN_DOMAIN                        = "@eaglewireless.com"

# Team password: only its PBKDF2-SHA256 hash is stored, the password itself is not in the tool.
# New password -> new hash (run inside the src folder, the password is typed hidden):
#   python -c "import gui, getpass; print(gui.login_password_hash(getpass.getpass()))"
LOGIN_PASSWORD_SALT                 = "7de5cc815dc8ce73c7935c9d7013fcfb"
LOGIN_PASSWORD_ITERATIONS           = 600000
LOGIN_PASSWORD_HASH                 = "ffd95546ae16773b9666df8375fce15637084c90a93b9034433caac4ed8a75c4"

# JIRA ID project -> module (same rule as main.py)
JIRA_PROJECTS                       = {"QCX216": "QCX216", "SW9X07": "MDM9X07", "SDX35": "SDX35"}

# JIRA ticket link: JIRA_BROWSE_URL + JIRA ID
JIRA_BASE_URL                       = "https://eagleelectronics.atlassian.net"
JIRA_BROWSE_URL                     = JIRA_BASE_URL + "/browse/"
JIRA_TIMEOUT_SEC                    = 30

# ID JIRA can add to an attachment name, e.g. "Name (16c488dd-2a91-4bd5-8559-efc4a6361bf3)"
RE_ATTACHMENT_ID                    = re.compile(r"\s*\([0-9a-f]{8}(?:-[0-9a-f]{4}){3}-[0-9a-f]{12}\)$", re.IGNORECASE)

# Online mode only runs on a pre-test ticket: the JIRA topic must contain Pretest (Pre-test, Pre test)
RE_PRETEST                          = re.compile(r"pre[\s_-]?test", re.IGNORECASE)
NOT_PRETEST_REASON                  = "not a Pretest JIRA"
API_TOKEN_URL                       = "https://id.atlassian.com/manage-profile/security/api-tokens"

# Online mode internet check: retry the JIRA site until it answers or the time is up
INTERNET_CHECK_SEC                  = 30
INTERNET_RETRY_MS                   = 3000
INTERNET_PROBE_TIMEOUT_SEC          = 5
SPINNER_FRAMES                      = "◐◓◑◒"
SPINNER_MS                          = 150

# API token: checked after the internet check (and after typing), stored in Windows Credential Manager
TOKEN_CHECK_DELAY_MS                = 800
CREDENTIAL_TARGET                   = "EaglePretest JIRA API Token ({username})"

##
# @brief Directory used by the pre-test modules for the Excel sheet and logs.
#
# @return Application directory
##
def app_directory():
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).parent

LOGO_FILE                           = "EagleWireless_Logo.jpg"
ICON_SIZES                          = (16, 24, 32, 48, 64, 128, 256)

##
# @brief Locate a resource next to the script / executable, or bundled
#        inside the PyInstaller executable.
#
# @param [in] name Resource file name
#
# @return Resource path, or None if not found
##
def resource_path(name):
    for base in (app_directory(), Path(getattr(sys, "_MEIPASS", app_directory()))):
        if (base / name).is_file():
            return base / name
    return None

##
# @brief Build the window icon from the Eagle Wireless logo: the eagle emblem
#        (square at the top of the artwork, above the text) on a transparent
#        background.
#
# @return PIL RGBA image, or None if the logo / Pillow is not available
##
def load_logo_emblem():
    path = resource_path(LOGO_FILE)
    if path is None:
        return None
    try:
        from PIL import Image
        logo = Image.open(path).convert("RGB")
    except Exception:
        return None

    from PIL import ImageFilter, ImageOps

    # Label the dark shapes on a reduced copy of the artwork
    scale = 4
    width, height = logo.width // scale, logo.height // scale
    small = logo.convert("L").resize((width, height))
    dark = [v < 160 for v in small.tobytes()]
    labels = [0] * (width * height)
    shapes = []
    for start in range(width * height):
        if not dark[start] or labels[start]:
            continue
        label = len(shapes) + 1
        labels[start] = label
        stack, pixels = [start], []
        while stack:
            index = stack.pop()
            pixels.append(index)
            x, y = index % width, index // width
            for nx, ny in ((x + 1, y), (x - 1, y), (x, y + 1), (x, y - 1)):
                if 0 <= nx < width and 0 <= ny < height:
                    neighbor = ny * width + nx
                    if dark[neighbor] and not labels[neighbor]:
                        labels[neighbor] = label
                        stack.append(neighbor)
        ys = [index // width for index in pixels]
        shapes.append((pixels, min(ys), max(ys)))
    if not shapes:
        return None

    # The text starts at the top of the first tall letter in the lower part;
    # every shape starting from there down is text, everything above is the emblem
    artwork_top = min(top for _, top, _ in shapes)
    artwork_bottom = max(bottom for _, _, bottom in shapes)
    span = artwork_bottom - artwork_top
    letter_tops = [top for _, top, bottom in shapes
                   if top > artwork_top + 0.55 * span and bottom - top > 0.08 * span]
    text_top = min(letter_tops, default=artwork_bottom + 1)
    text_mask = Image.new("L", (width, height), 0)
    text_pixels = text_mask.load()
    emblem_box = None
    for pixels, top, bottom in shapes:
        if top >= text_top:
            for index in pixels:
                text_pixels[index % width, index // width] = 255
        else:
            xs = [index % width for index in pixels]
            box = (min(xs), top, max(xs) + 1, bottom + 1)
            emblem_box = box if emblem_box is None else (min(emblem_box[0], box[0]), min(emblem_box[1], box[1]),
                                                         max(emblem_box[2], box[2]), max(emblem_box[3], box[3]))
    if emblem_box is None:
        return None

    # Erase the text, then crop the emblem into a square with a small margin
    text_mask = text_mask.resize(logo.size, Image.NEAREST).filter(ImageFilter.MaxFilter(9))
    logo.paste((255, 255, 255), mask=text_mask)
    left, top, right, bottom = (value * scale for value in emblem_box)
    side = int(max(right - left, bottom - top) * 1.04)
    center_x, center_y = (left + right) // 2, (top + bottom) // 2
    padded = ImageOps.expand(logo, border=side, fill=(255, 255, 255))   # Crop may reach past the edges
    left, top = center_x - side // 2 + side, center_y - side // 2 + side
    emblem = padded.crop((left, top, left + side, top + side))

    # White background -> transparent, keeping anti-aliased edges
    alpha = emblem.convert("L").point(lambda v: max(0, min(255, (235 - v) * 2)))
    emblem = emblem.convert("RGBA")
    emblem.putalpha(alpha)
    return emblem

# Module profiles

##
# @brief Module specific settings used by the GUI.
##
@dataclass(frozen=True)
class ModuleProfile:
    key                 :   str
    technology          :   str
    excel_platform      :   str
    accent              :   str
    service             :   object
    entry               :   str
    reset_delay         :   Optional[int]
    excel_keys          :   tuple
    qgmr_rule           :   str
    notes               :   tuple
    is_unsupported      :   Callable = lambda test_case: False

    ##
    # @brief Run the pre-test of this module (entry function of the module script).
    #
    # @param [in] username Login username
    # @param [in] port COM port, e.g. COM5
    # @param [in] ccid SIM CCID
    # @param [in] imei SIM IMEI
    # @param [in] imsi SIM IMSI
    #
    # @return None
    ##
    def run(self, username, port, ccid, imei, imsi):
        getattr(self.service, self.entry)(username, port, ccid, imei, imsi)

##
# @brief Check if a test case is skipped by SDX35 (same rule as sdx35.at_execute_command).
#
# @param [in] test_case Test case
#
# @return True if the test case is skipped, otherwise False
##
def _sdx35_unsupported(test_case):
    # Mirrors the skip rule inside sdx35.at_execute_command
    return ("Pretest_LTE_CD-04" in test_case.test_case_id) or ("AT+QGSN" in test_case.at_command)

MODULE_PROFILES = {
    "MDM9X07": ModuleProfile(
        key             =   "MDM9X07",
        technology      =   "LTE",
        excel_platform  =   "MDM9207",
        accent          =   "#1d4ed8",
        service         =   mdm9x07,
        entry           =   "pretest_mdm9x07",
        reset_delay     =   const.RESET_DELAY_MDM9X07_SEC,
        excel_keys      =   (("Platform", "platformName"), ("Branch Name", "branch_name"),
                             ("Cert Version", "certversion"), ("Sub-Version", "subversion"),
                             ("SVN", "svn"), ("Version", "version"), ("Version Info", "version_info"),
                             ("OEM", "oem_id"), ("Modem App", "modemapp"), ("Linux App", "linuxapp")),
        qgmr_rule       =   "certversion _modemapp .linuxapp[0:2]",
        notes           =   ("ATI expects the build version",
                             "AT+QSVN=99 expects ERROR",
                             f"Modem reset wait: {const.RESET_DELAY_MDM9X07_SEC} s"),
    ),
    "QCX216": ModuleProfile(
        key             =   "QCX216",
        technology      =   "LTE",
        excel_platform  =   "QCX216",
        accent          =   "#0f766e",
        service         =   qcx216,
        entry           =   "pretest_qcx216",
        reset_delay     =   const.RESET_DELAY_QCX216_SEC,
        excel_keys      =   (("Platform", "platformName"), ("Branch Name", "branch_name"),
                             ("Cert Version", "certversion"), ("Sub-Version", "subversion"),
                             ("SVN", "svn"), ("Version", "version"), ("Version Info", "version_info"),
                             ("OEM", "OEM"), ("Modem App", "modemapp"), ("Linux App", "linuxapp")),
        qgmr_rule       =   "version _OEM _modemapp.modemapp",
        notes           =   ("ATI expects the cert version",
                             "AT+QSVN=99 expects OK",
                             f"Modem reset wait: {const.RESET_DELAY_QCX216_SEC} s"),
    ),
    "SDX35": ModuleProfile(
        key             =   "SDX35",
        technology      =   "NR",
        excel_platform  =   "",
        accent          =   "#6d28d9",
        service         =   sdx35,
        entry           =   "pretest_sdx35",
        reset_delay     =   None,
        excel_keys      =   (("Branch Name", "branchname"), ("Cert Version", "certversion"),
                             ("Sub-Version", "subversion"), ("SVN", "svn"), ("Version", "version"),
                             ("Version Info", "version_info"), ("Modem App", "modemapp"),
                             ("Linux App", "linuxapp")),
        qgmr_rule       =   "certversion _modemapp .linuxapp[0:2]",
        notes           =   ("CD-04 (QSVN) and AT+QGSN are skipped",
                             "Version checks use the build version",
                             "No modem reset handling"),
        is_unsupported  =   _sdx35_unsupported,
    ),
}

##
# @brief Map the JIRA ID to a module key (same rule as main.py).
#
# @param [in] jira_id JIRA ID, e.g. SDX35-112
#
# @return Module key, or None if the JIRA ID has no known project
##
def module_from_jira(jira_id):
    for project, key in JIRA_PROJECTS.items():
        if project in jira_id:
            return key
    return None

# Worker (child process)

##
# @brief Entry point of the child process. Runs one module pre-test and
#        streams its console output to the parent through stdout.
#
# @param [in] argv [module key, username, COM port, SIM CCID, SIM IMEI, SIM IMSI]
#
# @return None
##
def run_worker(argv):
    if sys.stdout is None:
        sys.stdout = open(1, "w", encoding="utf-8", errors="replace", closefd=False)
    else:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
    if sys.stderr is None:
        sys.stderr = sys.stdout
    else:
        sys.stderr.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)

    module_key, username, port, ccid, imei, imsi = argv

    # Every time a module opens the port (start of the run and after each
    # AT+CFUN=1,1 reset), wait until the modem answers before its commands are sent
    ##
    # @brief Serial port that waits until the modem answers AT every time it is opened.
    ##
    class ATReadySerial(serial.Serial):
        ##
        # @brief Open the port, then wait until the modem answers AT.
        #
        # @return None
        ##
        def open(self):
            super().open()
            started = time.monotonic()
            is_ready, _ = wait_for_at_ready(self, AT_READY_TIMEOUT_SEC)
            waited = time.monotonic() - started
            if is_ready:
                print(f"[Port] {self.port} ready after {waited:.1f} s")
            else:
                print(f"[Port] {self.port} did not answer AT within {AT_READY_TIMEOUT_SEC} s - continuing")
            self.reset_input_buffer()

    serial.Serial = ATReadySerial

    try:
        MODULE_PROFILES[module_key].run(username, port, ccid, imei, imsi)
    except Exception as e:
        print(f"Error: {e}")
    finally:
        for handler in logging.getLogger().handlers:
            if isinstance(handler, logging.FileHandler):
                handler.flush()
                print(MARKER_LOG_FILE + handler.baseFilename)
        print(MARKER_EXIT)
        sys.stdout.flush()

# AT port check (in the UI process, before the pre-test opens the port)

PROBE_TIMEOUT_SEC                   = 5               # Port check before a run
AT_READY_TIMEOUT_SEC                = 20           # Each port open during a run (incl. after AT+CFUN=1,1)
AT_READY_RETRY_SEC                  = 0.5
PORT_CHECK_DELAY_MS                 = 400
NON_AT_PORT_PATTERN                 = re.compile(r"\b(DM|NMEA|GPS|Diagnostics?)\b", re.IGNORECASE)

# Only one port check at a time, so a port is never opened twice
probe_lock = threading.Lock()

##
# @brief Check that a COM port is a modem AT port: send "AT" and wait for "OK".
#
# @param [in] port Serial port
#
# @return (True, "") if the port answered, otherwise (False, reason)
##
def probe_at_port(port):
    with probe_lock:
        return _probe_at_port(port)

##
# @brief Wait until the modem answers "AT" on an open port.
#
# Some modems (e.g. QCX216) drop everything sent during roughly the first
# second after the port is opened, so "AT" is repeated until "OK" arrives.
#
# @param [in] serial_port Open serial port
# @param [in] timeout Maximum wait in seconds
#
# @return (True, response) once "OK" was received, otherwise (False, response)
##
def wait_for_at_ready(serial_port, timeout):
    previous_timeout = serial_port.timeout
    serial_port.timeout = 0.1
    response = ""
    try:
        deadline = time.monotonic() + timeout
        while time.monotonic() < deadline:
            serial_port.reset_input_buffer()
            serial_port.write(b"AT\r\n")
            retry_at = min(deadline, time.monotonic() + AT_READY_RETRY_SEC)
            while time.monotonic() < retry_at:
                data = serial_port.read(serial_port.in_waiting or 1)
                if data:
                    response += data.decode(errors="ignore")
                    if "OK" in response:
                        return True, response
        return False, response
    finally:
        serial_port.timeout = previous_timeout

##
# @brief Open the COM port and check that the modem answers AT with OK (probe_lock held).
#
# @param [in] port COM port, e.g. COM5
#
# @return (True, "") if the port answered, otherwise (False, reason)
##
def _probe_at_port(port):
    try:
        # Open Serial COM port
        with serial.Serial(port, const.BAUD_RATE, timeout=0.2, write_timeout=1) as serial_port:

            # Wait until module responds AT OK
            is_ready, response = wait_for_at_ready(serial_port, PROBE_TIMEOUT_SEC)
            if is_ready:
                return True, ""
        reply = " ".join(response.split())
        detail = f"It replied: {reply[:60]}" if reply else "There was no reply."
        return False, f"{port} did not answer the AT command with OK. {detail}"
    except serial.SerialTimeoutException:
        return False, f"{port} did not accept the AT command (write timed out)."
    except serial.SerialException as e:
        return False, f"Cannot open {port}: {e}"

# Excel helpers (read-only, in the UI process)

##
# @brief Excel sheet found in the application directory.
##
@dataclass
class ExcelInfo:
    path            :   Optional[Path]
    platform_name   :   str = "None"
    branchname      :   str = "None"
    error          :   str = ""
    warning         :   str = ""

##
# @brief List the Excel sheets stored in the application directory (lock files excluded).
#
# @return List of Excel sheet paths
##
def stored_excel_files():
    return [path for path in app_directory().glob("*.xlsx") if not path.name.startswith("~$")]

##
# @brief Find the Excel sheet the pre-test modules will use and read its platform.
#
# @return Excel information (error is set if the sheet cannot be used)
##
def detect_excel():
    # The pre-test modules use the first *.xlsx found in the application directory
    excel_path = next(app_directory().glob("*.xlsx"), None)
    if excel_path is None:
        return ExcelInfo(path=None, error="No Excel sheet found - upload the Jenkins Build Parameter sheet")

    info = ExcelInfo(path=excel_path)
    if excel_path.name.startswith("~$"):
        info.error = "Excel sheet is open in Microsoft Excel (lock file found) - close it first"
        return info
    try:
        info.platform_name = pretest_main.read_jenkins_value(excel_path, "platformName")
        info.branchname = pretest_main.read_jenkins_value(excel_path, "branchname")
    except Exception as e:
        info.error = f"Cannot read Excel: {e}"
    # The module is selected from the JIRA ID, not from the Excel sheet
    return info

##
# @brief Read several Jenkins parameter values from the Excel sheet.
#
# @param [in] excel_path Excel path
# @param [in] keys [(label, parameter name)]
#
# @return {parameter name: value}, "?" if a value cannot be read
##
def read_excel_values(excel_path, keys):
    values = {}
    for _, key in keys:
        try:
            values[key] = pretest_main.read_jenkins_value(excel_path, key)
        except Exception:
            values[key] = "?"
    return values

# JIRA helpers (in a background thread)

##
# @brief JIRA request failed (the message is shown to the user).
##
class JiraError(Exception):
    pass

##
# @brief Send a JIRA REST request with the login username and API token.
#
# @param [in] url Request URL
# @param [in] username Login username
# @param [in] token Atlassian API token
# @param [in] data Request body, or None for a GET request
# @param [in] headers Extra request headers, or None
#
# @return HTTP response (close it after use)
##
def jira_open(url, username, token, data=None, headers=None):
    credentials = base64.b64encode(f"{username}:{token}".encode()).decode()
    request = urllib.request.Request(url, data=data, headers={"Authorization": f"Basic {credentials}",
                                                              "Accept": "application/json", **(headers or {})})
    return urllib.request.urlopen(request, timeout=JIRA_TIMEOUT_SEC)

##
# @brief Post the pre-test result to a JIRA ticket: attach the files, then add
#        a comment that links them.
#
# @param [in] jira_id JIRA ID, e.g. SDX35-112
# @param [in] username Atlassian account (login username)
# @param [in] token Atlassian API token
# @param [in] files Files to attach (report and log)
# @param [in] comment Comment text (JIRA wiki markup, [^name] links an attachment)
#
# @return None
##
def jira_post_result(jira_id, username, token, files, comment):
    issue_url = f"{JIRA_BASE_URL}/rest/api/2/issue/{urllib.parse.quote(jira_id)}"
    for path in files:
        boundary = f"EaglePretest{os.urandom(16).hex()}"
        body = (f"--{boundary}\r\nContent-Disposition: form-data; name=\"file\"; filename=\"{path.name}\"\r\n"
                f"Content-Type: application/octet-stream\r\n\r\n").encode() + path.read_bytes() + \
               f"\r\n--{boundary}--\r\n".encode()
        try:
            jira_open(f"{issue_url}/attachments", username, token, body,
                      {"Content-Type": f"multipart/form-data; boundary={boundary}",
                       "X-Atlassian-Token": "no-check"}).close()
        except urllib.error.HTTPError as e:
            if e.code in (401, 403, 404):
                raise JiraError(f"Cannot attach {path.name} to {jira_id} - {username} has no permission, "
                                f"or the API token is invalid or expired.")
            raise JiraError(f"Cannot attach {path.name} to {jira_id} (HTTP {e.code}).")
        except (urllib.error.URLError, OSError) as e:
            raise JiraError(f"Cannot attach {path.name} to {jira_id}: {getattr(e, 'reason', e)}")

    try:
        jira_open(f"{issue_url}/comment", username, token, json.dumps({"body": comment}).encode(),
                  {"Content-Type": "application/json"}).close()
    except urllib.error.HTTPError as e:
        if e.code in (401, 403, 404):
            raise JiraError(f"Cannot comment on {jira_id} - {username} has no permission, "
                            f"or the API token is invalid or expired.")
        raise JiraError(f"Cannot comment on {jira_id} (HTTP {e.code}).")
    except (urllib.error.URLError, OSError) as e:
        raise JiraError(f"Cannot comment on {jira_id}: {getattr(e, 'reason', e)}")

##
# @brief Check that the API token can log in to JIRA.
#
# @param [in] username Login username
# @param [in] token Atlassian API token
#
# @return ("ok", ""), ("bad", reason) if JIRA rejected the token, or ("error", reason)
##
def jira_check_token(username, token):
    try:
        jira_open(f"{JIRA_BASE_URL}/rest/api/3/myself", username, token).close()
    except urllib.error.HTTPError as e:
        if e.code in (401, 403):
            return "bad", "invalid or expired"
        return "error", f"HTTP {e.code}"
    except (urllib.error.URLError, OSError) as e:
        return "error", str(getattr(e, "reason", e))
    return "ok", ""

##
# @brief Check that the JIRA ticket is a pre-test ticket (topic contains Pretest).
#
# @param [in] jira_id JIRA ID, e.g. SDX35-112
# @param [in] username Login username
# @param [in] token Atlassian API token
#
# @return ("ok", ""), ("bad", reason) if the ticket cannot be used, or ("error", reason)
##
def jira_check_topic(jira_id, username, token):
    url = f"{JIRA_BASE_URL}/rest/api/3/issue/{urllib.parse.quote(jira_id)}?fields=summary"
    try:
        with jira_open(url, username, token) as response:
            summary = json.load(response).get("fields", {}).get("summary") or ""
    except urllib.error.HTTPError as e:
        if e.code in (403, 404):
            return "bad", "JIRA not found"
        return "error", f"cannot check the JIRA topic (HTTP {e.code})"
    except (urllib.error.URLError, OSError, ValueError) as e:
        return "error", f"cannot check the JIRA topic ({getattr(e, 'reason', e)})"
    if not RE_PRETEST.search(summary):
        return "bad", NOT_PRETEST_REASON
    return "ok", ""

# Windows Credential Manager (API token store, one generic credential per login username)

CRED_TYPE_GENERIC                   = 1
CRED_PERSIST_LOCAL_MACHINE          = 2      # Kept for this Windows user on this PC

##
# @brief CREDENTIALW structure of the Windows credential functions.
##
class Credential(ctypes.Structure):
    _fields_ = [("Flags", wintypes.DWORD), ("Type", wintypes.DWORD), ("TargetName", wintypes.LPWSTR),
                ("Comment", wintypes.LPWSTR), ("LastWritten", wintypes.FILETIME),
                ("CredentialBlobSize", wintypes.DWORD), ("CredentialBlob", ctypes.POINTER(ctypes.c_ubyte)),
                ("Persist", wintypes.DWORD), ("AttributeCount", wintypes.DWORD), ("Attributes", ctypes.c_void_p),
                ("TargetAlias", wintypes.LPWSTR), ("UserName", wintypes.LPWSTR)]

##
# @brief Load advapi32.dll with the argument types of the credential functions.
#
# @return advapi32 library
##
def _advapi32():
    advapi = ctypes.WinDLL("advapi32", use_last_error=True)
    advapi.CredReadW.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD,
                                 ctypes.POINTER(ctypes.POINTER(Credential))]
    advapi.CredWriteW.argtypes = [ctypes.POINTER(Credential), wintypes.DWORD]
    advapi.CredFree.argtypes = [ctypes.c_void_p]
    return advapi

##
# @brief Read the API token of a login username from Windows Credential Manager.
#
# @param [in] username Login username
#
# @return Token, or "" if none is stored
##
def load_token(username):
    if os.name != "nt":
        return ""
    advapi = _advapi32()
    stored = ctypes.POINTER(Credential)()
    if not advapi.CredReadW(CREDENTIAL_TARGET.format(username=username), CRED_TYPE_GENERIC, 0,
                            ctypes.byref(stored)):
        return ""
    try:
        credential = stored.contents
        return ctypes.string_at(credential.CredentialBlob, credential.CredentialBlobSize).decode("utf-16-le")
    finally:
        advapi.CredFree(stored)

##
# @brief Store (or replace) the API token of a login username in Windows Credential Manager.
#
# @param [in] username Login username
# @param [in] token Atlassian API token
#
# @return None
##
def save_token(username, token):
    if os.name != "nt":
        return
    blob = token.encode("utf-16-le")
    buffer = (ctypes.c_ubyte * len(blob)).from_buffer_copy(blob)
    credential = Credential(Type=CRED_TYPE_GENERIC, TargetName=CREDENTIAL_TARGET.format(username=username),
                            CredentialBlobSize=len(blob),
                            CredentialBlob=ctypes.cast(buffer, ctypes.POINTER(ctypes.c_ubyte)),
                            Persist=CRED_PERSIST_LOCAL_MACHINE, UserName=username)
    if not _advapi32().CredWriteW(ctypes.byref(credential), 0):
        raise ctypes.WinError(ctypes.get_last_error())

##
# @brief Check the internet connection by reaching the JIRA site (no login needed).
#
# @return True if the JIRA site answered (any HTTP status), otherwise False
##
def jira_reachable():
    try:
        urllib.request.urlopen(JIRA_BASE_URL, timeout=INTERNET_PROBE_TIMEOUT_SEC).close()
    except urllib.error.HTTPError:
        pass    # The site answered, only the page needs a login
    except (urllib.error.URLError, OSError):
        return False
    return True

##
# @brief Download the newest Excel (.xlsx) attachment of a JIRA ticket.
#
# @param [in] jira_id JIRA ID, e.g. SDX35-112
# @param [in] username Atlassian account (login username)
# @param [in] token Atlassian API token
# @param [in] folder Folder the attachment is saved to
#
# @return Path of the downloaded file
##
def jira_download_excel(jira_id, username, token, folder):
    # Check the account first, so an invalid token is not reported as a missing ticket
    try:
        jira_open(f"{JIRA_BASE_URL}/rest/api/3/myself", username, token).close()
    except urllib.error.HTTPError as e:
        if e.code in (401, 403):
            raise JiraError(f"JIRA login failed - the API token is invalid or expired, "
                            f"or it does not belong to {username}.")
        raise JiraError(f"JIRA login failed (HTTP {e.code}).")
    except (urllib.error.URLError, OSError) as e:
        raise JiraError(f"Cannot reach {JIRA_BASE_URL}: {getattr(e, 'reason', e)}")

    url = f"{JIRA_BASE_URL}/rest/api/3/issue/{urllib.parse.quote(jira_id)}?fields=attachment,summary"
    try:
        with jira_open(url, username, token) as response:
            issue = json.load(response)
    except urllib.error.HTTPError as e:
        if e.code in (403, 404):
            raise JiraError("JIRA not found.")
        raise JiraError(f"Cannot open {jira_id} (HTTP {e.code}).")
    except (urllib.error.URLError, OSError, ValueError) as e:
        raise JiraError(f"Cannot open {jira_id}: {getattr(e, 'reason', e)}")
    if not RE_PRETEST.search(issue.get("fields", {}).get("summary") or ""):
        raise JiraError(f"{jira_id} is {NOT_PRETEST_REASON}.")

    sheets = [attachment for attachment in issue.get("fields", {}).get("attachment") or []
              if attachment.get("filename", "").lower().endswith(".xlsx")]
    if not sheets:
        raise JiraError(f"{jira_id} has no Excel (.xlsx) attachment.")
    sheet = max(sheets, key=lambda attachment: attachment.get("created", ""))

    # Saved without the ID JIRA can add to the attachment name
    name = Path(sheet["filename"])
    target = Path(folder) / f"{RE_ATTACHMENT_ID.sub('', name.stem)}{name.suffix}"
    try:
        with jira_open(sheet["content"], username, token) as response, open(target, "wb") as file:
            shutil.copyfileobj(response, file)
    except (urllib.error.URLError, OSError) as e:
        raise JiraError(f"Cannot download {sheet['filename']}: {getattr(e, 'reason', e)}")
    return target

##
# @brief Preview the expected responses a module will use by running its
#        Excel extraction in-process (output discarded, nothing is logged
#        because the UI process never configures the log file).
#
# @param [in] profile Module profile
#
# @return Derived QGMR version, or "" if extraction failed
##
def preview_expected(profile):
    try:
        with contextlib.redirect_stdout(io.StringIO()):
            if not profile.service.excel_extract_parameter():
                return ""
        return at_cmd_general_info_list[0].version_qgmr
    except Exception:
        return ""

# Output parser

RE_INDEX                            = re.compile(r"^\[(\d+)/(\d+)\]$")
RE_TEST_CASE                        = re.compile(r"^Test Case : (.+) \[(\d+)\]$")
RE_SEND                             = re.compile(r"^Send      : (.*)$")
RE_RECEIVE                          = re.compile(r"^Receive   : (.*)$")
RE_EXPECTED                         = re.compile(r"^Expected  : (.*)$")
RE_RESULT                           = re.compile(r"^Result    : .* : (PASS|FAIL|UNSUPPORTED)$")
RE_TOTAL                            = re.compile(r"^Total : (\d+) \| Pass : (\d+) \| Fail : (\d+)")

##
# @brief Command, response and result of one test case of a run.
##
@dataclass
class TestRecord:
    index       :   int
    send        :   str = ""
    receive     :   str = ""
    expected    :   str = ""
    result      :   str = ""

# Theme

COLORS = {
    "bg": "#f3f4f6",
    "card": "#ffffff",
    "border": "#e2e4e9",
    "text": "#1f2937",
    "muted": "#6b7280",
    "pass": "#15803d",
    "pass_bg": "#e3f4e8",
    "fail": "#b91c1c",
    "fail_bg": "#fbe4e4",
    "skip": "#6b7280",
    "skip_bg": "#eeeff2",
    "warn": "#b45309",
    "online": "#14532d",            # Online / Offline button (dark, so it is not too bright)
    "offline": "#7f1d1d",
    "default_accent": "#003e71",     # Eagle Wireless logo navy
    "console_bg": "#111827",
    "console_fg": "#d1d5db",
}

COLUMN_LINE_COLOR                   = "#d9dce2"

FONT                                = ("Segoe UI", 10)
FONT_BOLD                           = ("Segoe UI Semibold", 10)
FONT_SMALL                          = ("Segoe UI", 9)
FONT_TITLE                          = ("Segoe UI Semibold", 15)
FONT_BIG                            = ("Segoe UI Semibold", 13)
FONT_MONO                           = ("Consolas", 10)

# Pixel sizes below are designed for 100 % display scaling (96 DPI) and multiplied
# by this factor, so the layout keeps its proportions at 125 % / 150 % scaling
UI_SCALE                            = 1.0

##
# @brief Scale a pixel size designed for 100 % display scaling to the current scaling.
#
# @param [in] value Pixel size at 100 % scaling
#
# @return Scaled pixel size
##
def px(value):
    return round(value * UI_SCALE)

##
# @brief Tell Windows the GUI draws itself at the real screen resolution.
#
# Without this, Windows renders the window at 100 % and stretches the bitmap
# on screens with display scaling above 100 %, which makes text look blurry.
#
# @return None
##
def enable_dpi_awareness():
    if os.name != "nt":
        return
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)      # System DPI aware (Windows 8.1+)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass

# Report (PDF)

REPORT_RESULT_TEXT                  = {"PASS": "PASS", "FAIL": "FAIL", "UNSUPPORTED": "SKIP", "N/A": "N/A", "": "Not run"}

##
# @brief Write the pre-test report: run information and the Test Cases table
#        with the actual modem response of every command.
#
# @param [in] path PDF file to write
# @param [in] jira_id JIRA ID shown on the right of the title
# @param [in] left [(label, value)] run information, left column
# @param [in] right [(label, value)] run information, right column (below the JIRA link)
# @param [in] rows [(no, test case, step, AT command, expected, response, result)]
#
# @return None
##
def write_report_pdf(path, jira_id, left, right, rows):
    from xml.sax.saxutils import escape
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.platypus import Image, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

    accent = colors.HexColor(COLORS["default_accent"])
    text_style = ParagraphStyle("text", fontName="Helvetica", fontSize=7.5, leading=9.5)
    mono_style = ParagraphStyle("mono", parent=text_style, fontName="Courier", fontSize=7, leading=9)
    head_style = ParagraphStyle("head", parent=text_style, fontName="Helvetica-Bold", textColor=colors.white)
    key_style = ParagraphStyle("key", parent=text_style, fontName="Helvetica-Bold", fontSize=8.5, leading=11,
                               textColor=colors.HexColor(COLORS["muted"]))
    value_style = ParagraphStyle("value", parent=text_style, fontSize=8.5, leading=11)
    title_style = ParagraphStyle("title", fontName="Helvetica-Bold", fontSize=16, leading=20, textColor=accent)

    ##
    # @brief Table cell text: one line per non-empty line of the value.
    #
    # @param [in] value Cell value
    # @param [in] style Paragraph style
    #
    # @return Paragraph of the cell
    ##
    def cell(value, style=text_style):
        lines = [escape(line.strip()) for line in str(value).splitlines() if line.strip()]
        return Paragraph("<br/>".join(lines) or "-", style)

    ##
    # @brief Page footer: tool name and version on the left, page number on the right.
    #
    # @param [in] canvas PDF canvas
    # @param [in] doc PDF document
    #
    # @return None
    ##
    def footer(canvas, doc):
        canvas.saveState()
        canvas.setFont("Helvetica", 7.5)
        canvas.setFillColor(colors.HexColor(COLORS["muted"]))
        canvas.drawString(doc.leftMargin, 7 * mm, f"{APP_TITLE} v{APP_VERSION} - Anas Akasyah")
        canvas.drawRightString(doc.pagesize[0] - doc.rightMargin, 7 * mm, f"Page {doc.page}")
        canvas.restoreState()

    story = []
    logo = resource_path(LOGO_FILE)
    if logo is not None:
        # A reduced copy of the logo keeps the PDF small
        from PIL import Image as PILImage
        with PILImage.open(logo) as image:
            ratio = image.width / image.height
            small = image.convert("RGB")
            small.thumbnail((round(240 * ratio), 240))
        data = io.BytesIO()
        small.save(data, "JPEG", quality=90)
        data.seek(0)
        story.append(Image(data, width=16 * mm * ratio, height=16 * mm, hAlign="LEFT"))
        story.append(Spacer(1, 3 * mm))
    # Title on the left, JIRA ID on the right
    jira_style = ParagraphStyle("jira", parent=title_style, alignment=2)
    title = Table([[Paragraph(f"{escape(APP_TITLE)} - AT Command Pre-Test Report", title_style),
                    Paragraph(escape(jira_id), jira_style)]],
                  colWidths=[180 * mm, 93 * mm])
    title.setStyle(TableStyle([("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                               ("TOPPADDING", (0, 0), (-1, -1), 0), ("BOTTOMPADDING", (0, 0), (-1, -1), 0)]))
    story.append(title)
    story.append(Spacer(1, 4 * mm))

    # Run information, two label / value pairs per row
    # Run information in two columns; the JIRA link is clickable but kept black and bold
    url = escape(JIRA_BROWSE_URL + jira_id)
    link = Paragraph(f'<a href="{url}"><b>{url}</b></a>', value_style)
    left_pairs = [(Paragraph(escape(label), key_style), cell(value, value_style)) for label, value in left]
    right_pairs = [(Paragraph("JIRA Link", key_style), link)]
    right_pairs += [(Paragraph(escape(label), key_style), cell(value, value_style)) for label, value in right]
    info_table = Table([list(l) + list(r) for l, r in itertools.zip_longest(left_pairs, right_pairs,
                                                                            fillvalue=("", ""))],
                       colWidths=[28 * mm, 108 * mm, 30 * mm, 107 * mm])
    info_table.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"),
                                    ("LINEBELOW", (0, 0), (-1, -1), 0.25, colors.HexColor(COLORS["border"])),
                                    ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 3)]))
    story.append(info_table)
    story.append(Spacer(1, 6 * mm))

    # Test Cases table (same columns as the GUI, plus the actual response)
    headers = ("#", "Test Case", "Step", "AT Command", "Expected", "Actual Response", "Result")
    data = [[Paragraph(text, head_style) for text in headers]]
    style = [("BACKGROUND", (0, 0), (-1, 0), accent),
             ("VALIGN", (0, 0), (-1, -1), "TOP"),
             ("ALIGN", (0, 0), (0, -1), "CENTER"), ("ALIGN", (2, 0), (2, -1), "CENTER"),
             ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor(COLUMN_LINE_COLOR)),
             ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 3)]
    result_colors = {"PASS": ("pass", "pass_bg"), "FAIL": ("fail", "fail_bg"), "UNSUPPORTED": ("skip", "skip_bg")}
    for row, (number, case, step, command, expected, response, result) in enumerate(rows, start=1):
        result_style = text_style
        if result in result_colors:
            fg, bg = result_colors[result]
            result_style = ParagraphStyle(f"result_{result}", parent=text_style, fontName="Helvetica-Bold",
                                          textColor=colors.HexColor(COLORS[fg]))
            style.append(("BACKGROUND", (6, row), (6, row), colors.HexColor(COLORS[bg])))
        data.append([cell(number), cell(case), cell(step), cell(command, mono_style), cell(expected, mono_style),
                     cell(response, mono_style), Paragraph(REPORT_RESULT_TEXT.get(result, result), result_style)])
    table = Table(data, colWidths=[8 * mm, 40 * mm, 10 * mm, 45 * mm, 60 * mm, 92 * mm, 18 * mm], repeatRows=1)
    table.setStyle(TableStyle(style))
    story.append(table)

    doc = SimpleDocTemplate(str(path), pagesize=landscape(A4), leftMargin=12 * mm, rightMargin=12 * mm,
                            topMargin=12 * mm, bottomMargin=14 * mm, title=f"{APP_TITLE} - AT Command Pre-Test Report",
                            author=APP_TITLE)
    doc.build(story, onFirstPage=footer, onLaterPages=footer)

# Application

TOOLTIP_DELAY_MS                    = 1000

SIM_FIELD_HELP = {
    "SIM CCID": ("Integrated Circuit Card Identifier",
                 "The SIM card's unique serial number (20 characters, may end in a letter, e.g. ...556F). "
                 "Printed on the SIM card; read it from the modem with AT+QCCID (or AT+CCID)."),
    "SIM IMEI": ("International Mobile Equipment Identity",
                 "The 15-digit number that identifies the module itself (the device, not the SIM). "
                 "Read it from the modem with AT+GSN."),
    "SIM IMSI": ("International Mobile Subscriber Identity",
                 "The 15-digit subscriber number stored on the SIM, identifying the operator and the "
                 "subscription. Read it from the modem with AT+CIMI."),
}

##
# @brief Tooltip shown when the cursor rests on a widget for TOOLTIP_DELAY_MS.
##
class Tooltip:

    ##
    # @brief Attach the tooltip to a widget.
    #
    # @param [in] widget Widget the tooltip belongs to
    # @param [in] title Tooltip title
    # @param [in] text Tooltip text
    ##
    def __init__(self, widget, title, text):
        self.widget, self.title, self.text = widget, title, text
        self.job = None
        self.window = None
        widget.bind("<Enter>", self._schedule, add="+")
        widget.bind("<Leave>", self._hide, add="+")
        widget.bind("<ButtonPress>", self._hide, add="+")

    ##
    # @brief Mouse entered the widget: show the tooltip after TOOLTIP_DELAY_MS.
    #
    # @param [in] _event Tk event (not used)
    #
    # @return None
    ##
    def _schedule(self, _event=None):
        self._hide()
        self.job = self.widget.after(TOOLTIP_DELAY_MS, self._show)

    ##
    # @brief Show the tooltip next to the mouse pointer.
    #
    # @return None
    ##
    def _show(self):
        self.job = None
        x = self.widget.winfo_pointerx() + px(14)
        y = self.widget.winfo_pointery() + px(18)
        self.window = tk.Toplevel(self.widget)
        self.window.wm_overrideredirect(True)
        self.window.attributes("-topmost", True)
        frame = tk.Frame(self.window, bg="#1f2937", padx=px(10), pady=px(7))
        frame.pack()
        tk.Label(frame, text=self.title, font=FONT_BOLD, fg="white", bg="#1f2937",
                 justify="left").pack(anchor="w")
        tk.Label(frame, text=self.text, font=FONT_SMALL, fg="#d1d5db", bg="#1f2937",
                 wraplength=px(300), justify="left").pack(anchor="w", pady=(2, 0))
        self.window.wm_geometry(f"+{x}+{y}")

    ##
    # @brief Hide the tooltip, or cancel one that is about to be shown.
    #
    # @param [in] _event Tk event (not used)
    #
    # @return None
    ##
    def _hide(self, _event=None):
        if self.job is not None:
            self.widget.after_cancel(self.job)
            self.job = None
        if self.window is not None:
            self.window.destroy()
            self.window = None

##
# @brief Hash a password with the login salt (PBKDF2-SHA256), the same way as
#        LOGIN_PASSWORD_HASH was made.
#
# @param [in] password Password
#
# @return Password hash (hex)
##
def login_password_hash(password):
    password_hash = hashlib.pbkdf2_hmac("sha256", password.encode(), bytes.fromhex(LOGIN_PASSWORD_SALT),
                                        LOGIN_PASSWORD_ITERATIONS)
    return password_hash.hex()

##
# @brief Login page shown before the pre-test tool. The tool only opens once
#        the username and password are entered correctly.
##
class LoginWindow(tk.Tk):

    ##
    # @brief Create the login page.
    ##
    def __init__(self):
        super().__init__()
        global UI_SCALE
        UI_SCALE = self.winfo_fpixels("1i") / 96
        self.title(f"{APP_TITLE} v{APP_VERSION}")
        # Same window size as the pre-test tool page
        PretestApp._fit_window(self, px(1320), px(840), px(1100), px(700))
        self.configure(bg=COLORS["bg"])

        self.is_authenticated = False
        self.username = ""
        self.username_var = tk.StringVar()
        self.password_var = tk.StringVar()

        self._init_style()
        self._build_layout()
        self.username_entry.focus_set()

    ##
    # @brief Set the ttk styles of the login page.
    #
    # @return None
    ##
    def _init_style(self):
        accent = COLORS["default_accent"]
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure(".", font=FONT, background=COLORS["bg"], foreground=COLORS["text"])
        style.configure("TEntry", padding=6)
        style.configure("Accent.TButton", background=accent, foreground="white", font=FONT_BOLD,
                        bordercolor=accent, lightcolor=accent, darkcolor=accent, padding=(14, 7))
        style.map("Accent.TButton", background=[("active", accent)])

    ##
    # @brief Build the login page: logo, username, password and Log In button.
    #
    # @return None
    ##
    def _build_layout(self):
        # Login card centered in the window
        outer = tk.Frame(self, bg=COLORS["card"], highlightthickness=1,
                         highlightbackground=COLORS["border"], highlightcolor=COLORS["border"])
        outer.place(relx=0.5, rely=0.5, anchor="center")
        # Author information at the bottom right
        author = tk.Frame(self, bg=COLORS["bg"])
        author.place(relx=1, rely=1, anchor="se", x=-px(16), y=-px(12))
        for row, (key, value) in enumerate((("Author", "Anas Akasyah"), ("Date", "21-10-2026"))):
            for column, text in enumerate((key, ":", value)):
                tk.Label(author, text=text, font=FONT_SMALL, fg=COLORS["muted"],
                         bg=COLORS["bg"]).grid(row=row, column=column, sticky="w", padx=(0, 4))
        card = tk.Frame(outer, bg=COLORS["card"], padx=px(32), pady=px(24))
        card.pack()

        # Logo: the eagle emblem (also used as the window icon)
        emblem = load_logo_emblem()
        if emblem is not None:
            from PIL import Image, ImageTk
            # Keep references, Tk does not hold on to PhotoImage objects
            self.icon_images = [ImageTk.PhotoImage(emblem.resize((size, size), Image.LANCZOS))
                                for size in ICON_SIZES]
            self.iconphoto(True, *self.icon_images)
            self.logo_image = ImageTk.PhotoImage(emblem.resize((px(120), px(120)), Image.LANCZOS))
            tk.Label(card, image=self.logo_image, bg=COLORS["card"]).pack()

        tk.Label(card, text="Eagle Wireless", font=FONT_TITLE, fg=COLORS["default_accent"],
                 bg=COLORS["card"]).pack(pady=(px(6), 0))
        tk.Label(card, text="Pre-Test", font=FONT_BIG, fg=COLORS["muted"], bg=COLORS["card"]).pack(pady=(0, px(18)))

        form = tk.Frame(card, bg=COLORS["card"])
        form.pack(fill="x")
        tk.Label(form, text="Username", font=FONT_BOLD, bg=COLORS["card"]).pack(anchor="w")
        self.username_entry = ttk.Entry(form, textvariable=self.username_var, width=32)
        self.username_entry.pack(fill="x", pady=(2, 0))
        tk.Label(form, text=f"e.g. name{LOGIN_DOMAIN}", font=FONT_SMALL, fg=COLORS["muted"],
                 bg=COLORS["card"]).pack(anchor="w", pady=(0, 8))
        tk.Label(form, text="Password", font=FONT_BOLD, bg=COLORS["card"]).pack(anchor="w")
        self.password_entry = ttk.Entry(form, textvariable=self.password_var, show="•", width=32)
        self.password_entry.pack(fill="x", pady=(2, 4))
        self.error_label = tk.Label(form, text="", font=FONT_SMALL, fg=COLORS["fail"], bg=COLORS["card"])
        self.error_label.pack(anchor="w", pady=(0, 8))

        ttk.Button(form, text="Login", style="Accent.TButton", command=self.login).pack(fill="x")
        tk.Label(card, text=f"v{APP_VERSION}", font=FONT_SMALL, fg=COLORS["muted"],
                 bg=COLORS["card"]).pack(pady=(px(14), 0))

        self.username_entry.bind("<Return>", lambda e: self.password_entry.focus_set())
        self.password_entry.bind("<Return>", lambda e: self.login())
        self.username_var.trace_add("write", lambda *_: self.error_label.configure(text=""))
        self.password_var.trace_add("write", lambda *_: self.error_label.configure(text=""))

    ##
    # @brief Log In: check the username (same rule as main.py) and the password, then close the login page so
    #        the pre-test tool opens.
    #
    # @return None
    ##
    def login(self):
        # Same username check as main.py
        username = self.username_var.get().strip().lower()
        if not username.endswith(LOGIN_DOMAIN) or username == LOGIN_DOMAIN:
            self.error_label.configure(text=f"✗ Username must end with {LOGIN_DOMAIN}")
            self.username_entry.focus_set()
            return
        if not hmac.compare_digest(login_password_hash(self.password_var.get()), LOGIN_PASSWORD_HASH):
            self.password_var.set("")
            self.error_label.configure(text="✗ Invalid password")
            self.password_entry.focus_set()
            return
        self.is_authenticated = True
        self.username = username
        self.destroy()

##
# @brief Main window of the pre-test tool.
##
class PretestApp(tk.Tk):

    ##
    # @brief Create the main window of the pre-test tool.
    #
    # @param [in] username Login username (Reporter of the runs)
    ##
    def __init__(self, username):
        super().__init__()
        self.username = username
        global UI_SCALE
        UI_SCALE = self.winfo_fpixels("1i") / 96
        self.title(f"{APP_TITLE} v{APP_VERSION}")
        self._set_window_icon()
        self._fit_window(px(1320), px(840), px(1100), px(700))
        self.configure(bg=COLORS["bg"])

        self.process = None
        self.output_queue = queue.Queue()
        self.is_running = False
        self.is_aborted = False
        self.excel = ExcelInfo(path=None)
        self.profile = None
        self.records = {}
        self.current = None
        self.is_in_receive = False
        self.is_summary_seen = False
        self.error_lines = []
        self.log_file_path = None
        self.start_time = None
        self.run_started_at = 0.0
        self.run_info = {}
        self.run_token = ""                 # API token of the run, for posting the result (memory only)
        self.is_checking_port = False
        self.port_state = ("", "none", "")     # (port, none | checking | ok | bad, reason)
        self.port_check_id = 0
        self.port_check_job = None
        self.token_state = ("", "none", "")    # (token, none | checking | ok | bad | error, message)
        self.token_check_id = 0
        self.token_check_job = None
        self.is_token_from_store = False           # Token box filled from Windows Credential Manager
        self.topic_state = ("", "none", "")    # (JIRA ID, none | checking | ok | bad | error, reason)
        self.topic_check_id = 0
        self.topic_check_job = None
        self.reset_deadline = None
        self.counts = {"PASS": 0, "FAIL": 0, "UNSUPPORTED": 0}

        self.module_var = tk.StringVar(value="SDX35")
        self.jira_var = tk.StringVar()
        self.token_var = tk.StringVar()     # JIRA API token, kept in memory only
        self.jira_excel = None              # (JIRA ID, path) of the sheet downloaded in online mode
        self.is_downloading = False
        self.is_online_pending_shown = False
        self.is_internet_ok = False            # Online mode: the JIRA site was reached
        self.is_internet_checking = False
        self.internet_check_id = 0
        self.internet_deadline = 0.0
        self.jira_id = ""
        self.port_var = tk.StringVar()
        self.ccid_var = tk.StringVar()
        self.imei_var = tk.StringVar()
        self.imsi_var = tk.StringVar()
        self.autoscroll_var = tk.BooleanVar(value=True)

        self._init_style()
        self._build_layout()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

        self._clear_previous_session()
        self.refresh_ports()
        self.reload_excel()
        self.after(50, self._drain_output)
        self.after(500, self._tick)

    # Window icon

    ##
    # @brief Size the window to fit the screen area above the taskbar.
    #
    # At high display scaling (e.g. 150 % on a laptop) the full layout is larger
    # than the screen, so the window opens maximized instead of being cut off.
    #
    # @param [in] width Window width
    # @param [in] height Window height
    # @param [in] min_width Minimum window width
    # @param [in] min_height Minimum window height
    #
    # @return None
    ##
    def _fit_window(self, width, height, min_width, min_height):
        left, top, right, bottom = 0, 0, self.winfo_screenwidth(), self.winfo_screenheight()
        if os.name == "nt":
            try:
                rect = wintypes.RECT()
                if ctypes.windll.user32.SystemParametersInfoW(0x0030, 0, ctypes.byref(rect), 0):  # SPI_GETWORKAREA
                    left, top, right, bottom = rect.left, rect.top, rect.right, rect.bottom
            except Exception:
                pass
        frame = px(40)      # Title bar and window borders
        area_w, area_h = right - left, bottom - top - frame
        self.minsize(min(min_width, area_w), min(min_height, area_h))
        if width > area_w or height > area_h:
            self.geometry(f"{min(width, area_w)}x{min(height, area_h)}+{left}+{top}")
            self.state("zoomed")
        else:
            self.geometry(f"{width}x{height}+{left + (area_w - width) // 2}+{top + (area_h - height) // 2}")

    ##
    # @brief Set the window icon to the eagle emblem of the Eagle Wireless logo.
    #
    # @return None
    ##
    def _set_window_icon(self):
        emblem = load_logo_emblem()
        if emblem is None:
            return  # Keep the default Tk icon
        from PIL import Image, ImageTk
        # Keep references, Tk does not hold on to PhotoImage objects
        self.icon_images = [ImageTk.PhotoImage(emblem.resize((size, size), Image.LANCZOS))
                            for size in ICON_SIZES]
        self.iconphoto(True, *self.icon_images)

    # Style

    ##
    # @brief Set the ttk styles of the main window.
    #
    # @return None
    ##
    def _init_style(self):
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure(".", font=FONT, background=COLORS["bg"], foreground=COLORS["text"])
        style.configure("Card.TFrame", background=COLORS["card"])
        style.configure("Card.TLabel", background=COLORS["card"])
        style.configure("Muted.TLabel", background=COLORS["card"], foreground=COLORS["muted"], font=FONT_SMALL)
        style.configure("Key.TLabel", background=COLORS["card"], foreground=COLORS["muted"])
        style.configure("Value.TLabel", background=COLORS["card"], font=FONT_BOLD)
        style.configure("Section.TLabel", background=COLORS["card"], font=FONT_BOLD)
        style.configure("Big.TLabel", background=COLORS["card"], font=FONT_BIG)
        style.configure("Card.TCheckbutton", background=COLORS["card"])
        style.configure("TEntry", padding=4)
        style.configure("TCombobox", padding=4)
        style.configure("TButton", padding=(10, 5))
        style.configure("Tool.TButton", padding=(2, 3))
        style.configure("Treeview", rowheight=px(24), font=FONT_SMALL, background=COLORS["card"],
                        fieldbackground=COLORS["card"], bordercolor=COLORS["border"])
        style.configure("Treeview.Heading", font=FONT_BOLD, background="#eef0f3", relief="flat")
        style.map("Treeview", background=[("selected", "#c7d2fe")], foreground=[("selected", COLORS["text"])])
        style.configure("Run.Horizontal.TProgressbar", thickness=px(14), troughcolor="#e5e7eb", borderwidth=0)
        style.configure("Reset.Horizontal.TProgressbar", thickness=px(14), troughcolor="#e5e7eb",
                        borderwidth=0, background=COLORS["warn"], lightcolor=COLORS["warn"], darkcolor=COLORS["warn"])
        self.style = style

    ##
    # @brief Apply the accent color of the selected module (default: Eagle Wireless navy).
    #
    # @return None
    ##
    def _apply_accent(self):
        accent = self.profile.accent if self.profile else COLORS["default_accent"]
        style = self.style
        style.configure("Accent.TButton", background=accent, foreground="white", font=FONT_BOLD,
                        bordercolor=accent, lightcolor=accent, darkcolor=accent, padding=(14, 7))
        style.map("Accent.TButton",
                  background=[("disabled", "#c4c8cf"), ("active", accent)],
                  foreground=[("disabled", "#f3f4f6")])
        style.configure("Run.Horizontal.TProgressbar", background=accent, lightcolor=accent, darkcolor=accent)
        self.header.configure(bg=accent)
        for widget in self.header_widgets:
            widget.configure(bg=accent)
        self.tech_badge.configure(fg=accent)
        for key, button in self.module_buttons.items():
            selected = self.profile is not None and key == self.profile.key
            button.configure(
                bg=accent if selected else COLORS["card"],
                fg="white" if selected else "#9ca3af",
                highlightbackground=accent if selected else COLORS["border"],
            )

    # Layout

    ##
    # @brief Add a card (white frame with a border) with an optional title.
    #
    # @param [in] parent Parent widget
    # @param [in] title Card title, or None
    # @param [in] pack Pack options of the card
    #
    # @return Content frame of the card
    ##
    def _card(self, parent, title=None, **pack):
        outer = tk.Frame(parent, bg=COLORS["card"], highlightthickness=1,
                         highlightbackground=COLORS["border"], highlightcolor=COLORS["border"])
        outer.pack(**pack)
        inner = ttk.Frame(outer, style="Card.TFrame", padding=(14, 10))
        inner.pack(fill="both", expand=True)
        if title:
            ttk.Label(inner, text=title, style="Section.TLabel").pack(anchor="w", pady=(0, 6))
        content = ttk.Frame(inner, style="Card.TFrame")
        content.pack(fill="both", expand=True)
        return content

    ##
    # @brief Build the main window: header, inputs on the left, progress, Test Cases table and live log on
    #        the right.
    #
    # @return None
    ##
    def _build_layout(self):
        # Header
        self.header = tk.Frame(self, height=64)
        self.header.pack(fill="x")
        title = tk.Label(self.header, text="Eagle Wireless  ·  EaglePretest",
                         font=FONT_TITLE, fg="white")
        title.pack(side="left", padx=18, pady=14)
        self.tech_badge = tk.Label(self.header, text="LTE", font=FONT_BOLD, bg="white", padx=10, pady=2)
        self.tech_badge.pack(side="right", padx=(6, 18))
        self.header_module = tk.Label(self.header, text="", font=FONT_BIG, fg="white")
        self.header_module.pack(side="right")
        self.header_widgets = [title, self.header_module]

        # Module selector
        selector = tk.Frame(self, bg=COLORS["bg"])
        selector.pack(fill="x", padx=16, pady=(12, 0))
        tk.Label(selector, text="Module", font=FONT_BOLD, bg=COLORS["bg"]).pack(side="left", padx=(0, 10))
        self.module_buttons = {}
        for key in MODULE_PROFILES:
            button = tk.Label(selector, text=f"  {key}  ", font=FONT_BOLD, bd=0,
                              highlightthickness=1, padx=14, pady=5)
            button.pack(side="left", padx=(0, 6))
            self.module_buttons[key] = button
        self.detect_label = tk.Label(selector, text="", font=FONT_SMALL, bg=COLORS["bg"], fg=COLORS["muted"])
        self.detect_label.pack(side="left", padx=12)
        self.upload_button = ttk.Button(selector, text="⬆  Upload Excel", command=self.upload_excel)
        self.upload_button.pack(side="right")
        self.reset_button = ttk.Button(selector, text="↺  Reset View", command=self.reset_view)
        self.reset_button.pack(side="right", padx=(0, 6))
        self._build_search_bar(selector)

        body = tk.Frame(self, bg=COLORS["bg"])
        body.pack(fill="both", expand=True, padx=16, pady=12)

        # Left column
        left = tk.Frame(body, bg=COLORS["bg"], width=px(360))
        left.pack(side="left", fill="y", padx=(0, 12))
        left.pack_propagate(False)

        actions = tk.Frame(left, bg=COLORS["bg"])
        actions.pack(fill="x", side="bottom")
        self.start_button = ttk.Button(actions, text="▶  Start Pre-Test", style="Accent.TButton", command=self.start_test)
        self.start_button.pack(side="left", fill="x", expand=True)
        self.stop_button = ttk.Button(actions, text="■  Stop", command=self.stop_test, state="disabled")
        self.stop_button.pack(side="left", padx=(8, 0))

        precheck = tk.Frame(left, bg=COLORS["bg"])
        precheck.pack(fill="x", side="bottom", pady=(0, 8))
        tk.Label(precheck, text="Pre-check", font=FONT_BOLD, bg=COLORS["bg"]).pack(anchor="w", pady=(0, 4))
        chips = tk.Frame(precheck, bg=COLORS["bg"])
        chips.pack(fill="x")
        self.precheck_chips = {}
        for column, name in enumerate(("JIRA", "AT port", "Excel", "CCID", "IMEI", "IMSI")):
            chips.columnconfigure(column, weight=1, uniform="precheck")
            chip = tk.Label(chips, text=name, font=FONT_SMALL, padx=2, pady=3)
            chip.grid(row=0, column=column, sticky="ew", padx=(0 if column == 0 else 4, 0))
            self.precheck_chips[name] = chip

        self._build_jira_card(left)
        self._build_connection_card(left)
        tabs = ttk.Notebook(left)
        tabs.pack(fill="both", expand=True, pady=(0, 10))
        self._build_excel_card(tabs)
        self._build_profile_card(tabs)

        # Right column
        right = tk.Frame(body, bg=COLORS["bg"])
        right.pack(side="left", fill="both", expand=True)
        self._build_progress_card(right)

        panes = ttk.PanedWindow(right, orient="vertical")
        panes.pack(fill="both", expand=True, pady=(12, 0))
        panes.add(self._build_table(panes), weight=3)
        panes.add(self._build_console(panes), weight=2)

    ##
    # @brief Search bar beside Reset View. Hidden until a pre-test has finished.
    #
    # @param [in] parent Parent widget
    #
    # @return None
    ##
    def _build_search_bar(self, parent):
        self.search_var = tk.StringVar()
        self.search_hits = []
        self.search_index = -1
        self.search_job = None
        self.search_frame = tk.Frame(parent, bg=COLORS["bg"])
        tk.Label(self.search_frame, text="Search", font=FONT_BOLD, bg=COLORS["bg"]).pack(side="left", padx=(0, 6))
        self.search_entry = ttk.Entry(self.search_frame, textvariable=self.search_var, width=22)
        self.search_entry.pack(side="left")
        for text, command in (("▲", lambda: self._search_step(-1)), ("▼", lambda: self._search_step(1)),
                              ("✕", self.clear_search)):
            ttk.Button(self.search_frame, text=text, width=2, style="Tool.TButton", command=command).pack(side="left", padx=(4, 0))
        self.search_count = tk.Label(self.search_frame, text="", font=FONT_SMALL, bg=COLORS["bg"],
                                     fg=COLORS["muted"], anchor="w")
        self.search_count.pack(side="left", padx=(8, 0))

        self.search_var.trace_add("write", lambda *_: self._schedule_search())
        self.search_entry.bind("<Return>", lambda e: self._search_step(1))
        self.search_entry.bind("<Shift-Return>", lambda e: self._search_step(-1))
        self.search_entry.bind("<Escape>", lambda e: self.clear_search())

    ##
    # @brief Show the search bar (after a pre-test has finished).
    #
    # @return None
    ##
    def _show_search(self):
        self.search_frame.pack(side="right", padx=(0, 10), after=self.reset_button)

    ##
    # @brief Hide the search bar and clear the search.
    #
    # @return None
    ##
    def _hide_search(self):
        self.clear_search()
        self.search_frame.pack_forget()

    ##
    # @brief Clear the search text and show every row again.
    #
    # @return None
    ##
    def clear_search(self):
        self.search_var.set("")
        self._apply_search()

    ##
    # @brief Search after a short pause in typing.
    #
    # @return None
    ##
    def _schedule_search(self):
        if self.search_job is not None:
            self.after_cancel(self.search_job)
        self.search_job = self.after(200, self._apply_search)

    ##
    # @brief Filter the test case table to the matching rows and highlight
    #        every match in the live log (case-insensitive).
    #
    # @return None
    ##
    def _apply_search(self):
        self.search_job = None
        query = self.search_var.get().strip().lower()

        # Table: matching rows only, original order kept
        visible = 0
        for number in sorted(self.records):
            iid = str(number)
            if not self.tree.exists(iid):
                continue
            record = self.records[number]
            text = " ".join(str(value) for value in self.tree.item(iid, "values")) + " " + record.receive
            if not query or query in text.lower():
                self.tree.move(iid, "", visible)
                visible += 1
            else:
                self.tree.detach(iid)

        # Live log: highlight every occurrence
        self.console.tag_remove("search", "1.0", "end")
        self.console.tag_remove("search_current", "1.0", "end")
        self.search_hits, self.search_index = [], -1
        if query:
            length = tk.IntVar()
            position = "1.0"
            while True:
                position = self.console.search(query, position, "end", nocase=True, count=length)
                if not position:
                    break
                end = f"{position}+{length.get()}c"
                self.console.tag_add("search", position, end)
                self.search_hits.append((position, end))
                position = end

        if not query:
            self.search_count.configure(text="", fg=COLORS["muted"])
        elif not visible and not self.search_hits:
            self.search_count.configure(text="No matches", fg=COLORS["fail"])
        else:
            self._update_search_count()
            if self.search_hits:
                self._search_step(1)

    ##
    # @brief Jump to the next or previous search match in the live log.
    #
    # @param [in] direction 1 for the next match, -1 for the previous match
    #
    # @return None
    ##
    def _search_step(self, direction):
        if not self.search_hits:
            return
        self.console.tag_remove("search_current", "1.0", "end")
        self.search_index = (self.search_index + direction) % len(self.search_hits)
        start, end = self.search_hits[self.search_index]
        self.console.tag_add("search_current", start, end)
        self.console.see(start)
        self._update_search_count()

    ##
    # @brief Show the position of the current search match, e.g. 2/5.
    #
    # @return None
    ##
    def _update_search_count(self):
        # Position of the current live log match, e.g. "2/15"
        self.search_count.configure(text=f"{self.search_index + 1}/{len(self.search_hits)}", fg=COLORS["muted"])

    ##
    # @brief Build the Connection & SIM card: COM port, SIM CCID / IMEI / IMSI.
    #
    # @param [in] parent Parent widget
    #
    # @return None
    ##
    def _build_connection_card(self, parent):
        card = self._card(parent, "Connection & SIM", fill="x", pady=(0, 10))
        card.columnconfigure(1, weight=1)

        ttk.Label(card, text="COM Port", style="Key.TLabel").grid(row=1, column=0, sticky="w", pady=3)
        port_row = ttk.Frame(card, style="Card.TFrame")
        port_row.grid(row=1, column=1, columnspan=2, sticky="ew", pady=3)
        self.port_combo = ttk.Combobox(port_row, textvariable=self.port_var, width=18)
        self.port_combo.pack(side="left", fill="x", expand=True)
        ttk.Button(port_row, text="⟳", width=3, command=self.refresh_ports).pack(side="left", padx=(6, 0))
        self.port_desc = ttk.Label(card, text="", style="Muted.TLabel", wraplength=px(250))
        self.port_desc.grid(row=2, column=1, columnspan=2, sticky="w")
        self.port_var.trace_add("write", lambda *_: self._on_port_changed())

        self.length_labels = {}
        fields = (("SIM CCID", self.ccid_var, const.SIM_CCID_LEN),
                  ("SIM IMEI", self.imei_var, const.SIM_IMEI_LEN),
                  ("SIM IMSI", self.imsi_var, const.SIM_IMSI_LEN))
        for row, (label, var, length) in enumerate(fields, start=3):
            name = ttk.Label(card, text=label, style="Key.TLabel", cursor="question_arrow")
            name.grid(row=row, column=0, sticky="w", pady=3, padx=(0, 8))
            Tooltip(name, *SIM_FIELD_HELP[label])
            entry = ttk.Entry(card, textvariable=var, font=FONT_MONO)
            entry.grid(row=row, column=1, sticky="ew", pady=3)
            counter = ttk.Label(card, text=f"0/{length}", style="Muted.TLabel", width=6, anchor="e")
            counter.grid(row=row, column=2, sticky="e")
            self.length_labels[label] = (counter, var, length)
            var.trace_add("write", lambda *_: self._update_lengths())
        self.input_widgets = [self.jira_entry, self.token_entry, self.port_combo] + [w for w in card.winfo_children()
                                                                    if isinstance(w, ttk.Entry)]
        self.input_widgets.append(self.internet_button)

    ##
    # @brief Online / Offline button. Starts Offline, every click toggles the mode.
    #        In online mode the ⟳ button next to it checks the internet connection.
    #
    # @param [in] parent Parent widget
    #
    # @return Frame with the buttons
    ##
    def _build_connection_button(self, parent):
        self.online_var = tk.BooleanVar(value=False)
        frame = ttk.Frame(parent, style="Card.TFrame")
        self.connection_button = tk.Label(frame, font=FONT_BOLD, fg="white", width=8, pady=2, cursor="hand2")
        self.connection_button.pack(side="left", fill="y")
        self.connection_button.bind("<Button-1>", lambda e: self._toggle_connection())
        Tooltip(self.connection_button, "Connection mode",
                "Click to switch between Online and Offline. Online turns green once the internet "
                "connection is confirmed.")
        # A box the same height as the Online button (a ttk button is taller)
        self.internet_button = tk.Label(frame, text="⟳", font=("Segoe UI Symbol", 10), width=3, pady=2, bd=0,
                                        bg=COLORS["card"], fg=COLORS["text"], cursor="hand2",
                                        highlightthickness=1, highlightbackground=COLORS["muted"])
        self.internet_button.bind("<Button-1>", lambda e: self._start_internet_check()
                                  if str(self.internet_button.cget("state")) != "disabled" else None)
        Tooltip(self.internet_button, "Internet connection",
                f"Check that the JIRA site ({urllib.parse.urlparse(JIRA_BASE_URL).hostname}) can be reached. "
                f"Retries for up to {INTERNET_CHECK_SEC} seconds.")
        # Shown next to ⟳ when online mode could not connect
        self.help_button = tk.Label(frame, text="?", font=FONT_BOLD, width=3, pady=2, bd=0,
                                    bg=COLORS["card"], fg=COLORS["text"], cursor="hand2",
                                    highlightthickness=1, highlightbackground=COLORS["muted"])
        self.help_button.bind("<Button-1>", lambda e: self._show_connection_help())
        Tooltip(self.help_button, "Cannot connect?", "Click to see what to check.")
        self._update_connection_button()
        return frame

    ##
    # @brief Show or hide the ? box (online mode only).
    #
    # @param [in] is_visible True to show the ? box
    #
    # @return None
    ##
    def _set_connection_help(self, is_visible):
        if is_visible and self.online_var.get():
            self.help_button.pack(side="left", fill="y", padx=(0, 6), before=self.internet_button)
        else:
            self.help_button.pack_forget()

    ##
    # @brief Show what to check when online mode cannot connect.
    #
    # @return None
    ##
    def _show_connection_help(self):
        messagebox.showinfo("Cannot connect to JIRA",
                            "Online mode could not connect. Please check:\n\n"
                            "1. Turn ON the internet connection.\n"
                            f"2. Use your JIRA login username ({self.username}).\n"
                            "3. Use an API token that is still alive (not expired or revoked).\n"
                            "4. Use the correct JIRA ID (e.g. SW9X07-10).\n\n"
                            f"Create API token here : {API_TOKEN_URL}")

    ##
    # @brief Online mode takes the build parameters from the JIRA link, so the
    #        Excel upload is removed (top right and Build Parameters tab).
    #
    # @return None
    ##
    def _toggle_connection(self):
        if self.is_running or self.is_checking_port or self.is_downloading:
            return
        self.online_var.set(not self.online_var.get())
        # Every switch cancels a running internet check
        self.is_internet_ok = False
        self.is_internet_checking = False
        self.internet_check_id += 1
        self._cancel_token_check()
        self._set_token_state(self.token_var.get().strip(), "none", "")
        self._cancel_topic_check()
        self._set_topic_state("", "none", "")
        self._set_connection_help(False)
        if self.online_var.get():
            self.upload_button.pack_forget()
            self.internet_button.pack(side="left", fill="y", padx=(0, 6), before=self.connection_button)
            self.token_label.grid(row=3, column=0, sticky="w", pady=3, padx=(0, 8))
            self.token_entry.grid(row=3, column=1, sticky="ew", pady=3)
            self.token_desc.grid(row=4, column=1, sticky="w")
        else:
            self.upload_button.pack(side="right", before=self.reset_button)
            self.internet_button.pack_forget()
            for widget in (self.token_label, self.token_entry, self.token_desc):
                widget.grid_remove()
        self._update_connection_button()
        self._update_detect_label()
        self._render_excel_values()
        self._update_start_state()
        if self.online_var.get():
            self._start_internet_check()

    ##
    # @brief Set the Online / Offline button text and color.
    #
    # @return None
    ##
    def _update_connection_button(self):
        # Offline: red, Online: grey until the internet connection is confirmed, then green
        is_online = self.online_var.get()
        if not is_online:
            color = COLORS["offline"]
        else:
            color = COLORS["online"] if self.is_internet_ok else COLORS["muted"]
        self.connection_button.configure(text="Online" if is_online else "Offline", bg=color)
        if not self.is_internet_checking:
            self.internet_button.configure(text="⟳")

    ##
    # @brief Online mode: check the internet connection by reaching the JIRA site.
    #        Retries until it answers, or reports no internet after INTERNET_CHECK_SEC.
    #
    # @return None
    ##
    def _start_internet_check(self):
        if not self.online_var.get() or self.is_internet_checking:
            return
        self.is_internet_ok = False
        self.is_internet_checking = True
        self.internet_check_id += 1
        self.internet_deadline = time.monotonic() + INTERNET_CHECK_SEC
        self._update_connection_button()
        self._update_start_state()
        self._spin_internet_button(self.internet_check_id, 0)
        self._try_internet(self.internet_check_id)

    ##
    # @brief Turn the spinner in the ⟳ box while the internet check runs.
    #
    # @param [in] check_id ID of the check, an older check is ignored
    # @param [in] frame Spinner frame number
    #
    # @return None
    ##
    def _spin_internet_button(self, check_id, frame):
        if check_id != self.internet_check_id or not self.is_internet_checking:
            return
        self.internet_button.configure(text=SPINNER_FRAMES[frame % len(SPINNER_FRAMES)])
        self.after(SPINNER_MS, self._spin_internet_button, check_id, frame + 1)

    ##
    # @brief Try once to reach the JIRA site (in a background thread).
    #
    # @param [in] check_id ID of the check, an older check is ignored
    #
    # @return None
    ##
    def _try_internet(self, check_id):
        if check_id != self.internet_check_id:
            return
        result = {}
        threading.Thread(target=lambda: result.update(ok=jira_reachable()), daemon=True).start()
        self._wait_internet_check(result, check_id)

    ##
    # @brief Wait for the internet check result: retry until INTERNET_CHECK_SEC, then show the result.
    #
    # @param [in] result Result filled in by the background thread
    # @param [in] check_id ID of the check, an older check is ignored
    #
    # @return None
    ##
    def _wait_internet_check(self, result, check_id):
        if check_id != self.internet_check_id:
            return  # Switched to Offline, or a newer check replaced this one
        if "ok" not in result:
            self.after(100, self._wait_internet_check, result, check_id)
            return
        if not result["ok"] and time.monotonic() < self.internet_deadline:
            self.after(INTERNET_RETRY_MS, self._try_internet, check_id)
            return
        self.is_internet_checking = False
        self.is_internet_ok = result["ok"]
        self._update_connection_button()
        self._update_start_state()
        self._set_connection_help(not self.is_internet_ok)
        if self.is_internet_ok:
            self._start_token_check()
        else:
            reason = (f"Internet is not available - {JIRA_BASE_URL} could not be reached for "
                      f"{INTERNET_CHECK_SEC} seconds.")
            self._console_write(f"[GUI] {reason}\n", "error")
            messagebox.showerror("No internet connection",
                                 f"{reason}\n\nCheck the network (or VPN) and press ⟳ to check again.")

    ##
    # @brief JIRA ticket of the build. The module (platform) is selected from
    #        the JIRA ID, the same way as main.py.
    #
    # @param [in] parent Parent widget
    #
    # @return None
    ##
    def _build_jira_card(self, parent):
        card = self._card(parent, None, fill="x", pady=(0, 10))
        card.columnconfigure(1, weight=1)

        # Title with the Online / Offline button on the right
        ttk.Label(card, text="JIRA Ticket", style="Section.TLabel").grid(row=0, column=0, columnspan=2,
                                                                       sticky="w", pady=(0, 6))
        self._build_connection_button(card).grid(row=0, column=1, sticky="e", pady=(0, 6))

        name = ttk.Label(card, text="JIRA ID", style="Key.TLabel", cursor="question_arrow")
        name.grid(row=1, column=0, sticky="w", pady=3, padx=(0, 8))
        Tooltip(name, "JIRA ID", "The JIRA ticket of the build, e.g. SW9X07-10. The module is selected "
                                 "from it: QCX216 → QCX216, SW9X07 → MDM9X07, SDX35 → SDX35.")
        self.jira_entry = ttk.Entry(card, textvariable=self.jira_var, font=FONT_MONO)
        self.jira_entry.grid(row=1, column=1, sticky="ew", pady=3)
        self.jira_desc = ttk.Label(card, text="", style="Muted.TLabel", wraplength=px(250))
        self.jira_desc.grid(row=2, column=1, sticky="w")
        self.jira_var.trace_add("write", lambda *_: self._on_jira_typed())
        self._update_jira_desc()

        # API token, only shown in online mode (used with the login username)
        self.token_label = ttk.Label(card, text="API Token", style="Key.TLabel", cursor="question_arrow")
        Tooltip(self.token_label, "JIRA API Token",
                "Create it at id.atlassian.com → Security → API tokens. It is used with your login "
                f"username ({self.username}) to open the JIRA ticket.")
        self.token_entry = ttk.Entry(card, textvariable=self.token_var, font=FONT_MONO, show="•",
                                     exportselection=False)
        # The token cannot be copied out of the box
        for sequence in ("<<Copy>>", "<<Cut>>"):
            self.token_entry.bind(sequence, lambda e: "break")
        self.token_desc = ttk.Label(card, text=f"Used with {self.username}", style="Muted.TLabel",
                                    wraplength=px(250))
        self.token_var.trace_add("write", lambda *_: self._on_token_changed())

    ##
    # @brief Module selected by the JIRA ID in the box.
    #
    # @return Module key, or None if the JIRA ID has no known project
    ##
    def _jira_module(self):
        return module_from_jira(self.jira_var.get().strip().upper())

    ##
    # @brief The JIRA ID changed: update its description, the pre-check and the Excel view.
    #
    # @return None
    ##
    def _on_jira_changed(self):
        self._update_jira_desc()
        if hasattr(self, "precheck_chips"):
            self._update_start_state()
            if self._online_pending() != self.is_online_pending_shown:
                # A sheet downloaded for another JIRA ID is not used
                self._update_detect_label()
                self._render_excel_values()

    ##
    # @brief Show the module of the JIRA ID, or the JIRA topic check result, below the JIRA ID box.
    #
    # @return None
    ##
    def _update_jira_desc(self):
        key = self._jira_module()
        topic_id, topic_status, topic_reason = self.topic_state
        if not self.jira_var.get().strip():
            text, color = "e.g. SW9X07-10", COLORS["muted"]
        elif key is None:
            text, color = f"✗ Must contain {', '.join(JIRA_PROJECTS)}", COLORS["fail"]
        elif topic_id != self.jira_var.get().strip().upper() or topic_status == "none":
            text, color = f"✓ {key} ({MODULE_PROFILES[key].technology})", COLORS["pass"]
        elif topic_status == "checking":
            text, color = "Checking the JIRA topic…", COLORS["muted"]
        elif topic_status == "ok":
            text, color = f"✓ {key} ({MODULE_PROFILES[key].technology}) - Pretest JIRA", COLORS["pass"]
        else:
            text, color = f"✗ {topic_reason[0].upper()}{topic_reason[1:]}", COLORS["fail"]
        self.jira_desc.configure(text=text, foreground=color)

    ##
    # @brief The JIRA ID was typed in: check its topic (after a short pause in typing).
    #
    # @return None
    ##
    def _on_jira_typed(self):
        self._start_topic_check(TOKEN_CHECK_DELAY_MS)
        self._on_jira_changed()

    ##
    # @brief Online mode: check that the JIRA ticket is a pre-test ticket, once
    #        the internet connection and the API token are OK.
    #
    # @param [in] delay Delay before the check in ms (0: check now)
    #
    # @return None
    ##
    def _start_topic_check(self, delay=0):
        self._cancel_topic_check()
        jira_id = self.jira_var.get().strip().upper()
        if not (self.online_var.get() and self.is_internet_ok and self._token_check()[0] == "ok"
                and self._jira_module()):
            self._set_topic_state(jira_id, "none", "")
            return
        self._set_topic_state(jira_id, "checking", "")
        self.topic_check_job = self.after(delay, self._check_topic, jira_id, self.token_var.get().strip(),
                                          self.topic_check_id)

    ##
    # @brief Cancel a pending or running JIRA topic check.
    #
    # @return None
    ##
    def _cancel_topic_check(self):
        self.topic_check_id += 1
        if self.topic_check_job is not None:
            self.after_cancel(self.topic_check_job)
            self.topic_check_job = None

    ##
    # @brief Check the JIRA topic in a background thread.
    #
    # @param [in] jira_id JIRA ID, e.g. SDX35-112
    # @param [in] token Atlassian API token
    # @param [in] check_id ID of the check, an older check is ignored
    #
    # @return None
    ##
    def _check_topic(self, jira_id, token, check_id):
        self.topic_check_job = None
        result = {}
        threading.Thread(target=lambda: result.update(check=jira_check_topic(jira_id, self.username, token)),
                         daemon=True).start()
        self._wait_topic_check(result, jira_id, check_id)

    ##
    # @brief Wait for the JIRA topic check result and show it.
    #
    # @param [in] result Result filled in by the background thread
    # @param [in] jira_id JIRA ID, e.g. SDX35-112
    # @param [in] check_id ID of the check, an older check is ignored
    #
    # @return None
    ##
    def _wait_topic_check(self, result, jira_id, check_id):
        if check_id != self.topic_check_id:
            return  # JIRA ID changed, or switched to Offline
        if "check" not in result:
            self.after(100, self._wait_topic_check, result, jira_id, check_id)
            return
        status, reason = result["check"]
        self._set_topic_state(jira_id, status, reason)
        self._set_connection_help(status != "ok")

    ##
    # @brief Pre-check of the JIRA topic (online mode). Only checked once the API
    #        token is OK - Start waits for the token anyway.
    #
    # @return (state, reason) with state "ok", "checking" or "bad"
    ##
    def _topic_check(self):
        if not (self.online_var.get() and self.is_internet_ok and self._token_check()[0] == "ok"):
            return "ok", ""
        jira_id, status, reason = self.topic_state
        if jira_id != self.jira_var.get().strip().upper() or status in ("none", "checking"):
            return "checking", "checking the JIRA topic"
        if status == "ok":
            return "ok", ""
        return "bad", reason

    ##
    # @brief Store the JIRA topic check state and update the view.
    #
    # @param [in] jira_id JIRA ID, e.g. SDX35-112
    # @param [in] status none, checking, ok, bad or error
    # @param [in] reason Reason the JIRA cannot be used, "" if none
    #
    # @return None
    ##
    def _set_topic_state(self, jira_id, status, reason):
        self.topic_state = (jira_id, status, reason)
        self._update_jira_desc()
        self._update_start_state()

    ##
    # @brief API token check, right after the internet check: use the token in the
    #        box, otherwise the one stored in Windows Credential Manager.
    #
    # @return None
    ##
    def _start_token_check(self):
        self._cancel_token_check()
        token = self.token_var.get().strip()
        if token:
            self._set_token_state(token, "checking", "Checking the API token…")
            self._check_token(token, self.token_check_id, self.is_token_from_store)
            return
        try:
            stored = load_token(self.username)
        except (OSError, UnicodeDecodeError):
            stored = ""
        if not stored:
            self._set_token_state("", "none", "No token found. Please input API token.")
            return
        self._set_token_state("", "checking", "Checking the API token…")
        self._check_token(stored, self.token_check_id, True)

    ##
    # @brief Cancel a pending or running API token check.
    #
    # @return None
    ##
    def _cancel_token_check(self):
        self.token_check_id += 1
        if self.token_check_job is not None:
            self.after_cancel(self.token_check_job)
            self.token_check_job = None

    ##
    # @brief The token was typed in: check it (after a short pause in typing).
    #
    # @return None
    ##
    def _on_token_changed(self):
        token = self.token_var.get().strip()
        if token and token == self.token_state[0]:
            self._on_jira_changed()
            return  # Filled in from Windows Credential Manager, already checked
        self.is_token_from_store = False
        self._cancel_token_check()
        if not (self.online_var.get() and self.is_internet_ok):
            self._set_token_state(token, "none", "")    # Checked after the internet check
        elif not token:
            self._set_token_state("", "none", "Please input API token.")
        else:
            self._set_token_state(token, "checking", "Checking the API token…")
            self.token_check_job = self.after(TOKEN_CHECK_DELAY_MS, self._check_token, token,
                                              self.token_check_id, False)
        self._on_jira_changed()

    ##
    # @brief Check the API token in a background thread.
    #
    # @param [in] token Atlassian API token
    # @param [in] check_id ID of the check, an older check is ignored
    # @param [in] is_stored True if the token is the remembered one
    #
    # @return None
    ##
    def _check_token(self, token, check_id, is_stored):
        self.token_check_job = None
        result = {}
        threading.Thread(target=lambda: result.update(check=jira_check_token(self.username, token)),
                         daemon=True).start()
        self._wait_token_check(result, token, check_id, is_stored)

    ##
    # @brief Wait for the API token check result: fill in, remember or reject the token.
    #
    # @param [in] result Result filled in by the background thread
    # @param [in] token Atlassian API token
    # @param [in] check_id ID of the check, an older check is ignored
    # @param [in] is_stored True if the token is the remembered one
    #
    # @return None
    ##
    def _wait_token_check(self, result, token, check_id, is_stored):
        if check_id != self.token_check_id:
            return  # Token changed, or switched to Offline
        if "check" not in result:
            self.after(100, self._wait_token_check, result, token, check_id, is_stored)
            return
        status, reason = result["check"]
        if is_stored:
            if status == "ok":
                # Auto-fill: the state is set first, so the box change is not checked again
                self.is_token_from_store = True
                self._set_token_state(token, "ok", "✓ Token valid")
                self.token_var.set(token)
            else:
                if self.token_var.get():
                    self.token_var.set("")      # Remove the auto-filled token that no longer works
                if status == "bad":
                    self._set_token_state("", "none", "Token expired. Please input API token.")
                else:
                    self._set_token_state("", "none", f"Cannot check the stored token ({reason}). "
                                                      "Please input API token.")
        elif status == "ok":
            try:
                save_token(self.username, token)
            except OSError:
                pass    # Still usable, only not filled in next time
            self._set_token_state(token, "ok", "✓ Token valid")
        elif status == "bad":
            self._set_token_state(token, "bad", "✗ Token is invalid or expired. Please input API token.")
        else:
            self._set_token_state(token, "error", f"✗ Cannot check the token ({reason})")
        self._set_connection_help(self.token_state[1] != "ok")
        if self.token_state[1] == "ok":
            self._start_topic_check()

    ##
    # @brief Pre-check of the API token in the box (online mode, internet OK).
    #
    # @return (state, reason) with state "ok", "checking" or "bad"
    ##
    def _token_check(self):
        token = self.token_var.get().strip()
        checked, status, _ = self.token_state
        if status == "checking":
            return "checking", "checking the API token"
        if not token:
            return "bad", "enter the JIRA API token"
        if checked != token or status == "none":
            return "checking", "checking the API token"
        if status == "ok":
            return "ok", ""
        return "bad", "the API token is invalid or expired" if status == "bad" else "the API token cannot be checked"

    ##
    # @brief Store the API token check state and show its message below the API token box.
    #
    # @param [in] token Atlassian API token
    # @param [in] status none, checking, ok, bad or error
    # @param [in] message Message below the box, "" for the default text
    #
    # @return None
    ##
    def _set_token_state(self, token, status, message):
        self.token_state = (token, status, message)
        color = {"ok": COLORS["pass"], "checking": COLORS["muted"]}.get(status, COLORS["fail"])
        if not message:
            message, color = f"Used with {self.username}", COLORS["muted"]
        self.token_desc.configure(text=message, foreground=color)
        self._update_start_state()

    ##
    # @brief Add a notebook tab whose content scrolls vertically when it is
    #        taller than the tab (scrollbar only shown when needed).
    #
    # @param [in] tabs Notebook the tab is added to
    # @param [in] title Tab title
    #
    # @return (canvas, content frame)
    ##
    def _scrollable_tab(self, tabs, title):
        outer = ttk.Frame(tabs, style="Card.TFrame")
        tabs.add(outer, text=title)
        canvas = tk.Canvas(outer, bg=COLORS["card"], highlightthickness=0, bd=0, yscrollincrement=20)
        scrollbar = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        card = ttk.Frame(canvas, style="Card.TFrame", padding=(12, 10))
        window = canvas.create_window(0, 0, window=card, anchor="nw")

        ##
        # @brief Fit the content to the tab width and show the scrollbar only when the content does not fit.
        #
        # @param [in] _event Tk event (not used)
        #
        # @return None
        ##
        def refresh(_event=None):
            canvas.itemconfigure(window, width=canvas.winfo_width())
            canvas.configure(scrollregion=(0, 0, card.winfo_reqwidth(), card.winfo_reqheight()))
            # Show the scrollbar only when the content does not fit
            is_scroll_needed = card.winfo_reqheight() > canvas.winfo_height() > 1
            if is_scroll_needed and not scrollbar.winfo_ismapped():
                scrollbar.pack(side="right", fill="y", before=canvas)
            elif not is_scroll_needed and scrollbar.winfo_ismapped():
                scrollbar.pack_forget()
                canvas.yview_moveto(0)
            else:
                pass    # Scrollbar already right

        ##
        # @brief Scroll the tab with the mouse wheel (only when it can scroll).
        #
        # @param [in] event Tk mouse wheel event
        #
        # @return None
        ##
        def on_wheel(event):
            if scrollbar.winfo_ismapped():
                canvas.yview_scroll(int(-event.delta / 120), "units")

        card.bind("<Configure>", refresh)
        canvas.bind("<Configure>", refresh)
        canvas.bind("<Enter>", lambda e: canvas.bind_all("<MouseWheel>", on_wheel))
        canvas.bind("<Leave>", lambda e: canvas.unbind_all("<MouseWheel>"))
        return canvas, card

    ##
    # @brief Build the Build Parameters tab.
    #
    # @param [in] tabs Notebook the tab is added to
    #
    # @return None
    ##
    def _build_excel_card(self, tabs):
        # Scrollable: the parameter list can be taller than the tab
        self.excel_canvas, card = self._scrollable_tab(tabs, "  Build Parameters  ")
        self.excel_file_label = ttk.Label(card, text="", style="Muted.TLabel", wraplength=px(300))
        self.excel_file_label.pack(anchor="w", pady=(0, 6))
        self.excel_grid = ttk.Frame(card, style="Card.TFrame")
        self.excel_grid.pack(fill="x")
        self.excel_grid.columnconfigure(1, weight=1)

    ##
    # @brief Build the Module Profile tab.
    #
    # @param [in] tabs Notebook the tab is added to
    #
    # @return None
    ##
    def _build_profile_card(self, tabs):
        self.profile_canvas, self.profile_frame = self._scrollable_tab(tabs, "  Module Profile  ")

    ##
    # @brief Build the progress card: status, progress bar, PASS / FAIL / SKIP counts and run time.
    #
    # @param [in] parent Parent widget
    #
    # @return None
    ##
    def _build_progress_card(self, parent):
        card = self._card(parent, None, fill="x")
        top = ttk.Frame(card, style="Card.TFrame")
        top.pack(fill="x")
        self.status_label = ttk.Label(top, text="Ready", style="Big.TLabel")
        self.status_label.pack(side="left")
        self.percent_label = ttk.Label(top, text="0 %", style="Big.TLabel")
        self.percent_label.pack(side="right")

        self.progress = ttk.Progressbar(card, style="Run.Horizontal.TProgressbar", mode="determinate")
        self.progress.pack(fill="x", pady=(8, 6))

        bottom = ttk.Frame(card, style="Card.TFrame")
        bottom.pack(fill="x")
        self.detail_label = ttk.Label(bottom, text="Fill in the SIM information and press Start.", style="Muted.TLabel", wraplength=px(520))
        self.detail_label.pack(side="left")
        self.chips = {}
        for key, text, fg, bg in (("time", "00:00", COLORS["text"], "#eef0f3"),
                                  ("UNSUPPORTED", "SKIP 0", COLORS["skip"], COLORS["skip_bg"]),
                                  ("FAIL", "FAIL 0", COLORS["fail"], COLORS["fail_bg"]),
                                  ("PASS", "PASS 0", COLORS["pass"], COLORS["pass_bg"])):
            chip = tk.Label(bottom, text=text, font=FONT_BOLD, fg=fg, bg=bg, padx=10, pady=2)
            chip.pack(side="right", padx=(6, 0))
            self.chips[key] = chip

    ##
    # @brief Build the Test Cases table and the details of the selected command.
    #
    # @param [in] parent Parent widget
    #
    # @return Frame of the table
    ##
    def _build_table(self, parent):
        outer = tk.Frame(parent, bg=COLORS["card"], highlightthickness=1, highlightbackground=COLORS["border"])
        frame = ttk.Frame(outer, style="Card.TFrame", padding=(10, 8))
        frame.pack(fill="both", expand=True)

        head = ttk.Frame(frame, style="Card.TFrame")
        head.pack(fill="x", pady=(0, 6))
        ttk.Label(head, text="Test Cases", style="Section.TLabel").pack(side="left")
        self.table_hint = ttk.Label(head, text="Select a row to see the modem response", style="Muted.TLabel")
        self.table_hint.pack(side="right")

        table_area = ttk.Frame(frame, style="Card.TFrame")
        table_area.pack(fill="both", expand=True)
        columns = ("no", "case", "idx", "command", "expected", "result")
        self.tree = ttk.Treeview(table_area, columns=columns, show="headings", selectmode="browse")
        for col, text, width, anchor, stretch in (("no", "#", 40, "center", False),
                                                   ("case", "Test Case", 190, "w", False),
                                                   ("idx", "Step", 45, "center", False),
                                                   ("command", "AT Command", 260, "w", True),
                                                   ("expected", "Expected", 260, "w", True),
                                                   ("result", "Result", 110, "center", False)):
            self.tree.heading(col, text=text)
            self.tree.column(col, width=px(width), anchor=anchor, stretch=stretch)
        scroll = ttk.Scrollbar(table_area, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set)
        # Scrollbar first, so a wide table at high display scaling cannot push it out of view
        scroll.pack(side="right", fill="y")
        self.tree.pack(side="left", fill="both", expand=True)

        self.tree.tag_configure("PASS", foreground=COLORS["pass"], background=COLORS["pass_bg"])
        self.tree.tag_configure("FAIL", foreground=COLORS["fail"], background=COLORS["fail_bg"])
        self.tree.tag_configure("UNSUPPORTED", foreground=COLORS["skip"], background=COLORS["skip_bg"])
        self.tree.tag_configure("NA", foreground="#9ca3af")
        self.tree.tag_configure("RUNNING", background="#fef3c7")
        self.tree.bind("<<TreeviewSelect>>", lambda e: self._show_selected_detail())
        self.tree.bind("<Configure>", self._fit_columns, add="+")

        # Faint vertical lines between columns (ttk.Treeview has no column grid lines)
        self.column_lines = []
        for sequence in ("<Configure>", "<B1-Motion>", "<ButtonRelease-1>"):
            self.tree.bind(sequence, lambda e: self.after_idle(self._draw_column_lines), add="+")

        self.detail_text = tk.Text(frame, height=5, font=FONT_MONO, relief="flat", bg=COLORS["console_bg"],
                                   fg=COLORS["console_fg"], insertbackground="white", wrap="word",
                                   padx=10, pady=6, state="disabled")
        # Packed ahead of the table so the table shrinks first, never this box
        self.detail_text.pack(side="bottom", fill="x", pady=(8, 0), before=table_area)
        # Same highlighting as the live log
        self.detail_text.tag_configure("PASS", foreground="#4ade80")
        self.detail_text.tag_configure("FAIL", foreground="#f87171")
        self.detail_text.tag_configure("UNSUPPORTED", foreground="#9ca3af")
        self.detail_text.tag_configure("title", foreground="#93c5fd", font=("Consolas", 10, "bold"))
        self.detail_text.tag_configure("label", foreground="#6b7280")
        return outer

    ##
    # @brief Build the live log console.
    #
    # @param [in] parent Parent widget
    #
    # @return Frame of the console
    ##
    def _build_console(self, parent):
        outer = tk.Frame(parent, bg=COLORS["card"], highlightthickness=1, highlightbackground=COLORS["border"])
        frame = ttk.Frame(outer, style="Card.TFrame", padding=(10, 8))
        frame.pack(fill="both", expand=True)

        head = ttk.Frame(frame, style="Card.TFrame")
        head.pack(fill="x", pady=(0, 6))
        ttk.Label(head, text="Live Log", style="Section.TLabel").pack(side="left")
        self.download_button = ttk.Button(head, text="Download Result", command=self._show_download_menu,
                                          state="disabled")
        self.download_button.pack(side="right")
        self.download_menu = tk.Menu(self, tearoff=0, font=FONT)
        self.download_menu.add_command(label="Log (.txt)", command=self.download_log)
        self.download_menu.add_command(label="Report (.pdf)", command=self.download_report)
        ttk.Button(head, text="Clear", command=self.clear_console).pack(side="right", padx=(0, 6))
        ttk.Checkbutton(head, text="Auto-scroll", variable=self.autoscroll_var,
                        style="Card.TCheckbutton").pack(side="right", padx=(0, 10))
        self.log_path_label = ttk.Label(head, text="", style="Muted.TLabel")
        self.log_path_label.pack(side="left", padx=12)

        self.console = ScrolledText(frame, font=FONT_MONO, bg=COLORS["console_bg"], fg=COLORS["console_fg"],
                                    insertbackground="white", relief="flat", padx=10, pady=8, height=10,
                                    state="disabled", wrap="none")
        self.console.pack(fill="both", expand=True)
        self.console.tag_configure("PASS", foreground="#4ade80")
        self.console.tag_configure("FAIL", foreground="#f87171")
        self.console.tag_configure("UNSUPPORTED", foreground="#9ca3af")
        self.console.tag_configure("index", foreground="#93c5fd", font=("Consolas", 10, "bold"))
        self.console.tag_configure("divider", foreground="#4b5563")
        self.console.tag_configure("error", foreground="#fca5a5", font=("Consolas", 10, "bold"))
        self.console.tag_configure("gui", foreground="#fcd34d")
        self.console.tag_configure("summary", foreground="#ffffff", font=("Consolas", 10, "bold"))
        self.console.tag_configure("search", background="#facc15", foreground="#111827")
        self.console.tag_configure("search_current", background="#fb923c", foreground="#111827")
        return outer

    # Module view

    ##
    # @brief Show the view of a module, or the default view.
    #
    # @param [in] key Module key (MDM9X07, QCX216, SDX35), or None for the default view
    #
    # @return None
    ##
    def select_module(self, key):
        if self.is_running:
            return
        self.profile = MODULE_PROFILES[key] if key else None
        self.module_var.set(key or "")
        if self.profile:
            self.header_module.configure(text=key, padx=0)
            self.tech_badge.configure(text=self.profile.technology)
            self.tech_badge.pack(side="right", padx=(6, 18), before=self.header_module)
        else:
            self.header_module.configure(text="Anas Akasyah", padx=18)
            self.tech_badge.pack_forget()
        self._apply_accent()
        self._update_detect_label()
        self._render_excel_values()
        self._render_profile()
        self._populate_table()
        self._reset_progress()

    ##
    # @brief Return to the default view: clears the previous run's results,
    #        log and module selection. COM port and SIM inputs are kept.
    #
    # @return None
    ##
    def reset_view(self):
        if self.is_running:
            return
        self._hide_search()
        self.clear_console()
        self.log_file_path = None
        self.log_path_label.configure(text="")
        self.download_button.configure(state="disabled")
        self.reload_excel()

    ##
    # @brief Read the stored Excel sheet again and return to the default view.
    #
    # @return None
    ##
    def reload_excel(self):
        if self.is_running:
            return
        self.excel = detect_excel()
        if self.excel.path is not None:
            self.excel_file_label.configure(text=self.excel.path.name)
        else:
            self.excel_file_label.configure(text="")
        # Default view until a pre-test is started
        self.select_module(None)
        self._update_start_state()

    ##
    # @brief Check that the Excel sheet can be used.
    #
    # @return True if the Excel sheet can be used, otherwise False
    ##
    def _excel_ready(self):
        return self.excel.path is not None and not self.excel.error

    ##
    # @brief Online mode only uses a sheet downloaded from the current JIRA ticket.
    #
    # @return True if online mode is still waiting for the JIRA sheet
    ##
    def _online_pending(self):
        if not self.online_var.get():
            return False
        current = (self.jira_var.get().strip().upper(), self.excel.path)
        return not (self._excel_ready() and self.jira_excel == current)

    ##
    # @brief Pre-check before a run: AT port, Excel sheet, SIM CCID / IMEI / IMSI.
    #
    # @return {name: (state, reason)} with state "ok", "checking" or "bad"
    ##
    def _precheck(self):
        port, status, reason = self.port_state
        current = self.port_var.get().strip()
        if self._port_ready():
            port_check = ("ok", "")
        elif status == "checking" and port == current:
            port_check = ("checking", f"checking {current}")
        else:
            port_check = ("bad", reason if (reason and port == current) else "select the modem AT port")
        jira_check = self._topic_check() if self._jira_module() else ("bad", "enter a valid JIRA ID")
        # Online mode: the Excel sheet is downloaded when the pre-test starts, the API token is checked instead
        if self.online_var.get():
            if self.is_internet_checking:
                excel_check = ("checking", "checking the internet connection")
            elif not self.is_internet_ok:
                excel_check = ("bad", "no internet connection - press ⟳ to check again")
            else:
                excel_check = self._token_check()
        elif self._excel_ready():
            excel_check = ("ok", "")
        else:
            excel_check = ("bad", self.excel.error or "no Excel sheet")
        checks = {"JIRA": jira_check,
                  "AT port": port_check,
                  "Excel": excel_check}
        for name, label in (("CCID", "SIM CCID"), ("IMEI", "SIM IMEI"), ("IMSI", "SIM IMSI")):
            length = self.length_labels[label][2]
            kind = "digits" if label != "SIM CCID" else "characters"
            checks[name] = ("ok", "") if self._sim_ok(label) else ("bad", f"{label} needs {length} {kind}")
        return checks

    ##
    # @brief Update the pre-check chips, the Start button, and the status that explains a disabled Start.
    #
    # @return None
    ##
    def _update_start_state(self):
        if not hasattr(self, "precheck_chips") or not hasattr(self, "length_labels"):
            return
        checks = self._precheck()
        for name, (state, _) in checks.items():
            symbol, fg, bg = {"ok": ("✓", COLORS["pass"], COLORS["pass_bg"]),
                              "checking": ("…", COLORS["muted"], COLORS["skip_bg"]),
                              "bad": ("✗", COLORS["fail"], COLORS["fail_bg"])}[state]
            label = "API" if (name == "Excel" and self.online_var.get()) else name
            self.precheck_chips[name].configure(text=f"{symbol} {label}", fg=fg, bg=bg)

        is_idle = not self.is_running and not self.is_checking_port and not self.is_downloading
        is_ready = is_idle and all(state == "ok" for state, _ in checks.values())
        self.start_button.configure(state="normal" if is_ready else "disabled")

        # Explain a disabled Start in the default view
        if is_idle and self.profile is None:
            pending = [(name, reason) for name, (state, reason) in checks.items() if state != "ok"]
            if not pending:
                self._set_status("Ready", "All checks passed - press Start Pre-Test")
            elif all(state == "checking" for name, (state, _) in checks.items() if state != "ok"):
                checking = []
                if checks["JIRA"][0] == "checking":
                    checking.append("Checking the JIRA topic…")
                if checks["AT port"][0] == "checking":
                    checking.append(f"Checking {self.port_var.get().strip()} for an AT port…")
                if checks["Excel"][0] == "checking":
                    checking.append(f"{checks['Excel'][1].capitalize()}…")
                self._set_status("Ready", "  ·  ".join(checking))
            else:
                details = ";  ".join(reason for _, reason in pending if reason)
                self._set_status("Not Ready", details, COLORS["fail"])

    ##
    # @brief Upload a new Jenkins Build Parameter sheet. The sheet is copied into
    #        the application directory (where the pre-test modules look for it)
    #        and the previously stored sheet(s) are deleted.
    #
    # @return None
    ##
    def upload_excel(self):
        if self.is_running or self.online_var.get():
            return
        selected = filedialog.askopenfilename(title="Select the Jenkins Build Parameter Excel sheet",
                                              filetypes=[("Excel workbook", "*.xlsx")])
        if not selected:
            return
        if self._store_excel(Path(selected), "Upload") is not None:
            self.reset_view()

    ##
    # @brief Copy a Jenkins Build Parameter sheet into the application directory
    #        and delete the previously stored sheet(s).
    #
    # @param [in] source Excel sheet to store
    # @param [in] action "Upload" or "Download", used in the messages
    #
    # @return Path of the stored sheet, or None if nothing was changed
    ##
    def _store_excel(self, source, action):
        # Validate before anything is replaced
        try:
            pretest_main.read_jenkins_value(source, "platformName")
        except KeyError:
            messagebox.showerror("Invalid Excel sheet",
                                 f"{source.name} has no 'Jenkins Parameter' / 'Value' columns.\n"
                                 "Nothing was changed.")
            return None
        except Exception as e:
            messagebox.showerror("Invalid Excel sheet", f"Cannot read {source.name}:\n{e}\n\nNothing was changed.")
            return None

        target = app_directory() / source.name
        old_sheets = [path for path in stored_excel_files() if path.resolve() != source.resolve()]
        if old_sheets:
            names = "\n".join(f"  •  {path.name}" for path in old_sheets)
            if not messagebox.askokcancel("Replace Excel sheet",
                                          f"{action} {source.name}?\n\nThe current sheet will be deleted:\n{names}"):
                return None

        try:
            if source.resolve() != target.resolve():
                shutil.copy2(source, target)
        except OSError as e:
            messagebox.showerror(f"{action} failed", f"Cannot copy {source.name}:\n{e}\n\nNothing was changed.")
            return None

        failed = []
        for path in old_sheets:
            if path.resolve() == target.resolve():
                continue    # Same name - already overwritten by the copy
            try:
                path.unlink()
            except OSError as e:
                failed.append(f"  •  {path.name}: {e}")
        if failed:
            messagebox.showwarning("Old sheet not deleted",
                                   f"The new sheet was {action.lower()}ed, but these could not be deleted "
                                   f"(close them in Excel and {action.lower()} again):\n" + "\n".join(failed))
        return target

    ##
    # @brief Online mode start: delete the previous Excel sheet(s), download the
    #        sheet attached to the JIRA ticket (JIRA_BROWSE_URL + JIRA ID) into
    #        the application directory, then run the pre-test.
    #
    # @return None
    ##
    def _download_and_launch(self):
        jira_id = self.jira_var.get().strip().upper()
        token = self.token_var.get().strip()
        url = JIRA_BROWSE_URL + jira_id

        # Only one Excel sheet in the folder at a time: delete the previous sheet(s) first
        failed = []
        for path in stored_excel_files():
            try:
                path.unlink()
                self._console_write(f"[GUI] Deleted previous Excel sheet: {path.name}\n", "gui")
            except OSError as e:
                failed.append(f"{path.name} ({e.strerror})")
        self.jira_excel = None
        self.reload_excel()
        if failed:
            reason = f"Cannot delete {', '.join(failed)} - close it in Excel and press Start again."
            self._set_status("Excel download failed", reason, COLORS["fail"])
            self._console_write(f"[GUI] {reason}\n", "error")
            messagebox.showerror("Excel download failed", reason)
            return

        self.is_downloading = True
        self._set_running_ui(True)
        self.stop_button.configure(state="disabled")
        self._set_status("Downloading Excel…", f"From {url}")
        self._console_write(f"[GUI] Downloading the Excel sheet from {url}\n", "gui")
        folder = tempfile.TemporaryDirectory()
        result = {}

        ##
        # @brief Download the Excel sheet (background thread).
        #
        # @return None
        ##
        def download():
            try:
                result["path"] = jira_download_excel(jira_id, self.username, token, folder.name)
            except JiraError as e:
                result["error"] = str(e)
            except Exception as e:
                result["error"] = f"Download failed: {e}"

        threading.Thread(target=download, daemon=True).start()
        self._wait_jira_download(result, jira_id, folder)

    ##
    # @brief Wait for the Excel download: store the sheet and run the pre-test, or report the error.
    #
    # @param [in] result Result filled in by the background thread
    # @param [in] jira_id JIRA ID, e.g. SDX35-112
    # @param [in] folder Temporary folder, deleted when done
    #
    # @return None
    ##
    def _wait_jira_download(self, result, jira_id, folder):
        if not result:
            self.after(100, self._wait_jira_download, result, jira_id, folder)
            return
        self.is_downloading = False
        self._set_running_ui(False)
        url = JIRA_BROWSE_URL + jira_id
        try:
            if "error" in result:
                self._set_status("Excel download failed", result["error"], COLORS["fail"])
                self._console_write(f"[GUI] {result['error']}\n", "error")
                self._set_connection_help(True)
                messagebox.showerror("Excel download failed", f"{result['error']}\n\nJIRA ticket: {url}")
                return
            self._set_connection_help(False)
            target = self._store_excel(result["path"], "Download")
            if target is None:
                self._set_status("Excel download failed", "The downloaded sheet is not a valid build parameter sheet",
                                 COLORS["fail"])
                return
        finally:
            folder.cleanup()

        self.jira_excel = (jira_id, target)
        self.reload_excel()
        self._launch_test()
        self._console_write(f"[GUI] Excel sheet {target.name} downloaded from {url}\n", "gui")

    ##
    # @brief Show which Excel sheet is used, or why it cannot be used.
    #
    # @return None
    ##
    def _update_detect_label(self):
        info = self.excel
        if self._online_pending():
            text, color = "Online mode - the build parameters will be extracted from the JIRA link", COLORS["muted"]
        elif info.error:
            text, color = f"⚠ {info.error}", COLORS["fail"]
        elif self.profile is None:
            text = "The module is selected from the JIRA ID when the pre-test starts"
            color = COLORS["muted"]
        else:
            text = f"✓ {self.profile.key} selected (JIRA ID = {self.jira_id})"
            color = COLORS["pass"]
        if info.warning:
            text, color = f"⚠ {info.warning}", COLORS["warn"]
        self.detect_label.configure(text=text, fg=color)

    ##
    # @brief Show the build parameters of the Excel sheet in the Build Parameters tab.
    #
    # @return None
    ##
    def _render_excel_values(self):
        self.is_online_pending_shown = self._online_pending()
        self.excel_canvas.yview_moveto(0)
        for child in self.excel_grid.winfo_children():
            child.destroy()
        if self._online_pending():
            self.excel_file_label.configure(text="")
            ttk.Label(self.excel_grid, text="In online mode, the build parameters will be extracted from the JIRA link.",
                      style="Muted.TLabel", wraplength=px(320)).grid(row=0, column=0, columnspan=2, sticky="w")
            self.qgmr = ""
            return
        self.excel_file_label.configure(text=self.excel.path.name if self.excel.path is not None else "")
        if self.excel.path is None or self.excel.error:
            banner = tk.Frame(self.excel_grid, bg=COLORS["fail_bg"], highlightthickness=1,
                              highlightbackground=COLORS["fail"])
            banner.grid(row=0, column=0, columnspan=2, sticky="ew")
            tk.Label(banner, text="⚠  Excel sheet not available", font=FONT_BOLD, fg=COLORS["fail"],
                     bg=COLORS["fail_bg"]).pack(anchor="w", padx=10, pady=(8, 2))
            tk.Label(banner, text=self.excel.error, font=FONT_SMALL, fg=COLORS["fail"], bg=COLORS["fail_bg"],
                     wraplength=px(290), justify="left").pack(anchor="w", padx=10)
            ttk.Button(banner, text="⬆  Upload Excel", command=self.upload_excel).pack(anchor="w", padx=10, pady=(6, 10))
            self.qgmr = ""
            return

        if self.profile is None:
            ttk.Label(self.excel_grid, text="Module specific build parameters are shown once the pre-test starts.",
                      style="Muted.TLabel", wraplength=px(320)).grid(row=0, column=0, columnspan=2, sticky="w")
            self.qgmr = ""
            return

        values = read_excel_values(self.excel.path, self.profile.excel_keys)
        self.qgmr = preview_expected(self.profile)
        rows = [(label, values[key]) for label, key in self.profile.excel_keys]
        rows.append(("Version (QGMR)", self.qgmr or "-"))
        for row, (label, value) in enumerate(rows):
            ttk.Label(self.excel_grid, text=label, style="Key.TLabel").grid(row=row, column=0, sticky="nw", padx=(0, 10), pady=1)
            style = "Muted.TLabel" if value == "None" else "Value.TLabel"
            ttk.Label(self.excel_grid, text=value, style=style, wraplength=px(195)).grid(row=row, column=1, sticky="w", pady=1)
        ttk.Label(self.excel_grid, text=f"QGMR rule: {self.profile.qgmr_rule}",
                  style="Muted.TLabel").grid(row=len(rows), column=0, columnspan=2, sticky="w", pady=(6, 0))

    ##
    # @brief Show the selected module profile in the Module Profile tab.
    #
    # @return None
    ##
    def _render_profile(self):
        self.profile_canvas.yview_moveto(0)
        frame = self.profile_frame
        for child in frame.winfo_children():
            child.destroy()
        if self.profile is None:
            ttk.Label(frame, text="Detected when the pre-test starts:", style="Value.TLabel").pack(anchor="w")
            for key, profile in MODULE_PROFILES.items():
                ttk.Label(frame, text=f"•  {key}  ({profile.technology})", style="Muted.TLabel").pack(anchor="w")
            return
        skipped = sum(1 for tc in at_cmd_test_case_list if self.profile.is_unsupported(tc))
        resets = sum(1 for tc in at_cmd_test_case_list if tc.at_command == const.RESET_AT_COMMAND)
        grid = ttk.Frame(frame, style="Card.TFrame")
        grid.pack(fill="x")
        rows = (("Technology", self.profile.technology),
                ("Test commands", f"{len(at_cmd_test_case_list) - skipped} run · {skipped} skipped"),
                ("Modem resets", f"{resets} × {self.profile.reset_delay} s" if self.profile.reset_delay else "Disabled"),
                ("Baud rate", str(const.BAUD_RATE)),
                ("Timeout", f"{const.SEND_TIMEOUT} s per command"))
        for row, (key, value) in enumerate(rows):
            ttk.Label(grid, text=key, style="Key.TLabel").grid(row=row, column=0, sticky="w", padx=(0, 10), pady=1)
            ttk.Label(grid, text=value, style="Value.TLabel").grid(row=row, column=1, sticky="w", pady=1)
        for note in self.profile.notes:
            ttk.Label(frame, text=f"•  {note}", style="Muted.TLabel").pack(anchor="w")

    ##
    # @brief Fill the Test Cases table with every test case of the selected module.
    #
    # @return None
    ##
    def _populate_table(self):
        self.tree.delete(*self.tree.get_children())
        self.records = {}
        self.detail_text.configure(state="normal")
        self.detail_text.delete("1.0", "end")
        self.detail_text.configure(state="disabled")
        for number, tc in enumerate(at_cmd_test_case_list, start=1):
            if self.profile and self.profile.is_unsupported(tc):
                result, tag = "N/A", "NA"
            else:
                result, tag = "", ""
            expected = self._expected_placeholder(number - 1, tc.expected_response) if self.profile else "-"
            case = tc.test_case_id.replace("Pretest_LTE_", "")
            self.tree.insert("", "end", iid=str(number), tags=(tag,) if tag else (),
                             values=(number, case, tc.at_cmd_index, tc.at_command, expected, result))
            self.records[number] = TestRecord(index=number, send=tc.at_command, expected=expected)
        self.tree.yview_moveto(0)

    ##
    # @brief Expected response shown before the run (the SIM values are only known when the run starts).
    #
    # @param [in] position Position in the test case list
    # @param [in] value Expected response of the test case
    #
    # @return Expected response, or a SIM placeholder
    ##
    @staticmethod
    def _expected_placeholder(position, value):
        # SIM based expectations are only known once the run starts
        from test_case import TestCaseIndex as Idx
        if position == Idx.CD_05_001_1:
            return "(SIM CCID)"
        if position in (Idx.CD_05_002_1, Idx.CD_05_002_2, Idx.CD_05_002_3):
            return "(SIM IMEI)"
        if position == Idx.CD_05_003_1:
            return "(SIM IMSI)"
        return value

    # Inputs

    ##
    # @brief List the COM ports, and select the modem AT port if none is selected.
    #
    # @return None
    ##
    def refresh_ports(self):
        self.ports = {p.device: p.description for p in serial.tools.list_ports.comports()}
        self.port_combo["values"] = sorted(self.ports, key=lambda d: (len(d), d))
        if not self.port_var.get() and self.ports:
            # Prefer a port the driver names as an AT port (e.g. "Eagle USB AT Port")
            at_ports = [device for device in self.port_combo["values"] if re.search(r"\bAT\b", self.ports[device])]
            self.port_var.set((at_ports or self.port_combo["values"])[0])
        self._on_port_changed()

    ##
    # @brief The selected port changed: check (after a short delay, so typing a
    #        port name does not open every partial name) whether it is an AT port.
    #
    # @return None
    ##
    def _on_port_changed(self):
        if self.is_running or self.is_checking_port:
            return
        if self.port_check_job is not None:
            self.after_cancel(self.port_check_job)
        port = self.port_var.get().strip()
        self.port_check_id += 1
        self.port_state = (port, "checking" if port else "none", "")
        self._update_port_desc()
        self._update_start_state()
        if port:
            self.port_check_job = self.after(PORT_CHECK_DELAY_MS, self._check_port, port, self.port_check_id)

    ##
    # @brief Check in a background thread whether a COM port is the modem AT port.
    #
    # @param [in] port COM port, e.g. COM5
    # @param [in] check_id ID of the check, an older check is ignored
    #
    # @return None
    ##
    def _check_port(self, port, check_id):
        self.port_check_job = None
        if NON_AT_PORT_PATTERN.search(self.ports.get(port, "")):
            # Known non-AT interface (DM / NMEA ...): reject without sending anything
            self._set_port_state(port, "bad", f"{port} is not the modem AT port", check_id)
            return
        result = {}
        threading.Thread(target=lambda: result.update(probe=probe_at_port(port)), daemon=True).start()
        self._wait_background_check(result, port, check_id)

    ##
    # @brief Wait for the AT port check result and show it.
    #
    # @param [in] result Result filled in by the background thread
    # @param [in] port COM port, e.g. COM5
    # @param [in] check_id ID of the check, an older check is ignored
    #
    # @return None
    ##
    def _wait_background_check(self, result, port, check_id):
        if "probe" not in result:
            self.after(100, self._wait_background_check, result, port, check_id)
            return
        is_at_port, reason = result["probe"]
        self._set_port_state(port, "ok" if is_at_port else "bad", reason, check_id)

    ##
    # @brief Store the AT port check state and update the view.
    #
    # @param [in] port COM port, e.g. COM5
    # @param [in] status none, checking, ok or bad
    # @param [in] reason Reason the port cannot be used
    # @param [in] check_id ID of the check (an older check is ignored), or None
    #
    # @return None
    ##
    def _set_port_state(self, port, status, reason, check_id=None):
        if check_id is not None and check_id != self.port_check_id:
            return  # A newer selection replaced this check
        self.port_state = (port, status, reason)
        self._update_port_desc()
        self._update_start_state()

    ##
    # @brief Check that the selected port is the modem AT port.
    #
    # @return True if the selected port is the modem AT port, otherwise False
    ##
    def _port_ready(self):
        port, status, _ = self.port_state
        return status == "ok" and port == self.port_var.get().strip()

    ##
    # @brief Show the AT port check state below the COM port box.
    #
    # @return None
    ##
    def _update_port_desc(self):
        port, status, _ = self.port_state
        desc = re.sub(r"\s*\(COM\d+\)$", "", self.ports.get(port, "")) or port
        if status == "none":
            text, color = ("No COM ports detected" if not self.ports else "Select the modem AT port"), COLORS["muted"]
        elif status == "checking":
            text, color = f"… {desc} - checking for an AT port", COLORS["muted"]
        elif status == "ok":
            text, color = f"✓ {desc} - AT port", COLORS["pass"]
        else:
            text, color = f"✗ {desc} - not an AT port", COLORS["fail"]
        self.port_desc.configure(text=text, foreground=color)

    ##
    # @brief Update the SIM length counters (n/length) and the pre-check.
    #
    # @return None
    ##
    def _update_lengths(self):
        for label, (counter, var, length) in self.length_labels.items():
            current = len(var.get().strip())
            if self._sim_ok(label):
                color = COLORS["pass"]
            elif current > length or (current and not self._sim_chars_ok(label)):
                color = COLORS["fail"]
            else:
                color = COLORS["muted"]
            counter.configure(text=f"{current}/{length}", foreground=color)
        if hasattr(self, "precheck_chips"):
            self._update_start_state()

    ##
    # @brief Check the characters of a SIM value: IMEI / IMSI digits only, the CCID may end with a hex digit.
    #
    # @param [in] label SIM CCID, SIM IMEI or SIM IMSI
    #
    # @return True if the characters are valid, otherwise False
    ##
    def _sim_chars_ok(self, label):
        value = self.length_labels[label][1].get().strip()
        # IMEI / IMSI are numeric, the CCID may end with a hex digit (e.g. ...614F)
        return value.isdigit() if label != "SIM CCID" else value.isalnum()

    ##
    # @brief Check the length and characters of a SIM value.
    #
    # @param [in] label SIM CCID, SIM IMEI or SIM IMSI
    #
    # @return True if the SIM value is valid, otherwise False
    ##
    def _sim_ok(self, label):
        _, var, length = self.length_labels[label]
        return len(var.get().strip()) == length and self._sim_chars_ok(label)

    ##
    # @brief Check the inputs before a run.
    #
    # @return List of problems, empty if the pre-test can start
    ##
    def _validate(self):
        problems = []
        if self._jira_module() is None:
            problems.append(f"JIRA ID must contain {', '.join(JIRA_PROJECTS)} (e.g. SW9X07-10).")
        if not self.port_var.get().strip():
            problems.append("Select a COM port.")
        for label, (_, var, length) in self.length_labels.items():
            if not self._sim_ok(label):
                kind = "digits" if label != "SIM CCID" else "characters"
                problems.append(f"{label} must be {length} {kind} (currently {len(var.get().strip())}).")
        if self.online_var.get():
            if not self.is_internet_ok:
                problems.append("Online mode: no internet connection - press ⟳ next to Online to check again.")
            elif self._token_check()[0] != "ok":
                problems.append(f"Online mode: API token - {self._token_check()[1]}.")
            elif self._topic_check()[0] != "ok":
                problems.append(f"Online mode: JIRA ID - {self._topic_check()[1]}.")
            else:
                pass    # Online checks OK
        elif self.excel.path is None or self.excel.error:
            problems.append(self.excel.error or "No Excel file found.")
        else:
            pass    # Excel sheet OK
        return problems

    # Run control

    ##
    # @brief Start Pre-Test: check the inputs and the COM port, then run the pre-test.
    #
    # @return None
    ##
    def start_test(self):
        if self.is_running:
            return
        self.excel = detect_excel()
        problems = self._validate()
        if problems:
            self.reload_excel()
            messagebox.showwarning("Cannot start", "\n".join(problems))
            return

        # Make sure the selected port is the modem AT port before starting
        port = self.port_var.get().strip()
        self.is_checking_port = True
        self._set_running_ui(True)
        self.stop_button.configure(state="disabled")
        self._set_status("Checking COM port…", f"Sending AT to {port}")
        result = {}
        threading.Thread(target=lambda: result.update(probe=probe_at_port(port)), daemon=True).start()
        self._wait_port_check(result, port)

    ##
    # @brief Wait for the COM port check of Start, then download the Excel sheet (online mode) or run the
    #        pre-test.
    #
    # @param [in] result Result filled in by the background thread
    # @param [in] port COM port, e.g. COM5
    #
    # @return None
    ##
    def _wait_port_check(self, result, port):
        if "probe" not in result:
            self.after(100, self._wait_port_check, result, port)
            return
        self.is_checking_port = False
        self._set_running_ui(False)
        is_at_port, reason = result["probe"]
        self._set_port_state(port, "ok" if is_at_port else "bad", reason)
        if not is_at_port:
            description = self.ports.get(port, "")
            port_name = f"{port} ({description})" if description else port
            self._set_status("Not an AT port", reason, COLORS["fail"])
            self._console_write(f"[GUI] {reason}\n", "error")
            messagebox.showerror("COM port is not an AT port",
                                 f"{port_name} is not responding as a modem AT port.\n\n{reason}\n\n"
                                 "Select the modem's AT port (e.g. \"... USB AT Port\") and press Start again.")
            return
        if self.online_var.get():
            self._download_and_launch()
            return
        self._launch_test()

    ##
    # @brief Run the pre-test of the module selected by the JIRA ID in a child process.
    #
    # @return None
    ##
    def _launch_test(self):
        self._hide_search()

        # Switch to the module selected by the JIRA ID
        self.jira_id = self.jira_var.get().strip().upper()
        self.excel_file_label.configure(text=self.excel.path.name)
        self.select_module(module_from_jira(self.jira_id))

        args = [self.profile.key, self.username, self.port_var.get().strip(), self.ccid_var.get().strip(),
                self.imei_var.get().strip(), self.imsi_var.get().strip()]
        # Run information for the PDF report
        self.run_info = {"module": f"{self.profile.key} ({self.profile.technology})", "jira": self.jira_id,
                         "ccid": args[3], "imei": args[4], "imsi": args[5], "firmware": RE_ATTACHMENT_ID.sub("", self.excel.path.stem),
                         "firmware_version": read_excel_values(self.excel.path, [("", "subversion")])["subversion"],
                         "mode": "Online" if self.online_var.get() else "Offline"}
        self.run_token = self.token_var.get().strip() if self.online_var.get() else ""
        if getattr(sys, "frozen", False):
            cmd = [sys.executable, WORKER_FLAG] + args
        else:
            cmd = [sys.executable, "-u", str(Path(__file__).resolve()), WORKER_FLAG] + args
        env = dict(os.environ, PYTHONUNBUFFERED="1", PYTHONIOENCODING="utf-8")
        flags = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0

        self._populate_table()
        self._reset_progress()
        self.clear_console()
        self.log_file_path = None
        self.log_path_label.configure(text="")
        self.download_button.configure(state="disabled")
        self._delete_old_logs()
        try:
            self.process = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                                            stdin=subprocess.DEVNULL, cwd=str(app_directory()), env=env,
                                            text=True, encoding="utf-8", errors="replace",
                                            creationflags=flags)
        except OSError as e:
            messagebox.showerror("Cannot start", f"Failed to start the pre-test process:\n{e}")
            return

        self.is_running = True
        self.is_aborted = False
        self.start_time = time.monotonic()
        self.run_started_at = time.time()
        self._set_running_ui(True)
        self._console_write(f"[GUI] Starting {self.profile.key} pre-test on {args[2]} "
                            f"(JIRA ID = {self.jira_id}, user = {self.username})\n", "gui")
        self._set_status("Starting…", f"Opening {args[2]} and reading build parameters")
        self._highlight_running(1)
        threading.Thread(target=self._reader, args=(self.process,), daemon=True).start()

    ##
    # @brief Stop the running pre-test (after the user confirms).
    #
    # @return None
    ##
    def stop_test(self):
        if not self.is_running or self.process is None:
            return
        if not messagebox.askyesno("Stop pre-test", "Stop the running pre-test?\nThe partial log file is kept."):
            return
        self.is_aborted = True
        self.process.terminate()

    ##
    # @brief Read the output of the pre-test process into the output queue (background thread).
    #
    # @param [in] process Pre-test process
    #
    # @return None
    ##
    def _reader(self, process):
        for line in process.stdout:
            self.output_queue.put(line.rstrip("\n"))
        process.wait()
        self.output_queue.put(None)

    ##
    # @brief Handle the queued output lines of the pre-test process (every 50 ms).
    #
    # @return None
    ##
    def _drain_output(self):
        try:
            for _ in range(500):
                line = self.output_queue.get_nowait()
                if line is None:
                    self._on_finished()
                    break
                self._handle_line(line)
        except queue.Empty:
            pass
        self.after(50, self._drain_output)

    ##
    # @brief Handle one output line of the pre-test process: console, table, progress and summary.
    #
    # @param [in] line Output line
    #
    # @return None
    ##
    def _handle_line(self, line):
        if line.startswith(MARKER_LOG_FILE):
            self.log_file_path = Path(line[len(MARKER_LOG_FILE):].strip())
            self.log_path_label.configure(text=self.log_file_path.name)
            self.download_button.configure(state="normal")
            return
        if line == MARKER_EXIT:
            return

        stripped = line.strip()
        self._console_write(line + "\n", self._console_tag(stripped))

        match = RE_INDEX.match(stripped)
        if match:
            self.is_in_receive = False
            self.current = self.records.get(int(match.group(1)))
            self.progress.configure(maximum=int(match.group(2)))
            return

        if self.current is not None:
            # Response lines until the Expected line
            if self.is_in_receive:
                if RE_EXPECTED.match(line):
                    self.is_in_receive = False
                else:
                    self.current.receive += line + "\n"
                    return

            # Command sent / first response line
            send_match = RE_SEND.match(line)
            receive_match = RE_RECEIVE.match(line)
            if send_match:
                self.current.send = send_match.group(1)
            elif receive_match:
                self.current.receive = receive_match.group(1) + "\n"
                self.is_in_receive = True
            else:
                pass    # Other line of the test case

            # Expected response / result
            expected_match = RE_EXPECTED.match(line)
            result_match = RE_RESULT.match(stripped)
            if expected_match:
                self.current.expected = expected_match.group(1)
            elif result_match:
                self._on_result(self.current, result_match.group(1))
                self.current = None
            else:
                pass    # Other line of the test case
            return

        # Summary and errors
        if stripped.startswith("Test Summary"):
            self.is_summary_seen = True
        elif stripped.startswith("Cannot open") or stripped.startswith("Error") or "Invalid SIM" in stripped:
            self.error_lines.append(stripped)
        else:
            pass    # Other output line

    ##
    # @brief Show the result of one command in the table, the progress and the status.
    #
    # @param [in] record Test record of the command
    # @param [in] result PASS, FAIL or UNSUPPORTED
    #
    # @return None
    ##
    def _on_result(self, record, result):
        record.result = result
        self.counts[result] += 1
        iid = str(record.index)
        values = list(self.tree.item(iid, "values"))
        values[4] = record.expected
        values[5] = {"PASS": "✓ PASS", "FAIL": "✗ FAIL", "UNSUPPORTED": "– SKIP"}[result]
        self.tree.item(iid, values=values, tags=(result,))
        if self.autoscroll_var.get():
            self.tree.see(iid)

        done = record.index
        total = int(float(self.progress.cget("maximum")))
        self.progress.configure(value=done)
        self.percent_label.configure(text=f"{int(done * 100 / total)} %")
        self._update_chips()

        if record.send == const.RESET_AT_COMMAND and self.profile.reset_delay:
            self.reset_deadline = time.monotonic() + self.profile.reset_delay
            self.progress.configure(style="Reset.Horizontal.TProgressbar")
            self._set_status(f"Modem resetting  ·  {done}/{total}", "")
        elif done < total:
            nxt = at_cmd_test_case_list[done]
            self._highlight_running(done + 1)
            self._set_status(f"Running  ·  {done}/{total} complete",
                             f"Sending {nxt.at_command}  ({nxt.test_case_id} [{nxt.at_cmd_index}])")
        else:
            self._set_status(f"Finishing  ·  {done}/{total} complete", "Writing summary")

    ##
    # @brief Mark the command that is running in the table.
    #
    # @param [in] number Command number
    #
    # @return None
    ##
    def _highlight_running(self, number):
        iid = str(number)
        if self.tree.exists(iid) and not self.tree.item(iid, "tags"):
            self.tree.item(iid, tags=("RUNNING",))
            if self.autoscroll_var.get():
                self.tree.see(iid)

    ##
    # @brief The pre-test process ended: show the final status, and post the result to JIRA (online mode).
    #
    # @return None
    ##
    def _on_finished(self):
        if not self.is_running:
            return
        self.is_running = False
        self.reset_deadline = None
        self.process = None
        self._set_running_ui(False)
        self.progress.configure(style="Run.Horizontal.TProgressbar")
        if self.log_file_path is None:
            self.log_file_path = self._find_run_log()
            if self.log_file_path is not None:
                self.log_path_label.configure(text=self.log_file_path.name)
                self.download_button.configure(state="normal")
        for iid in self.tree.get_children():
            if "RUNNING" in self.tree.item(iid, "tags"):
                self.tree.item(iid, tags=())

        passed, failed, skipped = self.counts["PASS"], self.counts["FAIL"], self.counts["UNSUPPORTED"]
        if self.is_aborted:
            self._set_status("Stopped by user", f"{passed} passed · {failed} failed before stopping", COLORS["warn"])
            self._console_write("[GUI] Pre-test stopped by user.\n", "gui")
        elif not self.is_summary_seen:
            reason = self.error_lines[-1] if self.error_lines else "The pre-test ended unexpectedly - see the log."
            self._set_status("Pre-test did not complete", reason, COLORS["fail"])
            self._console_write(f"[GUI] {reason}\n", "error")
        elif failed == 0:
            self._set_status("Completed  ·  ALL PASS",
                             f"{passed} passed" + (f", {skipped} skipped" if skipped else ""), COLORS["pass"])
        else:
            self._set_status(f"Completed  ·  {failed} FAILED",
                             f"{passed} passed, {failed} failed" + (f", {skipped} skipped" if skipped else ""),
                             COLORS["fail"])
            first_fail = next((r for r in self.records.values() if r.result == "FAIL"), None)
            if first_fail:
                self.tree.selection_set(str(first_fail.index))
                self.tree.see(str(first_fail.index))
        self._show_search()

        # Online mode: a completed pre-test is posted to the JIRA ticket
        if self.run_token and not self.is_aborted and self.is_summary_seen:
            self._post_result_to_jira()

    ##
    # @brief Enable or disable the inputs while a pre-test or a check is running.
    #
    # @param [in] is_running True while running
    #
    # @return None
    ##
    def _set_running_ui(self, is_running):
        state = "disabled" if is_running else "normal"
        self._update_start_state()
        self.stop_button.configure(state="normal" if is_running else "disabled")
        self.reset_button.configure(state=state)
        self.upload_button.configure(state=state)
        for widget in self.input_widgets:
            widget.configure(state=state)

    ##
    # @brief Reset the progress, the counts and the run state for a new run.
    #
    # @return None
    ##
    def _reset_progress(self):
        self.counts = {"PASS": 0, "FAIL": 0, "UNSUPPORTED": 0}
        self.is_summary_seen = False
        self.error_lines = []
        self.current = None
        self.is_in_receive = False
        self.start_time = None
        self.progress.configure(value=0, maximum=len(at_cmd_test_case_list), style="Run.Horizontal.TProgressbar")
        self.percent_label.configure(text="0 %")
        self.chips["time"].configure(text="00:00")
        self._update_chips()
        if not self._excel_ready():
            self._set_status("No Excel sheet", "Upload the Jenkins Build Parameter Excel sheet to enable Start",
                             COLORS["fail"])
        elif self.profile is None:
            self._set_status("Ready", "Fill in the COM port and SIM information, then press Start")
        else:
            self._set_status("Ready", f"{self.profile.key} selected  ·  {len(at_cmd_test_case_list)} test commands")

    ##
    # @brief Show the PASS / FAIL / SKIP counts.
    #
    # @return None
    ##
    def _update_chips(self):
        self.chips["PASS"].configure(text=f"PASS {self.counts['PASS']}")
        self.chips["FAIL"].configure(text=f"FAIL {self.counts['FAIL']}")
        self.chips["UNSUPPORTED"].configure(text=f"SKIP {self.counts['UNSUPPORTED']}")

    ##
    # @brief Show the status and its detail.
    #
    # @param [in] status Status text
    # @param [in] detail Detail text
    # @param [in] color Status color, or None for the default text color
    #
    # @return None
    ##
    def _set_status(self, status, detail, color=None):
        self.status_label.configure(text=status, foreground=color or COLORS["text"])
        self.detail_label.configure(text=detail)

    ##
    # @brief Every 500 ms: update the run time and the modem reset countdown, and reload the Excel sheet if
    #        it changed.
    #
    # @return None
    ##
    def _tick(self):
        if not self.is_running and not self.is_checking_port and not self.is_downloading:
            current = next(app_directory().glob("*.xlsx"), None)
            if current != self.excel.path:
                self.reset_view()
        if self.is_running and self.start_time is not None:
            elapsed = int(time.monotonic() - self.start_time)
            self.chips["time"].configure(text=f"{elapsed // 60:02d}:{elapsed % 60:02d}")
        if self.is_running and self.reset_deadline is not None:
            remaining = self.reset_deadline - time.monotonic()
            if remaining > 0:
                self.detail_label.configure(
                    text=f"AT+CFUN=1,1 sent - waiting {int(remaining) + 1} s for the modem to reboot and reconnect")
            else:
                self.reset_deadline = None
                self.progress.configure(style="Run.Horizontal.TProgressbar")
                done = int(float(self.progress.cget("value")))
                total = int(float(self.progress.cget("maximum")))
                if done < total:
                    nxt = at_cmd_test_case_list[done]
                    self._highlight_running(done + 1)
                    self._set_status(f"Running  ·  {done}/{total} complete",
                                     f"Reconnected - sending {nxt.at_command}")
        self.after(500, self._tick)

    # Details & console

    ##
    # @brief Stretch the table columns to the width of the table.
    #
    # @param [in] event Tk configure event of the table
    #
    # @return None
    ##
    def _fit_columns(self, event):
        # Size the AT Command / Expected columns so all columns fit the table width,
        # otherwise the Result column is cut off at high display scaling
        stretch = [column for column in self.tree["columns"] if self.tree.column(column, "stretch")]
        total = sum(int(self.tree.column(column, "width")) for column in self.tree["columns"])
        extra = event.width - total
        if not stretch or extra == 0:
            return
        for index, column in enumerate(stretch):
            share = extra // len(stretch) + (1 if index < extra % len(stretch) else 0)
            width = int(self.tree.column(column, "width"))
            self.tree.column(column, width=max(px(100), width + share))

    ##
    # @brief Draw the column lines of the table.
    #
    # @return None
    ##
    def _draw_column_lines(self):
        columns = self.tree["columns"]
        children = self.tree.get_children()
        # Start below the headings, so the heading resize handles stay usable
        first_row = self.tree.bbox(children[0]) if children else None
        top = first_row[1] if first_row else 26
        x = 0
        for index, column in enumerate(columns[:-1]):
            x += int(self.tree.column(column, "width"))
            if index == len(self.column_lines):
                self.column_lines.append(tk.Frame(self.tree, width=1, bg=COLUMN_LINE_COLOR))
            self.column_lines[index].place(x=x - 1, y=top, width=1, relheight=1, height=-top)

    ##
    # @brief Show the details of the selected command below the table.
    #
    # @return None
    ##
    def _show_selected_detail(self):
        selection = self.tree.selection()
        if not selection:
            return
        record = self.records.get(int(selection[0]))
        values = self.tree.item(selection[0], "values")
        lines = [line.strip() for line in record.receive.splitlines() if line.strip()] if record else []
        receive = " ⏎ ".join(lines) or "-"
        result_tag = record.result if record and record.result else "label"
        self.detail_text.configure(state="normal")
        self.detail_text.delete("1.0", "end")
        self.detail_text.insert("end", f"{values[1]} [{values[2]}]", "title")
        self.detail_text.insert("end", "   Result: ", "label")
        self.detail_text.insert("end", values[5] or "not run", result_tag)
        for label, value in (("Send     : ", values[3]), ("Expected : ", values[4]), ("Receive  : ", receive)):
            self.detail_text.insert("end", "\n" + label, "label")
            self.detail_text.insert("end", value)
        self.detail_text.configure(state="disabled")

    ##
    # @brief Console color tag of an output line.
    #
    # @param [in] stripped Output line without the surrounding spaces
    #
    # @return Tag name, or None for normal text
    ##
    @staticmethod
    def _console_tag(stripped):
        if stripped.endswith(": PASS"):
            return "PASS"
        if stripped.endswith(": FAIL") or stripped.startswith("FAIL :"):
            return "FAIL"
        if stripped.endswith(": UNSUPPORTED"):
            return "UNSUPPORTED"
        if RE_INDEX.match(stripped):
            return "index"
        if stripped.startswith("-----"):
            return "divider"
        if stripped.startswith(("Cannot open", "Error", "Invalid SIM", "Traceback")):
            return "error"
        if stripped.startswith("[Port]"):
            return "gui"
        if stripped.startswith(("Test Summary", "Total :")):
            return "summary"
        return None

    ##
    # @brief Add text to the live log console.
    #
    # @param [in] text Text to add
    # @param [in] tag Color tag, or None
    #
    # @return None
    ##
    def _console_write(self, text, tag=None):
        self.console.configure(state="normal")
        self.console.insert("end", text, (tag,) if tag else ())
        self.console.configure(state="disabled")
        if self.autoscroll_var.get():
            self.console.see("end")

    ##
    # @brief Clear the live log console.
    #
    # @return None
    ##
    def clear_console(self):
        self.console.configure(state="normal")
        self.console.delete("1.0", "end")
        self.console.configure(state="disabled")

    ##
    # @brief Save a copy of the current run's log file to a location chosen by the user.
    #
    # @return None
    ##
    def download_log(self):
        if not self.log_file_path or not self.log_file_path.exists():
            messagebox.showwarning("Download log", "The log file of this run is no longer available.")
            self.download_button.configure(state="disabled")
            return
        target = filedialog.asksaveasfilename(title="Save pre-test log file",
                                              initialfile=self._result_file_name(".txt"),
                                              defaultextension=".txt",
                                              filetypes=[("Text file", "*.txt"), ("All files", "*.*")])
        if not target:
            return
        try:
            shutil.copyfile(self.log_file_path, target)
        except OSError as e:
            messagebox.showerror("Download log", f"Cannot save the log file:\n{e}")
            return
        self._console_write(f"[GUI] Log file saved to {target}\n", "gui")

    ##
    # @brief File name for a saved log / report: EaglePretest_Report_<date>_<time>,
    #        same date and time (run start) for both.
    #
    # @param [in] extension File extension, e.g. .pdf
    #
    # @return File name with the given extension
    ##
    def _result_file_name(self, extension):
        return f"EaglePretest_Report_{time.strftime('%Y%m%d_%H%M%S', time.localtime(self.run_started_at))}{extension}"

    ##
    # @brief Download Result: choose the log (.txt) or the report (.pdf).
    #
    # @return None
    ##
    def _show_download_menu(self):
        # The report is written from the finished run
        self.download_menu.entryconfigure(1, state="disabled" if self.is_running else "normal")
        x = self.download_button.winfo_rootx()
        y = self.download_button.winfo_rooty() + self.download_button.winfo_height()
        try:
            self.download_menu.tk_popup(x, y)
        finally:
            self.download_menu.grab_release()

    ##
    # @brief Save a PDF report of the run: run information and the Test Cases
    #        table with the actual modem response of every command.
    #
    # @return None
    ##
    def download_report(self):
        if self.is_running or not self.records:
            return
        target = filedialog.asksaveasfilename(title="Save pre-test report",
                                              initialfile=self._result_file_name(".pdf"),
                                              defaultextension=".pdf",
                                              filetypes=[("PDF file", "*.pdf"), ("All files", "*.*")])
        if not target:
            return
        try:
            self._write_report(target)
        except Exception as e:
            messagebox.showerror("Download report", f"Cannot save the report:\n{e}")
            return
        self._console_write(f"[GUI] Report saved to {target}\n", "gui")

    ##
    # @brief Write the PDF report of the finished run.
    #
    # @param [in] target PDF file to write
    #
    # @return None
    ##
    def _write_report(self, target):
        passed, failed, skipped = self.counts["PASS"], self.counts["FAIL"], self.counts["UNSUPPORTED"]
        summary = f"{passed} passed, {failed} failed, {skipped} skipped of {len(self.records)} commands"
        if self.is_aborted:
            summary += " (stopped by user)"
        elif not self.is_summary_seen:
            summary += " (did not complete)"
        else:
            pass    # Completed
        run = self.run_info
        left = [("Reporter", self.username), ("Mode", run["mode"]),
                ("Date", time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(self.run_started_at))),
                ("Duration", self.chips["time"].cget("text")), ("Module", run["module"]), ("Summary", summary)]
        right = [("Firmware", run["firmware"]), ("Firmware Version", run["firmware_version"]),
                 ("SIM CCID", run["ccid"]), ("SIM IMSI", run["imsi"]), ("SIM IMEI", run["imei"])]
        rows = []
        for number, record in sorted(self.records.items()):
            values = self.tree.item(str(number), "values")
            result = record.result or ("N/A" if "NA" in self.tree.item(str(number), "tags") else "")
            rows.append((number, values[1], values[2], record.send, record.expected, record.receive, result))
        write_report_pdf(target, run["jira"], left, right, rows)

    ##
    # @brief Online mode: when the pre-test completes, comment on the JIRA ticket
    #        with the result and attach the PDF report and the log.
    #
    # @return None
    ##
    def _post_result_to_jira(self):
        jira_id = self.run_info["jira"]
        failed = self.counts["FAIL"]
        folder = tempfile.TemporaryDirectory()
        report = Path(folder.name) / self._result_file_name(".pdf")
        log = Path(folder.name) / self._result_file_name(".txt")
        try:
            self._write_report(report)
            files = [report]
            if self.log_file_path is not None and self.log_file_path.exists():
                shutil.copyfile(self.log_file_path, log)
                files.append(log)
        except Exception as e:
            folder.cleanup()
            self._jira_post_failed(jira_id, f"Cannot create the report: {e}")
            return

        result = "All pass, no fail" if failed == 0 else f"Not all pass, {failed} fail"
        lines = ["*EaglePretest*", "AT command pre-test has been completed. Here is the result:",
                 f"Result → {result}", f"Report → [^{report.name}]"]
        if log in files:
            lines.append(f"Log → [^{log.name}]")
        comment = "\n\n".join(lines)

        self._console_write(f"[GUI] Posting the result to {JIRA_BROWSE_URL + jira_id}\n", "gui")
        posted = {}

        ##
        # @brief Post the result to JIRA (background thread).
        #
        # @return None
        ##
        def post():
            try:
                jira_post_result(jira_id, self.username, self.run_token, files, comment)
                posted["ok"] = True
            except JiraError as e:
                posted["error"] = str(e)
            except Exception as e:
                posted["error"] = f"Posting failed: {e}"

        threading.Thread(target=post, daemon=True).start()
        self._wait_jira_post(posted, jira_id, folder)

    ##
    # @brief Wait for the result to be posted, then report it.
    #
    # @param [in] posted Post result filled in by the background thread
    # @param [in] jira_id JIRA ID, e.g. SDX35-112
    # @param [in] folder Temporary folder, deleted when done
    #
    # @return None
    ##
    def _wait_jira_post(self, posted, jira_id, folder):
        if not posted:
            self.after(100, self._wait_jira_post, posted, jira_id, folder)
            return
        folder.cleanup()
        if "error" in posted:
            self._jira_post_failed(jira_id, posted["error"])
            return
        self._console_write(f"[GUI] Result posted as a comment on {JIRA_BROWSE_URL + jira_id}\n", "gui")
        self.detail_label.configure(text=f"{self.detail_label.cget('text')}  ·  Result posted to {jira_id}")

    ##
    # @brief Report that the result could not be posted to JIRA.
    #
    # @param [in] jira_id JIRA ID, e.g. SDX35-112
    # @param [in] reason Reason shown to the user
    #
    # @return None
    ##
    def _jira_post_failed(self, jira_id, reason):
        self._console_write(f"[GUI] {reason}\n", "error")
        self._set_connection_help(True)
        messagebox.showerror("Cannot post to JIRA", f"The result was not posted to {jira_id}.\n\n{reason}\n\n"
                                                    "The report and log can still be saved with Download Result.")

    ##
    # @brief Delete the logs of previous runs so only the new run's log is stored.
    #
    # @return None
    ##
    def _clear_previous_session(self):
        # A new session starts clean: the sheet must be uploaded again and old logs are removed
        deleted, failed = [], []
        for path in stored_excel_files():
            try:
                path.unlink()
                deleted.append(path.name)
            except OSError as e:
                failed.append(f"{path.name} ({e.strerror})")
        for name in deleted:
            self._console_write(f"[GUI] Deleted Excel sheet from the previous session: {name}\n", "gui")
        for item in failed:
            self._console_write(f"[GUI] Could not delete Excel sheet {item}\n", "error")
        self._delete_old_logs()

    ##
    # @brief Delete the log files of previous runs.
    #
    # @return None
    ##
    def _delete_old_logs(self):
        deleted, failed = 0, []
        for path in app_directory().glob("Eagle_Wireless_SW_AT_Pretest_Log_*.txt"):
            try:
                path.unlink()
                deleted += 1
            except OSError as e:
                failed.append(f"{path.name} ({e.strerror})")
        if deleted:
            self._console_write(f"[GUI] Deleted {deleted} previous log file(s)\n", "gui")
        for item in failed:
            self._console_write(f"[GUI] Could not delete previous log {item}\n", "error")

    ##
    # @brief Find the log written by a run that ended without reporting it
    #        (e.g. stopped by the user).
    #
    # @return Path of the log, or None if not found
    ##
    def _find_run_log(self):
        logs = [path for path in app_directory().glob("Eagle_Wireless_SW_AT_Pretest_Log_*.txt")
                if path.stat().st_mtime >= self.run_started_at - 1]
        return max(logs, key=lambda path: path.stat().st_mtime, default=None)

    ##
    # @brief Close the window (a running pre-test is stopped first, after the user confirms).
    #
    # @return None
    ##
    def _on_close(self):
        if self.is_running:
            if not messagebox.askyesno("Pre-test running", "A pre-test is running. Stop it and exit?"):
                return
            self.is_aborted = True
            if self.process is not None:
                self.process.terminate()
        self.destroy()

##
# @brief Main application: run the worker (child process), or show the login page and then the pre-test tool.
#
# @return None
##
def main_application():

    # Child process: run one module pre-test
    if len(sys.argv) > 1 and sys.argv[1] == WORKER_FLAG:
        run_worker(sys.argv[2:])
        return

    # Show the window's own icon on the taskbar instead of the python.exe icon
    if os.name == "nt":
        try:
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("EagleWireless.ATPretest.GUI")
        except Exception:
            pass

    # Login page, then the pre-test tool
    enable_dpi_awareness()
    login = LoginWindow()
    login.mainloop()
    if login.is_authenticated:
        PretestApp(login.username).mainloop()

if __name__ == "__main__":
    main_application()
