##
# @file fake_worker.py
# @brief Stand-in for the pre-test worker process (no modem). Prints the same
#        console output as the module scripts (e.g. sdx35.py), so the GUI
#        output parser can be tested.
#
# Usage: python fake_worker.py <module> <scenario>
#   pass         every command passes
#   fail:N,M     commands N and M fail
#   crash        cannot open the port, no summary
#   stop:N       prints N commands, then waits to be stopped
##

import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

import constants as const
from test_case import at_cmd_test_case_list
from gui import MODULE_PROFILES, MARKER_LOG_FILE, MARKER_EXIT

module_key, scenario = sys.argv[1], sys.argv[2]
profile = MODULE_PROFILES[module_key]
total = len(at_cmd_test_case_list)
fails = {int(n) for n in scenario.split(":")[1].split(",")} if scenario.startswith("fail:") else set()
stop_after = int(scenario.split(":")[1]) if scenario.startswith("stop:") else None
log = Path.cwd() / "Eagle_Wireless_SW_AT_Pretest_Log_20261022_140000.txt"


def finish():
    log.write_text("fake log\n")
    print(MARKER_LOG_FILE + str(log))
    print(MARKER_EXIT)
    sys.stdout.flush()


print(const.STARTUP_MESSAGE_STRING)
if scenario == "crash":
    print("Cannot open COM99: could not open port 'COM99': FileNotFoundError")
    finish()
    sys.exit(0)

print(f"[ Extraction from Excel File : {log.parent / 'sheet.xlsx'} ]\n")
print(const.LINE_DIVISOR_STRING)
passed = failed = 0
for number, test_case in enumerate(at_cmd_test_case_list, start=1):
    name = f"{test_case.test_case_id} [{test_case.at_cmd_index}]"
    print(f"[{number}/{total}]\n")
    if profile.is_unsupported(test_case):
        result, response = "UNSUPPORTED", "-"
    else:
        result = "FAIL" if number in fails else "PASS"
        # Multi-line modem response, like the real response buffer
        response = f"{test_case.at_command}\n+RESP: {number}\n\n{'ERROR' if result == 'FAIL' else 'OK'}\n"
        passed += result == "PASS"
        failed += result == "FAIL"
    print(f"Test Case : {name}\nSend      : {test_case.at_command}\nReceive   : {response}\n"
          f"Expected  : {test_case.expected_response}\nResult    : {name} : {result}\n{const.LINE_DIVISOR_STRING}\n")
    sys.stdout.flush()
    if number == stop_after:
        time.sleep(60)      # Stopped by the GUI before this ends

print(f"Test Summary\nTotal : {total} | Pass : {passed} | Fail : {failed}\n")
for number in sorted(fails):
    test_case = at_cmd_test_case_list[number - 1]
    print(f"FAIL : {test_case.test_case_id} [{test_case.at_cmd_index}] : {test_case.at_command}")
finish()
