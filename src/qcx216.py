##
# @file main.py
# @brief LTE AT command pre-test automation tool for QCX216.
#
# Executes a predefined list of LTE AT command test cases against a modem
# through a serial port. The application records the modem responses,
# displays the test progress on the console, and generates a timestamped
# log file containing the test results.
#
# Features:
# - Executes predefined LTE AT commands
# - Logs all modem responses
# - Generates timestamped log files
# - Supports automatic modem reset handling (AT+CFUN=1,1)
#
# @author Anas Akasyah Bin Abdul Patas
# @version 1.0.0
# @date 2026-10-26
# @copyright Eagle Wireless
##

# Libraries
import re
import sys
import time
import serial
import logging
from pathlib import Path
from datetime import datetime
from openpyxl import load_workbook
from test_case import at_cmd_test_case_list
from test_case import at_cmd_general_info_list
from test_case import TestCaseIndex
import constants as const

general_info = at_cmd_general_info_list[0]

##
# @brief Configure log file.
#
# @return Created log file path
##
def log_configure():
    if getattr(sys, "frozen", False):
        script_dir = Path(sys.executable).parent
    else:
        script_dir = Path(__file__).parent

    log_file = script_dir / f"Eagle_Wireless_SW_AT_Pretest_Log_{datetime.now():%Y%m%d_%H%M%S}.txt"

    logging.basicConfig(
        filename=log_file,
        level=logging.INFO,
        format="%(asctime)s\r\n%(message)s",
        datefmt="%Y-%m-%d %H:%M:%S"
    )

    return log_file

##
# @brief Writes the log in a file.
#
# @param [in] text Text to write in log
#
# @return None
##
def log_write(text):
    logging.info(text)

# Function

##
# @brief Get SIM information via user input.
#
# @param [in] ccid SIM CCID
# @param [in] imei SIM IMEI
# @param [in] imsi SIM IMSI
#
# @return True if user input successful, otherwise False
##
def sim_get_user_input_information(ccid, imei, imsi):

    if len(ccid) != const.SIM_CCID_LEN or len(imei) != const.SIM_IMEI_LEN or len(imsi) != const.SIM_IMSI_LEN:
        print("Invalid SIM information length\n")
        return False

    # Map into general information list
    general_info.sim_ccid = ccid
    general_info.sim_imei = imei
    general_info.sim_imsi = imsi

    # SIM CCID
    at_cmd_test_case_list[TestCaseIndex.CD_05_001_1].expected_response = general_info.sim_ccid

    # SIM IMEI
    at_cmd_test_case_list[TestCaseIndex.CD_05_002_1].expected_response = general_info.sim_imei
    at_cmd_test_case_list[TestCaseIndex.CD_05_002_2].expected_response = general_info.sim_imei
    at_cmd_test_case_list[TestCaseIndex.CD_05_002_3].expected_response = general_info.sim_imei

    # SIM IMSI
    at_cmd_test_case_list[TestCaseIndex.CD_05_003_1].expected_response = general_info.sim_imsi

    return True

##
# @brief Read the jenkins value from the excel sheet.
#
# @param [in] excel_path Excel path
# @param [in] parameter_name Parameter name
#
# @return Jenkins parameter value
##
def read_jenkins_value(excel_path, parameter_name):

    workbook = load_workbook(excel_path, data_only=True)
    sheet = workbook.active

    header_row = 1

    # Find column indexes by header name
    headers = {
        cell.value: cell.column
        for cell in sheet[header_row]
        if cell.value is not None
    }

    param_col = headers["Jenkins Parameter"]
    value_col = headers["Value"]

    # Search parameter name
    for row in range(header_row + 1, sheet.max_row + 1):
        param = sheet.cell(row=row, column=param_col).value

        if param == parameter_name:
            jenkins_value = str(sheet.cell(row=row, column=value_col).value)
            workbook.close()
            return jenkins_value

    workbook.close()
    return "None"

##
# @brief Extract information from excel.
#
# @return True if extraction succeeds, otherwise False
##
def excel_extract_parameter():

    if getattr(sys, "frozen", False):
        script_dir = Path(sys.executable).parent
    else:
        script_dir = Path(__file__).parent

    excel_path = next(script_dir.glob("*.xlsx"), None)

    if excel_path is None:
        print("Error: No Excel file found.\n")
        return False

    if not excel_path.is_file():
        print(f"Error: Excel file not found: {excel_path}\n")
        return False

    # Empty general info list
    general_info.platform_name = const.PLATFORM_QCX216
    general_info.certversion = ""
    general_info.branch_name = ""
    general_info.sub_version = ""
    general_info.svn = ""
    general_info.version = ""
    general_info.version_info = ""
    general_info.modemapp = ""
    general_info.linuxapp = ""

    # Store excel sheet value in list
    general_info.certversion = read_jenkins_value(excel_path, "certversion")
    general_info.branch_name = read_jenkins_value(excel_path, "branch_name")
    general_info.sub_version = read_jenkins_value(excel_path, "subversion")
    general_info.svn = read_jenkins_value(excel_path, "svn")
    general_info.version = read_jenkins_value(excel_path, "version")
    general_info.version_info = read_jenkins_value(excel_path, "version_info")
    general_info.oem = read_jenkins_value(excel_path, "OEM")
    general_info.modemapp = read_jenkins_value(excel_path, "modemapp")
    general_info.linuxapp = read_jenkins_value(excel_path, "linuxapp")

    # QCX216 (QGMR) --> version + _OEM + _modemapp.modemapp
    if general_info.oem == "None":
        if general_info.modemapp == "None":
            general_info.version_qgmr = str(general_info.version)
        else:
            general_info.version_qgmr = str(general_info.version + "_" +  general_info.modemapp + "." + general_info.modemapp)
    elif general_info.oem in general_info.version:
        if general_info.modemapp == "None":
            general_info.version_qgmr = str(general_info.version)
        else:
            general_info.version_qgmr = str(general_info.version + "_" +  general_info.modemapp + "." + general_info.modemapp)
    else:
        if general_info.modemapp == "None":
            general_info.version_qgmr = str(general_info.version + "_" +  general_info.oem)
        else:
            general_info.version_qgmr = str(general_info.version + "_" +  general_info.oem + "_" +  general_info.modemapp + "." + general_info.modemapp)

    # Print & log
    print(f"[ Extraction from Excel File : {excel_path} ]\n")
    print(f"Platform     : {general_info.platform_name}\nBranch Name  : {general_info.branch_name}\nCert version : {general_info.certversion}\nSub-Version  : {general_info.sub_version}\nSVN          : {general_info.svn}\nVersion      : {general_info.version}\nVersion(QGMR): {general_info.version_qgmr}\nVersion Info : {general_info.version_info}\nSIM CCID     : {general_info.sim_ccid}\nSIM IMEI     : {general_info.sim_imei}\nSIM IMSI     : {general_info.sim_imsi}\n")
    print(const.LINE_DIVISOR_STRING)
    log_write(f"[ Extraction from Excel File : {excel_path} ]\nPlatform     : {general_info.platform_name}\nBranch Name  : {general_info.branch_name}\nCert version : {general_info.certversion}\nSub-Version  : {general_info.sub_version}\nSVN          : {general_info.svn}\nVersion      : {general_info.version}\nVersion(QGMR): {general_info.version_qgmr}\nVersion Info : {general_info.version_info}\nSIM CCID     : {general_info.sim_ccid}\nSIM IMEI     : {general_info.sim_imei}\nSIM IMSI     : {general_info.sim_imsi}\n")
    log_write(const.LINE_DIVISOR_STRING)

    # Check Manufacturer ID from Branch Name
    branch_name_lower_case = general_info.branch_name.lower()
    if "eagle" in branch_name_lower_case:
        general_info.branch_name = const.BRANCH_NAME_EAGLE
    elif "np" in branch_name_lower_case:
        general_info.branch_name = const.BRANCH_NAME_NETPRISMA
    else:
        general_info.branch_name = const.BRANCH_NAME_QUECTEL

    # Branch name
    at_cmd_test_case_list[TestCaseIndex.CD_01_001_1].expected_response = general_info.branch_name
    at_cmd_test_case_list[TestCaseIndex.CD_01_001_2].expected_response = general_info.branch_name

    # Certificate Version (ATI)
    at_cmd_test_case_list[TestCaseIndex.CD_01_003_1].expected_response = general_info.certversion
    at_cmd_test_case_list[TestCaseIndex.CD_01_003_2].expected_response = general_info.certversion
    at_cmd_test_case_list[TestCaseIndex.CD_01_003_3].expected_response = general_info.certversion
    at_cmd_test_case_list[TestCaseIndex.CD_01_003_4].expected_response = general_info.certversion

    # Version_OEM (QGMR)
    at_cmd_test_case_list[TestCaseIndex.CD_01_003_5].expected_response = general_info.version_qgmr

    # Sub-Version
    at_cmd_test_case_list[TestCaseIndex.CD_01_003_6].expected_response = general_info.sub_version

    # SVN
    at_cmd_test_case_list[TestCaseIndex.CD_02_001_1].expected_response = general_info.svn
    at_cmd_test_case_list[TestCaseIndex.CD_04_001_2].expected_response = general_info.svn

    # Check platform error
    at_cmd_test_case_list[TestCaseIndex.CD_04_002_1].expected_response = "OK"

    match = re.search(r'A(\d{2})', general_info.version)
    if match:
        at_cmd_test_case_list[TestCaseIndex.CD_04_003_13].expected_response = match.group(1)
        at_cmd_test_case_list[TestCaseIndex.CD_04_003_20].expected_response = match.group(1)

    return True

##
# @brief Sends an AT command and prints the modem response.
#
# @param [in] serial_port Serial Port
# @param [in] test_case Test case (current)
# @param [in] at_cmd_index AT command index
# @param [in] at_cmd_list AT command list
# @param [in] timeout Response timeout
#
# @return True if a response is available, otherwise False
##
def at_execute_command(serial_port, test_case, at_cmd_index, at_cmd_list, timeout):

    serial_port.reset_input_buffer()
    serial_port.write((test_case.at_command + "\r\n").encode())

    response_buffer = ""
    start_time = time.time()
    is_response_received = False
    is_response_the_same = False

    # Run serial port
    while time.time() - start_time < timeout:
        if serial_port.in_waiting:
            response_buffer += serial_port.read(serial_port.in_waiting).decode(errors="ignore")

            # Same response
            if test_case.expected_response in response_buffer:
                is_response_received = True
                break

        # Command delay
        time.sleep(const.COMMAND_DELAY_SEC)

    if is_response_received:
        if test_case.exact_match:
            lines = [line.strip() for line in response_buffer.splitlines() if line.strip()]

            if len(lines) >= const.AT_CMD_MIN_LINE_TO_PARSE:
                actual_response = lines[test_case.response_line]

                if test_case.has_prefix:
                    actual_response = actual_response.removeprefix(test_case.prefix_value).strip()

                if actual_response == test_case.expected_response:
                    is_response_the_same = True
            else:
                is_response_the_same = False
        else:
            is_response_the_same = True
    else:
        is_response_the_same = False

    if is_response_received and is_response_the_same:
        print(f"[{at_cmd_index}/{len(at_cmd_list)}]\n")
        print(f"Test Case : {test_case.test_case_id} [{test_case.at_cmd_index}]\nSend      : {test_case.at_command}\nReceive   : {response_buffer}\nExpected  : {test_case.expected_response}\nResult    : {test_case.test_case_id} [{test_case.at_cmd_index}] : PASS\n{const.LINE_DIVISOR_STRING}\n")
        log_write(f"[{at_cmd_index}/{len(at_cmd_list)}]\nTest Case : {test_case.test_case_id} [{test_case.at_cmd_index}]\nSend      : {test_case.at_command}\nReceive   : {response_buffer}Expected  : {test_case.expected_response}\nResult    : {test_case.test_case_id} [{test_case.at_cmd_index}] : PASS\n\n{const.LINE_DIVISOR_STRING}\n")
        return True
    else:
        print(f"[{at_cmd_index}/{len(at_cmd_list)}]\n")
        print(f"Test Case : {test_case.test_case_id} [{test_case.at_cmd_index}]\nSend      : {test_case.at_command}\nReceive   : {response_buffer}\nExpected  : {test_case.expected_response}\nResult    : {test_case.test_case_id} [{test_case.at_cmd_index}] : FAIL\n{const.LINE_DIVISOR_STRING}\n")
        log_write(f"[{at_cmd_index}/{len(at_cmd_list)}]\nTest Case : {test_case.test_case_id} [{test_case.at_cmd_index}]\nSend      : {test_case.at_command}\nReceive   : {response_buffer}Expected  : {test_case.expected_response}\nResult    : {test_case.test_case_id} [{test_case.at_cmd_index}] : FAIL\n\n{const.LINE_DIVISOR_STRING}\n")
        return False

##
# @brief Main pre-test application for MDM9X07.
#
# @param [in] username Username
# @param [in] port Serial port
# @param [in] sim_ccid CCID: Integrated Circuit Card Identifier
# @param [in] sim_imei IMEI: International Mobile Equipment Identity
# @param [in] sim_imsi IMSI: International Mobile Subscriber Identity
#
# @return None
##
def pretest_qcx216(username, port, sim_ccid, sim_imei, sim_imsi):
    try:
        pass_count = 0
        fail_count = 0
        failed_test_case_list = []

        # Configure log
        log_configure()

        # Startup message
        print(const.STARTUP_MESSAGE_STRING)
        print(f"Reporter : {username}\n\n")
        print(const.LINE_DIVISOR_STRING)
        log_write(const.STARTUP_MESSAGE_STRING)
        log_write(f"Reporter : {username}\n\n")
        log_write(const.LINE_DIVISOR_STRING)

        # SIM : User input
        if not sim_get_user_input_information(sim_ccid, sim_imei, sim_imsi):
            return

        # Excel : Extract
        if not excel_extract_parameter():
            return

        # Open Serial COM port
        with serial.Serial(port, const.BAUD_RATE, timeout=1) as serial_port:

            # Execute AT command
            for at_cmd_index, test_case in enumerate(at_cmd_test_case_list, start=1):
                result_ok = at_execute_command(serial_port, test_case, at_cmd_index, at_cmd_test_case_list, const.SEND_TIMEOUT)

                if result_ok:
                    pass_count += 1
                else:
                    fail_count += 1
                    failed_test_case_list.append(test_case)

                if test_case.at_command == const.RESET_AT_COMMAND:
                    serial_port.close()
                    time.sleep(const.RESET_DELAY_QCX216_SEC)
                    serial_port.open()
                    serial_port.reset_input_buffer()
                    serial_port.reset_output_buffer()
                else:
                    time.sleep(const.COMMAND_DELAY_SEC)

            print(f"Test Summary\nTotal : {at_cmd_index} | Pass : {pass_count} | Fail : {fail_count}\n")
            log_write(f"Test Summary\nTotal : {at_cmd_index} | Pass : {pass_count} | Fail : {fail_count}\n")

            # List all failed test cases
            for failed_test_case in failed_test_case_list:
                print(f"FAIL : {failed_test_case.test_case_id} [{failed_test_case.at_cmd_index}] : {failed_test_case.at_command}")
                log_write(f"FAIL : {failed_test_case.test_case_id} [{failed_test_case.at_cmd_index}] : {failed_test_case.at_command}")

    except serial.SerialException as e:
        print(f"Cannot open {port}: {e}")

if __name__ == "__main__":
    pretest_qcx216()
