from dataclasses import dataclass
from enum import IntEnum

@dataclass
class TestCaseIndex(IntEnum):
    CD_01_000_1             =   0
    CD_01_000_2             =   1
    CD_01_001_1             =   2
    CD_01_001_2             =   3
    CD_01_002_1             =   4
    CD_01_002_2             =   5
    CD_01_003_1             =   6
    CD_01_003_2             =   7
    CD_01_003_3             =   8
    CD_01_003_4             =   9
    CD_01_003_5             =   10
    CD_01_003_6             =   11
    CD_02_001_1             =   12
    CD_04_001_1             =   13
    CD_04_001_2             =   14
    CD_04_002_1             =   15
    CD_04_003_1             =   16
    CD_04_003_2             =   17
    CD_04_003_3             =   18
    CD_04_003_4             =   19
    CD_04_003_5             =   20
    CD_04_003_6             =   21
    CD_04_003_7             =   22
    CD_04_003_8             =   23
    CD_04_003_9             =   24
    CD_04_003_10            =   25
    CD_04_003_11            =   26
    CD_04_003_12            =   27
    CD_04_003_13            =   28
    CD_04_003_14            =   29
    CD_04_003_15            =   30
    CD_04_003_16            =   31
    CD_04_003_17            =   32
    CD_04_003_18            =   33
    CD_04_003_19            =   34
    CD_04_003_20            =   35
    CD_04_003_21            =   36
    CD_04_003_22            =   37
    CD_05_001_1             =   38
    CD_05_002_1             =   39
    CD_05_002_2             =   40
    CD_05_002_3             =   41
    CD_05_003_1             =   42
    CD_05_004_1             =   43

@dataclass
class ATCommonInfo:
    branch_name             :   str
    sub_version             :   str
    svn                     :   str
    version                 :   str
    version_info            :   str
    sim_ccid                :   str
    sim_imsi                :   str
    sim_imei                :   str
    oem                     :   str
    version_qgmr            :   str
    certversion             :   str
    platform_name           :   str
    modemapp                :   str
    linuxapp                :   str

at_cmd_general_info_list = [
    ATCommonInfo(
        branch_name         =   "R01_1075014_Eagle_73",
        sub_version         =   "V01",
        svn                 =   "02",
        version             =   "I4916QGLLGR01A02M04",
        version_info        =   "I4916QGLLGR01A02M04_ONT_A0.300.A0.300_V01",
        sim_ccid            =   "8960152241058850614F",
        sim_imsi            =   "502152107380061",
        sim_imei            =   "016580000000000",
        oem                 =   "ONT",
        version_qgmr        =   "I4916QGLLGR01A02M04_ONT",
        certversion         =   "I4916QGLLGR01A02M04",
        platform_name       =   "QCX216",
        modemapp            =   "A0.300",
        linuxapp            =   "1.100.100"
    )
]

@dataclass
class ATTestCase:
    test_case_id        :   str
    at_cmd_index        :   str
    at_command          :   str
    expected_response   :   str
    response_line       :   int     =   1
    prefix_value        :   str     =   None
    exact_match         :   bool    =   False
    has_prefix          :   bool    =   False

at_cmd_test_case_list = [

    ATTestCase(
        test_case_id        =   "Pretest_LTE_CD-01-000",
        at_cmd_index        =   "1",
        at_command          =   "AT",
        expected_response   =   "OK"
    ),

    ATTestCase(
        test_case_id        =   "Pretest_LTE_CD-01-000",
        at_cmd_index        =   "2",
        at_command          =   "ATE1",
        expected_response   =   "OK"
    ),

    # Pretest_LTE_CD-01-001

    ATTestCase(
        test_case_id        =   "Pretest_LTE_CD-01-001",
        at_cmd_index        =   "1",
        at_command          =   "AT+GMI",
        expected_response   =   "Eagle Wireless", #Branch Name
        exact_match         =   True
    ),

    ATTestCase(
        test_case_id        =   "Pretest_LTE_CD-01-001",
        at_cmd_index        =   "2",
        at_command          =   "AT+CGMI",
        expected_response   =   "Eagle Wireless", #Branch Name
        exact_match         =   True
    ),

    # Pretest_LTE_CD-01-002

    ATTestCase(
        test_case_id        =   "Pretest_LTE_CD-01-002",
        at_cmd_index        =   "1",
        at_command          =   "AT+GMM",
        expected_response   =   "OK",
    ),

    ATTestCase(
        test_case_id        =   "Pretest_LTE_CD-01-002",
        at_cmd_index        =   "2",
        at_command          =   "AT+CGMM",
        expected_response   =   "OK",
    ),

    # Pretest_LTE_CD-01-003

    ATTestCase(
        test_case_id        =   "Pretest_LTE_CD-01-003",
        at_cmd_index        =   "1",
        at_command          =   "AT+GMR",
        expected_response   =   "OK", #Version
    ),

    ATTestCase(
        test_case_id        =   "Pretest_LTE_CD-01-003",
        at_cmd_index        =   "2",
        at_command          =   "AT+CGMR",
        expected_response   =   "OK", #Version
    ),

    ATTestCase(
        test_case_id        =   "Pretest_LTE_CD-01-003",
        at_cmd_index        =   "3",
        at_command          =   "ATI",
        expected_response   =   "OK", #Version
        exact_match         =   True,
        has_prefix          =   True,
        prefix_value        =   "Revision: ", # Revision:
        response_line       =   3
    ),

    ATTestCase(
        test_case_id        =   "Pretest_LTE_CD-01-003",
        at_cmd_index        =   "4",
        at_command          =   "AT+CVERSION",
        expected_response   =   "OK", #Version
        exact_match         =   True,
        has_prefix          =   True,
        prefix_value        =   "VERSION: " # VERSION:
    ),

    ATTestCase(
        test_case_id        =   "Pretest_LTE_CD-01-003",
        at_cmd_index        =   "5",
        at_command          =   "AT+QGMR",
        expected_response   =   "OK", #Version QGMR
        exact_match         =   True
    ),

    ATTestCase(
        test_case_id        =   "Pretest_LTE_CD-01-003",
        at_cmd_index        =   "6",
        at_command          =   "AT+CSUB",
        expected_response   =   "OK" #Sub-Version
    ),

    # # Pretest_LTE_CD-02-001

    ATTestCase(
        test_case_id        =   "Pretest_LTE_CD-02-001",
        at_cmd_index        =   "1",
        at_command          =   "AT+EGMR=0,9",
        expected_response   =   "2" #SVN
    ),

    # Pretest_LTE_CD-04-001

    ATTestCase(
        test_case_id        =   "Pretest_LTE_CD-04-001",
        at_cmd_index        =   "1",
        at_command          =   "AT+QSVN=?",
        expected_response   =   "+QSVN: <svn>"
    ),

    ATTestCase(
        test_case_id        =   "Pretest_LTE_CD-04-001",
        at_cmd_index        =   "2",
        at_command          =   "AT+QSVN?",
        expected_response   =   "2" #SVN
    ),

    # Pretest_LTE_CD-04-002

    ATTestCase(
        test_case_id        =   "Pretest_LTE_CD-04-002",
        at_cmd_index        =   "1",
        at_command          =   "AT+QSVN=99",
        expected_response   =   "ERROR"
    ),

    # Pretest_LTE_CD-04-003

    ATTestCase(
        test_case_id        =   "Pretest_LTE_CD-04-003",
        at_cmd_index        =   "1",
        at_command          =   "AT+QSVN=\"13\"",
        expected_response   =   "OK"
    ),

    ATTestCase(
        test_case_id        =   "Pretest_LTE_CD-04-003",
        at_cmd_index        =   "2",
        at_command          =   "AT+CFUN=1,1",
        expected_response   =   "OK"
    ),

    ATTestCase(
        test_case_id        =   "Pretest_LTE_CD-04-003",
        at_cmd_index        =   "3",
        at_command          =   "AT+QSVN?",
        expected_response   =   "13"
    ),

    ATTestCase(
        test_case_id        =   "Pretest_LTE_CD-04-003",
        at_cmd_index        =   "4",
        at_command          =   "AT+EGMR=0,9",
        expected_response   =   "13"
    ),

    ATTestCase(
        test_case_id        =   "Pretest_LTE_CD-04-003",
        at_cmd_index        =   "5",
        at_command          =   "AT+QSVN=\"1a\"",
        expected_response   =   "ERROR"
    ),

    ATTestCase(
        test_case_id        =   "Pretest_LTE_CD-04-003",
        at_cmd_index        =   "6",
        at_command          =   "AT+QSVN=\"99\"",
        expected_response   =   "OK"
    ),

    ATTestCase(
        test_case_id        =   "Pretest_LTE_CD-04-003",
        at_cmd_index        =   "7",
        at_command          =   "AT+CFUN=1,1",
        expected_response   =   "OK"
    ),

    ATTestCase(
        test_case_id        =   "Pretest_LTE_CD-04-003",
        at_cmd_index        =   "8",
        at_command          =   "AT+QSVN?",
        expected_response   =   "99"
    ),

    ATTestCase(
        test_case_id        =   "Pretest_LTE_CD-04-003",
        at_cmd_index        =   "9",
        at_command          =   "AT+EGMR=0,9",
        expected_response   =   "99"
    ),

    ATTestCase(
        test_case_id        =   "Pretest_LTE_CD-04-003",
        at_cmd_index        =   "10",
        at_command          =   "AT+QSVN=\"FF\"",
        expected_response   =   "OK"
    ),

    ATTestCase(
        test_case_id        =   "Pretest_LTE_CD-04-003",
        at_cmd_index        =   "11",
        at_command          =   "AT+CFUN=1,1",
        expected_response   =   "OK"
    ),

    ATTestCase(
        test_case_id        =   "Pretest_LTE_CD-04-003",
        at_cmd_index        =   "12",
        at_command          =   "AT+QSVN?",
        expected_response   =   "FF"
    ),

    ATTestCase(
        test_case_id        =   "Pretest_LTE_CD-04-003",
        at_cmd_index        =   "13",
        at_command          =   "AT+EGMR=0,9",
        expected_response   =   "2"
    ),

    ATTestCase(
        test_case_id        =   "Pretest_LTE_CD-04-003",
        at_cmd_index        =   "14",
        at_command          =   "AT+QSVN=\"13\"",
        expected_response   =   "OK"
    ),

    ATTestCase(
        test_case_id        =   "Pretest_LTE_CD-04-003",
        at_cmd_index        =   "15",
        at_command          =   "AT+CFUN=1,1",
        expected_response   =   "OK"
    ),

    ATTestCase(
        test_case_id        =   "Pretest_LTE_CD-04-003",
        at_cmd_index        =   "16",
        at_command          =   "AT+EGMR=0,9",
        expected_response   =   "13"
    ),

    ATTestCase(
        test_case_id        =   "Pretest_LTE_CD-04-003",
        at_cmd_index        =   "17",
        at_command          =   "AT+QSVN=\"Ff\"",
        expected_response   =   "OK"
    ),

    ATTestCase(
        test_case_id        =   "Pretest_LTE_CD-04-003",
        at_cmd_index        =   "18",
        at_command          =   "AT+CFUN=1,1",
        expected_response   =   "OK"
    ),

    ATTestCase(
        test_case_id        =   "Pretest_LTE_CD-04-003",
        at_cmd_index        =   "19",
        at_command          =   "AT+QSVN?",
        expected_response   =   "Ff"
    ),

    ATTestCase(
        test_case_id        =   "Pretest_LTE_CD-04-003",
        at_cmd_index        =   "20",
        at_command          =   "AT+EGMR=0,9",
        expected_response   =   "2"
    ),

    ATTestCase(
        test_case_id        =   "Pretest_LTE_CD-04-003",
        at_cmd_index        =   "21",
        at_command          =   "AT+QSVN=\"05\"",
        expected_response   =   "OK"
    ),

    ATTestCase(
        test_case_id        =   "Pretest_LTE_CD-04-003",
        at_cmd_index        =   "22",
        at_command          =   "AT+QNVFD=\"/nv/item_files/eagle/qsvn\"",
        expected_response   =   "OK"
    ),

    # Pretest_LTE_CD-05-001

    ATTestCase(
        test_case_id        =   "Pretest_LTE_CD-05-001",
        at_cmd_index        =   "1",
        at_command          =   "AT+QCCID",
        expected_response   =   "+QCCID" #GET USER INPUT
    ),

    # Pretest_LTE_CD-05-002

    ATTestCase(
        test_case_id        =   "Pretest_LTE_CD-05-002",
        at_cmd_index        =   "1",
        at_command          =   "AT+GSN",
        expected_response   =   "OK" #GET USER INPUT
    ),

    ATTestCase(
        test_case_id        =   "Pretest_LTE_CD-05-002",
        at_cmd_index        =   "2",
        at_command          =   "AT+CGSN",
        expected_response   =   "OK" #GET USER INPUT
    ),

    ATTestCase(
        test_case_id        =   "Pretest_LTE_CD-05-002",
        at_cmd_index        =   "3",
        at_command          =   "AT+QGSN",
        expected_response   =   "OK" #GET USER INPUT
    ),

    # Pretest_LTE_CD-05-003

    ATTestCase(
        test_case_id        =   "Pretest_LTE_CD-05-003",
        at_cmd_index        =   "1",
        at_command          =   "AT+CIMI",
        expected_response   =   "OK" #GET USER INPUT
    ),

    # Pretest_LTE_CD-05-004

    ATTestCase(
        test_case_id        =   "Pretest_LTE_CD-05-004",
        at_cmd_index        =   "1",
        at_command          =   "AT+EGMR=?",
        expected_response   =   "+EGMR"
    ),
]
