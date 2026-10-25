##
# @file constants.py
# @brief Constants for LTE AT command pre-test automation tool.
##

# Software information (the only place the version number is written)
SOFTWARE_VERSION                    = "1.0.0"

# Serial configuration
BAUD_RATE                           = 115200
SEND_TIMEOUT                        = 5
COMMAND_DELAY_SEC                   = 0.1

# Reset handling
RESET_AT_COMMAND                    = "AT+CFUN=1,1"
RESET_DELAY_QCX216_SEC              = 30
RESET_DELAY_MDM9X07_SEC             = 30
RESET_DELAY_SDX35_SEC               = 90

# Formatting
LINE_DIVISOR_STRING = "------------------------------------------------------------------------"
STARTUP_MESSAGE_STRING = (
    "[ Eagle SW AT Command Pre-Test ]\n\n"
    "Author   : Anas Akasyah Bin Abdul Patas\n"
    f"Version  : {SOFTWARE_VERSION}\n"
    "Company  : Eagle Wireless SW Team\n\n"
)

# SIM information length
SIM_CCID_LEN                        = 20
SIM_IMEI_LEN                        = 15
SIM_IMSI_LEN                        = 15

# Branch display names
BRANCH_NAME_EAGLE                   = "Eagle Wireless"
BRANCH_NAME_NETPRISMA               = "NetPrisma"
BRANCH_NAME_QUECTEL                 = "Quectel"

# Maximum AT command parse line
AT_CMD_RESPONSE_LINE                =   1
AT_CMD_MIN_LINE_TO_PARSE            =   2

# Platform Name
PLATFORM_QCX216                     = "QCX216"
PLATFORM_MDM9X07                    = "MDM9X07"
PLATFORM_SDX35                      = "SDX35"
