# Eagle Wireless Pre-Test Software (EaglePretest)

## Software Information

* Author                : Anas Akasyah Bin Abdul Patas
* Software Name         : EaglePretest
* Software Version      : 1.0.0
* Release Date          : 26/10/2026
* User                  : Eagle Wireless Software Development Team
* Supported Modules     : QCX216, MDM9X07, SDX35

## Prerequisites

Before running the pre-test, make sure these are available:
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
* EaglePretest Executable (EaglePretest.exe)
* Eagle Wireless username (-@eaglewireless.com)
* Team password : eaglewireless@123
* Modem connected to the PC (AT port visible in Device Manager)
* SIM information list (.txt)
* Offline mode : Jenkins Build Parameter Excel Sheet
* Online mode  : Internet connection & Atlassian API token
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
* Create the Atlassian API token here
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
https://id.atlassian.com/manage-profile/security/api-tokens
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

## How to use

Before test

* Create a folder in any directory & name it "Pre-Test"
* Copy the EaglePretest executable (EaglePretest.exe) file inside the folder
* Copy the SIM information list (.txt) inside the folder
* Refer to your specific SIM card & module's information by running these AT commands beforehand:
    * AT+CCID = SIM CCID
    * AT+GSN  = SIM IMEI
    * AT+CIMI = SIM IMSI

Log in

* Double left-click on the EaglePretest executable (EaglePretest.exe) file
* Enter the username (-@eaglewireless.com) & the team password, then press Log In
* The password is only shared within the team. Ask your team lead if you do not have it.

During test

* Enter the JIRA ID of the build. The module is selected from the JIRA ID:
    * QCX216-xx = QCX216
    * SW9X07-xx = MDM9X07
    * SDX35-xx  = SDX35
* Select the mode with the Offline / Online button (top right of the JIRA Ticket box)
    * Offline (red)
        * Press "Upload Excel" & select the Jenkins Build Parameter Excel Sheet
    * Online (grey -> dark green)
        * The internet connection is checked first (spinning box). Press ⟳ to check again.
        * Enter the API token. A working API token is remembered for the next time.
        * The JIRA ticket must be a Pretest JIRA (the JIRA topic contains "Pretest")
        * The newest Excel sheet attached to the JIRA ticket is downloaded when the pre-test starts
        * If online mode cannot connect, press ? for what to check
* Select the COM port of the modem AT port (e.g. "... USB AT Port"). The port is checked automatically.
* Fill in the SIM information:
    * SIM CCID  = Integrated Circuit Card Identifier        = 20 characters
    * SIM IMEI  = International Mobile Equipment Identity   = 15 digits
    * SIM IMSI  = International Mobile Subscriber Identity  = 15 digits
* All of these information can be found in the SIM information list (.txt) file
* When every pre-check is ✓ (JIRA, AT port, Excel / API, CCID, IMEI, IMSI), press "Start Pre-Test"
* The Test Cases table, progress bar & live log are updated while the test runs
* Press "Stop" to stop the test. The partial log file is kept.

After test

* The result is shown in the Test Cases table. Select a row to see the response of the command.
* Use the search bar (beside Reset View) to find a test case or a text in the live log
* Press "Download Result" to save the result
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
* Log (.txt)    : Full log of the test
* Report (.pdf) : Run information & Test Cases table with the actual response of every command
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
* The downloaded log & report will be named as
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
EaglePretest_Report_YYYYMMDD_HHMMSS.txt
EaglePretest_Report_YYYYMMDD_HHMMSS.pdf
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
* Online mode : When the test completes, the result is posted to the JIRA ticket as a comment with the report (.pdf) & log (.txt) attached
* The log file of the latest test is kept in the folder until the tool is opened again
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Eagle_Wireless_SW_AT_Pretest_Log_YYYYMMDD_HHMMSS.txt
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

## Development

* Build the executable (inside the src folder)
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
python -m PyInstaller --noconfirm EaglePretest.spec
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
* The executable is created in src/dist/EaglePretest.exe
* Run the tests (inside the project folder). No modem & no real JIRA needed. See tests/README.md.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
python -m unittest discover -s tests -v
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

## Releases

### EaglePretest

* Version - 1.0.0
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
* Release Date - 26/10/2026
* New graphical tool (EaglePretest.exe), replaces the console pre-test
* Login page (Eagle Wireless username & team password)
* Module selected from the JIRA ID
* Offline mode : Jenkins Build Parameter Excel Sheet uploaded in the tool
* Online mode
    * Excel sheet downloaded from the JIRA ticket
    * Internet connection, API token & Pretest JIRA checked before the test
    * Result posted to the JIRA ticket as a comment (report & log attached)
* COM port checked automatically (modem AT port)
* Test Cases table, progress bar, live log & search
* Download Result : log (.txt) or report (.pdf)
* Branch name
    * If branch name has no indicators, default is "Quectel" (was "Eagle Wireless")
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

### Eagle SW AT Command Pre-Test (console, replaced by EaglePretest)

* Version - 1.0.0
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
* Release Date - 23/7/2026
* Initial release to the Software Development Team
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
* Version - 1.0.1
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
* Release Date - 24/7/2026
* Added fixes for QCX216
    * ATI - certversion
    * AT+QGMR - version + _OEM + _modemapp.modemapp
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
* Version - 1.0.2
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
* Release Date - 12/8/2026
* Added fixes for constants
    * "Eagle" change to "Eagle Wireless"
* Added fixes for test case
    * AT+GMI - Eagle Wireless
    * AT+CGMI - Eagle Wireless
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
* Version - 1.0.3
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
* Release Date - 27/8/2026
* Added fixes for branch name
    * If branch name is empty, default is "Eagle Wireless"
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
* Version - 1.0.4
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
* Release Date - 7/9/2026
* MDM9X07
    * Added modemapp & linuxapp
    * modemapp has "_" at the front while linuxapp has "." at the front
* Test case
    * Added ATE1 to activate command echo in response
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
* Version - 1.0.5
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
* Release Date - 2/10/2026
* SDX35
    * Added test script for SDX35 module
    * Specify supported & unsupported AT commands
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
