# EaglePretest GUI tests

Automatic checks for the EaglePretest GUI (`src/gui.py`). They run against a mock JIRA on this PC, so:

* no real JIRA ticket is changed
* no modem is needed
* the real tool folder is not used

Run from the project folder (`eagle_at_cmd_pre_test_2`):

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
python -m unittest discover -s tests -v
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Windows briefly open and close while the tests run. That is expected.

| File | Checks |
|---|---|
| test_login.py | Login page: username rule, hashed team password, version from constants.py |
| test_internet.py | Online / Offline button, internet check, ? box |
| test_token.py | API token check, remembered token, token cannot be copied |
| test_topic.py | Pretest JIRA check, JIRA not found |
| test_download.py | Excel download at Start (newest sheet, old sheets deleted, sheet open in Excel) |
| test_result.py | PDF report, result comment and attachments posted to JIRA |
| test_offline.py | Excel upload, SIM / JIRA ID / COM port checks, new session clean-up, full offline run |
| test_parser.py | Worker output read into the table: PASS / FAIL / SKIP, responses, progress, stop, modem reset, crash |
| support.py | Mock JIRA server and shared helpers (not a test) |
| fake_worker.py | Stand-in for the pre-test process, prints the module output without a modem (not a test) |
