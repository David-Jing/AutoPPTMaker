import datetime
import getpass
import configparser
import googleapiclient
import os

from enum import Enum
from apiclient import errors
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials


class Logging:
    class LogType(Enum):
        # Specifies the type of log, types consists of either "Verbose", "Warning", or "Error"
        #   Info - For debugging and stuff, no serious issues here
        #   Warning - Something to take note of, does not negatively affects operation
        #   Error   - Something that hinders operation
        Info = 0
        Warning = 1
        Error = 2

    # Update version number here
    VersionNumber = "2.0.1"

    sheetService: googleapiclient.discovery.Resource = None

    globalConfig = configparser.ConfigParser()
    globalConfig.read("Data/GlobalProperties.ini")

    @staticmethod
    def initializeLoggingService() -> None:
        if Logging.sheetService != None:
            return

        print("INITIALIZING LOGGING")
        # Refer to https://developers.google.com/slides/api/quickstart/python
        creds = None
        # The file token.pickle stores the user's access and refresh tokens, and is
        # created automatically when the authorization flow completes for the first
        # time.
        if os.path.exists("Data/token.pickle"):
            creds = Credentials.from_authorized_user_file("Data/token.pickle", ["https://www.googleapis.com/auth/spreadsheets"])
        # If there are no (valid) credentials available, let the user log in.
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    "Data/credentials.json", ["https://www.googleapis.com/auth/spreadsheets"])
                creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open("Data/token.pickle", "w") as token:
                token.write(creds.to_json())

        Logging.sheetService = build("sheets", "v4", credentials=creds)

    @staticmethod
    def writeLog(logType: LogType, msg: str) -> None:
        Logging.initializeLoggingService()

        # Log entries into a Google Sheet file; logType consists of either "Verbose", "Warning", or "Error"
        sheetID = Logging.globalConfig["SLIDE_MAKER_SHEET"]["SlideMakerSheetFileID"]
        loggingSheetID = Logging.globalConfig["SLIDE_MAKER_SHEET"]["SlideMakerSheetLoggingSheetID"]

        # Logging info
        date = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        user = getpass.getuser()
        msg = "=\"" + msg.replace("\n", "\"&char(10)&\"") + "\""

        # Create new row and insert data
        batchUpdateRequest = {
            "requests": [
                {
                    "insertRange": {
                        "range": {
                            "sheetId": loggingSheetID,
                            "startRowIndex": 1,
                            "endRowIndex": 2
                        },
                        "shiftDimension": "ROWS"
                    }
                },
                {
                    "pasteData": {
                        "data": f"{date}¬ {Logging.VersionNumber}¬ {logType.name}¬ {user}¬ {msg}",
                        "type": "PASTE_NORMAL",
                        "delimiter": "¬",
                        "coordinate": {
                            "sheetId": loggingSheetID,
                            "rowIndex": 1
                        }
                    }
                }
            ]
        }

        try:
            Logging.sheetService.spreadsheets().batchUpdate(spreadsheetId=sheetID, body=batchUpdateRequest).execute()
        except errors.HttpError as error:
            print(f"\tERROR: An error occurred on logging data; {error}")

# ==============================================================================================
# ============================================ TESTER ==========================================
# ==============================================================================================


if __name__ == "__main__":
    print(Logging.writeLog(Logging.LogType.Info, "TEST1"))
    print(Logging.writeLog(Logging.LogType.Info, "TEST2"))
