from __future__ import print_function
import configparser
from datetime import timedelta
import pickle
import sys
import os.path
import datetime
import webbrowser
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from apiclient import errors
from enum import Enum

class DateFormatMode(Enum):
    # Specifies the string format outputted by getFormattedNextSundayDate()
    Full = 0
    Short = 1


class PPTEditorTools:
    def __init__(self, type):
        # If modifying these scopes, delete the file token.json.
        self.scope = ['https://www.googleapis.com/auth/presentations',
                      'https://www.googleapis.com/auth/drive']

        # The ID of the source slide.
        config = configparser.ConfigParser()
        config.read(type + "SlideProperties.ini")
        self.sourceSlideID = config["SLIDE_PROPERTIES"][type + "SourceSlideID"]

        # Get access to the slide and drive
        [self.slideService, self.driveService] = self.getAPIServices()

        # Get rid of previously made slides
        self.removePreviousSlides()

        # Generate new slide and save its ID
        self.newSlideID = self.getDuplicatePresentation(type)

        # Slide change requests are appended here
        self.requests = []

        # Get access to slide data
        self.presentation = self.slideService.presentations().get(
            presentationId=self.newSlideID).execute()

    # ==========================================================================================
    # ======================================= API TOOLS ========================================
    # ==========================================================================================

    def getAPIServices(self):      
        # Refer to https://developers.google.com/slides/api/quickstart/python
        creds = None
        # The file token.json stores the user's access and refresh tokens, and is
        # created automatically when the authorization flow completes for the first
        # time.
        if os.path.exists('token.json'):
            creds = Credentials.from_authorized_user_file('token.json', self.scope)
        # If there are no (valid) credentials available, let the user log in.
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    'credentials.json', self.scope)
                creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open('token.json', 'w') as token:
                token.write(creds.to_json())

        slideService = build('slides', 'v1', credentials=creds, static_discovery=False)
        driveService = build('drive', 'v3', credentials=creds, static_discovery=False)

        return [slideService, driveService]

    def commitSlideChanges(self):
        # Commit changes to slides, reset requests, and update data
        successfulCommit = True
        if len(self.requests) > 0:
            try:
                self.slideService.presentations().batchUpdate(
                    presentationId=self.newSlideID, body={'requests': self.requests}).execute()
            except errors.HttpError as error:
                print(f"\tAn error occurred on slide commit: {error}")
                successfulCommit = False

        self.requests = []
        self.presentation = self.slideService.presentations().get(
            presentationId=self.newSlideID).execute()

        return successfulCommit

    # ==========================================================================================
    # =================================== SLIDE DATA GETTERS ===================================
    # ==========================================================================================

    def getTotalSlideNumber(self):
        # Returns the total number of slides
        return len(self.presentation.get('slides'))

    def getText(self, pageElement):
        # Returns the first block of formatted text from an element
        try:
            return pageElement['shape']['text']['textElements'][1]['textRun']['content']
        except:
            return ""

    def getSlideTextData(self, index):
        # Returns the objectID and text from the textboxes in the indexed slide
        textObjects = []
        for element in self.presentation.get('slides')[index]['pageElements']:
            textObject = [element.get('objectId'), self.getText(element)]
            if textObject[1] != "":
                textObjects.append(textObject)
        return textObjects

    def getSlideID(self, index):
        return self.presentation.get('slides')[index]['objectId']

    # ==========================================================================================
    # =================================== SLIDE MODIFIERS ======================================
    # ==========================================================================================

    def duplicateSlide(self, slideObjectID, newSlideObjectID):
        # Duplicates a slide, the duplicated slide is placed right after the source. "newSlideObjectID" must be unique.
        self.requests.append({"duplicateObject": {"objectId": slideObjectID,
                                                  "objectIds": {slideObjectID: newSlideObjectID}}})

    def deleteSlide(self, slideObjectID):
        # Delete the slide with the "slideObjectID"
        self.requests.append({"deleteObject": {"objectId": slideObjectID}})

    # ==========================================================================================
    # ================================= SLIDE FORMAT SETTERS ===================================
    # ==========================================================================================

    def setText(self, objectID, newText):
        # Delete existing text and insert new
        self.requests.append({"deleteText": {"objectId": objectID}})
        self.requests.append({"insertText": {"objectId": objectID, "text": newText}})

    def setBold(self, objectID, startIndex, endIndex):
        # Set bold to a select range of text
        self.requests.append({"updateTextStyle": {"objectId": objectID,
                                                  "textRange": {"type": "FIXED_RANGE", "startIndex": startIndex, "endIndex": endIndex},
                                                  "style": {"bold": True},
                                                  "fields": "bold"}})

    def setTextStyle(self, objectID, bold, italic, underline, size):
        # Sets bold, italic, or underline to "True" or "False" to enable or disable; also the entire object's font size.
        self.requests.append({"updateTextStyle": {"objectId": objectID,
                                                  "style": {"bold": bold, "italic": italic, "underline": underline, "fontSize": {"magnitude": size, "unit": "PT"}},
                                                  "fields": "bold, italic, underline, fontSize"}})

    def setParagraphStyle(self, objectID, spacing, alignment):
        # Sets the space between each line, default is 100.0; alignment options includes "START" (aka 'left'), "CENTER", "END" (aka 'right'), and "JUSTIFIED"
        self.requests.append({"updateParagraphStyle": {"objectId": objectID,
                                                       "style": {"lineSpacing": spacing, "alignment": alignment},
                                                       "fields": "lineSpacing, alignment"}})

    def setTextSuperScript(self, objectID, startIndex, endIndex):
        # Formats text indexed from 'startIndex' to 'endIndex' to superscript
        self.requests.append({"updateTextStyle": {"objectId": objectID,
                                                  "textRange": {"type": "FIXED_RANGE", "startIndex": startIndex, "endIndex": endIndex},
                                                  "style":  {"baselineOffset": "SUPERSCRIPT"},
                                                  "fields": "baselineOffset"}})

    def setSpaceAbove(self, objectID, paragraphIndex, units):
        # Adds spacing above a paragraph, any index within the desired line will work
        self.requests.append({"updateParagraphStyle": {"objectId": objectID,
                                                       "textRange": {"type": "FIXED_RANGE", "startIndex": paragraphIndex, "endIndex": paragraphIndex + 1},
                                                       "style": {"spaceAbove": {"magnitude": units, "unit": "PT"}},
                                                       "fields": "spaceAbove"}})

    # ==========================================================================================
    # ================================= SLIDE FORMAT UPDATERS ==================================
    # ==========================================================================================

    def updatePageElementTransform(self, objectID, scaleX=1, scaleY=1, translateX=0, translateY=0):
        # Horizontally and vertically scale or translate slide objects
        self.requests.append({"updatePageElementTransform": {"objectId": objectID,
                                                             "applyMode": "RELATIVE",
                                                             "transform": {"scaleX": scaleX, "scaleY": scaleY,
                                                                           "translateX": translateX, "translateY": translateY, "unit": "EMU"}}})

    # ==========================================================================================
    # ====================================== MISC GETTERS ======================================
    # ==========================================================================================

    def getNextSundayDate(self):
        # Get datetime variable with the date of next Sunday
        dt = datetime.date.today()
        dt += timedelta(days=(6 - dt.weekday()))
        return dt

    def getFormattedNextSundayDate(self, type):
        # Get formatted string with the date of next Sunday
        dt = self.getNextSundayDate()

        month = dt.strftime('%B') if type == DateFormatMode.Full else dt.strftime('%b')
        day = dt.strftime('%d')
        year = dt.strftime('%Y')

        # Get rid of leading zero in day
        if (day[0] == "0"):
            day = day[1]

        lastNum = day[len(day) - 1]
        if len(day) > 1 and day[0] != '1':
            ordinal = 'st' if lastNum == '1' else 'nd' if lastNum == '2' else 'rd' if lastNum == '3' else 'th'
        else:
            ordinal = 'th'
        return [f'{month} {day}{ordinal} {year}', len(month) + len(day) + 1]  # Return string and index of the ordinal for superscripting

    def getUpcomingSlideTitle(self, type):
        # Get slide file name for upcoming Sunday
        dt = self.getFormattedNextSundayDate(type)[0]
        return f'Sunday Worship Slides {dt} - {type}'

    def openSlideInBrowser(self):
        # Access may be denied due to slide permissions (HINT: Set source slide URL with editor permission)
        webbrowser.open(f'https://docs.google.com/presentation/d/{self.newSlideID}')

    # ==========================================================================================
    # =================================== DRIVE CHANGE TOOLS ===================================
    # ==========================================================================================

    def removePreviousSlides(self):
        # Remove all previously created slides
        if (os.path.exists("SlideIDList.txt")):
            f = open("SlideIDList.txt", "r+")
            lines = f.readlines()

            for line in lines:
                try:
                    # Get rid of the new line symbol in the ID
                    self.driveService.files().delete(fileId=line[:-1]).execute()
                except errors.HttpError as error:
                    print(f"An error occurred on slide removal: {error}")

            # Clear out the file
            f.truncate(0)

    def getDuplicatePresentation(self, type):
        # Generates a Duplicate Presentation
        body = {
            'name': self.getUpcomingSlideTitle(type)
        }

        drive_response = self.driveService.files().copy(
            fileId=self.sourceSlideID, body=body).execute()

        # Write ID to file, so it can be deleted later
        id = drive_response.get('id')
        f = open("SlideIDList.txt", "a")
        f.write(id + "\n")
        f.close()

        return id


if __name__ == '__main__':
    pe = PPTEditorTools('Stream')
    print(pe.presentation.get('slides')[20]['pageElements'])
