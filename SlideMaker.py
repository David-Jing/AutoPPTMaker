import configparser
import os
import sys
import time
import traceback

from typing import List
from enum import Enum

from HymnMaker import HymnMaker
from GoogleAPITools import GoogleAPITools
from ListMaker import ListMaker
from Logging import Logging
from Utility import Utility
from VerseMaker import VerseMaker

'''

Core class of the SlideMaker application, assembles all other classes to generate Google Slides from an user-defined input file.

'''


class SlideMaker:
    class PPTMode(Enum):
        # Specifies the PPT output format
        Null = -1
        Stream = 0
        Projected = 1
        Regular = 2

    def __init__(self) -> None:
        # Offset for when creating new slides
        self.slideOffset = 0

        # Access slide input data
        if not os.path.exists("SlideInputs.ini"):
            raise IOError(f"ERROR : SlideInputs.ini input file cannot be found.")
        self.input = configparser.ConfigParser()
        self.input.read("SlideInputs.ini")

    def setType(self, pptType: PPTMode) -> None:
        strType = pptType.name

        if not os.path.exists("Data/" + strType + "SlideProperties.ini"):
            raise IOError(f"ERROR : {strType}SlideProperties.ini config file cannot be found.")

        self.gEditor = GoogleAPITools(strType)
        self.verseMaker = VerseMaker(strType)
        self.hymnMaker = HymnMaker(strType)

        # Initialize logging
        Logging.initializeLoggingService()

        # Access slide property data
        self.config = configparser.ConfigParser()
        self.config.read("Data/" + strType + "SlideProperties.ini")

        # General linespacing for all slides
        self.lineSpacing = int(self.config["SLIDE_PROPERTIES"]["SlideLineSpacing"])

    # ======================================================================================================
    # ========================================== SLIDE MAKERS ==============================================
    # ======================================================================================================

    def sundayServiceSlide(self) -> bool:
        title = self.input["SUNDAY_SERVICE_HEADER"]["SundayServiceHeaderTitle"].upper()

        return self._titleSlide(title, "SUNDAY_SERVICE_HEADER_PROPERTIES", "SundayServiceHeader")

    def monthlyScriptureSlide(self) -> bool:
        title = self.input["MONTHLY_SCRIPTURE"]["MonthlyScriptureTitle"].upper()
        source = self.input["MONTHLY_SCRIPTURE"]["MonthlyScriptureSource"].upper()
        maxLineLength = int(self.config["MONTHLY_SCRIPTURE_PROPERTIES"]["MonthlyScriptureMaxLineLength"])

        return self._scriptureSingleSlide(title, source, maxLineLength,
                                          "MONTHLY_SCRIPTURE_PROPERTIES", "MonthlyScripture")

    def announcementSlide(self) -> bool:
        title = self.input["ANNOUNCEMENTS"]["AnnouncementsTitle"].upper()
        announcementList = self.gEditor.getAnnouncements()

        return self._textMultiSlide(title, announcementList, "ANNOUNCEMENTS_PROPERTIES", "Announcements")

    def bibleVerseMemorizationSlide(self, nextWeek: bool = False) -> bool:
        lastWeekString = "LastWeek" if not nextWeek else ""
        title = self.input["BIBLE_MEMORIZATION"]["BibleMemorization" + lastWeekString + "Title"].upper()
        source = self.input["BIBLE_MEMORIZATION"]["BibleMemorization" + lastWeekString + "Source"].upper()
        maxLineLength = int(self.config["BIBLE_MEMORIZATION_PROPERTIES"]["BibleMemorizationMaxLineLength"])

        return self._scriptureSingleSlide(title, source, maxLineLength,
                                          "BIBLE_MEMORIZATION_PROPERTIES", "BibleMemorization", nextWeek)

    def catechismSlide(self, nextWeek: bool = False) -> bool:
        title = self.input["CATECHISM"]["Catechism" + ("LastWeek" if not nextWeek else "") + "Title"].upper()
        return self._headerOnlySlide(title, "CATECHISM_PROPERTIES", "Catechism", nextWeek)

    def supplicationSlide(self) -> bool:
        enabled = self.input["SUPPLICATIONS"]["SupplicationsEnabled"].upper()
        title = self.input["SUPPLICATIONS"]["SupplicationsTitle"].upper()
        supplicationList = self.gEditor.getSupplications()

        if (enabled == "TRUE"):
            return self._numListMultiSlide(title, supplicationList, "SUPPLICATIONS_PROPERTIES", "Supplications")
        else:
            slideIndex = int(self.config["SUPPLICATIONS_PROPERTIES"]["SupplicationsIndex"]) + self.slideOffset
            return self._deleteSlide(slideIndex)

    def worshipSlide(self) -> bool:
        title = self.input["WORSHIP_HEADER"]["WorshipHeaderTitle"].upper()

        return self._titleSlide(title, "WORSHIP_HEADER_PROPERTIES", "WorshipHeader")

    def callToWorshipSlide(self) -> bool:
        source = self.input["CALL_TO_WORSHIP"]["CallToWorshipSource"].upper()
        maxLineLength = int(self.config["CALL_TO_WORSHIP_PROPERTIES"]["CallToWorshipMaxLineLength"])
        maxLinesPerSlide = int(self.config["CALL_TO_WORSHIP_PROPERTIES"]["CallToWorshipMaxLines"])

        # Delete section if source is empty
        if (source == ""):
            return self._deleteSlide(int(self.config["CALL_TO_WORSHIP_PROPERTIES"]["CallToWorshipIndex"]) + self.slideOffset)

        return self._scriptureMultiSlide(source, maxLineLength, maxLinesPerSlide,
                                         "CALL_TO_WORSHIP_PROPERTIES", "CallToWorship")

    def prayerOfConfessionSlide(self) -> bool:
        title = self.input["PRAYER_OF_CONFESSION"]["PrayerOfConfessionTitle"].upper()
        source = self.input["PRAYER_OF_CONFESSION"]["PrayerOfConfessionSource"].upper()
        maxLineLength = int(self.config["PRAYER_OF_CONFESSION_PROPERTIES"]["PrayerOfConfessionMaxLineLength"])

        # Delete section if source is empty
        if (source == ""):
            return self._deleteSlide(int(self.config["PRAYER_OF_CONFESSION_PROPERTIES"]["PrayerOfConfessionIndex"]) + self.slideOffset)

        return self._scriptureSingleSlide(title, source, maxLineLength,
                                          "PRAYER_OF_CONFESSION_PROPERTIES", "PrayerOfConfession")

    def holyCommunionSlide(self) -> bool:
        enabled = self.input["HOLY_COMMUNION"]["HolyCommunionSlidesEnabled"].upper()
        slideIndex = int(self.config["HOLY_COMMUNION_PROPERTIES"]["HolyCommunionIndex"]) + self.slideOffset
        numOfSlides = int(self.config["HOLY_COMMUNION_PROPERTIES"]["HolyCommunionSlides"])

        # If not enabled, delete section
        if (enabled != "TRUE"):
            for i in range(slideIndex, slideIndex + numOfSlides):
                sourceSlideID = self.gEditor.getSlideID(i)
                self.gEditor.deleteSlide(sourceSlideID)

            if (not self.gEditor.commitSlideChanges()):
                return False

            self.slideOffset -= numOfSlides

            # Disabled status
            return False

        return True

    def sermonHeaderSlide(self) -> bool:
        title = self.input["SERMON_HEADER"]["SermonHeaderTitle"].upper()
        speaker = self.input["SERMON_HEADER"]["SermonHeaderSpeaker"]

        return self._sermonHeaderSlide(title, speaker, "SERMON_HEADER_PROPERTIES", "SermonHeader")

    def sermonVerseSlide(self) -> bool:
        # Allow for multiple unique sources separated by ","
        sources = self.input["SERMON_VERSE"]["SermonVerseSource"].upper()
        sourceList = sources.split(",")
        maxLineLength = int(self.config["SERMON_VERSE_PROPERTIES"]["SermonVerseMaxLineLength"])
        maxLinesPerSlide = int(self.config["SERMON_VERSE_PROPERTIES"]["SermonVerseMaxLines"])

        # Generate templates for each verse source
        slideIndex = int(self.config["SERMON_VERSE_PROPERTIES"]["SermonVerseIndex"]) + self.slideOffset
        if not self._duplicateSlide(len(sourceList) - 1, slideIndex):
            return False

        # Iterate and generate the slides
        output = True
        for source in sourceList:
            if (not self._scriptureMultiSlide(source.strip(), maxLineLength, maxLinesPerSlide, "SERMON_VERSE_PROPERTIES", "SermonVerse")
                    and output):
                output = False
            self.slideOffset += 1
        self.slideOffset -= 1

        return output

    def hymnSlides(self, number: int) -> bool:
        hymnSource = self.input["HYMN"][f"Hymn{number}Source"]
        hymnIndex = int(self.config["HYMN_PROPERTIES"][f"Hymn{number}Index"]) + self.slideOffset

        # Delete section if source is empty
        if (hymnSource == ""):
            return self._deleteSlide(hymnIndex)

        return self._hymnSlide(hymnSource, hymnIndex)

    # ======================================================================================================
    # ===================================== SLIDE MAKER IMPLEMENTATIONS ====================================
    # ======================================================================================================

    def _titleSlide(self, title: str, propertyName: str, dataNameHeader: str) -> bool:
        (dateString, cordinalIndex) = self.gEditor.getFormattedNextSundayDate(self.gEditor.DateFormatMode.Full)

        checkPoint = [False, False]
        try:
            data = self.gEditor.getSlideTextData(int(self.config[propertyName][dataNameHeader + "Index"]) + self.slideOffset)
            for item in data:
                if '{Title}' in item[1]:
                    checkPoint[0] = True
                    self._insertText(
                        objectID=item[0],
                        text=title.upper(),
                        size=int(self.config[propertyName][dataNameHeader + "TitleTextSize"]),
                        bold=self._str2bool(self.config[propertyName][dataNameHeader + "TitleBolded"]),
                        italic=self._str2bool(self.config[propertyName][dataNameHeader + "TitleItalicized"]),
                        underlined=self._str2bool(self.config[propertyName][dataNameHeader + "TitleUnderlined"]),
                        alignment=self.config[propertyName][dataNameHeader + "TitleAlignment"])
                elif '{Text}' in item[1]:
                    checkPoint[1] = True
                    self._insertText(
                        objectID=item[0],
                        text=dateString[:-5] + "," + dateString[-5:],   # Add comma between year and date
                        size=int(self.config[propertyName][dataNameHeader + "DateTextSize"]),
                        bold=self._str2bool(self.config[propertyName][dataNameHeader + "DateBolded"]),
                        italic=self._str2bool(self.config[propertyName][dataNameHeader + "DateItalicized"]),
                        underlined=self._str2bool(self.config[propertyName][dataNameHeader + "DateUnderlined"]),
                        alignment=self.config[propertyName][dataNameHeader + "DateAlignment"])

                    self.gEditor.setTextSuperScript(item[0], cordinalIndex, cordinalIndex + 2)
        except Exception:
            print(f"\tERROR : {traceback.format_exc()}")
            return False

        if (False in checkPoint):
            print(f"\tERROR : {checkPoint.count(True)} out of {len(checkPoint)} text placeholders found.")
            return False

        return self.gEditor.commitSlideChanges()

    def _hymnSlide(self, source: str, slideIndex: int) -> bool:
        if (not self.hymnMaker.setSource(source)):
            print(f"\tERROR : Hymn [{source}] not found.")
            return False

        [titleList, formattedLyricsList, lookupSource] = self.hymnMaker.getContent()

        # Log if is sourced from the web
        if lookupSource == "Web":
            Logging.writeLog(Logging.LogType.Info, f"SlideMaker - Hymn web source lookup: [{source}]")

        # Title font size adjustments and decide if the lyrics text box needs to be shifted
        titleFontSize = int(self.config["HYMN_PROPERTIES"]["HymnTitleTextSize"])
        maxCharUnitLength = int(self.config["HYMN_PROPERTIES"]["HymnTitleMaxUnitLength"])
        minCharUnitLength = int(self.config["HYMN_PROPERTIES"]["HymnTitleMinUnitLength"])
        titleLength = Utility.getVisualLength(titleList[0])
        multiLineTitle = False
        if (titleLength > maxCharUnitLength):                                         # Heuristic : Shift the lyrics text block down for better aesthetics for multi-line titles
            multiLineTitle = True
            if ((titleLength - maxCharUnitLength) > 1.1*maxCharUnitLength):           # Heuristic : It's better when there're more text in the first line than second
                print(f"  MULTILINE FORMAT 1 : {titleLength} < {maxCharUnitLength} (line 2 too long)")
                titleFontSize = int(self.config["HYMN_PROPERTIES"]["HymnTitleMultiLineTextSize"])
            elif((titleLength - maxCharUnitLength) < minCharUnitLength):               # Heuristic : It looks terrible when's only a tiny bit of text on the 2nd line, reduce to one line
                print(f"  MULTILINE FORMAT 2 : {titleLength} - {maxCharUnitLength} < {minCharUnitLength} (line 2 too short)")
                titleFontSize = int(self.config["HYMN_PROPERTIES"]["HymnTitleMultiLineTextSize"])
                multiLineTitle = False
            else:
                print(f"  MULTILINE FORMAT DEFAULT : {titleLength}, {maxCharUnitLength}, {minCharUnitLength}")

        # Generate create duplicate request and commit it
        if not self._duplicateSlide(len(titleList) - 1, slideIndex):
            return False

        # Insert title and lyrics data
        for i in range(slideIndex, slideIndex + len(titleList)):
            checkPoint = [False, False]
            data = self.gEditor.getSlideTextData(i)
            try:
                for item in data:
                    if ('{Title}' in item[1]):
                        checkPoint[0] = True
                        self._insertText(
                            objectID=item[0],
                            text=titleList[i - slideIndex],
                            size=titleFontSize,
                            bold=self._str2bool(self.config["HYMN_PROPERTIES"]["HymnTitleBolded"]),
                            italic=self._str2bool(self.config["HYMN_PROPERTIES"]["HymnTitleItalicized"]),
                            underlined=self._str2bool(self.config["HYMN_PROPERTIES"]["HymnTitleUnderlined"]),
                            alignment=self.config["HYMN_PROPERTIES"]["HymnTitleAlignment"])
                    elif ('{Text}' in item[1]):
                        checkPoint[1] = True
                        self._insertText(
                            objectID=item[0],
                            text=formattedLyricsList[i - slideIndex],
                            size=int(self.config["HYMN_PROPERTIES"]["HymnTextSize"]),
                            bold=self._str2bool(self.config["HYMN_PROPERTIES"]["HymnBolded"]),
                            italic=self._str2bool(self.config["HYMN_PROPERTIES"]["HymnItalicized"]),
                            underlined=self._str2bool(self.config["HYMN_PROPERTIES"]["HymnUnderlined"]),
                            alignment=self.config["HYMN_PROPERTIES"]["HymnAlignment"])

                        if (multiLineTitle):
                            self.gEditor.updatePageElementTransform(item[0], translateY=int(self.config["HYMN_PROPERTIES"]["HymnLoweredUnitHeight"]))
            except Exception:
                print(f"\tERROR : {traceback.format_exc()}")
                return False

            if (False in checkPoint):
                print(f"\tERROR : {checkPoint.count(True)} out of {len(checkPoint)} text placeholders found.")
                return False

        return self.gEditor.commitSlideChanges()

    def _scriptureSingleSlide(self, title: str, source: str, charPerLine: int, propertyName: str, dataNameHeader: str, nextWeek: bool = False) -> bool:
        # Assumes monthly scripture is short enough to fit in one slide
        nextWeekString = "NextWeek" if nextWeek else ""

        # Generate slide anyway if source verse cannot be found
        if not self.verseMaker.setSource(source, charPerLine):
            print(f"\tWARNING: Verse [{source}] not found.")
            verseString = "{Text}"
            ssIndexList = []
        else:
            [verseString, ssIndexList] = self.verseMaker.getVerseString()

        checkPoint = [False, False]
        try:
            data = self.gEditor.getSlideTextData(int(self.config[propertyName][dataNameHeader + nextWeekString + "Index"]) + self.slideOffset)
            for item in data:
                if '{Title}' in item[1]:
                    checkPoint[0] = True
                    self._insertText(
                        objectID=item[0],
                        text=title,
                        size=int(self.config[propertyName][dataNameHeader + nextWeekString + "TitleTextSize"]),
                        bold=self._str2bool(self.config[propertyName][dataNameHeader + "TitleBolded"]),
                        italic=self._str2bool(self.config[propertyName][dataNameHeader + "TitleItalicized"]),
                        underlined=self._str2bool(self.config[propertyName][dataNameHeader + "TitleUnderlined"]),
                        alignment=self.config[propertyName][dataNameHeader + "TitleAlignment"])
                elif '{Text}' in item[1]:
                    checkPoint[1] = True
                    self._insertText(
                        objectID=item[0],
                        text=verseString,
                        size=int(self.config[propertyName][dataNameHeader + "TextSize"]),
                        bold=self._str2bool(self.config[propertyName][dataNameHeader + "Bolded"]),
                        italic=self._str2bool(self.config[propertyName][dataNameHeader + "Italicized"]),
                        underlined=self._str2bool(self.config[propertyName][dataNameHeader + "Underlined"]),
                        alignment=self.config[propertyName][dataNameHeader + "Alignment"])

                    # Superscript and bold the verse numbers, and set paragraph spacing between the verses
                    for ssRange in ssIndexList:
                        self.gEditor.setTextSuperScript(item[0], ssRange[1], ssRange[2])
                        self.gEditor.setBold(item[0], ssRange[1], ssRange[2])
                        if (ssRange[1] > 0 and ssRange[0] == self.verseMaker.SSType.VerseNumber):  # No need for paragraph spacing in initial line
                            self.gEditor.setSpaceAbove(item[0], ssRange[1], 10)
        except Exception:
            print(f"\tERROR : {traceback.format_exc()}")
            return False

        if (False in checkPoint):
            print(f"\tERROR : {checkPoint.count(True)} out of {len(checkPoint)} text placeholders found.")
            return False

        return self.gEditor.commitSlideChanges()

    def _scriptureMultiSlide(self, source: str, maxLineLength: int, maxLinesPerSlide: int, propertyName: str, dataNameHeader: str) -> bool:
        # Assumes monthly scripture is short enough to fit in one slide
        if (not self.verseMaker.setSource(source, maxLineLength)):
            print(f"\tERROR : Verse {source} not found.")
            return False

        [title, slideVersesList, slideSSIndexList] = self.verseMaker.getVerseStringMultiSlide(maxLinesPerSlide)

        # Generate create duplicate request and commit it
        slideIndex = int(self.config[propertyName][dataNameHeader + "Index"]) + self.slideOffset
        if not self._duplicateSlide(len(slideVersesList) - 1, slideIndex):
            return False

        # Get units of paragraph that separates verse numbers
        paragraphSpace = int(self.config["VERSE_PROPERTIES"]["VerseParagraphSpace"])

        # Insert title and verses data
        for i in range(slideIndex, slideIndex + len(slideVersesList)):
            checkPoint = [False, False]
            try:
                data = self.gEditor.getSlideTextData(i)
                for item in data:
                    if '{Title}' in item[1]:
                        checkPoint[0] = True
                        self._insertText(
                            objectID=item[0],
                            text=title,
                            size=int(self.config[propertyName][dataNameHeader + "TitleTextSize"]),
                            bold=self._str2bool(self.config[propertyName][dataNameHeader + "TitleBolded"]),
                            italic=self._str2bool(self.config[propertyName][dataNameHeader + "TitleItalicized"]),
                            underlined=self._str2bool(self.config[propertyName][dataNameHeader + "TitleUnderlined"]),
                            alignment=self.config[propertyName][dataNameHeader + "TitleAlignment"])
                    elif '{Text}' in item[1]:
                        checkPoint[1] = True
                        self._insertText(
                            objectID=item[0],
                            text=slideVersesList[i - slideIndex],
                            size=int(self.config[propertyName][dataNameHeader + "TextSize"]),
                            bold=self._str2bool(self.config[propertyName][dataNameHeader + "Bolded"]),
                            italic=self._str2bool(self.config[propertyName][dataNameHeader + "Italicized"]),
                            underlined=self._str2bool(self.config[propertyName][dataNameHeader + "Underlined"]),
                            alignment=self.config[propertyName][dataNameHeader + "Alignment"])

                        # Superscript and bold the verse numbers, and set paragraph spacing between the verses
                        for ssRange in slideSSIndexList[i - slideIndex]:
                            self.gEditor.setTextSuperScript(item[0], ssRange[1], ssRange[2])
                            self.gEditor.setBold(item[0], ssRange[1], ssRange[2])
                            if (ssRange[1] > 0 and ssRange[0] == self.verseMaker.SSType.VerseNumber):  # No need for paragraph spacing in initial line
                                self.gEditor.setSpaceAbove(item[0], ssRange[1], paragraphSpace)
            except Exception:
                print(f"\tERROR : {traceback.format_exc()}")
                return False

            if (False in checkPoint):
                print(f"\tERROR : {checkPoint.count(True)} out of {len(checkPoint)} text placeholders found.")
                return False

        return self.gEditor.commitSlideChanges()

    def _sermonHeaderSlide(self, title: str, speaker: str, propertyName: str, dataNameHeader: str) -> bool:
        checkPoint = [False, False]
        try:
            data = self.gEditor.getSlideTextData(int(self.config[propertyName][dataNameHeader + "Index"]) + self.slideOffset)
            for item in data:
                if '{Title}' in item[1]:
                    checkPoint[0] = True
                    self._insertText(
                        objectID=item[0],
                        text=title.upper(),
                        size=int(self.config[propertyName][dataNameHeader + "TitleTextSize"]),
                        bold=self._str2bool(self.config[propertyName][dataNameHeader + "TitleBolded"]),
                        italic=self._str2bool(self.config[propertyName][dataNameHeader + "TitleItalicized"]),
                        underlined=self._str2bool(self.config[propertyName][dataNameHeader + "TitleUnderlined"]),
                        alignment=self.config[propertyName][dataNameHeader + "TitleAlignment"])
                elif '{Text}' in item[1]:
                    checkPoint[1] = True
                    self._insertText(
                        objectID=item[0],
                        text=speaker.title(),
                        size=int(self.config[propertyName][dataNameHeader + "SpeakerTextSize"]),
                        bold=self._str2bool(self.config[propertyName][dataNameHeader + "SpeakerBolded"]),
                        italic=self._str2bool(self.config[propertyName][dataNameHeader + "SpeakerItalicized"]),
                        underlined=self._str2bool(self.config[propertyName][dataNameHeader + "SpeakerUnderlined"]),
                        alignment=self.config[propertyName][dataNameHeader + "SpeakerAlignment"])
        except Exception:
            print(f"\tERROR : {traceback.format_exc()}")
            return False

        if (False in checkPoint):
            print(f"\tERROR : {checkPoint.count(True)} out of {len(checkPoint)} text placeholders found.")
            return False

        return self.gEditor.commitSlideChanges()

    def _headerOnlySlide(self, title: str, propertyName: str, dataNameHeader: str, nextWeek: bool = False) -> bool:
        nextWeekString = "NextWeek" if nextWeek else ""
        checkPoint = [False]
        try:
            data = self.gEditor.getSlideTextData(int(self.config[propertyName][dataNameHeader + nextWeekString + "Index"]) + self.slideOffset)
            for item in data:
                if '{Title}' in item[1]:
                    checkPoint[0] = True
                    self._insertText(
                        objectID=item[0],
                        text=title.upper(),
                        size=int(self.config[propertyName][dataNameHeader + nextWeekString + "TitleTextSize"]),
                        bold=self._str2bool(self.config[propertyName][dataNameHeader + "TitleBolded"]),
                        italic=self._str2bool(self.config[propertyName][dataNameHeader + "TitleItalicized"]),
                        underlined=self._str2bool(self.config[propertyName][dataNameHeader + "TitleUnderlined"]),
                        alignment=self.config[propertyName][dataNameHeader + "TitleAlignment"])
        except Exception:
            print(f"\tERROR : {traceback.format_exc()}")
            return False

        if (False in checkPoint):
            print(f"\tERROR : {checkPoint.count(True)} out of {len(checkPoint)} text placeholders found.")
            return False

        return self.gEditor.commitSlideChanges()

    def _textMultiSlide(self, title: str, textContentList: List[str], propertyName: str, dataNameHeader: str) -> bool:
        # For announcement slides
        slideIndex = int(self.config[propertyName][dataNameHeader + "Index"]) + self.slideOffset

        # Generate create duplicate request and commit it
        if not self._duplicateSlide(len(textContentList) - 1, slideIndex):
            return False

        # Insert title and verses data
        for i in range(slideIndex, slideIndex + len(textContentList)):
            checkPoint = [False, False]
            try:
                data = self.gEditor.getSlideTextData(i)
                for item in data:
                    if '{Title}' in item[1]:
                        checkPoint[0] = True
                        self._insertText(
                            objectID=item[0],
                            text=title.upper(),
                            size=int(self.config[propertyName][dataNameHeader + "TitleTextSize"]),
                            bold=self._str2bool(self.config[propertyName][dataNameHeader + "TitleBolded"]),
                            italic=self._str2bool(self.config[propertyName][dataNameHeader + "TitleItalicized"]),
                            underlined=self._str2bool(self.config[propertyName][dataNameHeader + "TitleUnderlined"]),
                            alignment=self.config[propertyName][dataNameHeader + "TitleAlignment"])
                    elif '{Text}' in item[1]:
                        checkPoint[1] = True
                        self._insertText(
                            objectID=item[0],
                            text=textContentList[i - slideIndex],
                            size=int(self.config[propertyName][dataNameHeader + "TextSize"]),
                            bold=self._str2bool(self.config[propertyName][dataNameHeader + "Bolded"]),
                            italic=self._str2bool(self.config[propertyName][dataNameHeader + "Italicized"]),
                            underlined=self._str2bool(self.config[propertyName][dataNameHeader + "Underlined"]),
                            alignment=self.config[propertyName][dataNameHeader + "Alignment"])
            except Exception:
                print(f"\tERROR : {traceback.format_exc()}")
                return False

            if (False in checkPoint):
                print(f"\tERROR : {checkPoint.count(True)} out of {len(checkPoint)} text placeholders found.")
                return False

        return self.gEditor.commitSlideChanges()

    def _numListMultiSlide(self, title: str, contentList: List[str], propertyName: str, dataNameHeader: str) -> bool:
        # For supplication slides
        slideIndex = int(self.config[propertyName][dataNameHeader + "Index"]) + self.slideOffset
        paragraphSpace = int(self.config[propertyName][dataNameHeader + "ParagraphSpace"])

        # Get slide separated list
        [slideContentList, listStartIndex] = ListMaker.getFormattedListElementsBySlide(contentList,
                                                                                       indentSpace=int(self.config[propertyName][dataNameHeader + "IndentSpace"]),
                                                                                       maxLineLength=int(self.config[propertyName][dataNameHeader + "TextMaxLineLength"]),
                                                                                       maxLinesPerSlide=int(self.config[propertyName][dataNameHeader + "TextMaxLines"]))

        # Generate create duplicate request and commit it
        if not self._duplicateSlide(len(slideContentList) - 1, slideIndex):
            return False

        # Insert title and verses data
        for i in range(slideIndex, slideIndex + len(slideContentList)):
            checkPoint = [False, False]
            try:
                data = self.gEditor.getSlideTextData(i)
                for item in data:
                    if '{Title}' in item[1]:
                        checkPoint[0] = True
                        self._insertText(
                            objectID=item[0],
                            text=title.upper(),
                            size=int(self.config[propertyName][dataNameHeader + "TitleTextSize"]),
                            bold=self._str2bool(self.config[propertyName][dataNameHeader + "TitleBolded"]),
                            italic=self._str2bool(self.config[propertyName][dataNameHeader + "TitleItalicized"]),
                            underlined=self._str2bool(self.config[propertyName][dataNameHeader + "TitleUnderlined"]),
                            alignment=self.config[propertyName][dataNameHeader + "TitleAlignment"])
                    elif '{Text}' in item[1]:
                        checkPoint[1] = True
                        self._insertText(
                            objectID=item[0],
                            text=slideContentList[i - slideIndex],
                            size=int(self.config[propertyName][dataNameHeader + "TextSize"]),
                            bold=self._str2bool(self.config[propertyName][dataNameHeader + "Bolded"]),
                            italic=self._str2bool(self.config[propertyName][dataNameHeader + "Italicized"]),
                            underlined=self._str2bool(self.config[propertyName][dataNameHeader + "Underlined"]),
                            alignment=self.config[propertyName][dataNameHeader + "Alignment"])

                        # Add bullets
                        for lineStartIndex in listStartIndex[i - slideIndex]:
                            if lineStartIndex > 0:  # No need for paragraph spacing in initial line
                                self.gEditor.setSpaceAbove(item[0], lineStartIndex, paragraphSpace)

            except Exception:
                print(f"\tERROR : {traceback.format_exc()}")
                return False

            if (False in checkPoint):
                print(f"\tERROR : {checkPoint.count(True)} out of {len(checkPoint)} text placeholders found.")
                return False

        return self.gEditor.commitSlideChanges()

    # ======================================================================================================
    # ============================================ TOOLS ===================================================
    # ======================================================================================================

    def _deleteSlide(self, slideIndex: int) -> bool:
        sourceSlideID = self.gEditor.getSlideID(slideIndex)
        self.gEditor.deleteSlide(sourceSlideID)

        if (not self.gEditor.commitSlideChanges()):
            return False

        self.slideOffset -= 1

        # Disabled status
        return False

    def _str2bool(self, boolString: str) -> bool:
        return boolString.lower() in ("yes", "true", "t", "1")

    def _openSlideInBrowser(self) -> None:
        self.gEditor.openSlideInBrowser()

    def _insertText(self, objectID: str, text: str, size: int, bold: bool, italic: bool, underlined: bool, alignment: str, linespacing: int = -1) -> None:
        self.gEditor.setText(objectID, text)
        self.gEditor.setTextStyle(objectID, bold, italic, underlined, size)
        self.gEditor.setParagraphStyle(objectID, self.lineSpacing if linespacing < 0 else linespacing, alignment)

    def _duplicateSlide(self, dupCount: int, slideIndex: int) -> bool:
        # Generate extra duplicate slides (total number of slides = [dupCount + 1]) and commit it
        sourceSlideID = self.gEditor.getSlideID(slideIndex)
        for i in range(dupCount):  # Recall that the original source slide remains
            self.gEditor.duplicateSlide(
                sourceSlideID, sourceSlideID + '__' + str(i))
        if (not self.gEditor.commitSlideChanges()):
            return False

        # Update offset
        self.slideOffset += dupCount

        return True


if __name__ == '__main__':
    print("====================================================================")
    print(f"\t\t\tSlideMaker v{Logging.VersionNumber}")
    print("====================================================================")
    print("\nInput Options:\n  Stream Slides: \ts\n  Projected Slides: \tp\n  Regular Slides: \tr\n  Quit: \t\tq\n")
    print("====================================================================\n")

    mode = ""
    while (mode != "q"):
        pptType = SlideMaker.PPTMode.Null
        mode = input("Input: ")
        if (len(mode) > 0):
            if (mode == "s"):
                pptType = SlideMaker.PPTMode.Stream
            elif (mode == "p"):
                pptType = SlideMaker.PPTMode.Projected
            elif (mode == "r"):
                pptType = SlideMaker.PPTMode.Regular
            elif (mode == "-t"):
                inputStr = r' '.join(sys.argv[2:])
                print(f"String Visual Length of '{inputStr}' = {Utility.getVisualLength(inputStr)} Units")

        if (pptType != SlideMaker.PPTMode.Null):
            start = time.time()

            print("\nINITIALIZING...")
            sm: SlideMaker = SlideMaker()

            sm.setType(pptType)

            strType = pptType.name

            print(f"CREATING {strType.upper()} SLIDES...")
            print("  sundayServiceSlide() : ", sm.sundayServiceSlide())
            print("  monthlyScriptureSlide() : ", sm.monthlyScriptureSlide())
            print("  announcementSlide() : ", sm.announcementSlide())
            print("  bibleVerseMemorizationSlide() : ", sm.bibleVerseMemorizationSlide())
            print("  catechismSlide() : ", sm.catechismSlide())
            print("  bibleVerseMemorizationNextWeekSlide() : ", sm.bibleVerseMemorizationSlide(True))
            print("  catechismNextWeekSlide() : ", sm.catechismSlide(True))
            print("  supplicationSlide() : ", sm.supplicationSlide())
            print("  worshipSlide() : ", sm.worshipSlide())
            print("  callToWorshipSlide() : ", sm.callToWorshipSlide())
            print("  hymnSlides(1) : ", sm.hymnSlides(1))
            print("  prayerOfConfessionSlide() : ", sm.prayerOfConfessionSlide())
            print("  hymnSlides(2) : ", sm.hymnSlides(2))
            print("  holyCommunionSlide() : ", sm.holyCommunionSlide())
            print("  sermonHeaderSlide() : ", sm.sermonHeaderSlide())
            print("  sermonVerseSlide() : ", sm.sermonVerseSlide())
            print("  hymnSlides(3) : ", sm.hymnSlides(3))
            print("  hymnSlides(4) : ", sm.hymnSlides(4))

            print(f"OPENING {strType.upper()} SLIDES...")
            sm._openSlideInBrowser()

            print(f"\nTask completed in {(time.time() - start):.2f} seconds.\n")
            print("====================================================================\n")
