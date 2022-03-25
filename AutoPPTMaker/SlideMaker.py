import configparser
import os
import sys
import time
import traceback
import numpy

from typing import Dict, List, Tuple
from enum import Enum

from HymnMaker import HymnMaker
from GoogleAPITools import GoogleAPITools
from ListMaker import ListMaker
from Logging import Logging
from Utility import Utility
from VerseMaker import VerseMaker

"""

Core class of the SlideMaker application, assembles all other classes to generate Google Slides from an user-defined input file.

"""


class SlideMaker:
    class PPTMode(Enum):
        # Specifies the PPT output format
        Null = -1
        Stream = 0
        Projected = 1
        Regular = 2

    class ModuleStatus(Enum):
        Done = 0
        Failed = 1
        Disabled = 2

    def __init__(self, seed: int) -> None:
        # Seeded RNG for consistency within this run instance
        self.rng = numpy.random.RandomState(seed)

        # Access slide input data
        if not os.path.exists("SlideInputs.ini"):
            raise IOError("ERROR : SlideInputs.ini input file cannot be found.")
        self.input = configparser.ConfigParser()
        self.input.read("SlideInputs.ini")

        self.slideComponentCaller = {
            "SundayServiceHeader":          self.sundayServiceSlide,
            "MonthlyScripture":             self.monthlyScriptureSlide,
            "Announcements":                self.announcementSlide,
            "BibleMemorization":            self.bibleVerseMemorizationSlide,
            "Catechism":                    self.catechismSlide,
            "BibleMemorizationNextWeek":    self.bibleVerseMemorizationSlide,
            "CatechismNextWeek":            self.catechismSlide,
            "Supplications":                self.supplicationSlide,
            "WorshipHeader":                self.worshipSlide,
            "CallToWorship":                self.callToWorshipSlide,
            "Hymn1":                        self.hymnSlide,
            "PrayerOfConfession":           self.prayerOfConfessionSlide,
            "LordsPrayer":                  self.lordsPrayerSlide,
            "Hymn2":                        self.hymnSlide,
            "HolyCommunion":                self.holyCommunionSlide,
            "ApostleCreed":                 self.apostleCreedSlide,
            "SermonHeader":                 self.sermonHeaderSlide,
            "SermonVerse":                  self.sermonVerseSlide,
            "Hymn3":                        self.hymnSlide,
            "Offering":                     self.offeringSlide,
            "Hymn4":                        self.hymnSlide,
            "Doxology":                     self.doxologySlide,
            "Benediction":                  self.benedictionSlide,
            "NextWeekSchedule":             self.nextWeekScheduleSlide,
        }

        # Slide component should have 0 or 1 arguments
        self.slideComponentArgument = {
            "BibleMemorizationNextWeek":    True,
            "CatechismNextWeek":            True,
            "Hymn1":                        1,
            "Hymn2":                        2,
            "Hymn3":                        3,
            "Hymn4":                        4,
        }

    def setType(self, pptType: PPTMode) -> None:
        strType = pptType.name

        self.gEditor = GoogleAPITools(strType)
        self.verseMaker = VerseMaker(strType)
        self.hymnMaker = HymnMaker(strType)

        # Initialize logging
        Logging.initializeLoggingService()

        if not os.path.exists("Data/" + strType + "SlideProperties.ini"):
            raise IOError(f"ERROR : {strType}SlideProperties.ini config file cannot be found.")

        # Access slide property data
        self.config = configparser.ConfigParser()
        self.config.read("Data/" + strType + "SlideProperties.ini")

        # General linespacing for all slides
        self.lineSpacing = int(self.config["SLIDE_PROPERTIES"]["SlideLineSpacing"])

        # Hymn slide randomizer
        self.hymnSlideIndex = numpy.arange(int(self.config["HYMN_PROPERTIES"]["HymnStartIndex"]), self.gEditor.getPresentationLength())
        self.rng.shuffle(self.hymnSlideIndex)

        print(f"CREATING {strType.upper()} SLIDES...")

    # ======================================================================================================
    # ========================================= MASTER METHOD ==============================================
    # ======================================================================================================

    def createSlide(self) -> bool:
        # Access global data
        globalConfig = configparser.ConfigParser()
        globalConfig.read("Data/GlobalProperties.ini")

        # Get ordering mode
        slideOrderModeValue = self.input["SLIDE_ORDERING_MODE"]["SlideOrderMode"]

        # Check if ordering mode exists
        if not globalConfig.has_section(slideOrderModeValue):
            raise IOError(f"ERROR : Slide ordering mode [{slideOrderModeValue}] cannot be found.")

        # Read the entire global config section
        slideOrdering = dict(globalConfig.items(slideOrderModeValue))

        for i in range(len(slideOrdering.keys()), 0, -1):
            # Check if key number exists
            if not str(i) in slideOrdering:
                raise IOError(f"ERROR : Ordering number [{i}] does not exist.")

            slideType = slideOrdering[str(i)]
            slideIDList = []
            status = self.ModuleStatus.Failed

            # Check if slideType exists:
            if not slideType in self.slideComponentCaller:
                raise IOError(f"ERROR : Slide type [{slideType}] does not exist.")

            # Check for extra method arguments
            if slideType in self.slideComponentArgument:
                status, slideIDList = self.slideComponentCaller[slideType](self.slideComponentArgument[slideType])  # type: ignore
            else:
                status, slideIDList = self.slideComponentCaller[slideType]()  # type: ignore

            # Move slides to beginning
            if slideIDList:
                self.gEditor.moveSlideSet(slideIDList, 0)

            print(f"  {slideType} : ".ljust(45), status.name)

        # Get rid of template slides (recall commit doesn't been called yet)
        for i in range(self.gEditor.getPresentationLength()):
            self.gEditor.deleteSlide(self.gEditor.getSlideID(i))

        # Commit slide modifications
        return self.gEditor.commitSlideChanges()

    # ======================================================================================================
    # ====================================== SLIDE COMPONENT MAKERS ========================================
    # ======================================================================================================

    def sundayServiceSlide(self) -> Tuple[ModuleStatus, List[str]]:
        title = self.input["SUNDAY_SERVICE_HEADER"]["SundayServiceHeaderTitle"].upper()

        slideIDList = self._titleSlide(title, "SUNDAY_SERVICE_HEADER_PROPERTIES", "SundayServiceHeader")
        status = self.ModuleStatus.Done if slideIDList else self.ModuleStatus.Failed
        return (status, slideIDList)

    def monthlyScriptureSlide(self) -> Tuple[ModuleStatus, List[str]]:
        title = self.input["MONTHLY_SCRIPTURE"]["MonthlyScriptureTitle"].upper()
        source = self.input["MONTHLY_SCRIPTURE"]["MonthlyScriptureSource"].upper()
        maxLineLength = int(self.config["MONTHLY_SCRIPTURE_PROPERTIES"]["MonthlyScriptureMaxLineLength"])

        slideIDList = self._scriptureSingleSlide(title, source, maxLineLength,
                                                 "MONTHLY_SCRIPTURE_PROPERTIES", "MonthlyScripture")
        status = self.ModuleStatus.Done if slideIDList else self.ModuleStatus.Failed
        return (status, slideIDList)

    def announcementSlide(self) -> Tuple[ModuleStatus, List[str]]:
        autoRetrieve = self.input["ANNOUNCEMENTS"]["AnnouncementsAutoRetrieve"].upper()
        title = self.input["ANNOUNCEMENTS"]["AnnouncementsTitle"].upper()

        if (autoRetrieve != "TRUE"):
            slideIDList = self._headerOnlySlide(title, "ANNOUNCEMENTS_PROPERTIES", "Announcements")
            status = self.ModuleStatus.Done if slideIDList else self.ModuleStatus.Failed
            return (status, slideIDList)
        else:
            announcementList = self.gEditor.getAnnouncements()

            slideIDList = self._textMultiSlide(title, announcementList, "ANNOUNCEMENTS_PROPERTIES", "Announcements")
            status = self.ModuleStatus.Done if slideIDList else self.ModuleStatus.Failed
            return (status, slideIDList)

    def bibleVerseMemorizationSlide(self, nextWeek: bool = False) -> Tuple[ModuleStatus, List[str]]:
        lastWeekString = "LastWeek" if not nextWeek else ""
        title = self.input["BIBLE_MEMORIZATION"]["BibleMemorization" + lastWeekString + "Title"].upper()
        source = self.input["BIBLE_MEMORIZATION"]["BibleMemorization" + lastWeekString + "Source"].upper()
        maxLineLength = int(self.config["BIBLE_MEMORIZATION_PROPERTIES"]["BibleMemorizationMaxLineLength"])

        slideIDList = self._scriptureSingleSlide(title, source, maxLineLength,
                                                 "BIBLE_MEMORIZATION_PROPERTIES", "BibleMemorization", nextWeek)
        status = self.ModuleStatus.Done if slideIDList else self.ModuleStatus.Failed
        return (status, slideIDList)

    def catechismSlide(self, nextWeek: bool = False) -> Tuple[ModuleStatus, List[str]]:
        title = self.input["CATECHISM"]["Catechism" + ("LastWeek" if not nextWeek else "") + "Title"].upper()

        slideIDList = self._headerOnlySlide(title, "CATECHISM_PROPERTIES", "Catechism", nextWeek)
        status = self.ModuleStatus.Done if slideIDList else self.ModuleStatus.Failed
        return (status, slideIDList)

    def supplicationSlide(self) -> Tuple[ModuleStatus, List[str]]:
        enabled = self.input["SUPPLICATIONS"]["SupplicationsEnabled"].upper()
        autoRetrieve = self.input["SUPPLICATIONS"]["SupplicationAutoRetrieve"].upper()
        title = self.input["SUPPLICATIONS"]["SupplicationsTitle"].upper()

        if (enabled != "TRUE"):
            return (self.ModuleStatus.Disabled, [])

        if (autoRetrieve != "TRUE"):
            slideIDList = self._headerOnlySlide(title, "SUPPLICATIONS_PROPERTIES", "Supplications")
            status = self.ModuleStatus.Done if slideIDList else self.ModuleStatus.Failed
            return (status, slideIDList)
        else:
            supplicationList = self.gEditor.getSupplications()

            slideIDList = self._numListMultiSlide(title, supplicationList, "SUPPLICATIONS_PROPERTIES", "Supplications")
            status = self.ModuleStatus.Done if slideIDList else self.ModuleStatus.Failed
            return (status, slideIDList)

    def worshipSlide(self) -> Tuple[ModuleStatus, List[str]]:
        title = self.input["WORSHIP_HEADER"]["WorshipHeaderTitle"].upper()

        slideIDList = self._titleSlide(title, "WORSHIP_HEADER_PROPERTIES", "WorshipHeader")
        status = self.ModuleStatus.Done if slideIDList else self.ModuleStatus.Failed
        return (status, slideIDList)

    def callToWorshipSlide(self) -> Tuple[ModuleStatus, List[str]]:
        source = self.input["CALL_TO_WORSHIP"]["CallToWorshipSource"].upper()
        enabled = self.input["CALL_TO_WORSHIP"]["CallToWorshipEnabled"].upper()
        maxLineLength = int(self.config["CALL_TO_WORSHIP_PROPERTIES"]["CallToWorshipMaxLineLength"])
        maxLinesPerSlide = int(self.config["CALL_TO_WORSHIP_PROPERTIES"]["CallToWorshipMaxLines"])

        # Do not generate section if not enabled
        if (enabled != "TRUE" or source == ""):
            return (self.ModuleStatus.Disabled, [])

        slideIDList = self._scriptureMultiSlide(source, maxLineLength, maxLinesPerSlide,
                                                "CALL_TO_WORSHIP_PROPERTIES", "CallToWorship")
        status = self.ModuleStatus.Done if slideIDList else self.ModuleStatus.Failed
        return (status, slideIDList)

    def lordsPrayerSlide(self) -> Tuple[ModuleStatus, List[str]]:
        slideIDList = self._staticSlides("MISC_PROPERTIES", "LordsPrayer")
        return (self.ModuleStatus.Done, slideIDList)

    def prayerOfConfessionSlide(self) -> Tuple[ModuleStatus, List[str]]:
        title = self.input["PRAYER_OF_CONFESSION"]["PrayerOfConfessionTitle"].upper()
        source = self.input["PRAYER_OF_CONFESSION"]["PrayerOfConfessionSource"].upper()
        enabled = self.input["PRAYER_OF_CONFESSION"]["PrayerOfConfessionEnabled"].upper()
        maxLineLength = int(self.config["PRAYER_OF_CONFESSION_PROPERTIES"]["PrayerOfConfessionMaxLineLength"])

        # Do not generate section if not enabled
        if (enabled != "TRUE" or source == ""):
            return (self.ModuleStatus.Disabled, [])

        slideIDList = self._scriptureSingleSlide(title, source, maxLineLength,
                                                 "PRAYER_OF_CONFESSION_PROPERTIES", "PrayerOfConfession")
        status = self.ModuleStatus.Done if slideIDList else self.ModuleStatus.Failed
        return (status, slideIDList)

    def holyCommunionSlide(self) -> Tuple[ModuleStatus, List[str]]:
        enabled = self.input["HOLY_COMMUNION"]["HolyCommunionSlidesEnabled"].upper()

        # Do not generate section if not enabled
        if (enabled != "TRUE"):
            return (self.ModuleStatus.Disabled, [])

        slideIDList = self._staticSlides("MISC_PROPERTIES", "HolyCommunion")
        return (self.ModuleStatus.Done, slideIDList)

    def apostleCreedSlide(self) -> Tuple[ModuleStatus, List[str]]:
        slideIDList = self._staticSlides("MISC_PROPERTIES", "ApostleCreed")
        return (self.ModuleStatus.Done, slideIDList)

    def sermonHeaderSlide(self) -> Tuple[ModuleStatus, List[str]]:
        title = self.input["SERMON_HEADER"]["SermonHeaderTitle"].upper()
        speaker = self.input["SERMON_HEADER"]["SermonHeaderSpeaker"]

        slideIDList = self._sermonHeaderSlide(title, speaker, "SERMON_HEADER_PROPERTIES", "SermonHeader")
        status = self.ModuleStatus.Done if slideIDList else self.ModuleStatus.Failed
        return (status, slideIDList)

    def sermonVerseSlide(self) -> Tuple[ModuleStatus, List[str]]:
        # Allow for multiple unique sources separated by ","
        sources = self.input["SERMON_VERSE"]["SermonVerseSource"].upper()
        sourceList = sources.split(",")
        maxLineLength = int(self.config["SERMON_VERSE_PROPERTIES"]["SermonVerseMaxLineLength"])
        maxLinesPerSlide = int(self.config["SERMON_VERSE_PROPERTIES"]["SermonVerseMaxLines"])

        # Iterate and generate the slides by each unique source
        result = True
        slideIDList: List[str] = []
        for i in range(len(sourceList) - 1, -1, -1):
            slideIDListSub = self._scriptureMultiSlide(sourceList[i].strip(), maxLineLength, maxLinesPerSlide, "SERMON_VERSE_PROPERTIES", "SermonVerse")

            if (not slideIDListSub and result):
                result = False

            slideIDList = slideIDListSub + slideIDList

        status = self.ModuleStatus.Done if result else self.ModuleStatus.Failed
        return (status, slideIDList)

    def offeringSlide(self) -> Tuple[ModuleStatus, List[str]]:
        slideIDList = self._staticSlides("MISC_PROPERTIES", "Offering")
        return (self.ModuleStatus.Done, slideIDList)

    def hymnSlide(self, number: int) -> Tuple[ModuleStatus, List[str]]:
        source = self.input["HYMN"][f"Hymn{number}Source"]
        enabled = self.input["HYMN"][f"Hymn{number}Enabled"].upper()

        # Slide randomizer
        slideIndex = self.hymnSlideIndex[number - 1]

        # Do not generate section if not enabled
        if (enabled != "TRUE" or source == ""):
            return (self.ModuleStatus.Disabled, [])

        slideIDList = self._hymnSlide(source, slideIndex)
        status = self.ModuleStatus.Done if slideIDList else self.ModuleStatus.Failed
        return (status, slideIDList)

    def doxologySlide(self) -> Tuple[ModuleStatus, List[str]]:
        slideIDList = self._staticSlides("MISC_PROPERTIES", "Doxology")
        return (self.ModuleStatus.Done, slideIDList)

    def benedictionSlide(self) -> Tuple[ModuleStatus, List[str]]:
        slideIDList = self._staticSlides("MISC_PROPERTIES", "Benediction")
        return (self.ModuleStatus.Done, slideIDList)

    def nextWeekScheduleSlide(self) -> Tuple[ModuleStatus, List[str]]:
        propertyName = "NEXT_WEEK_SCHEDULE_PROPERTIES"
        dataNameHeader = "NextWeekSchedule"

        slideIndex = int(self.config[propertyName][dataNameHeader + "Index"])
        tableIDList = self.gEditor.getTableID(slideIndex)

        if len(tableIDList) != 1:
            raise AssertionError(f"\tERROR : Unusual Next Week's Schedule Slide detected with [{len(tableIDList)}] tables")

        # Update date cell
        (dateString, cordinalIndex) = self.gEditor.getFormattedNextSundayDate(self.gEditor.DateFormatMode.Full, True)

        self._insertTextInTable(tableID=tableIDList[0],
                                text=dateString[:-5] + "," + dateString[-5:],
                                size=int(self.config["NEXT_WEEK_SCHEDULE_PROPERTIES"][dataNameHeader + "DateTextSize"]),
                                bold=self._str2bool(self.config[propertyName][dataNameHeader + "DateBolded"]),
                                italic=self._str2bool(self.config[propertyName][dataNameHeader + "DateItalicized"]),
                                underlined=self._str2bool(self.config[propertyName][dataNameHeader + "DateUnderlined"]),
                                alignment=self.config[propertyName][dataNameHeader + "DateAlignment"],
                                rowIndex=0,
                                colIndex=1)
        self.gEditor.setTextSuperScriptInTable(tableIDList[0], cordinalIndex, cordinalIndex + 2, 0, 1)

        slideIDList = self._staticSlides(propertyName, dataNameHeader)
        return (self.ModuleStatus.Done, slideIDList)

    # ======================================================================================================
    # ===================================== SLIDE MAKER IMPLEMENTATIONS ====================================
    # ======================================================================================================

    def _titleSlide(self, title: str, propertyName: str, dataNameHeader: str) -> List[str]:
        (dateString, cordinalIndex) = self.gEditor.getFormattedNextSundayDate(self.gEditor.DateFormatMode.Full)
        slideIndex = int(self.config[propertyName][dataNameHeader + "Index"])

        # Generate create duplicate request
        objIDMappingList = self._duplicateSlide(1, slideIndex)

        refData = self.gEditor.getSlideTextData(slideIndex)
        refSlideID = self.gEditor.getSlideID(slideIndex)
        slideIDList = [objIDMappingList[0][refSlideID]]

        checkPoint = [False, False]
        try:
            for item in refData:
                if "{Title}" in item[1]:
                    checkPoint[0] = True
                    self._insertText(
                        objectID=objIDMappingList[0][item[0]],
                        text=title.upper(),
                        size=int(self.config[propertyName][dataNameHeader + "TitleTextSize"]),
                        bold=self._str2bool(self.config[propertyName][dataNameHeader + "TitleBolded"]),
                        italic=self._str2bool(self.config[propertyName][dataNameHeader + "TitleItalicized"]),
                        underlined=self._str2bool(self.config[propertyName][dataNameHeader + "TitleUnderlined"]),
                        alignment=self.config[propertyName][dataNameHeader + "TitleAlignment"])
                elif "{Text}" in item[1]:
                    checkPoint[1] = True
                    self._insertText(
                        objectID=objIDMappingList[0][item[0]],
                        text=dateString[:-5] + "," + dateString[-5:],   # Add comma between year and date
                        size=int(self.config[propertyName][dataNameHeader + "DateTextSize"]),
                        bold=self._str2bool(self.config[propertyName][dataNameHeader + "DateBolded"]),
                        italic=self._str2bool(self.config[propertyName][dataNameHeader + "DateItalicized"]),
                        underlined=self._str2bool(self.config[propertyName][dataNameHeader + "DateUnderlined"]),
                        alignment=self.config[propertyName][dataNameHeader + "DateAlignment"])

                    self.gEditor.setTextSuperScript(objIDMappingList[0][item[0]], cordinalIndex, cordinalIndex + 2)

        except Exception:
            print(f"\tERROR : {traceback.format_exc()}")
            return []

        if (False in checkPoint):
            print(f"\tERROR : {checkPoint.count(True)} out of {len(checkPoint)} text placeholders found.")
            return []

        return slideIDList

    def _hymnSlide(self, source: str, slideIndex: int) -> List[str]:
        if (not self.hymnMaker.setSource(source)):
            print(f"\tERROR : Hymn [{source}] not found.")
            return []

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

        # Generate create duplicate request
        objIDMappingList = self._duplicateSlide(len(titleList), slideIndex)

        refData = self.gEditor.getSlideTextData(slideIndex)
        refSlideID = self.gEditor.getSlideID(slideIndex)
        slideIDList = []

        # Insert title and lyrics data
        for i in range(len(titleList)):
            # Save slide ID
            slideIDList.append(objIDMappingList[i][refSlideID])

            checkPoint = [False, False]
            try:
                for item in refData:
                    if ("{Title}" in item[1]):
                        checkPoint[0] = True
                        self._insertText(
                            objectID=objIDMappingList[i][item[0]],
                            text=titleList[i],
                            size=titleFontSize,
                            bold=self._str2bool(self.config["HYMN_PROPERTIES"]["HymnTitleBolded"]),
                            italic=self._str2bool(self.config["HYMN_PROPERTIES"]["HymnTitleItalicized"]),
                            underlined=self._str2bool(self.config["HYMN_PROPERTIES"]["HymnTitleUnderlined"]),
                            alignment=self.config["HYMN_PROPERTIES"]["HymnTitleAlignment"])
                    elif ("{Text}" in item[1]):
                        checkPoint[1] = True
                        self._insertText(
                            objectID=objIDMappingList[i][item[0]],
                            text=formattedLyricsList[i],
                            size=int(self.config["HYMN_PROPERTIES"]["HymnTextSize"]),
                            bold=self._str2bool(self.config["HYMN_PROPERTIES"]["HymnBolded"]),
                            italic=self._str2bool(self.config["HYMN_PROPERTIES"]["HymnItalicized"]),
                            underlined=self._str2bool(self.config["HYMN_PROPERTIES"]["HymnUnderlined"]),
                            alignment=self.config["HYMN_PROPERTIES"]["HymnAlignment"])

                        if (multiLineTitle):
                            self.gEditor.updatePageElementTransform(objIDMappingList[i][item[0]], translateY=int(self.config["HYMN_PROPERTIES"]["HymnLoweredUnitHeight"]))
            except Exception:
                print(f"\tERROR : {traceback.format_exc()}")
                return []

            if (False in checkPoint):
                print(f"\tERROR : {checkPoint.count(True)} out of {len(checkPoint)} text placeholders found.")
                return []

        return slideIDList

    def _scriptureSingleSlide(self, title: str, source: str, charPerLine: int, propertyName: str, dataNameHeader: str, nextWeek: bool = False) -> List[str]:
        # Assumes monthly scripture is short enough to fit in one slide
        nextWeekString = "NextWeek" if nextWeek else ""
        slideIndex = int(self.config[propertyName][dataNameHeader + nextWeekString + "Index"])

        # Generate slide anyway if source verse cannot be found
        if not self.verseMaker.setSource(source, charPerLine):
            print(f"\tWARNING: Verse [{source}] not found.")
            verseString = "{Text}"
            ssIndexList = []
        else:
            [verseString, ssIndexList] = self.verseMaker.getVerseString()

        # Generate create duplicate request
        objIDMappingList = self._duplicateSlide(1, slideIndex)

        refData = self.gEditor.getSlideTextData(slideIndex)
        refSlideID = self.gEditor.getSlideID(slideIndex)
        slideIDList = [objIDMappingList[0][refSlideID]]

        checkPoint = [False, False]
        try:
            for item in refData:
                if "{Title}" in item[1]:
                    checkPoint[0] = True
                    self._insertText(
                        objectID=objIDMappingList[0][item[0]],
                        text=title,
                        size=int(self.config[propertyName][dataNameHeader + nextWeekString + "TitleTextSize"]),
                        bold=self._str2bool(self.config[propertyName][dataNameHeader + "TitleBolded"]),
                        italic=self._str2bool(self.config[propertyName][dataNameHeader + "TitleItalicized"]),
                        underlined=self._str2bool(self.config[propertyName][dataNameHeader + "TitleUnderlined"]),
                        alignment=self.config[propertyName][dataNameHeader + "TitleAlignment"])
                elif "{Text}" in item[1]:
                    checkPoint[1] = True
                    self._insertText(
                        objectID=objIDMappingList[0][item[0]],
                        text=verseString,
                        size=int(self.config[propertyName][dataNameHeader + "TextSize"]),
                        bold=self._str2bool(self.config[propertyName][dataNameHeader + "Bolded"]),
                        italic=self._str2bool(self.config[propertyName][dataNameHeader + "Italicized"]),
                        underlined=self._str2bool(self.config[propertyName][dataNameHeader + "Underlined"]),
                        alignment=self.config[propertyName][dataNameHeader + "Alignment"])

                    # Superscript and bold the verse numbers, and set paragraph spacing between the verses
                    for ssRange in ssIndexList:
                        self.gEditor.setTextSuperScript(objIDMappingList[0][item[0]], ssRange[1], ssRange[2])
                        self.gEditor.setBold(objIDMappingList[0][item[0]], ssRange[1], ssRange[2])
                        if (ssRange[1] > 0 and ssRange[0] == self.verseMaker.SSType.VerseNumber):  # No need for paragraph spacing in initial line
                            self.gEditor.setSpaceAbove(objIDMappingList[0][item[0]], ssRange[1], 10)
        except Exception:
            print(f"\tERROR : {traceback.format_exc()}")
            return []

        if (False in checkPoint):
            print(f"\tERROR : {checkPoint.count(True)} out of {len(checkPoint)} text placeholders found.")
            return []

        return slideIDList

    def _scriptureMultiSlide(self, source: str, maxLineLength: int, maxLinesPerSlide: int, propertyName: str, dataNameHeader: str) -> List[str]:
        # Assumes monthly scripture is short enough to fit in one slide
        if (not self.verseMaker.setSource(source, maxLineLength)):
            print(f"\tERROR : Verse {source} not found.")
            return []

        [title, slideVersesList, slideSSIndexList] = self.verseMaker.getVerseStringMultiSlide(maxLinesPerSlide)

        # Generate create duplicate request and commit it
        slideIndex = int(self.config[propertyName][dataNameHeader + "Index"])

        # Generate create duplicate request
        objIDMappingList = self._duplicateSlide(len(slideVersesList), slideIndex)

        refData = self.gEditor.getSlideTextData(slideIndex)
        refSlideID = self.gEditor.getSlideID(slideIndex)
        slideIDList = []

        # Get units of paragraph that separates verse numbers
        paragraphSpace = int(self.config["VERSE_PROPERTIES"]["VerseParagraphSpace"])

        # Insert title and verses data
        for i in range(len(slideVersesList)):
            # Save slide ID
            slideIDList.append(objIDMappingList[i][refSlideID])

            checkPoint = [False, False]
            try:
                for item in refData:
                    if "{Title}" in item[1]:
                        checkPoint[0] = True
                        self._insertText(
                            objectID=objIDMappingList[i][item[0]],
                            text=title,
                            size=int(self.config[propertyName][dataNameHeader + "TitleTextSize"]),
                            bold=self._str2bool(self.config[propertyName][dataNameHeader + "TitleBolded"]),
                            italic=self._str2bool(self.config[propertyName][dataNameHeader + "TitleItalicized"]),
                            underlined=self._str2bool(self.config[propertyName][dataNameHeader + "TitleUnderlined"]),
                            alignment=self.config[propertyName][dataNameHeader + "TitleAlignment"])
                    elif "{Text}" in item[1]:
                        checkPoint[1] = True
                        self._insertText(
                            objectID=objIDMappingList[i][item[0]],
                            text=slideVersesList[i],
                            size=int(self.config[propertyName][dataNameHeader + "TextSize"]),
                            bold=self._str2bool(self.config[propertyName][dataNameHeader + "Bolded"]),
                            italic=self._str2bool(self.config[propertyName][dataNameHeader + "Italicized"]),
                            underlined=self._str2bool(self.config[propertyName][dataNameHeader + "Underlined"]),
                            alignment=self.config[propertyName][dataNameHeader + "Alignment"])

                        # Superscript and bold the verse numbers, and set paragraph spacing between the verses
                        for ssRange in slideSSIndexList[i]:
                            self.gEditor.setTextSuperScript(objIDMappingList[i][item[0]], ssRange[1], ssRange[2])
                            self.gEditor.setBold(objIDMappingList[i][item[0]], ssRange[1], ssRange[2])
                            if (ssRange[1] > 0 and ssRange[0] == self.verseMaker.SSType.VerseNumber):  # No need for paragraph spacing in initial line
                                self.gEditor.setSpaceAbove(objIDMappingList[i][item[0]], ssRange[1], paragraphSpace)
            except Exception:
                print(f"\tERROR : {traceback.format_exc()}")
                return []

            if (False in checkPoint):
                print(f"\tERROR : {checkPoint.count(True)} out of {len(checkPoint)} text placeholders found.")
                return []

        return slideIDList

    def _sermonHeaderSlide(self, title: str, speaker: str, propertyName: str, dataNameHeader: str) -> List[str]:
        slideIndex = int(self.config[propertyName][dataNameHeader + "Index"])

        # Generate create duplicate request
        objIDMappingList = self._duplicateSlide(1, slideIndex)

        refData = self.gEditor.getSlideTextData(slideIndex)
        refSlideID = self.gEditor.getSlideID(slideIndex)
        slideIDList = [objIDMappingList[0][refSlideID]]

        checkPoint = [False, False]
        try:
            for item in refData:
                if "{Title}" in item[1]:
                    checkPoint[0] = True
                    self._insertText(
                        objectID=objIDMappingList[0][item[0]],
                        text=title.upper(),
                        size=int(self.config[propertyName][dataNameHeader + "TitleTextSize"]),
                        bold=self._str2bool(self.config[propertyName][dataNameHeader + "TitleBolded"]),
                        italic=self._str2bool(self.config[propertyName][dataNameHeader + "TitleItalicized"]),
                        underlined=self._str2bool(self.config[propertyName][dataNameHeader + "TitleUnderlined"]),
                        alignment=self.config[propertyName][dataNameHeader + "TitleAlignment"])
                elif "{Text}" in item[1]:
                    checkPoint[1] = True
                    self._insertText(
                        objectID=objIDMappingList[0][item[0]],
                        text=speaker.title(),
                        size=int(self.config[propertyName][dataNameHeader + "SpeakerTextSize"]),
                        bold=self._str2bool(self.config[propertyName][dataNameHeader + "SpeakerBolded"]),
                        italic=self._str2bool(self.config[propertyName][dataNameHeader + "SpeakerItalicized"]),
                        underlined=self._str2bool(self.config[propertyName][dataNameHeader + "SpeakerUnderlined"]),
                        alignment=self.config[propertyName][dataNameHeader + "SpeakerAlignment"])
        except Exception:
            print(f"\tERROR : {traceback.format_exc()}")
            return []

        if (False in checkPoint):
            print(f"\tERROR : {checkPoint.count(True)} out of {len(checkPoint)} text placeholders found.")
            return []

        return slideIDList

    def _headerOnlySlide(self, title: str, propertyName: str, dataNameHeader: str, nextWeek: bool = False) -> List[str]:
        nextWeekString = "NextWeek" if nextWeek else ""
        slideIndex = int(self.config[propertyName][dataNameHeader + nextWeekString + "Index"])

        # Generate create duplicate request
        objIDMappingList = self._duplicateSlide(1, slideIndex)

        refData = self.gEditor.getSlideTextData(slideIndex)
        refSlideID = self.gEditor.getSlideID(slideIndex)
        slideIDList = [objIDMappingList[0][refSlideID]]

        checkPoint = [False]
        try:
            for item in refData:
                if "{Title}" in item[1]:
                    checkPoint[0] = True
                    self._insertText(
                        objectID=objIDMappingList[0][item[0]],
                        text=title.upper(),
                        size=int(self.config[propertyName][dataNameHeader + nextWeekString + "TitleTextSize"]),
                        bold=self._str2bool(self.config[propertyName][dataNameHeader + "TitleBolded"]),
                        italic=self._str2bool(self.config[propertyName][dataNameHeader + "TitleItalicized"]),
                        underlined=self._str2bool(self.config[propertyName][dataNameHeader + "TitleUnderlined"]),
                        alignment=self.config[propertyName][dataNameHeader + "TitleAlignment"])
        except Exception:
            print(f"\tERROR : {traceback.format_exc()}")
            return []

        if (False in checkPoint):
            print(f"\tERROR : {checkPoint.count(True)} out of {len(checkPoint)} text placeholders found.")
            return []

        return slideIDList

    def _textMultiSlide(self, title: str, textContentList: List[str], propertyName: str, dataNameHeader: str) -> List[str]:
        # For announcement slides
        slideIndex = int(self.config[propertyName][dataNameHeader + "Index"])

        # Generate create duplicate request
        objIDMappingList = self._duplicateSlide(len(textContentList), slideIndex)

        refData = self.gEditor.getSlideTextData(slideIndex)
        refSlideID = self.gEditor.getSlideID(slideIndex)
        slideIDList = []

        # Insert title and verses data
        for i in range(len(textContentList)):
            # Save slide ID
            slideIDList.append(objIDMappingList[i][refSlideID])

            checkPoint = [False, False]
            try:
                for item in refData:
                    if "{Title}" in item[1]:
                        checkPoint[0] = True
                        self._insertText(
                            objectID=objIDMappingList[i][item[0]],
                            text=title.upper(),
                            size=int(self.config[propertyName][dataNameHeader + "TitleTextSize"]),
                            bold=self._str2bool(self.config[propertyName][dataNameHeader + "TitleBolded"]),
                            italic=self._str2bool(self.config[propertyName][dataNameHeader + "TitleItalicized"]),
                            underlined=self._str2bool(self.config[propertyName][dataNameHeader + "TitleUnderlined"]),
                            alignment=self.config[propertyName][dataNameHeader + "TitleAlignment"])
                    elif "{Text}" in item[1]:
                        checkPoint[1] = True
                        self._insertText(
                            objectID=objIDMappingList[i][item[0]],
                            text=textContentList[i],
                            size=int(self.config[propertyName][dataNameHeader + "TextSize"]),
                            bold=self._str2bool(self.config[propertyName][dataNameHeader + "Bolded"]),
                            italic=self._str2bool(self.config[propertyName][dataNameHeader + "Italicized"]),
                            underlined=self._str2bool(self.config[propertyName][dataNameHeader + "Underlined"]),
                            alignment=self.config[propertyName][dataNameHeader + "Alignment"])
            except Exception:
                print(f"\tERROR : {traceback.format_exc()}")
                return []

            if (False in checkPoint):
                print(f"\tERROR : {checkPoint.count(True)} out of {len(checkPoint)} text placeholders found.")
                return []

        return slideIDList

    def _numListMultiSlide(self, title: str, contentList: List[str], propertyName: str, dataNameHeader: str) -> List[str]:
        # For supplication slides
        slideIndex = int(self.config[propertyName][dataNameHeader + "Index"])
        paragraphSpace = int(self.config[propertyName][dataNameHeader + "ParagraphSpace"])

        # Get slide separated list
        [slideContentList, listStartIndex] = ListMaker.getFormattedListElementsBySlide(contentList,
                                                                                       indentSpace=int(self.config[propertyName][dataNameHeader + "IndentSpace"]),
                                                                                       maxLineLength=int(self.config[propertyName][dataNameHeader + "TextMaxLineLength"]),
                                                                                       maxLinesPerSlide=int(self.config[propertyName][dataNameHeader + "TextMaxLines"]))

        # Generate create duplicate request
        objIDMappingList = self._duplicateSlide(len(slideContentList), slideIndex)

        refData = self.gEditor.getSlideTextData(slideIndex)
        refSlideID = self.gEditor.getSlideID(slideIndex)
        slideIDList = []

        # Insert title and verses data
        for i in range(len(slideContentList)):
            # Save slide ID
            slideIDList.append(objIDMappingList[i][refSlideID])

            checkPoint = [False, False]
            try:
                for item in refData:
                    if "{Title}" in item[1]:
                        checkPoint[0] = True
                        self._insertText(
                            objectID=objIDMappingList[i][item[0]],
                            text=title.upper(),
                            size=int(self.config[propertyName][dataNameHeader + "TitleTextSize"]),
                            bold=self._str2bool(self.config[propertyName][dataNameHeader + "TitleBolded"]),
                            italic=self._str2bool(self.config[propertyName][dataNameHeader + "TitleItalicized"]),
                            underlined=self._str2bool(self.config[propertyName][dataNameHeader + "TitleUnderlined"]),
                            alignment=self.config[propertyName][dataNameHeader + "TitleAlignment"])
                    elif "{Text}" in item[1]:
                        checkPoint[1] = True
                        self._insertText(
                            objectID=objIDMappingList[i][item[0]],
                            text=slideContentList[i],
                            size=int(self.config[propertyName][dataNameHeader + "TextSize"]),
                            bold=self._str2bool(self.config[propertyName][dataNameHeader + "Bolded"]),
                            italic=self._str2bool(self.config[propertyName][dataNameHeader + "Italicized"]),
                            underlined=self._str2bool(self.config[propertyName][dataNameHeader + "Underlined"]),
                            alignment=self.config[propertyName][dataNameHeader + "Alignment"])

                        # Add bullets
                        for lineStartIndex in listStartIndex[i]:
                            if lineStartIndex > 0:  # No need for paragraph spacing in initial line
                                self.gEditor.setSpaceAbove(objIDMappingList[i][item[0]], lineStartIndex, paragraphSpace)

            except Exception:
                print(f"\tERROR : {traceback.format_exc()}")
                return []

            if (False in checkPoint):
                print(f"\tERROR : {checkPoint.count(True)} out of {len(checkPoint)} text placeholders found.")
                return []

        return slideIDList

    def _staticSlides(self, propertyName: str, dataNameHeader: str) -> List[str]:
        slideIndexList = [int(index) for index in self.config[propertyName][dataNameHeader + "Index"].split(", ")]

        slideIDList = []
        for slideIndex in slideIndexList:
            # Generate create duplicate request
            objIDMappingList = self._duplicateSlide(1, slideIndex)
            refSlideID = self.gEditor.getSlideID(slideIndex)

            # Save slide ID
            slideIDList.append(objIDMappingList[0][refSlideID])

        return slideIDList

    # ======================================================================================================
    # ============================================ TOOLS ===================================================
    # ======================================================================================================

    def _deleteSlide(self, slideIndexList: List[int]) -> bool:
        for slideIndex in slideIndexList:
            sourceSlideID = self.gEditor.getSlideID(slideIndex)
            self.gEditor.deleteSlide(sourceSlideID)

        if (not self.gEditor.commitSlideChanges()):
            return False

        return True

    def _str2bool(self, boolString: str) -> bool:
        return boolString.lower() in ("yes", "true", "t", "1")

    def _openSlideInBrowser(self) -> None:
        self.gEditor.openSlideInBrowser()

    def _insertText(self, objectID: str, text: str, size: int, bold: bool, italic: bool, underlined: bool, alignment: str,
                    linespacing: int = -1, rgbColor: Tuple[float, float, float] = (1.0, 1.0, 1.0)) -> None:
        self.gEditor.setText(objectID, text)
        self.gEditor.setTextStyle(objectID, bold, italic, underlined, size)
        self.gEditor.setParagraphStyle(objectID, self.lineSpacing if linespacing < 0 else linespacing, alignment)
        self.gEditor.setTextColor(objectID, rgbColor)

    def _insertTextInTable(self, tableID: str, text: str, size: int, bold: bool, italic: bool, underlined: bool, alignment: str,
                           rowIndex: int, colIndex: int, linespacing: int = -1, rgbColor: Tuple[float, float, float] = (1.0, 1.0, 1.0)) -> None:
        self.gEditor.setTextInTable(tableID, text, rowIndex, colIndex)
        self.gEditor.setTextStyleInTable(tableID, bold, italic, underlined, size, rowIndex, colIndex)
        self.gEditor.setParagraphStyleInTable(tableID, self.lineSpacing if linespacing < 0 else linespacing, alignment, rowIndex, colIndex)
        self.gEditor.setTextColorInTable(tableID, rgbColor, rowIndex, colIndex)

    def _duplicateSlide(self, dupCount: int, slideIndex: int) -> List[Dict[str, str]]:
        # Generate extra duplicate slides, return object ID mapping from original to duplicated
        objIDMappingList: List[Dict[str, str]] = []
        for i in range(dupCount):
            objIDMappingList.insert(0, self.gEditor.duplicateSlide(slideIndex))

        return objIDMappingList

# ==============================================================================================
# ======================================= APPLICATION RUNNER ===================================
# ==============================================================================================


if __name__ == "__main__":
    print(f"""
====================================================================
                        SlideMaker v{Logging.VersionNumber}
====================================================================

Input Options:
  Livestreaming Slides:                 s
  Main Sanctuary Projector Slides:      p
  Lower Auditorium Projector Slides:    r
  Quit:                                 q

====================================================================
    """)

    mode = ""

    # Set seed so RNG is consist between instances
    seed = int(time.time())

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
                inputStr = r" ".join(sys.argv[2:])
                print(f"String Visual Length of '{inputStr}' = {Utility.getVisualLength(inputStr)} Units")

        if (pptType != SlideMaker.PPTMode.Null):
            start = time.time()

            print("\nINITIALIZING...")

            sm: SlideMaker = SlideMaker(seed)
            sm.setType(pptType)
            sm.createSlide()

            print(f"OPENING {pptType.name.upper()} SLIDES...")
            sm._openSlideInBrowser()

            print(f"\nTask completed in {(time.time() - start):.2f} seconds.\n")
            print("====================================================================\n")
