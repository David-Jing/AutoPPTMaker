import configparser
import os
import sys
import time
import matplotlib

from HymnMaker import HymnMaker
from PPTEditorTools import DateFormatMode
from PPTEditorTools import PPTEditorTools
from VerseMaker import SSType
from VerseMaker import VerseMaker
from enum import Enum
from matplotlib.afm import AFM


class PPTMode(Enum):
    # Specifies the PPT output format
    Stream = 0
    Projected = 1
    Regular = 2


class SlideMaker:
    def __init__(self):
        # Offset for when creating new slides
        self.slideOffset = 0

        # Access slide input data
        self.input = configparser.ConfigParser()
        self.input.read("SlideInputs.ini")

        # For finding visual lengths of text strings
        afm_filename = os.path.join(matplotlib.get_data_path(), 'fonts', 'afm', 'ptmr8a.afm')
        self.afm = AFM(open(afm_filename, "rb"))

    def setType(self, type):
        if (type == PPTMode.Stream):
            type = "Stream"
        elif (type == PPTMode.Projected):
            type = "Projected"
        elif (type == PPTMode.Regular):
            type = "Regular"
        else:
            raise ValueError(f"ERROR : PPTMode not recognized.")

        if (not os.path.exists(type + "SlideProperties.ini")):
            raise IOError(f"ERROR : {type}SlideProperties.ini config file does not exist.")

        self.pptEditor = PPTEditorTools(type)
        self.verseMaker = VerseMaker(type)
        self.hymnMaker = HymnMaker(type)

        # Access slide property data
        self.config = configparser.ConfigParser()
        self.config.read(type + "SlideProperties.ini")

        # General linespacing for all slides
        self.lineSpacing = self.config["SLIDE_PROPERTIES"]["SlideLineSpacing"]

    # ======================================================================================================
    # ========================================== SLIDE MAKERS ==============================================
    # ======================================================================================================

    def sundayServiceSlide(self):
        title = self.input["SUNDAY_SERVICE_HEADER"]["SundayServiceHeaderTitle"].upper()

        return self._titleSlide(title, "SUNDAY_SERVICE_HEADER_PROPERTIES", "SundayServiceHeader")

    def monthlyScriptureSlide(self):
        title = self.input["MONTHLY_SCRIPTURE"]["MonthlyScriptureTitle"].upper()
        source = self.input["MONTHLY_SCRIPTURE"]["MonthlyScriptureSource"].upper()
        maxLineLength = int(self.config["MONTHLY_SCRIPTURE_PROPERTIES"]["MonthlyScriptureMaxLineLength"])

        return self._scriptureSingleSlide(title, source, maxLineLength,
                                          "MONTHLY_SCRIPTURE_PROPERTIES", "MonthlyScripture")

    def announcementSlide(self):
        title = self.input["ANNOUNCEMENTS"]["AnnouncementsTitle"].upper()
        return self._headerOnlySlide(title, "ANNOUNCEMENTS_PROPERTIES", "Announcements")

    def bibleVerseMemorizationSlide(self, nextWeek=False):
        lastWeekString = "LastWeek" if not nextWeek else ""
        title = self.input["BIBLE_MEMORIZATION"]["BibleMemorization" + lastWeekString + "Title"].upper()
        source = self.input["BIBLE_MEMORIZATION"]["BibleMemorization" + lastWeekString + "Source"].upper()
        maxLineLength = int(self.config["BIBLE_MEMORIZATION_PROPERTIES"]["BibleMemorizationMaxLineLength"])

        return self._scriptureSingleSlide(title, source, maxLineLength,
                                          "BIBLE_MEMORIZATION_PROPERTIES", "BibleMemorization", nextWeek)

    def catechismSlide(self, nextWeek=False):
        title = self.input["CATECHISM"]["Catechism" + ("LastWeek" if not nextWeek else "") + "Title"].upper()
        return self._headerOnlySlide(title, "CATECHISM_PROPERTIES", "Catechism", nextWeek)

    def worshipSlide(self):
        title = self.input["WORSHIP_HEADER"]["WorshipHeaderTitle"].upper()

        return self._titleSlide(title, "WORSHIP_HEADER_PROPERTIES", "WorshipHeader")

    def callToWorshipSlide(self):
        source = self.input["CALL_TO_WORSHIP"]["CallToWorshipSource"].upper()
        maxLineLength = int(self.config["CALL_TO_WORSHIP_PROPERTIES"]["CallToWorshipMaxLineLength"])
        maxLinesPerSlide = int(self.config["CALL_TO_WORSHIP_PROPERTIES"]["CallToWorshipMaxLines"])

        # Delete section if source is empty
        if (source == ""):
            return self._deleteSlide(int(self.config["CALL_TO_WORSHIP_PROPERTIES"]["CallToWorshipIndex"]) + self.slideOffset)

        return self._scriptureMultiSlide(source, maxLineLength, maxLinesPerSlide,
                                         "CALL_TO_WORSHIP_PROPERTIES", "CallToWorship")

    def prayerOfConfessionSlide(self):
        title = self.input["PRAYER_OF_CONFESSION"]["PrayerOfConfessionTitle"].upper()
        source = self.input["PRAYER_OF_CONFESSION"]["PrayerOfConfessionSource"].upper()
        maxLineLength = int(self.config["PRAYER_OF_CONFESSION_PROPERTIES"]["PrayerOfConfessionMaxLineLength"])

        # Delete section if source is empty
        if (source == ""):
            return self._deleteSlide(int(self.config["PRAYER_OF_CONFESSION_PROPERTIES"]["PrayerOfConfessionIndex"]) + self.slideOffset)

        return self._scriptureSingleSlide(title, source, maxLineLength,
                                          "PRAYER_OF_CONFESSION_PROPERTIES", "PrayerOfConfession")

    def holyCommunionSlide(self):
        enabled = self.input["HOLY_COMMUNION"]["HolyCommunionEnabled"].upper()
        slideIndex = int(self.config["HOLY_COMMUNION_PROPERTIES"]["HolyCommunionIndex"]) + self.slideOffset
        numOfSlides = int(self.config["HOLY_COMMUNION_PROPERTIES"]["HolyCommunionSlides"])

        # If not enabled, delete section
        if (enabled != "TRUE"):
            for i in range(slideIndex, slideIndex + numOfSlides):
                sourceSlideID = self.pptEditor.getSlideID(i)
                self.pptEditor.deleteSlide(sourceSlideID)

            if (not self.pptEditor.commitSlideChanges()):
                return False

            self.slideOffset -= numOfSlides

            return "Disabled"

        return True

    def sermonHeaderSlide(self):
        title = self.input["SERMON_HEADER"]["SermonHeaderTitle"].upper()
        speaker = self.input["SERMON_HEADER"]["SermonHeaderSpeaker"]

        return self._sermonHeaderSlide(title, speaker, "SERMON_HEADER_PROPERTIES", "SermonHeader")

    def sermonVerseSlide(self):
        # Allow for multiple unique sources separated by ","
        sources = self.input["SERMON_VERSE"]["SermonVerseSource"].upper()
        sourceList = sources.split(",")
        maxLineLength = int(self.config["SERMON_VERSE_PROPERTIES"]["SermonVerseMaxLineLength"])
        maxLinesPerSlide = int(self.config["SERMON_VERSE_PROPERTIES"]["SermonVerseMaxLines"])

        # Generate templates for each verse source
        slideIndex = int(self.config["SERMON_VERSE_PROPERTIES"]["SermonVerseIndex"]) + self.slideOffset
        sourceSlideID = self.pptEditor.getSlideID(slideIndex)
        for i in range(len(sourceList) - 1):  # Recall that the original source slide remains
            self.pptEditor.duplicateSlide(
                sourceSlideID, sourceSlideID + '_d' + str(i))
        if (not self.pptEditor.commitSlideChanges()):
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

    def hymnSlides(self, number):
        hymnSource = self.input["HYMN"][f"Hymn{number}Source"]
        hymnIndex = int(self.config["HYMN_PROPERTIES"][f"Hymn{number}Index"]) + self.slideOffset

        # Delete section if source is empty
        if (hymnSource == ""):
            return self._deleteSlide(hymnIndex)

        return self._hymnSlide(hymnSource, hymnIndex)

    # ======================================================================================================
    # ===================================== SLIDE MAKER IMPLEMENTATIONS ====================================
    # ======================================================================================================

    def _titleSlide(self, title, propertyName, dataNameHeader):
        [dateString, cordinalIndex] = self.pptEditor.getFormattedNextSundayDate(DateFormatMode.Full)

        checkPoint = [False, False]
        try:
            data = self.pptEditor.getSlideTextData(int(self.config[propertyName][dataNameHeader + "Index"]) + self.slideOffset)
            for item in data:
                if '{Title}' in item[1]:
                    checkPoint[0] = True
                    self._insertText(
                        objectID=item[0],
                        text=title.upper(),
                        size=int(self.config[propertyName][dataNameHeader + "TitleTextSize"]),
                        bold=self.config[propertyName][dataNameHeader + "TitleBolded"],
                        italic=self.config[propertyName][dataNameHeader + "TitleItalicized"],
                        underlined=self.config[propertyName][dataNameHeader + "TitleUnderlined"],
                        alignment=self.config[propertyName][dataNameHeader + "TitleAlignment"])
                elif '{Text}' in item[1]:
                    checkPoint[1] = True
                    self._insertText(
                        objectID=item[0],
                        text=dateString[:-5] + "," + dateString[-5:],   # Add comma between year and date
                        size=int(self.config[propertyName][dataNameHeader + "DateTextSize"]),
                        bold=self.config[propertyName][dataNameHeader + "DateBolded"],
                        italic=self.config[propertyName][dataNameHeader + "DateItalicized"],
                        underlined=self.config[propertyName][dataNameHeader + "DateUnderlined"],
                        alignment=self.config[propertyName][dataNameHeader + "DateAlignment"])

                    self.pptEditor.setTextSuperScript(item[0], cordinalIndex, cordinalIndex + 2)
        except:
            print(f"\tERROR: {os.system.exc_info()[0]}")
            return False

        if (False in checkPoint):
            print(f"\tERROR : {checkPoint.count(True)} out of {len(checkPoint)} text placeholders found.")
            return False

        return self.pptEditor.commitSlideChanges()

    def _hymnSlide(self, source, slideIndex):
        if (not self.hymnMaker.setSource(source)):
            print(f"\tERROR: Hymn [{source}] not found.")
            return False

        [titleList, lyricsList] = self.hymnMaker.getContent()

        # Title font size adjustments and decide if the lyrics text box needs to be shifted
        titleFontSize = int(self.config["HYMN_PROPERTIES"]["HymnTitleTextSize"])
        maxCharUnitLength = int(self.config["HYMN_PROPERTIES"]["HymnTitleMaxUnitLength"])
        minCharUnitLength = int(self.config["HYMN_PROPERTIES"]["HymnTitleMinUnitLength"])
        titleLength = self._getVisualLength(titleList[0])
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
        sourceSlideID = self.pptEditor.getSlideID(slideIndex)
        for i in range(len(titleList) - 1):  # Recall that the original source slide remains
            self.pptEditor.duplicateSlide(sourceSlideID, sourceSlideID + '_d' + str(i))
        if (not self.pptEditor.commitSlideChanges()):
            return False

        # Update offset
        self.slideOffset += len(titleList) - 1

        # Insert title and lyrics data
        for i in range(slideIndex, slideIndex + len(titleList)):
            checkPoint = [False, False]
            data = self.pptEditor.getSlideTextData(i)
            try:
                for item in data:
                    if ('{Title}' in item[1]):
                        checkPoint[0] = True
                        self._insertText(
                            objectID=item[0],
                            text=titleList[i - slideIndex],
                            size=titleFontSize,
                            bold=self.config["HYMN_PROPERTIES"]["HymnTitleBolded"],
                            italic=self.config["HYMN_PROPERTIES"]["HymnTitleItalicized"],
                            underlined=self.config["HYMN_PROPERTIES"]["HymnTitleUnderlined"],
                            alignment=self.config["HYMN_PROPERTIES"]["HymnTitleAlignment"])
                    elif ('{Text}' in item[1]):
                        checkPoint[1] = True
                        self._insertText(
                            objectID=item[0],
                            text=lyricsList[i - slideIndex],
                            size=int(self.config["HYMN_PROPERTIES"]["HymnTextSize"]),
                            bold=self.config["HYMN_PROPERTIES"]["HymnBolded"],
                            italic=self.config["HYMN_PROPERTIES"]["HymnItalicized"],
                            underlined=self.config["HYMN_PROPERTIES"]["HymnUnderlined"],
                            alignment=self.config["HYMN_PROPERTIES"]["HymnAlignment"])

                        if (multiLineTitle):
                            self.pptEditor.updatePageElementTransform(item[0], translateY=int(self.config["HYMN_PROPERTIES"]["HymnLoweredUnitHeight"]))
            except:
                print(f"\tERROR: {os.system.exc_info()[0]}")
                return False

            if (False in checkPoint):
                print(f"\tERROR: {checkPoint.count(True)} out of {len(checkPoint)} text placeholders found.")
                return False

        return self.pptEditor.commitSlideChanges()

    def _scriptureSingleSlide(self, title, source, charPerLine, propertyName, dataNameHeader, nextWeek=False):
        nextWeekString = "NextWeek" if nextWeek else ""

        # Assumes monthly scripture is short enough to fit in one slide
        if (not self.verseMaker.setSource(source, charPerLine)):
            print(f"\tERROR: Verse [{source}] not found.")
            return False

        [verseString, ssIndexList] = self.verseMaker.getVerseString()

        checkPoint = [False, False]
        try:
            data = self.pptEditor.getSlideTextData(int(self.config[propertyName][dataNameHeader + nextWeekString + "Index"]) + self.slideOffset)
            for item in data:
                if '{Title}' in item[1]:
                    checkPoint[0] = True
                    self._insertText(
                        objectID=item[0],
                        text=title,
                        size=int(self.config[propertyName][dataNameHeader + nextWeekString + "TitleTextSize"]),
                        bold=self.config[propertyName][dataNameHeader + "TitleBolded"],
                        italic=self.config[propertyName][dataNameHeader + "TitleItalicized"],
                        underlined=self.config[propertyName][dataNameHeader + "TitleUnderlined"],
                        alignment=self.config[propertyName][dataNameHeader + "TitleAlignment"])
                elif '{Text}' in item[1]:
                    checkPoint[1] = True
                    self._insertText(
                        objectID=item[0],
                        text=verseString,
                        size=int(self.config[propertyName][dataNameHeader + "TextSize"]),
                        bold=self.config[propertyName][dataNameHeader + "Bolded"],
                        italic=self.config[propertyName][dataNameHeader + "Italicized"],
                        underlined=self.config[propertyName][dataNameHeader + "Underlined"],
                        alignment=self.config[propertyName][dataNameHeader + "Alignment"])

                    # Superscript and bold the verse numbers, and set paragraph spacing between the verses
                    for ssRange in ssIndexList:
                        self.pptEditor.setTextSuperScript(item[0], ssRange[1], ssRange[2])
                        self.pptEditor.setBold(item[0], ssRange[1], ssRange[2])
                        if (ssRange[1] > 0 and ssRange[0] == SSType.VerseNumber):  # No need for paragraph spacing in initial line
                            self.pptEditor.setSpaceAbove(item[0], ssRange[1], 10)
        except:
            print(f"\tERROR: {os.system.exc_info()[0]}")
            return False

        if (False in checkPoint):
            print(f"\tERROR: {checkPoint.count(True)} out of {len(checkPoint)} text placeholders found.")
            return False

        return self.pptEditor.commitSlideChanges()

    def _scriptureMultiSlide(self, source, maxLineLength, maxLinesPerSlide, propertyName, dataNameHeader):
        # Assumes monthly scripture is short enough to fit in one slide
        if (not self.verseMaker.setSource(source, maxLineLength)):
            print(f"\tERROR: Verse {source} not found.")
            return False

        [title, slideVersesList, slideSSIndexList] = self.verseMaker.getVerseStringMultiSlide(maxLinesPerSlide)

        # Generate create duplicate request and commit it
        slideIndex = int(self.config[propertyName][dataNameHeader + "Index"]) + self.slideOffset
        sourceSlideID = self.pptEditor.getSlideID(slideIndex)
        for i in range(len(slideVersesList) - 1):  # Recall that the original source slide remains
            self.pptEditor.duplicateSlide(
                sourceSlideID, sourceSlideID + '_' + str(i))
        if (not self.pptEditor.commitSlideChanges()):
            return False

        # Update offset
        self.slideOffset += len(slideVersesList) - 1

        # Get units of paragraph that separates verse numbers
        paragraphSpace = self.config["VERSE_PROPERTIES"]["VerseParagraphSpace"]

        # Insert title and verses data
        for i in range(slideIndex, slideIndex + len(slideVersesList)):
            checkPoint = [False, False]
            try:
                data = self.pptEditor.getSlideTextData(i)
                for item in data:
                    if '{Title}' in item[1]:
                        checkPoint[0] = True
                        self._insertText(
                            objectID=item[0],
                            text=title,
                            size=int(self.config[propertyName][dataNameHeader + "TitleTextSize"]),
                            bold=self.config[propertyName][dataNameHeader + "TitleBolded"],
                            italic=self.config[propertyName][dataNameHeader + "TitleItalicized"],
                            underlined=self.config[propertyName][dataNameHeader + "TitleUnderlined"],
                            alignment=self.config[propertyName][dataNameHeader + "TitleAlignment"])
                    elif '{Text}' in item[1]:
                        checkPoint[1] = True
                        self._insertText(
                            objectID=item[0],
                            text=slideVersesList[i - slideIndex],
                            size=int(self.config[propertyName][dataNameHeader + "TextSize"]),
                            bold=self.config[propertyName][dataNameHeader + "Bolded"],
                            italic=self.config[propertyName][dataNameHeader + "Italicized"],
                            underlined=self.config[propertyName][dataNameHeader + "Underlined"],
                            alignment=self.config[propertyName][dataNameHeader + "Alignment"])

                        # Superscript and bold the verse numbers, and set paragraph spacing between the verses
                        for ssRange in slideSSIndexList[i - slideIndex]:
                            self.pptEditor.setTextSuperScript(item[0], ssRange[1], ssRange[2])
                            self.pptEditor.setBold(item[0], ssRange[1], ssRange[2])
                            if (ssRange[1] > 0 and ssRange[0] == SSType.VerseNumber):  # No need for paragraph spacing in initial line
                                self.pptEditor.setSpaceAbove(item[0], ssRange[1], paragraphSpace)
            except:
                print(f"\tERROR: {os.system.exc_info()[0]}")
                return False

            if (False in checkPoint):
                print(f"\tERROR: {checkPoint.count(True)} out of {len(checkPoint)} text placeholders found.")
                return False

        return self.pptEditor.commitSlideChanges()

    def _sermonHeaderSlide(self, title, speaker, propertyName, dataNameHeader):
        checkPoint = [False, False]
        try:
            data = self.pptEditor.getSlideTextData(int(self.config[propertyName][dataNameHeader + "Index"]) + self.slideOffset)
            for item in data:
                if '{Title}' in item[1]:
                    checkPoint[0] = True
                    self._insertText(
                        objectID=item[0],
                        text=title.upper(),
                        size=int(self.config[propertyName][dataNameHeader + "TitleTextSize"]),
                        bold=self.config[propertyName][dataNameHeader + "TitleBolded"],
                        italic=self.config[propertyName][dataNameHeader + "TitleItalicized"],
                        underlined=self.config[propertyName][dataNameHeader + "TitleUnderlined"],
                        alignment=self.config[propertyName][dataNameHeader + "TitleAlignment"])
                elif '{Text}' in item[1]:
                    checkPoint[1] = True
                    self._insertText(
                        objectID=item[0],
                        text=speaker.title(),
                        size=int(self.config[propertyName][dataNameHeader + "SpeakerTextSize"]),
                        bold=self.config[propertyName][dataNameHeader + "SpeakerBolded"],
                        italic=self.config[propertyName][dataNameHeader + "SpeakerItalicized"],
                        underlined=self.config[propertyName][dataNameHeader + "SpeakerUnderlined"],
                        alignment=self.config[propertyName][dataNameHeader + "SpeakerAlignment"])
        except:
            print(f"\tERROR: {os.system.exc_info()[0]}")
            return False

        if (False in checkPoint):
            print(f"\tERROR : {checkPoint.count(True)} out of {len(checkPoint)} text placeholders found.")
            return False

        return self.pptEditor.commitSlideChanges()

    def _headerOnlySlide(self, title, propertyName, dataNameHeader, nextWeek=False):
        nextWeekString = "NextWeek" if nextWeek else ""
        checkPoint = [False]
        try:
            data = self.pptEditor.getSlideTextData(int(self.config[propertyName][dataNameHeader + nextWeekString + "Index"]) + self.slideOffset)
            for item in data:
                if '{Title}' in item[1]:
                    checkPoint[0] = True
                    self._insertText(
                        objectID=item[0],
                        text=title.upper(),
                        size=int(self.config[propertyName][dataNameHeader + nextWeekString + "TitleTextSize"]),
                        bold=self.config[propertyName][dataNameHeader + "TitleBolded"],
                        italic=self.config[propertyName][dataNameHeader + "TitleItalicized"],
                        underlined=self.config[propertyName][dataNameHeader + "TitleUnderlined"],
                        alignment=self.config[propertyName][dataNameHeader + "TitleAlignment"])
        except:
            print(f"\tERROR: {os.system.exc_info()[0]}")
            return False

        if (False in checkPoint):
            print(f"\tERROR : {checkPoint.count(True)} out of {len(checkPoint)} text placeholders found.")
            return False

        return self.pptEditor.commitSlideChanges()

    # ======================================================================================================
    # ============================================ TOOLS ===================================================
    # ======================================================================================================

    def _deleteSlide(self, slideIndex):
        sourceSlideID = self.pptEditor.getSlideID(slideIndex)
        self.pptEditor.deleteSlide(sourceSlideID)

        if (not self.pptEditor.commitSlideChanges()):
            return False

        self.slideOffset -= 1

        return "Disabled"

    def _openSlideInBrowser(self):
        self.pptEditor.openSlideInBrowser()

    def _insertText(self, objectID, text, size, bold, italic, underlined, alignment, linespacing=-1):
        self.pptEditor.setText(objectID, text)
        self.pptEditor.setTextStyle(objectID, bold, italic, underlined, size)
        self.pptEditor.setParagraphStyle(objectID, self.lineSpacing if linespacing < 0 else linespacing, alignment)

    def _getVisualLength(self, text):
        # A precise measurement to indicate if text will align or will take up multiple lines
        text = ''.join([i if ord(i) < 128 else ' ' for i in text])  # Replace all non-ascii characters
        return int(self.afm.string_width_height(text)[0])


if __name__ == '__main__':
    print("====================================================================")
    print("\t\t\tSlideMaker v1.0.0")
    print("====================================================================")
    print("\nInput Options:\n  Stream Slides: \ts\n  Projected Slides: \tp\n  Regular Slides: \tr\n  Quit: \t\tq\n")
    print("====================================================================\n")

    mode = ""
    while (mode != "q"):
        type = -1
        mode = input("Input: ")
        if (len(mode) > 0):
            if (mode == "s"):
                type = PPTMode.Stream
            elif (mode == "p"):
                type = PPTMode.Projected
            elif (mode == "r"):
                type = PPTMode.Regular
            elif (mode == "-t"):
                sm = SlideMaker()
                input = r' '.join(sys.argv[2:])
                print(f"String Visual Length of '{input}' = {sm._getVisualLength(input)} Units")

        if (type != -1):
            start = time.time()

            print("\nINITIALIZING...")
            sm = SlideMaker()

            sm.setType(type)

            if (type == PPTMode.Stream):
                type = "Stream"
            elif (type == PPTMode.Projected):
                type = "Projected"
            elif (type == PPTMode.Regular):
                type = "Regular"

            print(f"CREATING {type.upper()} SLIDES...")
            print("  sundayServiceSlide() : ", sm.sundayServiceSlide())
            print("  monthlyScriptureSlide() : ", sm.monthlyScriptureSlide())
            print("  announcementSlide() : ", sm.announcementSlide())
            print("  bibleVerseMemorizationSlide() : ", sm.bibleVerseMemorizationSlide())
            print("  catechismSlide() : ", sm.catechismSlide())
            print("  bibleVerseMemorizationNextWeekSlide() : ", sm.bibleVerseMemorizationSlide(True))
            print("  catechismNextWeekSlide() : ", sm.catechismSlide(True))
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

            print(f"OPENING {type.upper()} SLIDES...")
            sm._openSlideInBrowser()

            print(f"\nTask completed in {(time.time() - start):.2f} seconds.\n")
            print("====================================================================\n")
