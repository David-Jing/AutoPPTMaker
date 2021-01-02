import configparser
import os
import sys
import time

from enum import Enum
from HymnMaker import HymnMaker
from PPTEditorTools import PPTEditorTools
from PPTEditorTools import DateFormatMode
from VerseMaker import VerseMaker
from VerseMaker import SSType


class PPTMode(Enum):
    # Specifies the PPT output format
    Stream = 0
    Projected = 1


class SlideMaker:
    def setType(self, type):
        if (type == PPTMode.Stream):
            type = "Stream"
        elif (type == PPTMode.Projected):
            type = "Projected"
        else:
            raise ValueError(f"ERROR : PPTMode not recognized.")

        if (not os.path.exists(type + "SlideProperties.ini")):
            raise IOError(f"ERROR : {type}SlideProperties.ini config file does not exist.")

        self.pptEditor = PPTEditorTools(type)
        self.verseMaker = VerseMaker(type)
        self.hymnMaker = HymnMaker(type)

        # Offset for when creating new slides
        self.slideOffset = 0

        # Access slide property data
        self.config = configparser.ConfigParser()
        self.config.read(type + "SlideProperties.ini")

        # Access slide input data
        self.input = configparser.ConfigParser()
        self.input.read("SlideInputs.ini")

    # ======================================================================================================
    # ========================================== SLIDE MAKER ===============================================
    # ======================================================================================================

    def sundayServiceSlide(self):
        title = self.input["SUNDAY_SERVICE_HEADER"]["SundayServiceHeaderTitle"].upper()

        return self._titleSlide(title, "SUNDAY_SERVICE_HEADER_PROPERTIES", "SundayServiceHeader")

    def monthlyScriptureSlide(self):
        title = self.input["MONTHLY_SCRIPTURE"]["MonthlyScriptureTitle"].upper()
        source = self.input["MONTHLY_SCRIPTURE"]["MonthlyScriptureSource"].upper()
        charPerLine = int(self.config["MONTHLY_SCRIPTURE_PROPERTIES"]["MonthlyScriptureCharPerLine"])

        return self._scriptureSingleSlide(title, source, charPerLine,
                                          "MONTHLY_SCRIPTURE_PROPERTIES", "MonthlyScripture")

    def announcementSlide(self):
        '''
        TODO
        '''

    def bibleVerseMemorizationSlide(self):
        title = self.input["BIBLE_MEMORIZATION"]["BibleMemorizationTitle"].upper()
        source = self.input["BIBLE_MEMORIZATION"]["BibleMemorizationSource"].upper()
        charPerLine = int(self.config["BIBLE_MEMORIZATION_PROPERTIES"]["BibleMemorizationCharPerLine"])

        return self._scriptureSingleSlide(title, source, charPerLine,
                                          "BIBLE_MEMORIZATION_PROPERTIES", "BibleMemorization")

    def catechismSlide(self):
        '''
        TODO
        '''

    def worshipSlide(self):
        title = self.input["WORSHIP_HEADER"]["WorshipHeaderTitle"].upper()

        return self._titleSlide(title, "WORSHIP_HEADER_PROPERTIES", "WorshipHeader")

    def callToWorshipSlide(self):
        source = self.input["CALL_TO_WORSHIP"]["CallToWorshipSource"].upper()
        charPerLine = int(self.config["CALL_TO_WORSHIP_PROPERTIES"]["CallToWorshipCharPerLine"])
        maxLinesPerSlide = int(self.config["CALL_TO_WORSHIP_PROPERTIES"]["CallToWorshipMaxLines"])

        return self._scriptureMultiSlide(source, charPerLine, maxLinesPerSlide,
                                         "CALL_TO_WORSHIP_PROPERTIES", "CallToWorship")

    def sermonHeaderSlide(self):
        title = self.input["SERMON_HEADER"]["SermonHeaderTitle"].upper()
        speaker = self.input["SERMON_HEADER"]["SermonHeaderSpeaker"]

        return self._sermonHeaderSlide(title, speaker, "SERMON_HEADER_PROPERTIES", "SermonHeader")

    def sermonVerseSlide(self):
        source = self.input["SERMON_VERSE"]["SermonVerseSource"].upper()
        charPerLine = int(self.config["SERMON_VERSE_PROPERTIES"]["SermonVerseCharPerLine"])
        maxLinesPerSlide = int(self.config["SERMON_VERSE_PROPERTIES"]["SermonVerseMaxLines"])

        return self._scriptureMultiSlide(source, charPerLine, maxLinesPerSlide,
                                         "SERMON_VERSE_PROPERTIES", "SermonVerse")

    def hymnSlides(self, number):
        hymnSource = self.input["HYMN"][f"Hymn{number}Source"]
        hymnIndex = int(self.config["HYMN_PROPERTIES"][f"Hymn{number}Index"]) + self.slideOffset

        return self._hymnSlide(hymnSource, hymnIndex)

    # ======================================================================================================
    # ======================================= GENERIC SLIDE MAKER ==========================================
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
                elif '{Date}' in item[1]:
                    checkPoint[1] = True
                    self._insertText(
                        objectID=item[0],
                        text=dateString,
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
        titleSize = int(self.config["HYMN_PROPERTIES"]["HymnTitleTextSize"])
        maxCharLength = int(self.config["HYMN_PROPERTIES"]["HymnTitleMaxChar"])
        multiLineTitle = False
        if (len(titleList[0]) > maxCharLength):                         # Heuristic : Shift the lyrics text block down for better aesthetics for multi-line titles
            multiLineTitle = True
            firstLineLength = titleList[0].rfind(" ", 0, maxCharLength)
            if (len(titleList[0]) - firstLineLength > firstLineLength):  # Heuristic : It's better when there're more text in the first line than second
                titleSize = int(self.config["HYMN_PROPERTIES"]["HymnTitleMultiLineTextSize"])

        # Generate create duplicate request and commit it
        sourceSlideID = self.pptEditor.getSlideID(slideIndex)
        for i in range(len(titleList) - 1):  # Recall that the original source slide remains
            self.pptEditor.getDuplicateSlide(sourceSlideID, sourceSlideID + '_d' + str(i))
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
                            size=titleSize,
                            bold=self.config["HYMN_PROPERTIES"]["HymnTitleBolded"],
                            italic=self.config["HYMN_PROPERTIES"]["HymnTitleItalicized"],
                            underlined=self.config["HYMN_PROPERTIES"]["HymnTitleUnderlined"],
                            alignment=self.config["HYMN_PROPERTIES"]["HymnTitleAlignment"])
                    elif ('{Lyrics}' in item[1]):
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
                            self.pptEditor.updatePageElementTransform(item[0], translateY=int(self.config["HYMN_PROPERTIES"]["HymnLoweredHeightAmount"]))
            except:
                print(f"\tERROR: {os.system.exc_info()[0]}")
                return False

            if (False in checkPoint):
                print(f"\tERROR: {checkPoint.count(True)} out of {len(checkPoint)} text placeholders found.")
                return False

        return self.pptEditor.commitSlideChanges()

    def _scriptureSingleSlide(self, title, source, charPerLine, propertyName, dataNameHeader):
        # Assumes monthly scripture is short enough to fit in one slide
        if (not self.verseMaker.setSource(source, charPerLine)):
            print(f"\tERROR: Verse [{source}] not found.")
            return False

        [verseString, ssIndexList] = self.verseMaker.getVerseString()

        checkPoint = [False, False]
        try:
            data = self.pptEditor.getSlideTextData(int(self.config[propertyName][dataNameHeader + "Index"]) + self.slideOffset)
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
                elif '{Verse}' in item[1]:
                    checkPoint[1] = True
                    self._insertText(
                        objectID=item[0],
                        text=verseString,
                        size=int(self.config[propertyName][dataNameHeader + "TextSize"]),
                        bold=self.config[propertyName][dataNameHeader + "Bolded"],
                        italic=self.config[propertyName][dataNameHeader + "Italicized"],
                        underlined=self.config[propertyName][dataNameHeader + "Underlined"],
                        alignment=self.config[propertyName][dataNameHeader + "Alignment"])

                    for ssRange in ssIndexList:
                        self.pptEditor.setTextSuperScript(item[0], ssRange[1], ssRange[2])
                        if (ssRange[1] > 0 and ssRange[0] == SSType.VerseNumber):  # No need for paragraph spacing in initial line
                            self.pptEditor.setSpaceAbove(item[0], ssRange[1], 10)
        except:
            print(f"\tERROR: {os.system.exc_info()[0]}")
            return False

        if (False in checkPoint):
            print(f"\tERROR: {checkPoint.count(True)} out of {len(checkPoint)} text placeholders found.")
            return False

        return self.pptEditor.commitSlideChanges()

    def _scriptureMultiSlide(self, source, charPerLine, maxLinesPerSlide, propertyName, dataNameHeader):
        # Assumes monthly scripture is short enough to fit in one slide
        if (not self.verseMaker.setSource(source, charPerLine)):
            print(f"\tERROR: Verse {source} not found.")
            return False

        [title, slideVersesList, slideSSIndexList] = self.verseMaker.getVerseStringMultiSlide(maxLinesPerSlide)

        # Generate create duplicate request and commit it
        slideIndex = int(self.config[propertyName][dataNameHeader + "Index"]) + self.slideOffset
        sourceSlideID = self.pptEditor.getSlideID(slideIndex)
        for i in range(len(slideVersesList) - 1):  # Recall that the original source slide remains
            self.pptEditor.getDuplicateSlide(
                sourceSlideID, sourceSlideID + '_' + str(i))
        if (not self.pptEditor.commitSlideChanges()):
            return False

        # Update offset
        self.slideOffset += len(slideVersesList) - 1

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
                    elif '{Verse}' in item[1]:
                        checkPoint[1] = True
                        self._insertText(
                            objectID=item[0],
                            text=slideVersesList[i - slideIndex],
                            size=int(self.config[propertyName][dataNameHeader + "TextSize"]),
                            bold=self.config[propertyName][dataNameHeader + "Bolded"],
                            italic=self.config[propertyName][dataNameHeader + "Italicized"],
                            underlined=self.config[propertyName][dataNameHeader + "Underlined"],
                            alignment=self.config[propertyName][dataNameHeader + "Alignment"])

                        for ssRange in slideSSIndexList[i - slideIndex]:
                            self.pptEditor.setTextSuperScript(item[0], ssRange[1], ssRange[2])
                            if (ssRange[1] > 0 and ssRange[0] == SSType.VerseNumber):  # No need for paragraph spacing in initial line
                                self.pptEditor.setSpaceAbove(item[0], ssRange[1], 10)
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
                elif '{Speaker}' in item[1]:
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

    # ======================================================================================================
    # ============================================ TOOLS ===================================================
    # ======================================================================================================

    def openSlideInBrowser(self):
        self.pptEditor.openSlideInBrowser()

    def _insertText(self, objectID, text, size, bold, italic, underlined, alignment, linespacing=110):
        self.pptEditor.setText(objectID, text)
        self.pptEditor.setTextStyle(objectID, bold, italic, underlined, size)
        self.pptEditor.setParagraphStyle(objectID, linespacing, alignment)


if __name__ == '__main__':
    start = time.time()

    type = -1
    if (len(sys. argv) > 1):
        if (sys.argv[1] == "-s"):
            type = PPTMode.Stream
        elif (sys.argv[1] == "-p"):
            type = PPTMode.Projected

    if (type != -1):
        print(f"INITIALIZING...")
        sm = SlideMaker()

        sm.setType(type)
        if (type == PPTMode.Stream):
            type = "Stream"
        elif (type == PPTMode.Projected):
            type = "Projected"

        print(f"CREATING {type.upper()} SLIDES...")
        print("  sundayServiceSlide() : ", sm.sundayServiceSlide())
        print("  monthlyScriptureSlide() : ", sm.monthlyScriptureSlide())
        print("  bibleVerseMemorizationSlide() : ", sm.bibleVerseMemorizationSlide())
        print("  worshipSlide() : ", sm.worshipSlide())
        print("  callToWorshipSlide() : ", sm.callToWorshipSlide())
        print("  hymnSlides(1) : ", sm.hymnSlides(1))
        print("  hymnSlides(2) : ", sm.hymnSlides(2))
        print("  sermonHeaderSlide() : ", sm.sermonHeaderSlide())
        print("  sermonVerseSlide() : ", sm.sermonVerseSlide())
        print("  hymnSlides(3) : ", sm.hymnSlides(3))

        print(f"OPENING {type.upper()} SLIDES...")
        sm.openSlideInBrowser()

    print(f"\nTask completed in {(time.time() - start):.2f} seconds.")
