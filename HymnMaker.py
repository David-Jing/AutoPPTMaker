import configparser
import math
import re
import sys
import os

from typing import Any, List

from LookupTools import LookupTools
from Utility import Utility

'''

Looks up inputted title and generates formatted hymn slide texts.
Please run setSource() before getContent().

titleList - list of title text corresponding to the order in formattedLyricsList
formattedLyricsList - list of formatted verses

'''


class HymnMaker:
    def __init__(self, type: str) -> None:
        # Get heuristics
        if (not os.path.exists("Data/" + type + "SlideProperties.ini")):
            raise IOError(f"ERROR : {type}SlideProperties.ini config file cannot be found.")
        config = configparser.ConfigParser()
        config.read("Data/" + type + "SlideProperties.ini")

        self.maxLines = int(config["HYMN_PROPERTIES"]["HymnMaxLines"])
        self.minLines = int(config["HYMN_PROPERTIES"]["HymnMinLines"])
        self.maxLineLength = int(config["HYMN_PROPERTIES"]["HymnMaxLineUnitLength"])

    def setSource(self, hymnTitle: str) -> bool:
        if (hymnTitle == ""):
            return False

        self.hymnTitle = hymnTitle
        self.hymn = LookupTools.getHymn(hymnTitle)

        return (self.hymn["title"] != "Not Found" and self.hymn["lyrics"] != "Not Found")

    def getContent(self) -> List[Any]:
        titleList = []
        formattedLyricsList = []

        if (self.hymn != ""):
            formattedLyricsList = self._getFormattedLyrics(self.hymn["lyrics"])

            # Generate titles and numbering
            slides = len(formattedLyricsList)
            for i in range(1, slides + 1):
                titleList.append(f'{self.hymnTitle.upper()} ({i}/{slides})')

        return [titleList, formattedLyricsList, self.hymn["source"]]

    # ==============================================================================================
    # ========================================= FORMATTER ==========================================
    # ==============================================================================================

    def _getFormattedLyrics(self, lyrics: str) -> List[str]:
        # Remove [...] headers and lyrics in brackets and split lyrics into its sections
        rawLyricsList = re.sub("(\[.*?\])|(\(.*?\))", "", lyrics).split("\n\n")

        # Remove starting and ending new lines
        for i in range(len(rawLyricsList)):
            rawLyricsList[i] = rawLyricsList[i].strip()

        # Remove duplicate blocks if it is under the maximum number of lines
        self._removeRepeatedAdjacentBlocks(rawLyricsList)

        # Split sections into lines and remove adjacent repeated lines that occur more than 3 times
        self._removeRepeatedLines(rawLyricsList)

        # Split long lines
        self._splitLongLines(rawLyricsList)

        # Herustics for a 'well-formatted' slides
        formattedLyricsList: List[str] = []
        i = 0
        while i < len(rawLyricsList):
            if rawLyricsList[i] == "\n":
                continue

            appended = False

            # Check if the verse has too few lines
            appended = self._formatShort(formattedLyricsList, i, rawLyricsList)
            i = i + 1 if appended else i

            # Check if the verse has too many lines
            if not appended:
                appended = self._formatLong(formattedLyricsList, i, rawLyricsList)

            if not appended:
                # Check if the verse is repeated, use '(x2)' to indicate repeat instead
                coreVerse = self._getPrincipalPeriod(rawLyricsList[i])
                if (coreVerse != ""):
                    formattedLyricsList.append(coreVerse[:-1] + f" (x{rawLyricsList[i].count(coreVerse)})\n")
                else:
                    formattedLyricsList.append(rawLyricsList[i])

            i += 1

        return formattedLyricsList

    def _formatLong(self, formattedLyricsList: List[str], rawLyricsIndex: int, rawLyricsList: List[str]) -> bool:
        # Recall that the last line does not have a new line symbol
        numOfLines = rawLyricsList[rawLyricsIndex].count("\n") + 1

        if (numOfLines > self.maxLines):
            # Find how many slides to split into
            numOfSlides = math.ceil(numOfLines / self.maxLines)
            linesPerSlide = round(numOfLines / numOfSlides)

            # Append splitted lyrics into list
            startIndex = 0
            endIndex = 0
            for i in range(numOfSlides):
                for _ in range(linesPerSlide):
                    endIndex = rawLyricsList[rawLyricsIndex].find("\n", endIndex + 1)

                # Add the entire remaining verses if on last slide
                if i == numOfSlides - 1 or endIndex == -1:
                    endIndex = len(rawLyricsList[rawLyricsIndex])

                formattedLyricsList.append(rawLyricsList[rawLyricsIndex][startIndex:endIndex])
                startIndex = endIndex

                # Exclude new line symbols
                if i != numOfSlides - 1 and endIndex != -1:
                    startIndex += 1

            return True
        return False

    def _formatShort(self, formattedLyricsList: List[str], rawLyricsIndex: int, rawLyricsList: List[str]) -> bool:
        # Combines this verse to one from another slide
        if rawLyricsIndex < len(rawLyricsList) - 1 \
                and rawLyricsList[rawLyricsIndex].count("\n") + rawLyricsList[rawLyricsIndex + 1].count("\n") + 3 <= self.maxLines:
            formattedLyricsList.append(rawLyricsList[rawLyricsIndex] + "\n\n" + rawLyricsList[rawLyricsIndex + 1])
            return True
        return False

    def _removeRepeatedAdjacentBlocks(self, lyricsList: List[str]) -> None:
        # Remove any repeated blocks of lyrics if under the max number of lines
        i = 0
        while i < len(lyricsList):
            # Recall that the last line does not have a new line symbol
            if lyricsList[i].count("\n") + 1 <= self.maxLines:
                numOfRepeats = 0

                while i < len(lyricsList) - 1 and lyricsList[i] == lyricsList[i + 1]:
                    numOfRepeats += 1
                    i += 1

                if (numOfRepeats > 0):
                    lyricsList[i] += f" (x{numOfRepeats + 1})"
                    for _ in range(i - numOfRepeats, i):
                        lyricsList.pop(i - numOfRepeats)

                    i -= numOfRepeats - 1
            i += 1

    def _removeRepeatedLines(self, lyricsList: List[str]) -> None:
        # Split sections into lines and remove adjacent repeated lines that occur more than 3 times
        for i in range(len(lyricsList)):
            lineList = lyricsList[i].split("\n")
            lineIndex = 0

            while lineIndex < len(lineList):
                refIndex = lineIndex  # The index the line first appears in
                numOfRepeats = 0

                lineIndex += 1
                while lineIndex < len(lineList) and lineList[lineIndex] == lineList[refIndex]:
                    numOfRepeats += 1
                    lineIndex += 1

                # If appeared more than 3 times
                if numOfRepeats > 2:
                    for _ in range(numOfRepeats):
                        lineList.pop(refIndex)

                    # Append repeat tag to a single instance of the repeated line
                    lineList[refIndex] = "\n" + lineList[refIndex] + f" (x{numOfRepeats + 1})" + "\n"
                else:
                    lineIndex -= numOfRepeats

            # Remove any appended new line for the last line
            if lineList[-1] == "":
                lineList.pop(-1)
            elif lineList[-1][-1] == "\n":
                lineList[-1] = lineList[-1][:-1]

            # Recombine lines
            lyricsList[i] = "\n".join(lineList)

    def _splitLongLines(self, lyricsList: List[str]) -> None:
        # Split lyrics lines that exceed max length (split verse between 20% to 80%)
        for i in range(len(lyricsList)):
            lineList = lyricsList[i].split("\n")

            j = 0
            while j < len(lineList):
                if Utility.getVisualLength(lineList[j]) > self.maxLineLength:
                    splitIndex = -1

                    # Split at comma
                    splitIndex = lineList[j].rfind(",", int(2*len(lineList[j])/5), int(4*len(lineList[j])/5))

                    # Or split at semicolon
                    splitIndex = lineList[j].rfind(";", int(2*len(lineList[j])/5), int(4*len(lineList[j])/5)) \
                        if splitIndex == -1 else splitIndex

                    # Or split at exclamation point
                    splitIndex = lineList[j].rfind("!", int(2*len(lineList[j])/5), int(4*len(lineList[j])/5)) \
                        if splitIndex == -1 else splitIndex

                    # Or split at period
                    splitIndex = lineList[j].rfind(".", int(2*len(lineList[j])/5), int(4*len(lineList[j])/5)) \
                        if splitIndex == -1 else splitIndex

                    # Or split at approximate half way
                    splitIndex = lineList[j].rfind(" ", int(1*len(lineList[j])/3), int(2*len(lineList[j])/3)) \
                        if splitIndex == -1 else splitIndex

                    # Split at index
                    if splitIndex > 0:
                        secondLineOffset = 1 if lineList[j][splitIndex] == " " else 2
                        lineList.insert(j + 1, lineList[j][splitIndex + secondLineOffset:])
                        lineList[j] = lineList[j][:splitIndex + 1]
                        j += 1
                j += 1

                # Recombine lines
            lyricsList[i] = "\n".join(lineList)

    # ==============================================================================================
    # =========================================== TOOLS ============================================
    # ==============================================================================================

    def _getPrincipalPeriod(self, verse: str) -> str:
        # Checks if the verse is periodic and is equal to a nontrivial rotation of itself
        i = (verse+verse).find(verse, 1, -1)
        return "" if i == -1 else verse[:i]

    def _cleanUpVerse(self, verse: str) -> str:
        # Gets rid of [...] header
        index = verse.find(']') + 1
        verse = verse if index <= 0 else verse[index:]

        # Remove new lines before and after text
        return verse.strip() + "\n"

# ==============================================================================================
# ============================================ TESTER ==========================================
# ==============================================================================================


if __name__ == '__main__':
    # python HymnMaker.py [type] [hymn name]
    type = ""
    if (len(sys. argv) > 2):
        if (sys.argv[1] == "-s"):
            type = "Stream"
        elif (sys.argv[1] == "-p"):
            type = "Projected"
        elif (sys.argv[1] == "-r"):
            type = "Regular"

    if (type != ""):
        hm = HymnMaker(type)

        if (hm.setSource(' '.join(sys.argv[2:]))):
            print(hm.hymn["title"] + "\n")
            [titleList, formattedLyricsList, _] = hm.getContent()
            for i in range(len(formattedLyricsList)):
                print("-----------" + titleList[i] + "-----------")
                print(formattedLyricsList[i])

        else:
            print("Song not found.")
