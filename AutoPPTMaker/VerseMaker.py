import configparser
import os
import sys
import math

from enum import Enum
from typing import Any, List

from ListMaker import ListMaker
from LookupTools import LookupTools
from Utility import Utility

"""

Looks up inputted biblical verse and generates formatted verse slide texts.
Please setSource() before running any getters.

titleList - list of title text corresponding to the order in lyricsList
verseList - list of formatted verses
ssIndexList - list of lists of character indexes that requires super scripting

Superscript alignment:
 - 1 superscripted character = 2 normal spaces converted to superscript
 - 2 superscripted characters = 1 normal space converted to superscript
 - 3 superscripted characters = 0 normal space converted to superscript

"""


class VerseMaker:
    class SSType(Enum):
        # Specifies if the superscripted range is superscripting empty spaces or verse numbers
        Space = 0
        VerseNumber = 1

    def __init__(self, type: str) -> None:
        self.verseSource = ""
        self.verses = ""

        # Get heuristics
        if not os.path.exists("Data/" + type + "SlideProperties.ini"):
            raise IOError(f"ERROR : {type}SlideProperties.ini config file cannot be found.")
        config = configparser.ConfigParser()
        config.read("Data/" + type + "SlideProperties.ini")

        self.maxLineLength = 0
        self.indentSpace = int(config["VERSE_PROPERTIES"]["VerseIndentSpace"])

    def setSource(self, verseSource: str, maxLineLength: int) -> bool:
        # Returns true if source is valid, false otherwise
        if (verseSource == ""):
            return False

        self.maxLineLength = maxLineLength
        self.verseSource = verseSource
        self.verses = LookupTools.getVerse(self.verseSource)

        return self.verses != "Not Found"

    def getContent(self) -> Any:
        title = self.verseSource
        verseList = []
        ssIndexList = []

        # Generate verse line formatting and indexes of superscripted characters
        if (self.verses != ""):
            verseList = self._getFormattedVerses(self.verses)
            ssIndexList = self._getSuperScriptIndexes(verseList)

        return [title, verseList, ssIndexList]

    # ==============================================================================================
    # ======================================= CUSTOM GETTERS =======================================
    # ==============================================================================================

    def getVerseStringSingleSlide(self) -> Any:
        # No title, but appends a "--[Verse Source]" on a new line at the end of the verse
        [title, verseList, ssIndexList] = self.getContent()

        # Convert verse to a single string
        verseString = ""
        for verse in verseList:
            for line in verse:
                verseString += line

        # Appends a "--[Verse Source]" to verses
        verseString += self._getFormattedVerseSource(title)

        return [verseString, ssIndexList]

    def getVerseStringMultiSlide(self, maxLinesPerSlide: int, appendSourceInVerse: bool = False) -> Any:
        # Splits verse into multiple components, each per slide
        [title, verseList, ssIndexList] = self.getContent()

        # Appends a "--[Verse Source]" to verses
        if appendSourceInVerse:
            verseList[-1].append(self._getFormattedVerseSource(title))

        # Calculate the number of slides needed
        lineCount = 0
        for verse in verseList:
            lineCount += len(verse)

        # Try to have even distribution of lines per slide
        numOfSlides = int(lineCount/maxLinesPerSlide) if lineCount % maxLinesPerSlide == 0 \
            else math.ceil(lineCount / maxLinesPerSlide + 0.1)                  # Heuristic : "Fudging" the statistic (guess-timation) for better aesthetics
        linesPerSlide = round((lineCount/numOfSlides + maxLinesPerSlide) / 2)   # Heuristic : Can't have slides be too empty

        # Split verses for each slide and recalculate ssIndexList
        slideVersesList = []
        slideSSIndexList = []   # Assumes each verse has only one symbol to superscript
        verseIndex = 0          # Tracking verse number
        charIndex = 0           # Tracking the number of character parsed, for adjusting ssIndexes
        ssIndex = 0             # Indexing the list of superscripted ranges
        ssIndexListLength = len(ssIndexList)
        numOfVerses = len(verseList)
        while(verseIndex < numOfVerses):
            verseString = ""
            currSSSlideIndexList = []
            currTotalLineCount = 0
            nextlineCount = 0 if verseIndex >= numOfVerses else len(verseList[verseIndex])

            # Asserts that a slide must contain at least one verse
            while (nextlineCount > 0 and (currTotalLineCount == 0 or currTotalLineCount + nextlineCount <= linesPerSlide)):
                isFirstLine = True
                for line in verseList[verseIndex]:
                    verseString += line
                    if (ssIndex < ssIndexListLength and (ssIndexList[ssIndex][0] == self.SSType.Space or isFirstLine)):
                        isFirstLine = False
                        currSSSlideIndexList.append([ssIndexList[ssIndex][0], ssIndexList[ssIndex][1] - charIndex, ssIndexList[ssIndex][2] - charIndex])
                        ssIndex += 1

                currTotalLineCount += nextlineCount
                verseIndex += 1
                nextlineCount = 0 if verseIndex >= numOfVerses else len(verseList[verseIndex])

            slideVersesList.append(verseString)
            slideSSIndexList.append(currSSSlideIndexList)
            charIndex += len(verseString)

        return [title, slideVersesList, slideSSIndexList]

    # ==============================================================================================
    # ========================================= FORMATTER ==========================================
    # ==============================================================================================

    def _getFormattedVerseSource(self, sourceStr: str) -> str:
        # Generates an indented "--[Verse Source]" string; for appending source after verse
        sourceStr = "—" + sourceStr.title()                                            # String.title() makes every first character capitalized
        sourceStrIndentSpace = int(0.93*self.maxLineLength/250) - len(sourceStr)       # Heuristic : It looks better if text is not right against the right wall (space, " ", is 250 units wide)
        sourceStrIndent = "".join(" " for _ in range(sourceStrIndentSpace if sourceStrIndentSpace > 0 else 0))

        return f"\n{sourceStrIndent}{sourceStr}"

    def _getFormattedVerses(self, verses: str) -> List[List[str]]:
        [verseList, numberList] = self._getVerseComponents(verses)

        return ListMaker.setLineLengthRestriction(verseList, numberList, self.indentSpace, self.maxLineLength)

    def _getSuperScriptIndexes(self, verseList: List[List[str]]) -> List[List[object]]:
        index = 0
        ssIndexList = []

        # Look for a verse line that starts with a number and those without
        currVerseNum = -1
        for verse in verseList:
            for line in verse:
                # Assumes each verse begins with a verse number
                if (line[0].isnumeric()):
                    numEndIndex = line.find(" ")
                    nextVerseNum = int(line[0:numEndIndex])
                    # Numbers either increment or reset to 1 on new chapter
                    if (numEndIndex != -1 and (currVerseNum < 0 or currVerseNum + 1 == nextVerseNum or nextVerseNum == 1)):
                        ssIndexList.append([self.SSType.VerseNumber, index, numEndIndex + index])
                        currVerseNum = nextVerseNum
                else:
                    ssAlignmentSpace = self._getSSAlignmentSpace(len(str(currVerseNum)))
                    if (ssAlignmentSpace > 0):
                        ssIndexList.append([self.SSType.Space, index, ssAlignmentSpace + index])

                index += len(line)

        return ssIndexList

    # ==============================================================================================
    # =========================================== TOOLS ============================================
    # ==============================================================================================

    def _getRightMostMaxIndex(self, verse: str) -> int:
        # Based on the max line unit length, find the rightmost character that is within the length
        leftIndex = 0
        rightIndex = len(verse)

        if (Utility.getVisualLength(verse) <= self.maxLineLength):
            return rightIndex

        # Iteratively narrow down the index position
        while(leftIndex + 1 < rightIndex):
            middleIndex = int((leftIndex + rightIndex) / 2)
            currUnitLineLength = Utility.getVisualLength(verse[:middleIndex])
            if (currUnitLineLength > self.maxLineLength):
                rightIndex = middleIndex
            else:
                leftIndex = middleIndex

        return leftIndex

    def _getVerseComponents(self, verses: str) -> List[List[str]]:
        # Split multiple verses into individuals
        verseNumberList = []
        verseList = []

        # Clean up empty spaces, new lines, and headers
        verses = " ".join(verses.split())

        # Iterate verses by search for "[" and "]"
        currIndex = 0
        prevIndex = verses.find("]", currIndex)
        while (True):
            currIndex = verses.find("[", currIndex)

            if (currIndex > prevIndex):
                verseList.append(verses[prevIndex:currIndex].strip())

            if (currIndex < 0):
                verseList.append(verses[prevIndex:].strip())
                break

            currIndex += 1
            prevIndex = verses.find("]", currIndex) + 1
            verseNumberList.append(verses[currIndex:prevIndex - 1])

        return [verseList, verseNumberList]

    def _getSSAlignmentSpace(self, ssCharLength: int) -> int:
        # Number of converted spaces need to align line with verse number and lines without;
        # assumes bible verse number does not exceed more than 3 digits
        if (ssCharLength == 3):
            return 0
        elif (ssCharLength == 2):
            return 1
        elif (ssCharLength == 1):
            return 2
        else:
            return 0

# ==============================================================================================
# ============================================ TESTER ==========================================
# ==============================================================================================


if __name__ == "__main__":
    # python VerseMaker.py [type] [verse source]

    type = ""
    if (len(sys. argv) > 2):
        if (sys.argv[1] == "-s"):
            type = "Stream"
        elif (sys.argv[1] == "-p"):
            type = "Projected"
        elif (sys.argv[1] == "-r"):
            type = "Regular"

    print(type)

    if (type != ""):
        vm = VerseMaker(type)
        config = configparser.ConfigParser()
        config.read("Data\\" + type + "SlideProperties.ini")

        if (vm.setSource(" ".join(sys.argv[2:]), int(config["SERMON_VERSE_PROPERTIES"]["SermonVerseMaxLineLength"]))):
            # GetContent() Test
            [title, verseList, ssIndexList] = vm.getContent()
            print("\n---------------" + title + "---------------")
            for verse in verseList:
                for line in verse:
                    print(line, end="")

            # GetVerseString() Test
            print("\n\n===========================================================\n")
            [verseString, ssIndexList] = vm.getVerseStringSingleSlide()
            print(verseString, end="")

            # getVerseStringMultiSlide() Test
            print("\n\n===========================================================\n")
            slideIndex = 0
            [title, slideVersesList, ssIndexList] = vm.getVerseStringMultiSlide(int(config["SERMON_VERSE_PROPERTIES"]["SermonVerseMaxLines"]))
            numOfVerses = len(slideVersesList)
            print("\n---------------" + title + "---------------")
            for i in range(numOfVerses):
                print(ssIndexList[slideIndex])
                slideIndex += 1
                if (i != numOfVerses - 1):
                    print(slideVersesList[i], end="")
                    print("\n---------------" + "".join("-" for _ in range(len(title))) + "---------------")
                else:
                    print(slideVersesList[i])

        else:
            print("Verse not found.")
