import configparser
import os
import sys
import math

from enum import Enum

import matplotlib
from LookupTools import LookupTools
from matplotlib.afm import AFM

'''

Looks up inputted biblical verse and generates formatted verse slide texts.
Please setSource() before running any getters.

titleList - list of title text corresponding to the order in lyricsList
verseList - list of formatted verses
ssIndexList - list of lists of character indexes that requires super scripting

Superscript alignment:
 - 1 superscripted character = 2 normal spaces converted to superscript
 - 2 superscripted characters = 1 normal space converted to superscript
 - 3 superscripted characters = 1 normal space converted to superscript
 
'''


class SSType(Enum):
    # Specifies if the superscripted range is superscripting empty spaces or verse numbers
    Space = 0
    VerseNumber = 1


class VerseMaker:
    def __init__(self, type):
        self.verseSource = ""
        self.verses = ""

        # Get heuristics
        if not os.path.exists("Data/" + type + "SlideProperties.ini"):
            raise IOError(f"ERROR : {type}SlideProperties.ini config file cannot be found.")
        config = configparser.ConfigParser()
        config.read("Data/" + type + "SlideProperties.ini")

        self.maxLineLength = 0
        self.indentSpace = int(config["VERSE_PROPERTIES"]["VerseIndentSpace"])

        # For finding visual lengths of text strings
        afm_filename = os.path.join(matplotlib.get_data_path(), 'fonts', 'afm', 'ptmr8a.afm')
        self.afm = AFM(open(afm_filename, "rb"))

    def setSource(self, verseSource, maxLineLength):
        # Returns true if source is valid, false otherwise
        if (verseSource == ""):
            return False

        self.maxLineLength = maxLineLength
        self.verseSource = verseSource
        self.verses = LookupTools.getVerse(self.verseSource)

        return self.verses != "Not Found"

    def getContent(self):
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

    def getVerseString(self):
        # No title, but appended with a '--[Verse Source]' on a new line
        [title, verseList, ssIndexList] = self.getContent()

        # Convert verse to a single string
        verseString = ""
        for verse in verseList:
            for line in verse:
                verseString += line

        # Appends title after verse
        title = "—" + title.title()
        sourceIndentSpace = int(0.93*self.maxLineLength/250) - len(title)          # Heuristic : It looks better if text is not right against the right wall (space, " ", is 250 units wide)
        titleIndent = "".join(" " for _ in range(sourceIndentSpace if sourceIndentSpace > 0 else 0))
        verseString += "\n" + titleIndent + title + "\n"

        return [verseString, ssIndexList]

    def getVerseStringMultiSlide(self, maxLinesPerSlide):
        # Splits verse into multiple components, each per slide
        [title, verseList, ssIndexList] = self.getContent()

        # Calculate the number of slides needed
        lineCount = 0
        for verse in verseList:
            lineCount += len(verse)

        # Try to have even distribution of lines per slide
        numOfSlides = int(lineCount/maxLinesPerSlide) if lineCount % maxLinesPerSlide == 0 \
            else math.ceil(lineCount / maxLinesPerSlide + 0.1)                  # Heuristic : "Fudging" the statistic (guess-timation) for better aesthetics
        linesPerSlide = round((lineCount/numOfSlides + maxLinesPerSlide) / 2)   # Heuristic : Can't have slides too empty

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
                    if (ssIndex < ssIndexListLength and (ssIndexList[ssIndex][0] == SSType.Space or isFirstLine)):
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

    def _getFormattedVerses(self, verses):
        [verseList, numberList] = self._getVerseComponents(verses)

        # A list of list of verse lines
        formatedVerses = []

        # Find max character of verse number
        maxDigits = 0
        for num in numberList:
            if len(num) > maxDigits:
                maxDigits = len(num)

        # Create indent spacing
        indent = "".join(" " for _ in range(self.indentSpace))

        for i in range(len(verseList)):
            # Split verse into separate lines that are within the character limit
            verseLines = []
            leftIndex = 0
            rightIndex = self._getRightMostMaxIndex(verseList[i])

            # Iteratively cut string into lengths less than maximum
            while (rightIndex < len(verseList[i])):
                # Find " " character left of the maxLineLength Index
                while (verseList[i][rightIndex] != " " and rightIndex >= 0):
                    rightIndex -= 1

                # Generate custom indent for line with verse numbers
                if (leftIndex < 1):
                    customIndentSpaces = self.indentSpace - len(numberList[i])
                    customIndent = "".join(" " for _ in range(customIndentSpaces if customIndentSpaces > 0 else 0))
                    verseLines.append(numberList[i] + customIndent + verseList[i][leftIndex:rightIndex] + "\n")
                else:
                    verseLines.append(indent + verseList[i][leftIndex:rightIndex] + "\n")

                leftIndex = rightIndex
                rightIndex += self._getRightMostMaxIndex(verseList[i][leftIndex:]) + 1

            # Append the remaining verses and add to list of formatted verses
            if (leftIndex < 1):
                customIndentSpaces = self.indentSpace - len(numberList[i])
                customIndent = "".join(" " for _ in range(customIndentSpaces if customIndentSpaces > 0 else 0))
                verseLines.append(numberList[i] + customIndent + verseList[i][leftIndex:rightIndex] + "\n")
            else:
                verseLines.append(indent + verseList[i][leftIndex:] + "\n")

            formatedVerses.append(verseLines)

        return formatedVerses

    def _getSuperScriptIndexes(self, verseList):
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
                        ssIndexList.append([SSType.VerseNumber, index, numEndIndex + index])
                        currVerseNum = nextVerseNum
                else:
                    ssAlignmentSpace = self._getSSAlignmentSpace(len(str(currVerseNum)))
                    if (ssAlignmentSpace > 0):
                        ssIndexList.append([SSType.Space, index, ssAlignmentSpace + index])

                index += len(line)

        return ssIndexList

    # ==============================================================================================
    # =========================================== TOOLS ============================================
    # ==============================================================================================

    def _getRightMostMaxIndex(self, verse):
        # Based on the max line unit length, find the rightmost character that is within the length
        leftIndex = 0
        rightIndex = len(verse)

        if (self._getVisualLength(verse) <= self.maxLineLength):
            return rightIndex

        # Iteratively narrow down the index position
        while(leftIndex + 1 < rightIndex):
            middleIndex = int((leftIndex + rightIndex) / 2)
            currUnitLineLength = self._getVisualLength(verse[:middleIndex])
            if (currUnitLineLength > self.maxLineLength):
                rightIndex = middleIndex
            else:
                leftIndex = middleIndex

        return leftIndex

    def _getVisualLength(self, text):
        # A precise measurement to indicate if text will align or will take up multiple lines
        text = ''.join([i if ord(i) < 128 else ' ' for i in text])  # Replace all non-ascii characters
        return int(self.afm.string_width_height(text)[0])

    def _getVerseComponents(self, verses):
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

    def _getSSAlignmentSpace(self, ssCharLength):
        # Number of converted spaces need to align line with verse number and lines without;
        # assumes bible verse number does not exceed 3 digits
        if (ssCharLength == 3 or ssCharLength == 2):
            return 1
        elif (ssCharLength == 1):
            return 2
        else:
            return 0

# ==============================================================================================
# ============================================ TESTER ==========================================
# ==============================================================================================


if __name__ == '__main__':
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
        config.read(type + "SlideProperties.ini")

        if (vm.setSource(' '.join(sys.argv[2:]), int(config["SERMON_VERSE_PROPERTIES"]["SermonVerseMaxLineLength"]))):
            # GetContent() Test
            [title, verseList, ssIndexList] = vm.getContent()
            print("\n---------------" + title + "---------------")
            for verse in verseList:
                for line in verse:
                    print(line, end="")

            # GetVerseString() Test
            print("\n\n===========================================================\n")
            [verseString, ssIndexList] = vm.getVerseString()
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
