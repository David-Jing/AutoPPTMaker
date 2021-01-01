import configparser
import sys
import math

from WebLookupTools import WebLookupTools

'''
Looks up inputted biblical verse and generates formatted verse slide texts.
Please setSource() before running any getters.

titleList - list of title text corresponding to the order in lyricsList
verseList - list of formatted verses
ssIndexList - list of lists of character indexes that requires super scripting
'''


class VerseMaker:
    def __init__(self, type):
        self.verseSource = ""
        self.verses = ""

        # Get heuristics
        config = configparser.ConfigParser()
        config.read(type + "SlideProperties.ini")

        self.maxChar = -1
        self.indentSpace = int(config["VERSE_PROPERTIES"]["VerseIndentSpace"])

    def setSource(self, verseSource, maxChar):
        # Returns true if source is valid, false otherwise
        if (verseSource == ""):
            return False

        self.maxChar = maxChar
        self.verseSource = verseSource
        self.verses = WebLookupTools.getVerse(self.verseSource)

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
        title = "—" + title.capitalize()
        sourceIndentSpace = int(1.5*self.maxChar) - len(title)
        titleIndent = "".join(" " for i in range(sourceIndentSpace if sourceIndentSpace > 0 else 0))
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
            else math.ceil(lineCount / maxLinesPerSlide + 0.1)                  # Heuristic : "Fudging" the statistic (guess-timation)
        linesPerSlide = round((lineCount/numOfSlides + maxLinesPerSlide) / 2)   # Heuristic : Can't have slides too empty

        # Split verses for each slide and recalculate ssIndexList
        slideVersesList = []
        slideSSIndexList = []   # Assumes each verse has only one symbol to superscript
        verseIndex = 0          # Tracking verse number
        charIndex = 0           # Tracking the number of character parsed, for adjusting ssIndexes
        numOfVerses = len(verseList)
        while(verseIndex < numOfVerses):
            verseString = ""
            currSSSlideIndexList = []
            currTotalLineCount = 0
            nextlineCount = 0 if verseIndex >= numOfVerses else len(verseList[verseIndex])

            # Asserts that a slide must contain at least one verse
            while (nextlineCount > 0 and (currTotalLineCount == 0 or currTotalLineCount + nextlineCount <= linesPerSlide)):
                for line in verseList[verseIndex]:
                    verseString += line

                currSSSlideIndexList.append([ssIndexList[verseIndex][0] - charIndex, ssIndexList[verseIndex][1] - charIndex])
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
        indent = "".join(" " for i in range(self.indentSpace))

        for i in range(len(verseList)):
            # Split verse into separate lines that are within the character limit
            verseLines = []
            leftIndex = 0
            rightIndex = self.maxChar

            # Iteratively cut string into lengths less than maximum
            while (rightIndex < len(verseList[i])):
                while (verseList[i][rightIndex] != " " and rightIndex >= 0):
                    rightIndex -= 1

                # Generate custom indent for line with verse numbers
                if (leftIndex == 0):
                    customIndentSpaces = self.indentSpace - len(numberList[i]) - 1
                    customIndent = "".join(" " for _ in range(customIndentSpaces if customIndentSpaces > 0 else 0))
                    verseLines.append(numberList[i] + customIndent + verseList[i][leftIndex:rightIndex] + "\n")
                else:
                    verseLines.append(indent + verseList[i][leftIndex:rightIndex] + "\n")

                leftIndex = rightIndex + 1
                rightIndex += self.maxChar

            # Append the remaining verses and add to list of formatted verses
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
                # Assumes verse numbers increments by 1
                if (line[0].isnumeric()):
                    numEndIndex = line.find(" ")
                    nextVerseNum = int(line[0:numEndIndex])
                    if (numEndIndex != -1 and (currVerseNum < 0 or currVerseNum + 1 == nextVerseNum)):
                        ssIndexList.append([index, numEndIndex + index])
                        currVerseNum = nextVerseNum

                index += len(line)

        return ssIndexList

    # ==============================================================================================
    # =========================================== TOOLS ============================================
    # ==============================================================================================

    def _getVerseComponents(self, verses):
        # Split multiple verses into individuals
        verseNumberList = []
        verseList = []

        # Clean up empty spaces and new lines
        verses = " ".join(verses.split())

        # Iterate verses by search for "[" and "]"
        currIndex = 0
        prevIndex = verses.find("]", currIndex)
        while (True):
            currIndex = verses.find("[", currIndex)

            if (currIndex > prevIndex):
                verseList.append(verses[prevIndex:currIndex].rstrip().lstrip())

            if (currIndex < 0):
                verseList.append(verses[prevIndex:].rstrip().lstrip())
                break

            currIndex += 1
            prevIndex = verses.find("]", currIndex) + 1
            verseNumberList.append(verses[currIndex:prevIndex - 1])

        return [verseList, verseNumberList]


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

    print(type)

    if (type != ""):
        vm = VerseMaker(type)
        config = configparser.ConfigParser()
        config.read(type + "SlideProperties.ini")

        if (vm.setSource(' '.join(sys.argv[2:]), int(config["SERMON_VERSE_PROPERTIES"]["SermonVerseCharPerLine"]))):
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
            [title, slideVersesList, ssIndexList] = vm.getVerseStringMultiSlide(int(config["SERMON_VERSE_PROPERTIES"]["SermonVerseMaxLines"]))
            numOfVerses = len(slideVersesList)
            print("\n---------------" + title + "---------------")
            for i in range(numOfVerses):
                if (i != numOfVerses - 1):
                    print(slideVersesList[i], end="")
                    print("\n---------------" + "".join("-" for _ in range(len(title))) + "---------------")
                else:
                    print(slideVersesList[i])

        else:
            print("Verse not found.")
