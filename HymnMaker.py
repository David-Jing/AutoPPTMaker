import configparser
import math
import re
import sys

from lyrics_extractor import lyrics

from WebLookupTools import WebLookupTools

'''
Looks up inputted title and generates formatted hymn slide texts.
Please run validate() before getContent().

titleList - list of title text corresponding to the order in lyricsList
lyricsList - list of formatted verses

'''


class HymnMaker:
    def __init__(self, type):
        # Get heuristics
        config = configparser.ConfigParser()
        config.read(type + "SlideProperties.ini")

        self.maxLines = int(config["HYMN_PROPERTIES"]["HymnMaxLines"])
        self.minLines = int(config["HYMN_PROPERTIES"]["HymnMinLines"])

    def setSource(self, hymnTitle):
        if (hymnTitle == ""):
            return False

        self.hymnTitle = hymnTitle
        self.hymn = WebLookupTools.getHymn(self.hymnTitle)

        return (self.hymn["title"] != "Not Found" and self.hymn["lyrics"] != "Not Found")

    def getContent(self):
        titleList = []
        lyricsList = []

        if (self.hymn != ""):
            lyricsList = self._getFormattedLyrics(self.hymn["lyrics"])

            # Generate titles and numbering
            slides = len(lyricsList)
            for i in range(1, slides + 1):
                titleList.append(f'{self.hymnTitle.upper()} ({i}/{slides})')

        return [titleList, lyricsList]

    # ==============================================================================================
    # ========================================= FORMATTER ==========================================
    # ==============================================================================================

    def _getFormattedLyrics(self, lyrics):
        # Remove [...] headers and lyrics in brackets
        lyrics = re.sub("(\[.*?\])|(\(.*?\))", "", lyrics)

        # Split by lyrics into its verses
        lyrics = lyrics.split("\n\n")

        # Count number of lines per verse
        lineCountList = []
        for i in range(len(lyrics)):
            lyrics[i] = self._cleanUpVerse(lyrics[i])
            lineCountList.append(lyrics[i].count("\n"))

        # Herustics for a 'well-formatted' slides
        lyricsList = []
        for verse in lyrics:
            if (verse == "\n"):
                continue

            appended = False
            # Check if the verse has too few lines
            if (verse.count("\n") < self.minLines):
                self._formatShort(lyricsList, verse, lyrics)
                appended = True

            # Check if the verse has too many lines
            if (verse.count("\n") > self.maxLines):
                self._formatLong(lyricsList, verse)
                appended = True

            if (not appended):
                # Check if the verse is repeated, use '(x2)' to indicate repeat instead
                coreVerse = self._getPrincipalPeriod(verse)
                if (coreVerse != ""):
                    lyricsList.append(coreVerse[:-1] + f" (x2)\n")
                else:
                    lyricsList.append(verse)

        return lyricsList

    def _formatLong(self, lyricsList, verse):
        # Find how many slides to split into
        numOfSlides = math.ceil(verse.count("\n") / self.maxLines)
        linesPerSlide = round(verse.count("\n")/numOfSlides)

        # Append splitted lyrics into list
        startIndex = 0
        endIndex = 0
        for _ in range(numOfSlides):
            for _ in range(linesPerSlide):
                endIndex = verse.find("\n", endIndex + 1)

            endIndex += 1  # Including the new line symbol
            lyricsList.append(verse[startIndex:endIndex])
            startIndex = endIndex

    def _formatShort(self, lyricsList, verse, lyrics):
        # Combines this verse to one from another slide
        if (len(lyricsList) < 1):
            lyrics[1] = verse + lyrics[1]
        else:
            preVerse = lyricsList.pop()
            newVerse = preVerse + verse

            # Combined verses can exceed maxLines
            if (newVerse.count("\n") > self.maxLines):
                self._formatLong(lyricsList, newVerse)
            else:
                lyricsList.append(newVerse)

    # ==============================================================================================
    # =========================================== TOOLS ============================================
    # ==============================================================================================

    def _getPrincipalPeriod(self, verse):
        # Checks if the verse is periodic and is equal to a nontrivial rotation of itself
        i = (verse+verse).find(verse, 1, -1)
        return "" if i == -1 else verse[:i]

    def _getMedian(self, lineCountList):
        # Finds the median number of new lines per verse in this song
        sortedLst = sorted(lineCountList)
        lstLen = len(lineCountList)
        index = (lstLen - 1) // 2

        if (lstLen % 2):
            return sortedLst[index]
        else:
            return (sortedLst[index] + sortedLst[index + 1])/2.0

    def _cleanUpVerse(self, verse):
        # Gets rid of [...] header
        index = verse.find(']') + 1
        verse = verse if index <= 0 else verse[index:]

        # Remove new lines before and after text
        return verse.rstrip().lstrip() + "\n"

# ==============================================================================================
# ============================================ TESTER ==========================================
# ==============================================================================================


if __name__ == '__main__':
    # python HymnMaker.py [hymn name]
    type = ""
    if (len(sys. argv) > 2):
        if (sys.argv[1] == "-s"):
            type = "Stream"
        elif (sys.argv[1] == "-p"):
            type = "Projected"

    if (type != ""):
        hm = HymnMaker(type)

        if (hm.setSource(' '.join(sys.argv[2:]))):
            print(hm.hymn["title"] + "\n")
            [titleList, lyricsList] = hm.getContent()
            for i in range(len(lyricsList)):
                print("-----------" + titleList[i] + "-----------")
                print(lyricsList[i])

        else:
            print("Song not found.")
