from typing import Any, List
from GoogleAPITools import GoogleAPITools

from Utility import Utility


class ListMaker:
    @staticmethod
    def getFormattedListElementsBySlide(listContent: List[str], indentSpace: int, maxLineLength: int, maxLinesPerSlide: int) -> List[List[Any]]:
        numberList = [str(i+1) + "." for i in range(len(listContent))]
        lineLengthRestrictedList = ListMaker.setLineLengthRestriction(listContent, numberList, indentSpace, maxLineLength)

        [slideContentList, listStartIndex] = ListMaker.setLinesPerSlideRestriction(lineLengthRestrictedList, maxLinesPerSlide)

        return [slideContentList, listStartIndex]

    # ==============================================================================================
    # ========================================= FORMATTER ==========================================
    # ==============================================================================================

    @staticmethod
    def setLineLengthRestriction(listContent: List[str], numberList: List[str], indentSpace: int, maxLineLength: int) -> List[List[str]]:
        formattedOut = []

        indent = "".join(" " for _ in range(indentSpace))

        for i in range(len(listContent)):
            # Split verse into separate lines that are within the character limit
            textLines = []
            leftIndex = 0
            rightIndex = ListMaker._getRightMostMaxIndex(listContent[i], maxLineLength)

            # Iteratively cut string into lengths less than maximum
            while (rightIndex < len(listContent[i])):
                # Find " " character left of the maxLineLength Index
                while (listContent[i][rightIndex] != " " and rightIndex >= 0):
                    rightIndex -= 1

                # Generate custom indent for line with numbers
                if (leftIndex < 1):
                    customIndentSpaces = indentSpace - len(numberList[i])
                    customIndent = "".join(" " for _ in range(customIndentSpaces if customIndentSpaces > 0 else 0))
                    textLines.append(numberList[i] + customIndent + listContent[i][leftIndex:rightIndex] + "\n")
                else:
                    textLines.append(indent + listContent[i][leftIndex:rightIndex] + "\n")

                leftIndex = rightIndex
                rightIndex += ListMaker._getRightMostMaxIndex(listContent[i][leftIndex:], maxLineLength) + 1

            # Append the remaining verses and add to list of formatted verses
            if (leftIndex < 1):
                customIndentSpaces = indentSpace - len(numberList[i])
                customIndent = "".join(" " for _ in range(customIndentSpaces if customIndentSpaces > 0 else 0))
                textLines.append(numberList[i] + customIndent + listContent[i][leftIndex:rightIndex] + "\n")
            else:
                textLines.append(indent + listContent[i][leftIndex:] + "\n")

            formattedOut.append(textLines)

        return formattedOut

    @staticmethod
    def setLinesPerSlideRestriction(lineLengthRestrictedContent: List[List[str]], maxLinesPerSlide: int) -> List[List[Any]]:
        # Use greedy method to fill slides
        slideContentList = []
        listStartIndex = []

        currLines = -1  # Offset initial paragraph spacing
        currIndex = 0
        currContentList = []
        currListStartIndex = []
        for content in lineLengthRestrictedContent:
            contentStr = "".join(content)

            if len(content) + currLines + 1 < maxLinesPerSlide:
                currContentList.append(contentStr)
                currListStartIndex.append(currIndex)

                currLines += 1
            else:
                currLines = -1

                slideContentList.append("".join(currContentList))
                listStartIndex.append(currListStartIndex)

                currContentList = [contentStr]
                currListStartIndex = [0]
                currIndex = 0

            currIndex += len(contentStr)
            currLines += len(content)

        # Append remaining currContentList
        if len(currContentList) > 0:
            slideContentList.append("".join(currContentList))
            listStartIndex.append(currListStartIndex)

        return [slideContentList, listStartIndex]

    # ==============================================================================================
    # =========================================== TOOLS ============================================
    # ==============================================================================================

    @staticmethod
    def _getRightMostMaxIndex(verse: str, maxLineLength: int) -> int:
        # Based on the max line unit length, find the rightmost character that is within the length
        leftIndex = 0
        rightIndex = len(verse)

        if (Utility.getVisualLength(verse) <= maxLineLength):
            return rightIndex

        # Iteratively narrow down the index position
        while(leftIndex + 1 < rightIndex):
            middleIndex = int((leftIndex + rightIndex) / 2)
            currUnitLineLength = Utility.getVisualLength(verse[:middleIndex])
            if (currUnitLineLength > maxLineLength):
                rightIndex = middleIndex
            else:
                leftIndex = middleIndex

        return leftIndex

# ==============================================================================================
# ============================================ TESTER ==========================================
# ==============================================================================================


if __name__ == "__main__":
    peS = GoogleAPITools("Stream")
    sups = peS.getSupplications()

    [slideContentList, listStartIndex] = ListMaker.getFormattedListElementsBySlide(sups, 4, 16000, 12)
    print(listStartIndex)
    for out in slideContentList:
        print(out)
        print("=========================")
