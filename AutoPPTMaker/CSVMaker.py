import configparser
import os
import time
import csv
import datetime

from enum import Enum

from datetime import timedelta
from HymnMaker import HymnMaker
from VerseMaker import VerseMaker

"""

Core class of the CSVMaker application

"""

class CSVMaker:
    class PPTMode(Enum):
        # Specifies the PPT output format
        Null = -1
        Stream = 0
        Projected = 1
        Regular = 2

    def __init__(self) -> None:
        # Access slide input data
        if not os.path.exists("SlideInputs.ini"):
            raise IOError("ERROR : SlideInputs.ini input file cannot be found.")
        self.input = configparser.ConfigParser()
        self.input.read("SlideInputs.ini")

    def setType(self, pptType: PPTMode) -> None:
        strType = pptType.name

        self.verseMaker = VerseMaker(strType)
        self.hymnMaker = HymnMaker(strType)

# ==============================================================================================
# ===================================== SLIDE MAKER IMPLEMENTATIONS ============================
# ==============================================================================================
    
    def _replaceChars(self, txt):
        txt = txt.replace(u'\u2005',' ')
        txt = txt.replace(u'\u205f',' ')
        txt = txt.replace(u'\u2013','-')
        txt = txt.replace(u'\u2014','-')
        txt = txt.replace(u'\u201C',u'\u0022')
        txt = txt.replace(u'\u201D',u'\u0022')
        txt = txt.replace(u'\u2018',u'\u0027')
        txt = txt.replace(u'\u2019',u'\u0027')
        txt = txt.replace(u'\u0435','e')
        return txt

    def _getCSVNextSundayDate(self, nextNextSundayDate: bool = False):
        # Get datetime variable with the date of next Sunday
        dt = datetime.date.today()

        if nextNextSundayDate:
            dt += timedelta(days=(13 - dt.weekday()))
        else:
            dt += timedelta(days=(6 - dt.weekday()))

        return dt.strftime('%Y-%m-%d')

# ==============================================================================================
# ======================================= APPLICATION RUNNER ===================================
# ==============================================================================================

if __name__ == "__main__":
    start = time.time()

    print("\nINITIALIZING...")
    cm: CSVMaker = CSVMaker()
    cm.setType(CSVMaker.PPTMode.Regular)

    filename = "service-guide.csv"
    csvfile = open(filename, 'w+',  newline='')
    fields = [  'Guide Title',\
                'Communion',\
                'Worship Leader',\
                'Date',\
                'Call to Worship Title',\
                'Call to Worship',\
                'Song 1 Title',\
                'Song 1',\
                'Song 2 Title',\
                'Song 2',\
                'Call to Confession Title',\
                'Call to Confession',\
                'Preacher',\
                'Sermon Title',\
                'Scripture Text Title',\
                'Scripture Text',\
                'Song 3 Title',\
                'Song 3',\
                'Song 4 Title',\
                'Song 4',\
                'Monthly Scripture Title',\
                'Monthly Scripture',\
                'Bible Memorization Title',\
                'Bible Memorization',
                'Catechism'\
            ]

    communion = cm.input["SLIDE_ORDERING_MODE"]["SlideOrderMode"] == "1"
    
    hymn1 = cm.input["HYMN"]["Hymn1Source"]
    hymn2 = cm.input["HYMN"]["Hymn2Source"]
    hymn3 = cm.input["HYMN"]["Hymn3Source"]
    hymn4 = cm.input["HYMN"]["Hymn4Source"]

    title = cm.input["SUNDAY_SERVICE_HEADER"]["SundayServiceHeaderTitle"]
    leader = cm.input["WORSHIP_LEADER"]["WorshipLeader"]
    calltoworship = cm.input["CALL_TO_WORSHIP"]["CallToWorshipSource"]
    calltoconfession = cm.input["PRAYER_OF_CONFESSION"]["PrayerOfConfessionSource"]
    preacher = cm.input["SERMON_HEADER"]["SermonHeaderSpeaker"]
    sermontitle = cm.input["SERMON_HEADER"]["SermonHeaderTitle"]
    sermonverses = cm.input["SERMON_VERSE"]["SermonVerseSource"]
    monthlyscripture = cm.input["MONTHLY_SCRIPTURE"]["MonthlyScriptureSource"]
    biblememorization = cm.input["BIBLE_MEMORIZATION"]["BibleMemorizationLastWeekSource"]
    catechism = cm.input["CATECHISM"]["CatechismLastWeekTitle"]

    date = cm._getCSVNextSundayDate()

    def getLyrics(source):
        cm.hymnMaker.setSource(source)
        lyrics = '\n\n'.join(cm.hymnMaker.getContent(False)[1])
        lyrics = cm._replaceChars(lyrics)
        return lyrics

    def getVerses(source):
        cm.verseMaker.setSource(source, 100000000)
        verses = ''.join(['\n'.join(lst) for lst in cm.verseMaker.getContent()[1]])
        verses = cm._replaceChars(verses)
        return verses
    
    if communion:
        row = [ title,\
                communion,\
                leader,\
                date,\
                '',\
                '',\
                hymn1,\
                getLyrics(hymn1),\
                '',\
                '',\
                '',\
                '',\
                preacher,\
                sermontitle,\
                sermonverses,\
                getVerses(sermonverses),\
                hymn2,\
                getLyrics(hymn2),\
                hymn3,\
                getLyrics(hymn3),\
                monthlyscripture,\
                getVerses(monthlyscripture),\
                biblememorization,\
                getVerses(biblememorization),\
                catechism.title()\
            ]
    else:
        row = [ title,\
                communion,\
                leader,\
                date,\
                calltoworship,\
                getVerses(calltoworship),\
                hymn1,\
                getLyrics(hymn1),\
                hymn2,\
                getLyrics(hymn2),\
                calltoconfession,\
                getVerses(calltoconfession),\
                preacher,\
                sermontitle,\
                sermonverses,\
                getVerses(sermonverses),\
                hymn3,\
                getLyrics(hymn3),\
                hymn4,\
                getLyrics(hymn4),\
                monthlyscripture,\
                getVerses(monthlyscripture),\
                biblememorization,\
                getVerses(biblememorization),\
                catechism.title()\
            ]

    csvwriter = csv.writer(csvfile)
    csvwriter.writerow(fields)
    csvwriter.writerow(row)

    print(f"\nTask completed in {(time.time() - start):.2f} seconds.\n")
    print("====================================================================\n")
