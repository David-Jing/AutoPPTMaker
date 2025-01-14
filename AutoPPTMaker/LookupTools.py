import sys
import requests
import sqlite3
from typing import Union

"""

Looks up ESV bible verses and song lyrics.
Song lyrics are picked from either a local SQL server or the web;
the SQL server provides better lookup accuracy than the web.

"""


class LookupTools:
    @staticmethod
    def getHymn(name: str) -> dict:
        start = 0
        end = 0
        hymnName = ""
        lyrics = ""

        # ==========================================================================================
        # ====================================== SQL LOOKUP ========================================
        # ==========================================================================================

        con = sqlite3.connect("Data/HymnDatabase.db")

        for row in con.execute(f"SELECT * FROM Hymn WHERE Replace(HymnName, ',', '') LIKE \"%{name.replace(',', '')}%\" \
                                                AND VERSION = 1 ORDER BY HymnName, Number"):
            hymnName = row[0]
            start = row[2]
            end = row[3]

            lyrics += row[4] + ("\n\n" if start != end else "")

            # In case there're multiple matches
            if start == end:
                break

        con.close()

        if start != end:
            print(f"ERROR: Database consistency error, {start}/{end} lyrics found.")
        elif len(lyrics) > 0:
            print("  LYRICS SOURCE: SQL")
            return {
                "source": "SQL",
                "title": hymnName,
                "lyrics": lyrics
            }

        # ==========================================================================================

        return {
            "source": "None",
            "title": name,
            "lyrics": ""
        }

    @staticmethod
    def getVerse(passage: str) -> str:
        # ESV Bible Verse Lookup ID
        # Refer to https://api.esv.org/ for more details
        ESV_API_KEY = "6d6a8ca8f166e35b2c0343bfcdada88bd0e7b161"
        ESV_API_URL = "https://api.esv.org/v3/passage/text/"

        params: dict[str, Union[bool, str]] = {
            "q": passage,
            "include-headings": False,
            "include-footnotes": False,
            "include-verse-numbers": True,
            "include-short-copyright": False,
            "include-passage-references": False
        }

        headers = {
            "Authorization": "Token %s" % ESV_API_KEY
        }

        response = requests.get(
            ESV_API_URL, params=params, headers=headers)

        passages = response.json()["passages"]

        return passages[0].strip() if passages else "Not Found"

# ==============================================================================================
# ============================================ TESTER ==========================================
# ==============================================================================================


if __name__ == "__main__":
    # python LookupTools.py -v [verse source]
    if (sys.argv[1] == "-v"):
        verse = " ".join(sys.argv[2:])
        if verse:
            print(LookupTools.getVerse(verse))
    # python LookupTools.py -h [hymn name]
    elif (sys.argv[1] == "-h"):
        name = " ".join(sys.argv[2:])
        if name:
            hymn = LookupTools.getHymn(name)
            print(hymn["title"] + "\n")
            print(hymn["lyrics"])
