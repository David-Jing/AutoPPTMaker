import sys
import requests
import sys
import os
import lyricsgenius as lg
from lyrics_extractor import SongLyrics


class WebLookupTools:

    def getHymn(name):
        # API Key and Engine ID of Google Custom Search JSON API
        # Refer to https://pypi.org/project/lyrics-extractor/ for more detail
        GCS_API_KEY = 'AIzaSyA8jw1Ws2yXn7BDqj4yYYJmE1BAK_J53zA'
        GCS_ENGINE_ID = '501493627fe694701'

        extract_lyrics = SongLyrics(GCS_API_KEY, GCS_ENGINE_ID)

        data = {
            'title': "Not Found",
            'lyrics': "Not Found"
        }

        try:
            data = extract_lyrics.get_lyrics(name)
        except:
            pass

        return data
        '''
        # Genius API
        # Refer to https://genius.com/api-clients for more details
        geniusService = lg.Genius('JTVclzLeCivYMkSaeLnDY1B9bN-wVOA9yeknV2VEiWblZ0X3FWeA0tYLmDo9LT7V',  # Client access token from Genius Client API page
                                  skip_non_songs=True, excluded_terms=["(Remix)", "(Live)"],
                                  remove_section_headers=True)

        # Block prints
        sys.stdout = open(os.devnull, 'w')

        hymn = geniusService.search_song(name)

        # Restore print
        sys.stdout = sys.__stdout__

        data = {
            'title': "Not Found",
            'lyrics': "Not Found"
        }

        if hymn is not None:
            data["title"] = hymn.title
            data["lyrics"] = hymn.lyrics

        return data
        '''

    def getVerse(passage):
        # ESV Bible Verse Lookup ID
        # Refer to https://api.esv.org/ for more details
        ESV_API_KEY = '6d6a8ca8f166e35b2c0343bfcdada88bd0e7b161'
        ESV_API_URL = 'https://api.esv.org/v3/passage/text/'

        params = {
            'q': passage,
            'include-headings': False,
            'include-footnotes': False,
            'include-verse-numbers': True,
            'include-short-copyright': False,
            'include-passage-references': False
        }

        headers = {
            'Authorization': 'Token %s' % ESV_API_KEY
        }

        response = requests.get(
            ESV_API_URL, params=params, headers=headers)

        passages = response.json()['passages']

        return passages[0].strip() if passages else 'Not Found'


if __name__ == '__main__':
    # python WebLookup.py -v [verse source]
    if (sys.argv[1] == '-v'):
        verse = ' '.join(sys.argv[2:])
        if verse:
            print(WebLookupTools.getVerse(verse))
    # python WebLookup.py -h [hymn name]
    elif (sys.argv[1] == '-h'):
        name = ' '.join(sys.argv[2:])
        if name:
            hymn = WebLookupTools.getHymn(name)
            print(hymn['title'] + f'\n')
            print(hymn['lyrics'])
