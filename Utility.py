import matplotlib
import os

from matplotlib.afm import AFM


class Utility:

    initialized = False
    afm: AFM = None

    @staticmethod
    def initializeUtility() -> None:
        if Utility.initialized:
            return
        Utility.initialized = True

        # For finding visual lengths of text strings
        afm_filename = os.path.join(matplotlib.get_data_path(), 'fonts', 'afm', 'ptmr8a.afm')
        Utility.afm = AFM(open(afm_filename, "rb"))

    @staticmethod
    def getVisualLength(text: str) -> int:
        Utility.initializeUtility()

        # A precise measurement to indicate if text will align or will take up multiple lines
        text = ''.join([i if ord(i) < 128 else ' ' for i in text])  # Replace all non-ascii characters
        return int(Utility.afm.string_width_height(text)[0])

# ==============================================================================================
# ============================================ TESTER ==========================================
# ==============================================================================================


if __name__ == '__main__':
    inputStr = ""
    while True:
        inputStr = str(input())

        if inputStr == 'q':
            break

        print(f"Length = {Utility.getVisualLength(inputStr)}")
