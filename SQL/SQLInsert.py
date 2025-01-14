import sqlite3

if __name__ == '__main__':
    try:
        # ============================== READ LYRICS FILE ==============================

        with open(r"SQL\Lyrics.txt", "r", encoding="utf8") as f:
            lyrics = f.read().strip().split("\n\n")

        if not lyrics:
            raise ValueError("The lyrics file is empty or improperly formatted.")

        hymnName = lyrics[0]
        length = len(lyrics)

		# ========================== INSERT LYRICS INTO DATABASE =======================

        # Connect to the SQLite database
        with sqlite3.connect(r'AutoPPTMaker\Data\HymnDatabase.db') as con:
            cursor = con.cursor()

			# Check if the hymn already exists in the database
            cursor.execute("SELECT * FROM Hymn WHERE hymnName = ?", (hymnName.upper(),))
            if cursor.fetchone():
                raise ValueError(f"The hymn [{hymnName}] already exists in the database.")

            # Insert verses into the database
            for i in range(1, length):
                cursor.execute(
                    "INSERT INTO Hymn (HymnName, Version, Number, End, Lyrics, Comments) "
                    "VALUES (?, ?, ?, ?, ?, ?)",
                    (hymnName.upper(), 1, i, length - 1, lyrics[i], "")
                )
            con.commit()

        print(f"Successfully inserted [{length - 1}] verses of [{hymnName}] into the database.")

    except FileNotFoundError:
        print("The specified lyrics file was not found.")
    except sqlite3.Error as e:
        print(f"Database error: {e}")
    except Exception as e:
        print(f"An error occurred: {e}")
