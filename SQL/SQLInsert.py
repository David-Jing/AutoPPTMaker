import sqlite3

if __name__ == '__main__':
	con = sqlite3.connect('../HymnDatabase.db')

	f = open("Lyrics.txt", encoding="utf8")
	lyrics = f.read()

	lyrics = lyrics.split("\n\n")

	length = len(lyrics) - 1
	hymnName = lyrics[0]
	for i in range(0, length):
		con.execute(f"INSERT INTO Hymn VALUES(\"{hymnName.upper()}\", 1, {i+1}, {length}, \"{lyrics[i+1]}\", \"\")")

	'''
	for row in con.execute("SELECT * FROM Hymn Where HymnName = 'All Glory Be To Christ'"):
	    print(row)
	'''

	con.commit()
	con.close()