import os
import music21 as m21


def load_songs_in_mid():
    songs = []

    for path, subdirs, files in os.walk("./"):
        for file in files:
            if file[-3:] == "mid":
                song = m21.converter.parse(os.path.join(path, file))
                songs.append(song)
    return songs


if __name__ == "__main__":
    songs = load_songs_in_mid()
    print(songs[0].flat.notesAndRests)
