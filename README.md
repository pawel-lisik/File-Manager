# File-Manager
File Manager with all the functions you need! For windows.

## Some of the functions:
* **Preview Many Files:** Provides instant floating previews, controlled via keyboard shortcuts like **Spacebar** or **Escape** to close.
* **Windows Dark/Light Theme Integration:** Reads the Windows Registry (`AppsUseLightTheme`) to automatically match the current system appearance.
* **Image & Animated GIF Viewing:** Renders various picture formats (`.jpg`, `.png`, `.webp`, `.bmp`, `.svg`) with smooth aspect-ratio scaling and plays `.gif` animations.
* **Video Playback:** Plays back common video formats (`.mp4`, `.avi`, `.mov`, `.mkv`) natively using embedded video widgets.
* **Audio Playback & Metadata Extraction:** Plays audio files (`.mp3`, `.wav`, `.flac`, `.m4a`, `.ogg`) while parsing embedded album art, song titles, artists, and album tags using `mutagen`.
* **Document & Text Preview:** Reads and renders plain text (`.txt`, `.rtf`), Microsoft Word documents (`.doc`, `.docx`), and renders PDF pages using `PyMuPDF`.
* **Folder & Disk Inspector:** Displays path details, drive type icons (HDD vs. USB), file counts, and asynchronously calculates directory sizes on a background thread (`QThread`).
* **External Application Launcher:** Includes an "Open in..." shortcut button to launch current files directly into pre-configured default desktop programs.
* **Files**: move, copy, paste, rename, send to trash, empty trash
* **4 File views**: grid, column, list and gallery.
* **Tags**: tag your files!
* ...and many more!

## Setup:

`pyinstaller --noupx --noconsole --icon="icon.ico" main.py`

Don't use `--onefile` flag, because it will slow down the app!

To open files in external default apps or to open folders in terminal, add the paths to the executables in the code! I use my custom programs, so I left the paths empty.


## Requirements:

* PyQt5
* pywin32
* psutil
* send2trash
* winshell
* Pillow
* PyMuPDF
* mutagen
* python-docx

## to do: 
create a helper method to get all the custom message boxes and simplify the code. But for now i'm not touching it because the code works :)

