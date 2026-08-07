import io
import sys
import winreg
import time
import pythoncom
import threading
import win32api
import win32con
import win32gui
import win32com.client
import os
import subprocess
from PyQt5.QtWidgets import (QApplication, QSystemTrayIcon, QMenu, QAction, QLabel, 
                            QMainWindow, QVBoxLayout, QHBoxLayout, QWidget, QPushButton,
                            QFileDialog, QSizePolicy, QScrollArea, QTextEdit)
from PyQt5.QtGui import QIcon, QPixmap, QMovie, QPainterPath, QRegion, QPainter, QImage
from PyQt5.QtCore import Qt, QObject, pyqtSignal, QSize, QUrl, QRectF, QThread
from PyQt5.QtMultimedia import QMediaPlayer, QMediaContent
from PyQt5.QtMultimediaWidgets import QVideoWidget

from mutagen import File
from mutagen.id3 import ID3, APIC, TIT2, TPE1, TALB
from mutagen.mp4 import MP4Cover

SUPPORTED_EXT = ('.jpg', '.jpeg', '.png', '.ico', '.svg', '.gif', '.mp4', '.avi', '.mov', '.mkv', '.pdf', '.txt', '.mp3', '.wav', '.flac', '.m4a', '.ogg', '.aif', '.aiff', '.webp', '.bmp', '.rtf', '.doc', '.docx')

# PUT YOUR FAV PROGRAM PATHS HERE!!:

PROGRAM_PATHS = {
    'pdf': r"C:",
    'photo': r"C:",
    'video': r"C:",
    'music': r"C:",
    'text': r"C:",
    'archieve': r"C:"
}

def is_dark_mode_enabled():
    """Checks if Windows is in dark mode by reading the registry."""
    try:
        reg_key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize")
        value, _ = winreg.QueryValueEx(reg_key, "AppsUseLightTheme")
        winreg.CloseKey(reg_key)
        return value == 0  # 0 means dark mode, 1 means light mode
    except Exception as e:
        return False
        




def extract_cover_pixmap(filepath):
    audio = File(filepath)

    if audio is None:
        return get_default_cover()

    try:
        # MP3 (ID3)
        if audio.tags and isinstance(audio.tags, ID3):
            for tag in audio.tags.values():
                if isinstance(tag, APIC):
                    cover_data = tag.data
                    image = QImage.fromData(cover_data)
                    return QPixmap.fromImage(image)

        # M4A, MP4
        elif hasattr(audio, "tags") and audio.tags and 'covr' in audio.tags:
            cover = audio.tags['covr'][0]
            if isinstance(cover, MP4Cover):
                image = QImage.fromData(bytes(cover))
                return QPixmap.fromImage(image)

    except Exception as e:
        print("Error loading cover:", e)

    return get_default_cover()




def get_default_cover():
    # Create new 600x600 default cover
    pixmap = QPixmap(600, 600)
    pixmap.fill(Qt.lightGray)
    
    # Draw musical note icon
    painter = QPainter(pixmap)
    painter.setPen(Qt.darkGray)
    font = painter.font()
    font.setPointSize(120)  # Larger font for 600x600 image
    painter.setFont(font)
    painter.drawText(pixmap.rect(), Qt.AlignCenter, "♫")
    painter.end() 
    return pixmap    
        

class FolderSizeWorker(QThread):
    size_computed = pyqtSignal(int)

    def __init__(self, path):
        super().__init__()
        self.path = path

    def run(self):
        total_size = 0
        for dirpath, _, filenames in os.walk(self.path):
            for f in filenames:
                fp = os.path.join(dirpath, f)
                if os.path.exists(fp):
                    try:
                        total_size += os.path.getsize(fp)
                    except:
                        pass  # Ignore files that can't be accessed
        self.size_computed.emit(total_size)



class Communicate(QObject):
    toggle_preview_signal = pyqtSignal(str)
    close_preview_signal = pyqtSignal()

class MediaPreview(QMainWindow):
    def __init__(self, path):
        super().__init__()
        self.setWindowTitle("Quick Look")
        self.setWindowIcon(QIcon("icon.ico"))
        self.setWindowFlags(Qt.WindowStaysOnTopHint | Qt.FramelessWindowHint)
        self.setWindowOpacity(0.99)
        self.setFocusPolicy(Qt.NoFocus)
        
        self.text_edit = QTextEdit()
        self.is_fullscreen = False
        self.original_geometry = None
        self.default_size = QSize(1200, 844)
        self.path = path
        self.media_player = None
        self.movie = None

        self.main = QWidget()
        main_layout = QVBoxLayout()
        self.main.setLayout(main_layout)
        preview = QVBoxLayout()

        bar = QHBoxLayout()

        self.open_in_btn = QPushButton("Open in...")
        self.open_in_btn.setFocusPolicy(Qt.NoFocus)
        self.open_in_btn.clicked.connect(self.open_in_external)

        self.fullscreen_btn = QPushButton()
        self.fullscreen_btn.setStyleSheet("background: none; border: none")
        self.fullscreen_btn.setIcon(QIcon("icons/media_preview/maximize.png"))
        self.fullscreen_btn.setIconSize(QSize(35, 35))
        self.fullscreen_btn.setFixedSize(44, 44)
        self.fullscreen_btn.clicked.connect(self.toggle_fullscreen)


        self.fullscreen_btn.setFocusPolicy(Qt.NoFocus)


        close_btn = QPushButton()
        close_btn.setFocusPolicy(Qt.NoFocus)
        
        close_btn.setStyleSheet("background: none; border: none")
        close_btn.setIcon(QIcon("icons/media_preview/close.png"))
        close_btn.setIconSize(QSize(35, 35))
        close_btn.setFixedSize(44, 44)
        close_btn.clicked.connect(self.close)
        
        spacer = QLabel()
        spacer.setFixedSize(14, 14)
        spacer.setStyleSheet("border: none; background: none")

        bar.addWidget(self.open_in_btn)
        bar.addStretch()
        bar.addWidget(self.fullscreen_btn)
        bar.addWidget(spacer)
        bar.addWidget(close_btn)

        main_layout.addLayout(bar)
        main_layout.addLayout(preview)

        self.media_widget = QWidget()
        self.media_widget.setStyleSheet("border: none")
        self.media_layout = QVBoxLayout(self.media_widget)
        preview.addWidget(self.media_widget)

        self.setCentralWidget(self.main)
        self.resize(self.default_size)
        
        self.load_media()
               
        if is_dark_mode_enabled(): 
            self.main.setStyleSheet("""
                background: #2a2a28;
                border: 3px solid #393939;
                border-radius: 12px;
                color: #c8c8c8;
            """)
            
            self.text_edit.setStyleSheet("""
                background-color: #1e1e1e;
                color: #aaaaaa;
                border-radius: 0px;
                font-size: 25px;
            """)

            self.open_in_btn.setStyleSheet("""
                QPushButton {
                    background-color: #61605f;
                    border: 1px solid #727170;
                    color: white;
                    border-radius: 8px;
                    padding: 7px;
                }

                QPushButton:pressed {
                    background-color: #0064e1; 
                    border: 1px solid #001c40;
                    color: white;
                }
            """)

        else:
            self.main.setStyleSheet("""
                background: #e0e0e0;
                border: 3px solid rgba(209, 209, 209, 175);
                border-radius: 12px;
                color: #3c3c3c;
            """)
            
            self.text_edit.setStyleSheet("""
                background-color: white;
                color: #3c3c3c;
                border-radius: 0px;
                font-size: 25px;
            """)
            
            self.open_in_btn.setStyleSheet("""
                QPushButton {
                    background-color: white;
                    border: 1px solid #cccccc;
                    color: black;
                    border-radius: 8px;
                    padding: 7px;
                }
                
                
                QPushButton:pressed {
                    background-color: #0064e1; 
                    border: 1px solid #001c40;
                    color: white;
                }
            """)


        self.setStyleSheet("""
            QScrollBar:horizontal {
                width: 0px;
                height: 0px;
            }
        
            QScrollBar:vertical {
                background-color: rgba(135, 135, 135, 0.00);
                width: 12px;
                margin: 1px;
            }

            QScrollBar::handle:vertical {
                background: rgba(180, 180, 180, 0.50);
                min-height: 20px;
                border-radius: 4px;
            }

            QScrollBar::add-line:vertical,
            QScrollBar::sub-line:vertical {
                height: 0px;
            }

            QScrollBar::add-page:vertical,
            QScrollBar::sub-page:vertical {
                background: none;
            }


        """)

        
        # Its semi transparent 
        self.setAttribute(Qt.WA_TranslucentBackground, False)
        self.setAttribute(Qt.WA_NoSystemBackground, False)
        
        # needed the mask, otherwise it was ugly 
        self.setMaskForRoundedCorners()



    def resizeEvent(self, event):
        super().resizeEvent(event)
        self.setMaskForRoundedCorners()
        if hasattr(self, 'pixmap'):
            self.label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
            self.label.setAlignment(Qt.AlignCenter)
            self.update_pixmap()
        if hasattr(self, 'label') and isinstance(self.label, QLabel):
            if hasattr(self, 'movie') and self.movie is not None:
                self.label.setFixedSize(self.media_widget.size())


    def open_in_external(self):
        """Open the file in its default external application based on its extension."""
        if not os.path.exists(self.path):
            return
            
        ext = os.path.splitext(self.path)[1].lower()
        
        if ext in ('.pdf'):
            program = PROGRAM_PATHS['pdf']
        elif ext in ('.jpg', '.jpeg', '.png', '.ico', '.svg', '.webp', '.bmp'):
            program = PROGRAM_PATHS['photo']
        elif ext in ('.mp4', '.avi', '.mov', '.mkv'):
            program = PROGRAM_PATHS['video']
        elif ext in ('.mp3', '.wav', '.flac', '.m4a', '.ogg', '.aif', '.aiff'):
            program = PROGRAM_PATHS['music']
        elif ext in ('.txt', '.rtf'):
            program = PROGRAM_PATHS['text']
        elif ext in ('.rar', '.zip'):
            program = PROGRAM_PATHS['archieve']                        
        else:
            return
            
        if os.path.exists(program):
            try:
                subprocess.Popen([program, self.path])
            except Exception as e:
                print(f"Error opening file: {e}")
                
                
    def is_supported_file(self, path):
        if os.path.isdir(path):
            return False
        return path.lower().endswith(SUPPORTED_EXT)

    def load_media(self):
        """Load media content based on current self.path"""
        # Clear previous content
        for i in reversed(range(self.media_layout.count())): 
            self.media_layout.itemAt(i).widget().setParent(None)
        
        if os.path.isdir(self.path):
            self.show_file_info()
        elif self.path.lower().endswith(('.jpg', '.jpeg', '.png', '.ico', '.svg', '.webp', '.bmp')):
            self.load_image()
        elif self.path.lower().endswith('.gif'):
            self.load_gif()
        elif self.path.lower().endswith(('.mp4', '.avi', '.mov', '.mkv')):
            self.load_video()
        elif self.path.lower().endswith(('.mp3', '.wav', '.flac', '.m4a', '.ogg', '.aif', '.aiff')):
            self.load_audio()
        elif self.path.lower().endswith('.pdf'):
            self.load_pdf()
        elif self.path.lower().endswith(('.txt', '.rtf')):
            self.load_txt()
        elif self.path.lower().endswith(('.doc', '.docx')):
            self.load_doc()
        else:
            self.show_file_info()  # Default view for unsupported files
            
    def load_doc(self):
        try:
            from docx import Document
            doc = Document(self.path)
            text = "\n".join([para.text for para in doc.paragraphs])
            
            
            self.text_edit.setPlainText(text)
            self.text_edit.setReadOnly(True)
            self.media_layout.addWidget(self.text_edit)
            self.open_in_btn.hide()
        except Exception as e:
            print(f"Nie można załadować DOC/DOCX: {e}")
            
    def show_file_info(self):
        container = QWidget()
        layout = QHBoxLayout(container)

        icon = QLabel()
        icon.setFixedSize(512, 512)

        if os.path.isdir(self.path):
            normalized_path = os.path.abspath(self.path).replace("/", "\\").lower()
            
            if normalized_path == "c:\\":
                icon_obj = QPixmap("icons/hdd.png")
            elif len(normalized_path) == 3 and normalized_path[1:] == ":\\":  # np. "d:\", "e:\"
                icon_obj = QPixmap("icons/usb.png")
            else:
                icon_obj = QPixmap("icons/media_preview/folder.png")
        elif self.path.lower().endswith(('.rar')):
            icon_obj = QPixmap("icons/media_preview/rar.png")
        elif self.path.lower().endswith(('.zip')):
            icon_obj = QPixmap("icons/media_preview/zip.png")
        else:
            icon_obj = QPixmap("icons/media_preview/file.png")

        scaled_pixmap = icon_obj.scaled(512, 512, Qt.KeepAspectRatio, Qt.SmoothTransformation)
        icon.setPixmap(scaled_pixmap)
        icon.setAlignment(Qt.AlignCenter)
        layout.addWidget(icon, 0, Qt.AlignCenter)

        info_widget = QWidget()
        info_widget.setFixedWidth(650)
        info_layout = QVBoxLayout(info_widget)
        info_layout.addStretch()

        name_label = QLabel(os.path.basename(self.path))
        name_label.setWordWrap(True)
        name_label.setStyleSheet("font-size: 34px; font-weight: bold")
        info_layout.addWidget(name_label)

        spacer4 = QLabel()
        info_layout.addWidget(spacer4)
        
        size_label = QLabel("Size: Calculating...")
        size_label.setStyleSheet("font-size: 28px")
        info_layout.addWidget(size_label)

        if os.path.isdir(self.path):
            file_count = self.count_files(self.path)
            if file_count <= 1000:
                self.size_worker = FolderSizeWorker(self.path)
                self.size_worker.size_computed.connect(lambda size: size_label.setText(f"Size: {self.format_size(size)}"))
                self.size_worker.start()
            else:
                size_label.setText("Size: Calculating...")  # na stałe
        else:
            size_bytes = os.path.getsize(self.path)
            size_label.setText(f"Size: {self.format_size(size_bytes)}")   

        spacer2 = QLabel()
        info_layout.addWidget(spacer2)
        
        path_label = QLabel(self.path)
        path_label.setStyleSheet("font-size: 28px")

        info_layout.addWidget(path_label)
        info_layout.addStretch()
        layout.addWidget(info_widget)

        self.media_layout.addWidget(container)
        if self.path.lower().endswith(('.rar', '.zip')):
            self.open_in_btn.setText("Open in Archive Utility")
            self.open_in_btn.show()
        else: 
            self.open_in_btn.hide()
        
    def get_folder_size(self, path):
        total_size = 0
        for dirpath, dirnames, filenames in os.walk(path):
            for f in filenames:
                fp = os.path.join(dirpath, f)
                if os.path.exists(fp):
                    total_size += os.path.getsize(fp)
        return total_size

    def format_size(self, size_bytes):
        for unit in ['B', 'KB', 'MB', 'GB', 'TB']:
            if size_bytes < 1024.0:
                return f"{size_bytes:.2f} {unit}"
            size_bytes /= 1024.0
        return f"{size_bytes:.2f} PB"
        
    def count_files(self, path, limit=10000):
        count = 0
        for _, _, files in os.walk(path):
            count += len(files)
            if count > limit:
                return count
        return count


            
    def load_txt(self):
        self.open_in_btn.setText("Open in Text Edit")
        try:
            with open(self.path, 'r', encoding='utf-8') as f:
                text = f.read()
        except UnicodeDecodeError:
            with open(self.path, 'r', encoding='latin1') as f:
                text = f.read()
 
               
        self.text_edit.setPlainText(text)
        self.text_edit.setReadOnly(True)
        self.media_layout.addWidget(self.text_edit)

            
    def load_image(self):
        self.label = QLabel(self)
        self.label.setMinimumSize(1200, 800)
        self.pixmap = QPixmap(self.path)

        
        self.original_pixmap = self.pixmap
        self.label.setAlignment(Qt.AlignCenter)
        self.label.setScaledContents(False)
        self.media_layout.addWidget(self.label)
        self.update_pixmap()
        self.open_in_btn.setText("Open in Preview")
        
    def load_gif(self):
        self.label = QLabel(self)
        self.movie = QMovie(self.path)
        self.label.setMovie(self.movie)
        self.label.setAlignment(Qt.AlignCenter)
        self.media_layout.addWidget(self.label)
        self.movie.start()
        

    def load_video(self):
        self.media_player = QMediaPlayer(None, QMediaPlayer.VideoSurface)
        video_widget = QVideoWidget()
        video_widget.setMinimumSize(1200, 800)
        
        video_container = QWidget()
        video_container.setStyleSheet("""
            background: black;
            border: none;
            border-radius: 0px;
        """)
        video_layout = QVBoxLayout(video_container)
        video_layout.setContentsMargins(0, 0, 0, 0)
        video_layout.addWidget(video_widget)
        
        self.media_player.setVideoOutput(video_widget)
        self.media_player.setMedia(QMediaContent(QUrl.fromLocalFile(self.path)))
        self.media_player.play()

        self.media_layout.addWidget(video_container)
        self.media_layout.setContentsMargins(0, 0, 0, 0)
        self.open_in_btn.setText("Open in QuickTime")
        
    def load_audio(self):        
        self.media_player = QMediaPlayer()
        self.media_player.setMedia(QMediaContent(QUrl.fromLocalFile(self.path)))
        self.media_player.play()

        self.music = QHBoxLayout()
        self.music.setSpacing(40)
        self.media_layout.addLayout(self.music)

        # cover art 
        cover_pixmap = extract_cover_pixmap(self.path)
        cover_label = QLabel()
        cover_label.setPixmap(cover_pixmap.scaled(600, 600, Qt.KeepAspectRatio, Qt.SmoothTransformation))
        cover_label.setAlignment(Qt.AlignCenter)
        self.music.addWidget(cover_label)

        # metadata of the song 
        audio = File(self.path)
        title = artist = album = "Unknown"
        if audio and audio.tags and isinstance(audio.tags, ID3):
            title = audio.tags.get('TIT2', TIT2(encoding=3, text='Unknown')).text[0]
            artist = audio.tags.get('TPE1', TPE1(encoding=3, text='Unknown')).text[0]
            album = audio.tags.get('TALB', TALB(encoding=3, text='Unknown')).text[0]

        info_layout = QVBoxLayout()

        title_label = QLabel(f"{title}")
        title_label.setStyleSheet("font-size: 34px; font-weight: bold")
        artist_label = QLabel(f"{artist}")
        artist_label.setStyleSheet("font-size: 28px;")
        album_label = QLabel(f"{album}")
        album_label.setStyleSheet("font-size: 28px;")
        info_layout.addStretch()
        for lbl in (title_label, artist_label, album_label):
            lbl.setAlignment(Qt.AlignLeft)
            info_layout.addWidget(lbl)

        info_layout.addStretch()
        self.music.addLayout(info_layout)

        self.open_in_btn.setText("Open in QuickTime")
        
    def load_pdf(self):
        self.open_in_btn.setText("Open in Preview")
        try:
            import fitz  # PyMuPDF
        except ImportError:
            label = QLabel("Aby wyświetlać PDF, zainstaluj PyMuPDF:\npip install PyMuPDF")
            label.setAlignment(Qt.AlignCenter)
            self.media_layout.addWidget(label)
            return

        try:
            doc = fitz.open(self.path)
            if doc.page_count > 0:
                # Create a container widget for the PDF view
                self.pdf_container = QWidget()
                pdf_layout = QVBoxLayout(self.pdf_container)
                pdf_layout.setContentsMargins(0, 0, 0, 0)
                
                # Create a scroll area for multi-page PDFs
                scroll_area = QScrollArea()
                scroll_area.setWidgetResizable(True)
                scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
                scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
                
                # Create a widget to hold all pages
                self.pdf_pages_widget = QWidget()
                self.pdf_pages_layout = QVBoxLayout(self.pdf_pages_widget)
                self.pdf_pages_layout.setAlignment(Qt.AlignCenter)
                self.pdf_pages_layout.setSpacing(20)
                
                # Load first page (or all pages if small)
                max_pages_to_load = 3 if doc.page_count > 1 else 1
                self.pdf_pixmaps = []
                
                for i in range(min(doc.page_count, max_pages_to_load)):
                    page = doc.load_page(i)
                    pix = page.get_pixmap(dpi=150)
                    img = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format_RGB888)
                    pixmap = QPixmap.fromImage(img)
                    self.pdf_pixmaps.append(pixmap)
                    
                    label = QLabel()
                    label.setPixmap(pixmap)
                    label.setAlignment(Qt.AlignCenter)
                    self.pdf_pages_layout.addWidget(label)
                
                # Add "Page X of Y" label if multi-page
                if doc.page_count > 1:
                    page_label = QLabel(f"{doc.page_count} Pages")
                    page_label.setAlignment(Qt.AlignCenter)
                    page_label.setStyleSheet("font-size: 23px;")
                    self.pdf_pages_layout.addWidget(page_label)
                
                scroll_area.setWidget(self.pdf_pages_widget)
                pdf_layout.addWidget(scroll_area)
                self.media_layout.addWidget(self.pdf_container)
                
                # Store original pixmaps for resizing
                self.original_pdf_pixmaps = self.pdf_pixmaps.copy()
                
                # Connect resize event
                self.pdf_container.resizeEvent = self.update_pdf_size
                
        except Exception as e:
            label = QLabel("Nie udało się załadować PDF:\n" + str(e))
            label.setAlignment(Qt.AlignCenter)
            self.media_layout.addWidget(label)

    def update_pdf_size(self, event):
        """Update PDF pixmaps maintaining aspect ratio"""
        if hasattr(self, 'pdf_pixmaps') and self.pdf_pixmaps:
            container_width = self.pdf_container.width() - 40  # Account for margins
            for i, (original_pixmap, label) in enumerate(zip(self.original_pdf_pixmaps, self.get_pdf_labels())):
                scaled_pixmap = original_pixmap.scaled(
                    container_width, 
                    self.pdf_container.height(), 
                    Qt.KeepAspectRatio, 
                    Qt.SmoothTransformation
                )
                label.setPixmap(scaled_pixmap)
                


    def get_pdf_labels(self):
        """Get all PDF page labels from the layout"""
        labels = []
        for i in range(self.pdf_pages_layout.count()):
            item = self.pdf_pages_layout.itemAt(i)
            if isinstance(item.widget(), QLabel) and item.widget().pixmap():
                labels.append(item.widget())
        return labels

        
    def update_pixmap(self):
        """Update image pixmap maintaining aspect ratio"""
        if hasattr(self, 'pixmap') and not self.pixmap.isNull():
            label_size = self.label.size()
            scaled_pixmap = self.original_pixmap.scaled(label_size, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            self.label.setPixmap(scaled_pixmap)
            


    def setMaskForRoundedCorners(self):
        """Set a mask for the window to have rounded corners. Otherwise its ugly."""
        if self.is_fullscreen:
            self.clearMask()  
            return
            
        radius = 17
        pixmap = QPixmap(self.size())
        pixmap.fill(Qt.transparent)

        painter = QPainter(pixmap)
        painter.setRenderHint(QPainter.Antialiasing)
        painter.setBrush(Qt.white)
        painter.setPen(Qt.NoPen)
        painter.drawRoundedRect(self.rect(), radius, radius)
        painter.end()

        mask = pixmap.createMaskFromColor(Qt.transparent, Qt.MaskInColor)
        self.setMask(mask)

    def toggle_fullscreen(self):
        if not self.is_fullscreen:
            self.original_geometry = self.frameGeometry()
            self.showFullScreen()
            self.fullscreen_btn.setIcon(QIcon("icons/media_preview/restore.png"))
            self.is_fullscreen = True
            if is_dark_mode_enabled(): 
                self.main.setStyleSheet("""
                    background: #2a2a28;
                    border: none;
                    border-radius: 0px;
                    color: #c8c8c8;
                """)
            else:
                self.main.setStyleSheet("""
                    background: #e0e0e0;
                    border: none;
                    border-radius: 0px;
                    color: #3c3c3c;
                """)
            self.clearMask()  # Usuń maskę
        else:
            self.showNormal()
            if self.original_geometry:
                self.setGeometry(self.original_geometry)
            self.fullscreen_btn.setIcon(QIcon("icons/media_preview/maximize.png"))
            self.is_fullscreen = False
            if is_dark_mode_enabled(): 
                self.main.setStyleSheet("""
                    background: #2a2a28;
                    border: 3px solid #393939;
                    border-radius: 12px;
                    color: #c8c8c8;
                """)
            else:
                self.main.setStyleSheet("""
                    background: #e0e0e0;
                    border: 3px solid rgba(209, 209, 209, 175);
                    border-radius: 12px;
                    color: #3c3c3c;
                """)
            self.setMaskForRoundedCorners()  # Przywróć maskę
        
        if hasattr(self, 'pixmap'):
            self.update_pixmap()

    def closeEvent(self, event):
        if hasattr(self, 'media_player') and self.media_player:
            self.media_player.stop()
        if hasattr(self, 'movie') and self.movie:
            self.movie.stop()
        super().closeEvent(event)

    def keyPressEvent(self, event):
        if event.key() == Qt.Key_Escape or event.key() == Qt.Key_Space:
            self.close()
            event.accept()  # Zawsze akceptuj te klawisze, aby nie propagować dalej
        else:
            super().keyPressEvent(event)