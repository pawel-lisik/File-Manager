# to do: create a helper method to get all the custom message boxes and simplify the code. But for now inm not touching it because the code works :)

import sys
import os
os.environ["QT_ACCESSIBILITY"] = "0"  # For faster startup, disabled accessibility features

import ctypes
import time
import json
from collections import defaultdict

# Basic imports for PyQt5
from PyQt5.QtWidgets import (
    QApplication, QWidget, QHBoxLayout, QVBoxLayout, QListView, QListWidget,
    QListWidgetItem, QFileSystemModel, QTreeView, QPushButton, QStackedWidget,
    QFileIconProvider, QSizePolicy, QToolButton, QButtonGroup, QLabel, 
    QLineEdit, QMenu, QMessageBox, QColumnView, QAction, QSpacerItem, 
    QStyledItemDelegate, QDialog, QStyle, QScrollArea, QHeaderView, QMenuBar  
)
from PyQt5.QtGui import QIcon, QKeySequence, QPixmap, QPainter, QColor, QBrush
from PyQt5.QtCore import Qt, QDir, QModelIndex, QSize, QFileInfo, QTimer, QObject, pyqtSignal, QThread, QRect, QSettings


# Lazy loading for modules that are not immediately needed, then the app starts faster!
LAZY_IMPORTS = {
    'psutil': None,
    'win32con': None,
    'win32api': None,
    'win32com': None,
    'shutil': None,
    'platform': None,
}

def lazy_import(name):
    if name in LAZY_IMPORTS and LAZY_IMPORTS[name] is None:
        if name == 'psutil':
            import psutil
            LAZY_IMPORTS[name] = psutil
        elif name == 'win32con':
            import win32con
            LAZY_IMPORTS[name] = win32con
        elif name == 'win32api':
            import win32api
            LAZY_IMPORTS[name] = win32api
        elif name == 'win32com':
            import win32com.client
            LAZY_IMPORTS[name] = win32com.client
        elif name == 'shutil':
            import shutil
            LAZY_IMPORTS[name] = shutil
        elif name == 'platform':
            import platform
            LAZY_IMPORTS[name] = platform            
    return LAZY_IMPORTS[name]



settings = QSettings(
    "HKEY_CURRENT_USER\\Software\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize",
    QSettings.NativeFormat
)
theme = settings.value("AppsUseLightTheme", True)

psutil = lazy_import('psutil')
win32com = lazy_import('win32com')



class file_manager(QWidget):
    def __init__(self, path=None):
        super().__init__()
        self.setWindowTitle(" ")
        self.setWindowIcon(QIcon("icon.ico"))
        self.setMinimumSize(1600, 1000)
        # Set the inital path
        self.initial_path = path if path else ""
        self.show_hidden_files = False
        self.model = CustomFileSystemModel()
        self.model.setReadOnly(False)
        self.model.setRootPath(QDir.rootPath())
        # Hide hidden files by default
        self.model.setNameFilters([".*", "~*"] if os.name != 'nt' else ["*.*"])
        self.model.setNameFilterDisables(False)
        
        # Set filter to show all files by default (control visibility via hidden flag)
        self.model.setFilter(QDir.AllDirs | QDir.Files | QDir.Drives | QDir.NoDotAndDotDot)
        self.update_hidden_files_filter()

        self.icon_provider = QFileIconProvider()

        self.current_path = ""
        self.current_tag = None 
        self.history = []
        self.history_index = -1


        # Predefined file/folder tags and their colors
        self.available_tags = {
            ' Red': "#ff5f5f",
            ' Orange': "#fba45b", 
            ' Yellow': "#f6cc67",
            ' Green': "#60cb68",
            ' Blue': "#33baef",
            ' Purple': "#d38adb",
            ' Gray': "#a4a4a7"
        }        
        self.tags_file = os.path.join(os.path.expanduser("~"), ".file_manager_tags.json")
        self.load_tags()

        self.init_ui()
        self.apply_theme()
        
        self.registry_monitor = RegistryMonitor()
        self.registry_monitor.connect_theme_changed(self.apply_theme)
        self.registry_monitor.start()
        
        self.clipboard = GlobalClipboard()
        self.copy_in_progress = False
        self.last_previewed_file = None        
        self.preview_window = None  # Track the preview window

        # Timer to check for drive changes every 5 seconds
        self.disk_check_timer = QTimer(self)
        self.disk_check_timer.timeout.connect(self.check_drives_changed)
        self.disk_check_timer.start(5000)
        self.last_drives = []  # Initialize as empty list
        self.check_drives_changed()  # Initial population of last_drives
        
        self._setup_drag_drop_configuration()
        self.scroll_positions = {}
        
        
    def init_ui(self):        
        self.setFocusPolicy(Qt.StrongFocus)
        self.setFocus()        
        main_layout = QVBoxLayout(self)
        self.setLayout(main_layout)
        main_layout.setSpacing(0)
        main_layout.setContentsMargins(0, 0, 0, 0)

        self._gallery_view_initialized = False
        self._column_view_initialized = False


        main_area = QHBoxLayout()
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setFixedWidth(370)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        scroll_area.setObjectName("sidebar2")

        scroll_content = QWidget()
        scroll_content.setObjectName("sidebar")  # styl jak wcześniej
        scroll_layout = QVBoxLayout()
        scroll_layout.setSpacing(6)
        scroll_layout.setContentsMargins(10, 10, 10, 10)
        scroll_content.setLayout(scroll_layout)

        scroll_area.setWidget(scroll_content)

        # iCloud
        header1 = QLabel("iCloud")
        header1.setObjectName("header")

        icloud_btn = QPushButton("iCloud Drive")
        icloud_btn.setStyleSheet("border: none; text-align: left; padding-left: 6px; font-family: 'Segoe UI', 'Helvetica Neue', sans-serif; font-size: 26px;")
        icloud_btn.setObjectName("tag_item")
        icloud_btn.setIcon(QIcon("icons\icloud.png"))
        icloud_btn.setIconSize(QSize(32, 32))
        icloud_btn.clicked.connect(self.open_icloud)

        # Favorites
        header2 = QLabel("Favorites")
        header2.setObjectName("header")

        self.sidebar = QListWidget()
        self.sidebar.setObjectName("sidebar")
        self.sidebar.setFocusPolicy(Qt.NoFocus)
        self.sidebar.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.sidebar.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)        
        self.sidebar.setSpacing(4)
        self.sidebar.setIconSize(QSize(32, 32))
        self.sidebar.setStyleSheet("QListWidget::item { padding: 6px; }")
        self.sidebar.setStyleSheet("background: transparent")
        self.sidebar.itemClicked.connect(self.on_sidebar_click)

        icon_map = {
            "Home": "icons\home.png",
            "Applications": "icons\apps.png",
            "Downloads": "icons\downloads.png",
            "Documents": "icons\documents.png",
            "Desktop": "icons\desktop.png",
            "Music": "icons\music.png",
            "Pictures": "icons\pictures.png",
            "Videos": "icons\videos.png",
        }

        self.favorites = {
            "Home": QDir.homePath(),
            "Applications": "C:\\MacOS\\Applications",
            "Downloads": os.path.join(QDir.homePath(), "Downloads"),
            "Documents": os.path.join(QDir.homePath(), "Documents"),
            "Desktop": os.path.join(QDir.homePath(), "Desktop"),
            "Music": os.path.join(QDir.homePath(), "Music"),
            "Pictures": os.path.join(QDir.homePath(), "Pictures"),
            "Videos": os.path.join(QDir.homePath(), "Videos"),
        }

        # Add a setting to track if Home is shown
        self.settings = QSettings("MacOS", "file_manager")
        self.show_home_in_sidebar = self.settings.value("show_home_in_sidebar", True, type=bool)

        # Modify the sidebar population to respect this setting
        self.update_sidebar()

        # Locations
        header3 = QLabel("Locations")
        header3.setObjectName("header")

        self.locations = QListWidget()
        self.locations.setObjectName("sidebar")
        self.locations.setFocusPolicy(Qt.NoFocus)
        self.locations.setSpacing(4)
        self.locations.itemClicked.connect(self.on_locations_click)
        self.locations.model().rowsInserted.connect(lambda *_: self.adjust_locations_height())
        self.locations.model().rowsRemoved.connect(lambda *_: self.adjust_locations_height())
        self.locations.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.locations.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.refresh_drives_list()

        # Tags
        header4 = QLabel("Tags")
        header4.setObjectName("header")

        self.sidebar_layout = QVBoxLayout()
        self.sidebar_layout.setContentsMargins(0, 0, 0, 0)
        self.sidebar_layout.setSpacing(4)

        tags_widget = QWidget()
        tags_widget.setStyleSheet("background: transparent; border: none; font-size: 26px; font-family: 'Segoe UI', 'Helvetica Neue', sans-serif;")
        tags_widget.setLayout(self.sidebar_layout)
        self.update_tags_sidebar()


        scroll_layout.addWidget(header1)
        scroll_layout.addWidget(icloud_btn)
        scroll_layout.addItem(QSpacerItem(40, 20, QSizePolicy.Expanding, QSizePolicy.Minimum))
        scroll_layout.addWidget(header2)
        scroll_layout.addWidget(self.sidebar)
        scroll_layout.addItem(QSpacerItem(40, 30, QSizePolicy.Expanding, QSizePolicy.Minimum))
        scroll_layout.addWidget(header3)
        scroll_layout.addWidget(self.locations)
        scroll_layout.addItem(QSpacerItem(40, 30, QSizePolicy.Expanding, QSizePolicy.Minimum))
        scroll_layout.addWidget(header4)
        scroll_layout.addWidget(tags_widget)
        scroll_layout.addStretch()


        main_area.addWidget(scroll_area)
                

        
        top = QWidget()
        top.setObjectName("top")
        top_bar = QHBoxLayout()
        top.setLayout(top_bar)
        self.path_bar = QHBoxLayout()
        self.path_bar.setSpacing(0)  


        self.back_button = QToolButton()
        self.back_button.setIcon(QIcon("icons\back.png"))
        self.back_button.setToolTip("Back")
        self.back_button.clicked.connect(self.go_back)

        self.forward_button = QToolButton()
        self.forward_button.setIcon(QIcon("icons\next.png"))
        self.forward_button.setToolTip("Forward")
        self.forward_button.clicked.connect(self.go_forward)

        top_bar.addWidget(self.back_button)
        top_bar.addWidget(self.forward_button)
        

        self.current_folder_label = QLabel()
        self.current_folder_label.setMinimumWidth(300)
        self.current_folder_label.setStyleSheet("font-weight: bold; font-size: 28px")
        top_bar.addWidget(self.current_folder_label)


        top_bar.addStretch()

        # View buttons
        self.view_buttons = QButtonGroup(self)
        self.icon_view_btn = QToolButton()
        self.icon_view_btn.setIcon(QIcon("icons\grid.png"))
        self.icon_view_btn.setCheckable(True)
        self.icon_view_btn.setToolTip("Icon view")

        self.list_view_btn = QToolButton()
        self.list_view_btn.setIcon(QIcon("icons\list.png"))
        self.list_view_btn.setCheckable(True)
        self.list_view_btn.setToolTip("List view")
        


        self.column_view_btn = QToolButton()
        self.column_view_btn.setIcon(QIcon("icons\column.png"))
        self.column_view_btn.setCheckable(True)
        self.column_view_btn.setToolTip("Column view")
        

        self.gallery_view_btn = QToolButton()
        self.gallery_view_btn.setIcon(QIcon("icons\gallery.png"))  # You'll need to add this icon
        self.gallery_view_btn.setCheckable(True)
        self.gallery_view_btn.setToolTip("Gallery view")


        self.view_buttons.addButton(self.icon_view_btn, 0)
        self.view_buttons.addButton(self.list_view_btn, 1)
        self.view_buttons.addButton(self.column_view_btn, 2)  
        self.view_buttons.addButton(self.gallery_view_btn, 3)  
        self.view_buttons.buttonClicked[int].connect(self.change_view)

        self.icon_view_btn.setChecked(True)

        top_bar.addWidget(self.icon_view_btn)
                
        self.line = QWidget()
        self.line.setFixedSize(2,36)
        self.line.setObjectName("line")
        top_bar.addWidget(self.line)
        
        top_bar.addWidget(self.list_view_btn)
        
        self.line2 = QWidget()
        self.line2.setFixedSize(2,36)
        self.line2.setObjectName("line")
        top_bar.addWidget(self.line2)

        top_bar.addWidget(self.column_view_btn)
        
        self.line3 = QWidget()
        self.line3.setFixedSize(2,36)
        self.line3.setObjectName("line")
        top_bar.addWidget(self.line3)        
        


        top_bar.addWidget(self.gallery_view_btn)
        
        
        
        top_bar.addItem(QSpacerItem(40, 20, QSizePolicy.Expanding, QSizePolicy.Minimum))

        self.sort_button = QToolButton()
        self.sort_button.setIcon(QIcon("icons\sort.png"))
        self.sort_button.setIconSize(QSize(64, 32))
        self.sort_button.setPopupMode(QToolButton.InstantPopup)
        self.sort_button.setToolTip("Sort")
        self.sort_button.setStyleSheet("QToolButton::menu-indicator { image: none; }")
        
        self.sort_actions = {}
        self.current_sort_key = "date"  # Default sort by date
        
        sort_menu = CustomMenu(parent=self)
        sort_options = {
            "name": ("Name", 0),
            "type": ("Kind", 2),
            "size": ("Size", 1),
            "date": ("Date", 3),
        }

        for key, (label, column) in sort_options.items():
            action = QAction(label, self)
            action.setCheckable(True)
            action.triggered.connect(lambda checked, k=key: self.sort_by(k))
            sort_menu.addAction(action)
            self.sort_actions[key] = action

        self.sort_actions["date"].setChecked(True)
        
      
        self.sort_button.setMenu(sort_menu)
        top_bar.addWidget(self.sort_button) 
        space = QLabel("")
        space.setFixedWidth(100)
        top_bar.addWidget(space)

        
        self.more_button = QToolButton()
        self.more_button.setIcon(QIcon("icons\more.png"))
        self.more_button.setIconSize(QSize(64, 32))
        self.more_button.setPopupMode(QToolButton.InstantPopup)
        self.more_button.setToolTip("More")
        self.more_button.setStyleSheet("QToolButton::menu-indicator { image: none; }")
        
        more_menu = CustomMenu(parent=self)
        
        new_window_action = more_menu.addAction("New file_manager Window")
        new_window_action.triggered.connect(self.open_new_window)
        
        open_trash_action = more_menu.addAction("Open Trash")
        open_trash_action.setShortcut('CTRL+T')
        open_trash_action.triggered.connect(self.show_recycle_bin)
        
        more_menu.addSeparator()
        
        show_hidden_action = more_menu.addAction("Show Hidden Files")
        show_hidden_action.setCheckable(True)
        show_hidden_action.setChecked(self.show_hidden_files)
        show_hidden_action.triggered.connect(self.toggle_hidden_files) 
        
        show_home_action = more_menu.addAction("Show Home")
        show_home_action.setCheckable(True)
        show_home_action.setChecked(self.show_home_in_sidebar)
        show_home_action.triggered.connect(self.toggle_home_in_sidebar)
        more_menu.addSeparator()

        self.preview_action = more_menu.addAction("Preview in Icon View")
        self.preview_action.setCheckable(True)
        self.preview_action.setChecked(False)  # Default unchecked (show icons)
        self.preview_action.setShortcut('CTRL+P')
        self.preview_action.triggered.connect(self.toggle_preview_mode)

        self.more_button.setMenu(more_menu)
        top_bar.addWidget(self.more_button)         

    
        self.search_bar = QLineEdit()
        self.search_bar.setPlaceholderText("Search")
        self.search_bar.setClearButtonEnabled(True)
        self.search_bar.setFixedWidth(300)
        
        self.search_bar.textChanged.connect(self.search_files)
        self.search_bar.hide()
        
        top_bar.addStretch()

        search_btn = QToolButton()
        search_btn.setIcon(QIcon("icons\search.png"))
        search_btn.clicked.connect(self.toggle_search_bar)
        
        go_to_folder_btn = QToolButton()
        go_to_folder_btn.setIcon(QIcon("icons\go.png"))
        go_to_folder_btn.clicked.connect(self.show_go_to_folder_dialog)  
        go_to_folder_btn.setToolTip("Go to Folder...")

        
        top_bar.addWidget(go_to_folder_btn)
        top_bar.addWidget(search_btn) 
        top_bar.addWidget(self.search_bar)


        main_layout.addWidget(top)
        main_layout.addLayout(main_area)

        self.stack = QStackedWidget()
        
        window = QWidget()
        window.setObjectName("window")
        window_layout = QVBoxLayout()
        window_layout.setContentsMargins(0, 0, 0, 0)
        window_layout.setSpacing(0)
        window.setLayout(window_layout)
        
        
        self.trash_bar = QWidget()
        self.trash_bar.setObjectName("trash")
        self.trash_bar_layout = QHBoxLayout()
        self.trash_bar_layout.setContentsMargins(10, 10, 10, 10)
        self.trash_bar.setLayout(self.trash_bar_layout)
        

        self.empty_trash_btn = QPushButton("Empty")
        self.empty_trash_btn.clicked.connect(self.empty_trash)
        
        self.restore_btn = QPushButton("Restore")
        self.restore_btn.clicked.connect(self.restore_from_recycle_bin)
        
        self.trash_bar_layout.addWidget(self.restore_btn)
        self.trash_bar_layout.addStretch()
        self.trash_bar_layout.addWidget(self.empty_trash_btn)
        
        self.trash_bar.hide()
        
        self.icon_view_container = QWidget()
        self.icon_view_layout = QVBoxLayout(self.icon_view_container)
        self.icon_view_layout.addWidget(self.stack)
        self.icon_view_layout.setContentsMargins(10, 10, 0, 0)  # lewy, góra, prawy, dół
        
        
        window_layout.addWidget(self.trash_bar)
        window_layout.addWidget(self.icon_view_container)
        
        main_area.addWidget(window)
        
        
        bottom = QWidget()
        bottom.setObjectName("bottom")
        bottom_bar = QHBoxLayout()
        bottom_bar.setContentsMargins(7, 7, 7, 7)
        bottom.setLayout(bottom_bar)
        bottom_bar.addLayout(self.path_bar)
        bottom_bar.addStretch()

        main_layout.addWidget(bottom)

        # Icon view
        self.icon_view = QListView()
        self.icon_view.setItemDelegate(GalleryItemDelegate(self.icon_view)) 
        self.icon_view.setViewMode(QListView.IconMode)
        self.icon_view.setIconSize(QSize(128, 128))
        self.icon_view.setResizeMode(QListView.Adjust)
        self.icon_view.setSpacing(12)
        self.icon_view.doubleClicked.connect(self.on_item_double_click)
        self.icon_view.setModel(self.model)
        self.icon_view.setDragEnabled(True)
        self.icon_view.setAcceptDrops(True)
        self.icon_view.setDropIndicatorShown(True)
        self.icon_view.setDragDropMode(QListView.InternalMove)
        
        self.icon_view.setFlow(QListView.LeftToRight) 
        self.icon_view.setWrapping(True)              
        self.icon_view.setGridSize(QSize(150, 180))

        # List view
        self.list_view = QTreeView()
        self.list_view.setModel(self.model)
        self.list_view.setRootIsDecorated(False)
        self.list_view.setItemsExpandable(False)
        self.list_view.setAnimated(True)
        self.list_view.setHeaderHidden(False)
        self.list_view.setSortingEnabled(True)
        self.list_view.doubleClicked.connect(self.on_item_double_click)
        self.list_view.setDragEnabled(True)
        self.list_view.setAcceptDrops(True)
        self.list_view.setDropIndicatorShown(True)
        self.list_view.setDragDropMode(QTreeView.InternalMove)
        self.list_view.setAlternatingRowColors(True)

        header = self.list_view.header()
        header.setFixedHeight(50)

        # Set resize modes for columns
        header.setSectionResizeMode(0, QHeaderView.Stretch)  # Name column takes remaining space
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)  # Size column
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)  # Type column
        header.setSectionResizeMode(3, QHeaderView.ResizeToContents)  # Date modified column

        # Optionally set minimum widths for other columns
        header.setMinimumSectionSize(80)  # Minimum width for all columns
        header.setStretchLastSection(False)  # Don't stretch the last section
        
        

        # Column view
        self.column_view = QColumnView()
        self.column_view.setAlternatingRowColors(True)
        self.column_view.setModel(self.model)
        self.column_view.doubleClicked.connect(self.on_item_double_click)
        self.column_view.setResizeGripsVisible(False)
        self.column_view.setColumnWidths([320] * self.model.columnCount())
        self.column_view.setStyleSheet('''
            QColumnView QListView::item:selected {
                corner-radius: 0px;"
            }
        ''')
        
        self.column_view.setItemDelegate(ColumnViewDelegate(self.column_view))
       
        # For icon view
        self.icon_view.setSelectionMode(QListView.ExtendedSelection)  # Allows multiple selection with Ctrl/Shift

        # For list view
        self.list_view.setSelectionMode(QTreeView.ExtendedSelection)

        # For column view
        self.column_view.setSelectionMode(QTreeView.ExtendedSelection)        
        
        
        
        # Replace the old gallery view with the new one
        self.gallery_view = GalleryView()
        self.gallery_view.preview_area.setMinimumHeight(400)  # Adjust as needed
        
        # Connect thumbnail selection to metadata update
        self.gallery_view.thumbnail_selected = self.update_metadata

        self.stack.setStyleSheet("border: none")

        
        self.stack.addWidget(self.icon_view) 
        self.stack.addWidget(self.list_view)
        self.stack.addWidget(self.column_view)
        self.stack.addWidget(self.gallery_view)
        self.search_active = False
        self.set_directory(self.current_path)
        
        
        
        if self.initial_path:
            # Check if the path is a directory
            if os.path.isdir(self.initial_path):
                self.set_directory(self.initial_path)
            else:
                # If it's a file, open its parent directory
                parent_dir = os.path.dirname(self.initial_path)
                if os.path.isdir(parent_dir):
                    self.set_directory(parent_dir)
        else:
            # Default view if no path is provided
            self.show_this_pc()      
        
        self.model.sort(3, Qt.DescendingOrder)
        
        if sys.platform == 'win32':
            myappid = 'mycompany.myproduct.subproduct.version'  # arbitrary string
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)

        
        self.icon_view.setContextMenuPolicy(Qt.CustomContextMenu)
        self.icon_view.customContextMenuRequested.connect(self.show_context_menu)
        self.list_view.setContextMenuPolicy(Qt.CustomContextMenu)
        self.list_view.customContextMenuRequested.connect(self.show_context_menu)
        self.sidebar.setContextMenuPolicy(Qt.CustomContextMenu)
        self.sidebar.customContextMenuRequested.connect(self.show_sidebar_context_menu)        
        self.locations.setContextMenuPolicy(Qt.CustomContextMenu)
        self.locations.customContextMenuRequested.connect(self.show_sidebar_context_menu)
        
        # rename
        self.rename_delegate = RenameDelegate(self)
        self.icon_view.setItemDelegate(self.rename_delegate)
        self.list_view.setItemDelegateForColumn(0, self.rename_delegate)
        self.column_view.setItemDelegateForColumn(0, self.rename_delegate)        
        # drag & drop

        self.icon_view.installEventFilter(self)


        # Key shortcuts
        self.copy_action = QAction("Copy", self)
        self.copy_action.setShortcut(QKeySequence.Copy)
        self.copy_action.triggered.connect(self.copy_selected)
        self.addAction(self.copy_action)

        self.cut_action = QAction("Cut", self)
        self.cut_action.setShortcut(QKeySequence.Cut)
        self.cut_action.triggered.connect(self.cut_selected)
        self.addAction(self.cut_action)

        self.paste_action = QAction("Paste", self)
        self.paste_action.setShortcut(QKeySequence.Paste)
        self.paste_action.triggered.connect(self.paste_files)
        self.addAction(self.paste_action)
        
        self.rename_enter_action = QAction(self)
        self.rename_enter_action.setShortcut(Qt.Key_Return)
        self.rename_enter_action.setShortcutContext(Qt.WidgetShortcut)
        self.rename_enter_action.triggered.connect(self.handle_enter_key)
        self.addAction(self.rename_enter_action)
        
                        
        self.quick_look = QAction("Quick Look", self)
        self.quick_look.setShortcut(Qt.Key_Space)
        self.quick_look.setShortcutContext(Qt.WidgetWithChildrenShortcut)  # Changed to ApplicationShortcut
        self.quick_look.triggered.connect(self.handle_space_key)
        self.addAction(self.quick_look)
        
        
        self.delete_action = QAction(self)
        self.delete_action.setShortcut(Qt.Key_Backspace)
        self.delete_action.setShortcutContext(Qt.WidgetShortcut)
        self.delete_action.triggered.connect(self.move_to_trash_selected)
        self.addAction(self.delete_action)
        
        self.go_to_folder_action = QAction("Go to Folder", self)
        self.go_to_folder_action.setShortcut("Ctrl+G")
        self.go_to_folder_action.triggered.connect(self.show_go_to_folder_dialog)
        self.addAction(self.go_to_folder_action)

        self.file_info_action = QAction("Get Info", self)
        self.file_info_action.setShortcut("Ctrl+I")
        self.file_info_action.triggered.connect(lambda _: self.show_file_or_drive_info())
        self.addAction(self.file_info_action)



        self.icon_view.setEditTriggers(QListView.NoEditTriggers)
        self.list_view.setEditTriggers(QTreeView.NoEditTriggers)
        self.column_view.setEditTriggers(QColumnView.NoEditTriggers) 




        # Menu bar
        self.menu_bar = QMenuBar()
        
        main_layout = self.layout()  
        main_layout.insertWidget(2, self.menu_bar)
        self.menu_bar.setFixedHeight(1)          
        self.menu_pliki = self.menu_bar.addMenu("File")
        self.menu_zaznacz = self.menu_bar.addMenu("Edit")
        self.menu_polecenia = self.menu_bar.addMenu("View")
        self.menu_siec = self.menu_bar.addMenu("Go")
        self.menu_widok = self.menu_bar.addMenu("Window")
        self.menu_konfiguracja = self.menu_bar.addMenu("Help")


        exit_action = QAction("Wyjście", self)
        exit_action.triggered.connect(self.close)
        self.menu_pliki.addAction(exit_action)
        
    def eventFilter(self, source, event):
        from PyQt5.QtCore import QEvent
        if source == self.icon_view:
            if event.type() == QEvent.KeyPress and event.key() in (Qt.Key_Return, Qt.Key_Enter):
                indexes = self.icon_view.selectedIndexes()
                if indexes:
                    self.icon_view.edit(indexes[0])
                    return True
                return False
            elif event.type() == QEvent.MouseButtonPress:
                delegate = self.icon_view.itemDelegate()
                if hasattr(delegate, 'current_editor') and delegate.current_editor:
                    # Use icon_view's viewport coordinates
                    pos_in_viewport = self.icon_view.viewport().mapFromGlobal(event.globalPos())
                    if not delegate.current_editor.geometry().contains(pos_in_viewport):
                        delegate.current_editor.clearFocus()
                        return True
                return False
        
        return super().eventFilter(source, event)
    


    def toggle_preview_mode(self):
        """Toggle between icon view and thumbnail preview mode"""
        if self.stack.currentIndex() == 0:  # Only affects icon view
            if self.preview_action.isChecked():
                # Enable preview mode
                self.icon_view.setItemDelegate(GalleryItemDelegate(self))
                self.icon_view.setIconSize(QSize(128, 128))
                self.icon_view.setGridSize(QSize(150, 180)) 
                self.icon_view.setSpacing(10)
            else:
                # Disable preview mode
                self.icon_view.setItemDelegate(self.rename_delegate)
                self.icon_view.setIconSize(QSize(128, 128))
                self.icon_view.setGridSize(QSize(150, 180))
            
            # Refresh the view
            if self.current_path and self.current_path not in ["This PC", "Trash"]:
                self.model.setRootPath(self.current_path)


    def ensure_column_view_initialized(self):
        if not hasattr(self, 'column_view') or not self.column_view:
            self.column_view = QColumnView()
            self.column_view.setModel(self.model)
            self.column_view.doubleClicked.connect(self.on_item_double_click)
            self.column_view.setAlternatingRowColors(True)
            self.column_view.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
            self.column_view.setResizeGripsVisible(False)
            self.column_view.setColumnWidths([320] * self.model.columnCount())
            self.column_view.setStyleSheet('''
                QColumnView QListView::item:selected {
                    corner-radius: 0px;
                }
            ''')
            self.column_view.setItemDelegate(ColumnViewDelegate(self.column_view))
            self.stack.addWidget(self.column_view)

    def _setup_drag_drop_configuration(self):
        """Configure custom drag & drop handling for all views"""
        # Main window - enable drops
        self.setAcceptDrops(True)
        from PyQt5.QtWidgets import QAbstractItemView
        # Common configuration for all views
        for view in [self.icon_view, self.list_view, self.column_view]:
            # Enable dragging from all views
            view.setDragEnabled(True)
            
            # Use our custom drop handling
            view.viewport().setAcceptDrops(True)
            view.setAcceptDrops(True)
            view.setDropIndicatorShown(True)
            
            # Set drag drop mode to enable both drag and drop
            view.setDragDropMode(QAbstractItemView.DragDrop)
            
            # Connect signals to handle drops properly
            view.dragEnterEvent = lambda e, v=view: self._handle_drag_enter_event(e, v)
            view.dragMoveEvent = lambda e, v=view: self._handle_drag_move_event(e, v)
            view.dropEvent = lambda e, v=view: self._handle_drop_event(e, v)
            
        # Special case for icon view to ensure proper item positioning
        self.icon_view.setMovement(QListView.Snap)  # Makes items snap to grid during drag
        self.icon_view.setDefaultDropAction(Qt.MoveAction)

    def _handle_drag_enter_event(self, event, view):
        """Handle drag entering any view"""
        if event.mimeData().hasUrls():
            if event.keyboardModifiers() & Qt.ShiftModifier:
                event.setDropAction(Qt.CopyAction)
            else:
                event.setDropAction(Qt.MoveAction)
            event.accept()
        else:
            event.ignore()

    def _handle_drag_move_event(self, event, view):
        """Handle drag moving over any view"""
        if event.mimeData().hasUrls():
            if event.keyboardModifiers() & Qt.ShiftModifier:
                event.setDropAction(Qt.CopyAction)
            else:
                event.setDropAction(Qt.MoveAction)
            event.accept()
        else:
            event.ignore()

    def _handle_drop_event(self, event, view):
        """Handle drop on any view"""
        # Determine target directory based on drop position
        pos = view.mapFromParent(event.pos())
        index = view.indexAt(pos)
        
        if index.isValid():
            item_path = self.model.filePath(index)
            dest_dir = item_path if os.path.isdir(item_path) else os.path.dirname(item_path)
        else:
            dest_dir = self.current_path
        
        # Check if it's copy or move
        is_copy = event.keyboardModifiers() & Qt.ShiftModifier
        operation = 'copy' if is_copy else 'move'
        
        src_paths = [url.toLocalFile() for url in event.mimeData().urls()]
        
        # Perform the file operations
        self._perform_file_operations(src_paths, dest_dir, operation)
        event.accept()

    def _perform_file_operations(self, src_paths, dest_dir, operation):
        import shutil
        from PyQt5.QtMultimedia import QSoundEffect
        from PyQt5.QtCore import QUrl
        """Wspólna metoda wykonująca operacje na plikach"""

        # Special treatment for Trash xD and my PC 
        if os.path.normcase(os.path.basename(dest_dir)) == "trash":
            operation = 'move2trash'
        elif "This PC" in os.path.normcase(dest_dir):
            return
        elif "tag:" in os.path.normcase(dest_dir):
            return

        # Automatically switch to copy if moving across different drives
        if operation == 'move':
            dest_drive = os.path.splitdrive(dest_dir)[0].lower()
            for src_path in src_paths:
                src_drive = os.path.splitdrive(src_path)[0].lower()
                if src_drive != dest_drive:
                    operation = 'copy'
                    break

        # Check for conflicts only if the operation is not 'move2trash'
        if operation != 'move2trash':
            conflicts = []
            for src_path in src_paths:
                if not os.path.exists(src_path):
                    continue

                base_name = os.path.basename(src_path)
                dest_path = os.path.join(dest_dir, base_name)

                if os.path.exists(dest_path):
                    if os.path.normcase(os.path.abspath(src_path)) == os.path.normcase(os.path.abspath(dest_path)):
                        continue
                    conflicts.append((src_path, dest_path))

            if conflicts:
                proceed, remaining_files = self._handle_copy_conflicts(conflicts)
                if not proceed or not remaining_files:
                    return
                # Filter src_paths to only include files that are in remaining_files
                src_paths = [f for f in src_paths if f in remaining_files]

        for i, src_path in enumerate(src_paths, 1):
            if not os.path.exists(src_path):
                continue

            base_name = os.path.basename(src_path)
            dest_path = os.path.join(dest_dir, base_name)

            if operation == 'move2trash':
                try:
                    norm_path = os.path.normpath(src_path)
                    drive_letter = os.path.splitdrive(norm_path)[0].upper()

                    if drive_letter == "C:":
                        from send2trash import send2trash
                        send2trash(norm_path)
                    else:
                        msg = self.create_message_box()
                        msg.setIcon(QMessageBox.Warning)
                        msg.setWindowTitle("Permanent deletion")
                          
                        msg.setText(
                            f"The file:\n{os.path.basename(norm_path)}\n"
                            f"is not on the C: drive and will be permanently deleted.\n\n"
                            f"Do you want to continue?"
                        )
                        yes_btn = msg.addButton("Yes", QMessageBox.YesRole)
                        no_btn = msg.addButton("No", QMessageBox.NoRole)

                        for btn in msg.buttons():
                            btn.setMinimumWidth(120)

                        msg.exec_()

                        if msg.clickedButton() != yes_btn:
                            continue  # Pomijamy usunięcie

                        if os.path.isdir(norm_path):
                            shutil.rmtree(norm_path)
                        else:
                            os.remove(norm_path)
                            
                    self.sound = QSoundEffect()
                    self.sound.setSource(QUrl.fromLocalFile("icons\remove.wav"))
                    self.sound.setVolume(1)
                    self.sound.play()

                except Exception as e:                   
                    msg = self.create_message_box()
                    msg.setIcon(QMessageBox.Warning)
                    msg.setWindowTitle("Error")
                      
                    msg.setText(f"Could not remove file:\n{e}")
                    ok_button = msg.addButton("OK", QMessageBox.AcceptRole)
                    ok_button.setMinimumWidth(120)
                     
                    msg.exec_()   


                continue

            if operation == 'move':
                if os.path.normcase(os.path.abspath(src_path)) == os.path.normcase(os.path.abspath(dest_path)):
                    continue
                try:
                    shutil.move(src_path, dest_path)
                except Exception as e:
                    print(f"Operacja przerwana: {e}")
                    return
            else:
                progress_dialog = CopyProgressDialog(
                    base_name,
                    QFileIconProvider().icon(QFileInfo(src_path)).pixmap(64, 64),
                    self
                )
                progress_dialog.setWindowTitle(" ")

                worker = CopyWorker(src_path, dest_path, operation)
                progress_dialog.set_worker(worker)

                worker.signals.progress.connect(progress_dialog.progress.setValue)
                worker.signals.finished.connect(progress_dialog.accept)
                worker.signals.error.connect(self.show_error_dialog)

                worker.start()

                if progress_dialog.exec_() == QDialog.Rejected:
                    try:
                        if os.path.exists(dest_path):
                            if os.path.isdir(dest_path):
                                shutil.rmtree(dest_path)
                            else:
                                os.remove(dest_path)
                    except Exception as e:                                        
                        msg = self.create_message_box()
                        msg.setIcon(QMessageBox.Warning)
                        msg.setWindowTitle("Error")
                          
                        msg.setText(f"Cleanup failed:\n{e}")
                        ok_button = msg.addButton("OK", QMessageBox.AcceptRole)
                        ok_button.setMinimumWidth(120)
                         
                        msg.exec_()   
                        
                    break

        if operation == 'copy':                           
            self.sound = QSoundEffect()
            self.sound.setSource(QUrl.fromLocalFile("icons\complete.wav"))
            self.sound.setVolume(0.8)
            self.sound.play()
            
        


    def _handle_copy_conflicts(self, conflicts):
        """Display a dialog to handle file copy conflicts and return the user's choice."""
        if len(conflicts) == 1:
            msg = f"A file named '{os.path.basename(conflicts[0][1])}' already exists in this location."
        else:
            msg = f"{len(conflicts)} files already exist in this location."

        dialog = QDialog(self)
        dialog.setWindowIcon(QIcon("icons\file_manager.ico"))
        dialog.setWindowTitle("Confirm File Replace")
        dialog.resize(600, 150)
        
        set_dark_title_bar(dialog, not self.theme)
        
        if self.theme:
            dialog.setStyleSheet(open("style.qss").read())
        else:
            dialog.setStyleSheet(open("style_dark.qss").read())

        layout = QVBoxLayout(dialog)

        msg_label = QLabel(msg)
        msg_label.setWordWrap(True)
        layout.addWidget(msg_label)

        details_button = QPushButton("Details")
        details_button.setCheckable(True)
        details_button.setChecked(False)
        details_button.setFixedWidth(120)

        details_area = QLabel()
        details_area.setVisible(False)
        details_area.setWordWrap(True)

        details_text = "\n".join(
            f"• {os.path.basename(src)} (Size: {format_size(os.path.getsize(src)) if os.path.exists(src) else 'Unknown'})"
            for src, dest in conflicts
        )
        details_area.setText(details_text)
        from PyQt5.QtWidgets import QDialogButtonBox
        button_box = QDialogButtonBox()
        skip_button = button_box.addButton("Skip", QDialogButtonBox.RejectRole)
        cancel_button = button_box.addButton("Cancel", QDialogButtonBox.DestructiveRole)
        replace_button = button_box.addButton("Replace", QDialogButtonBox.AcceptRole)

        for btn in [replace_button, skip_button, cancel_button]:
            btn.setMinimumWidth(120)

        layout.addWidget(details_button)
        layout.addWidget(details_area)
        layout.addWidget(button_box)

        details_button.toggled.connect(details_area.setVisible)

        result = {'action': None}

        def handle_response(button):
            if button == replace_button:
                result['action'] = 'replace'
                dialog.accept()
            elif button == skip_button:
                result['action'] = 'skip'
                dialog.accept()
            else:
                result['action'] = 'cancel'
                dialog.reject()

        button_box.clicked.connect(handle_response)

        # Use QTimer to ensure the title bar is set after the dialog is shown
        QTimer.singleShot(50, lambda: set_dark_title_bar(dialog, not self.theme))
        
        if dialog.exec_() != QDialog.Accepted or result['action'] == 'cancel':
            return False, []

        if result['action'] == 'skip':
            conflict_srcs = {src for src, dest in conflicts}
            remaining_files = [f for f in self.clipboard.files if f not in conflict_srcs]
            return bool(remaining_files), remaining_files

        if result['action'] == 'replace':
            return True, [src for src, _ in conflicts]




    def toggle_home_in_sidebar(self):
        """Toggle showing Home folder in sidebar"""
        self.show_home_in_sidebar = not self.show_home_in_sidebar
        self.settings.setValue("show_home_in_sidebar", self.show_home_in_sidebar)
        self.update_sidebar()
        
    def update_sidebar(self):
        """Update sidebar items based on settings"""
        self.sidebar.clear()
        icon_map = {
            "Home": "icons\home.png",
            "Applications": "icons\apps.png",
            "Downloads": "icons\downloads.png",
            "Documents": "icons\documents.png",
            "Desktop": "icons\desktop.png",
            "Music": "icons\music.png",
            "Pictures": "icons\pictures.png",
            "Videos": "icons\videos.png",
        }

        for name, path in self.favorites.items():
            if name == "Home" and not self.show_home_in_sidebar:
                continue
            icon = QIcon(icon_map.get(name, "icons\folder.png"))
            item = QListWidgetItem(icon, name)
            item.setData(Qt.UserRole, path)
            self.sidebar.addItem(item)
        
        row_height = 44
        count = self.sidebar.count()
        spacing = self.sidebar.spacing() * (count - 1)
        total_height = (row_height * count) + spacing + 2 * self.sidebar.frameWidth() 
        self.sidebar.setFixedHeight(total_height)


        

    def update_hidden_files_filter(self):
        """Update the model's filter based on hidden files setting"""
        filters = self.model.filter()
        if self.show_hidden_files:
            filters |= QDir.Hidden  # Show hidden files
        else:
            filters &= ~QDir.Hidden  # Hide hidden files
        self.model.setFilter(filters)        
      
    def apply_theme(self):
        settings = QSettings(
            "HKEY_CURRENT_USER\\Software\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize",
            QSettings.NativeFormat
        )
        self.theme = settings.value("AppsUseLightTheme", True)        

        if self.theme:
            self.setStyleSheet(open("style.qss").read())
            set_dark_title_bar(self, False)
        else:
            self.setStyleSheet(open("style_dark.qss").read())
            set_dark_title_bar(self, True)
        
        for widget in QApplication.topLevelWidgets():
            if isinstance(widget, (QMessageBox, ThemedMessageBox)):
                set_dark_title_bar(widget, not self.theme)

    def show_message_box(self, *args, **kwargs):
        """Helper method to create message boxes with proper theme"""
        msg = QMessageBox(*args, **kwargs)
         
        return msg  # Return the message box object, not the exec result
        
    def create_message_box(self, parent=None):
        """Creates a message box with the current theme"""
        msg = ThemedMessageBox(parent or self)
        set_dark_title_bar(msg, not self.theme)
        return msg

        
    def showEvent(self, event):
        """Ensure the widget gets focus when shown"""
        super().showEvent(event)
        self.setFocus()
        self.activateWindow()  # Add this line to ensure the window is active

    def check_drives_changed(self):
        """Checks if the list of drives has changed and refreshes if needed"""
        current_drives = self.list_drives()
        
        # Compare just the paths (not the full dictionaries)
        current_paths = {drive['path'] for drive in current_drives}
        last_paths = {drive['path'] for drive in self.last_drives}
        
        if current_paths != last_paths:
            self.refresh_drives_list()
            
            if self.current_path == "This PC":
                self.show_this_pc()

            self.last_drives = current_drives
        
    def adjust_locations_height(self):
        item_height = 40 + self.locations.spacing()  # 28 to wysokość z CSS, spacing = 4
        count = self.locations.count()
        total_height = count * item_height  # +2 na ramkę/padding
        self.locations.setFixedHeight(total_height)

    def check_existing_files(self, src_paths, dest_dir):
        """Check if any files in src_paths already exist in dest_dir"""
        conflicts = []
        for src_path in src_paths:
            base_name = os.path.basename(src_path)
            dest_path = os.path.join(dest_dir, base_name)
            if os.path.exists(dest_path):
                conflicts.append((src_path, dest_path))
        return conflicts



    def open_icloud(self):
        os.startfile("iCloud.lnk")
        
        
    def show_go_to_folder_dialog(self):
        """Show dialog for navigating to specific folder with theme support"""
        dialog = QDialog(self)
        dialog.setWindowTitle("Go to the Folder:")
        dialog.setWindowIcon(QIcon("icons/file_manager.ico"))
        dialog.setFixedSize(800, 120)
        
        # Apply current theme
        set_dark_title_bar(dialog, not self.theme)
        dialog.setStyleSheet(open("style.qss").read() 
                           if self.theme else 
                           open("style_dark.qss").read())

        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(15, 15, 15, 15)
        layout.setSpacing(15)
        # Path input
        path_input = QLineEdit()
        path_input.setPlaceholderText("Enter or paste folder path here")
        layout.addWidget(path_input)
        # Button box
        button_box = QHBoxLayout()
        button_box.addStretch()        
        ok_button = QPushButton("Go")
        ok_button.setStyleSheet("""
            QPushButton {
                background-color: #0263e4; 
                border: 1px solid #0b4ca5; 
                border-radius: 8px; 
                color: white;
                padding: 6px;
            }
            QPushButton:pressed {
                background-color: #0385ff;
            }
        """)
        ok_button.setFixedWidth(100)        
        button_box.addWidget(ok_button)
        layout.addLayout(button_box)

        def on_go_clicked():
            path = path_input.text().strip()
            if os.path.isdir(path):
                dialog.accept()
                self.set_directory(path)
            else:
                # Use our themed message box
                msg = self.create_message_box()
                msg.setIcon(QMessageBox.Warning)
                msg.setWindowTitle("Invalid Path")
                msg.setWindowIcon(QIcon("icons/file_manager.ico"))
                msg.setText("The specified path does not exist.")
                
                # Create styled OK button
                ok_btn = msg.addButton("OK", QMessageBox.AcceptRole)
                ok_btn.setMinimumWidth(120)                
                set_dark_title_bar(msg, not self.theme)
                msg.exec_()

        ok_button.clicked.connect(on_go_clicked)
        
        # Ensure theme is applied when dialog is shown
        dialog.showEvent = lambda e: set_dark_title_bar(dialog, not self.theme)
        
        dialog.exec_()

        

    def move_to_trash_selected(self):        
        """Moves selected files to trash or permanently deletes them if on non-C drive."""
        view = self.stack.currentWidget()
        indexes = view.selectedIndexes()

        # Handle multiple columns in list view
        if isinstance(view, QTreeView):
            rows = set()
            unique_indexes = []
            for index in indexes:
                if index.row() not in rows:
                    rows.add(index.row())
                    unique_indexes.append(index)
            indexes = unique_indexes

        if not indexes:
            return

        # Get unique paths
        paths = list({self.model.filePath(index) for index in indexes})

        # Confirmation dialog
        if len(paths) == 1:
            question = f"Are you sure you want to delete \n{os.path.basename(paths[0])}?"
        else:
            question = f"Are you sure you want to delete {len(paths)} items?"

        msg = self.create_message_box()
        msg.setIcon(QMessageBox.Question)
        msg.setWindowTitle("Confirm")
        msg.setWindowIcon(QIcon("icons/file_manager.ico"))
        msg.setText(question)

        yes_button = msg.addButton("Yes", QMessageBox.YesRole)
        no_button = msg.addButton("No", QMessageBox.NoRole)

        for btn in msg.buttons():
            btn.setMinimumWidth(120)

         

        msg.exec_()

        if msg.clickedButton() == yes_button:
            for path in paths:
                try:
                    norm_path = os.path.normpath(path)
                    drive_letter = os.path.splitdrive(norm_path)[0].upper()

                    if drive_letter == "C:":
                        from send2trash import send2trash
                        send2trash(norm_path)
                    else:
                        if os.path.isdir(norm_path):
                            import shutil
                            shutil.rmtree(norm_path)
                        else:
                            os.remove(norm_path)
                    
                    from PyQt5.QtMultimedia import QSoundEffect
                    from PyQt5.QtCore import QUrl
                    self.sound = QSoundEffect()
                    self.sound.setSource(QUrl.fromLocalFile("sounds/remove.wav"))
                    self.sound.setVolume(1)
                    self.sound.play()

                except Exception as e:
                    msg = self.create_message_box()
                    msg.setIcon(QMessageBox.Warning)
                    msg.setWindowTitle("Error")
                    msg.setWindowIcon(QIcon("icons/file_manager.ico"))
                    msg.setText(f"Could not delete: {path}\n{str(e)}")
                    button = msg.addButton("OK", QMessageBox.AcceptRole)
                    button.setMinimumWidth(120)
                     
                    msg.exec_()
        


    def handle_enter_key(self):
        """Handle Enter key - edit name or open"""
        if self.current_path == "This PC":
            view = self.stack.currentWidget()
            indexes = view.selectedIndexes()
            if indexes:
                drive_path = indexes[0].data(Qt.UserRole)
                if drive_path and os.path.exists(drive_path):
                    self.switch_to_normal_view(drive_path)
            return

        current_index = self.stack.currentIndex()
        if current_index == 3:
            return

        view = self.stack.currentWidget()
        indexes = view.selectedIndexes()

        # Pobierz tylko indeks z kolumny 0
        index = next((i for i in indexes if i.column() == 0), None)

        if index:
            file_path = self.model.filePath(index) if hasattr(index, 'model') and index.model() == self.model else None

            if file_path:
                if isinstance(view, QTreeView):
                    view.edit(index)
                else:
                    self.rename_item(file_path)


    def handle_space_key(self):
        """Handle space key press - show/hide media_preview preview"""
        # Only handle if this window is active
        if not self.isActiveWindow():
            return
            
        if self.current_path == "This PC":  # Skip for "This PC" view
            return
            
        current_index = self.stack.currentIndex()
        if current_index == 3:  # Skip GalleryView
            return
            
        view = self.stack.currentWidget()
        indexes = view.selectedIndexes()
        
        if isinstance(view, QTreeView):
            rows = set()
            unique_indexes = []
            for index in indexes:
                if index.row() not in rows:
                    rows.add(index.row())
                    unique_indexes.append(index)
            indexes = unique_indexes
        
        if len(indexes) == 1:
            file_path = self.model.filePath(indexes[0]) if hasattr(indexes[0], 'model') and indexes[0].model() == self.model else None
            
            if file_path and os.path.exists(file_path):
                if hasattr(self, 'preview_window') and self.preview_window and self.last_previewed_file == file_path:
                    self.preview_window.close()
                    self.preview_window = None
                    self.last_previewed_file = None
                else:
                    self.show_media_preview(file_path)
                    self.last_previewed_file = file_path

    def keyPressEvent(self, event):
        """Handle keyboard events"""
        # Only handle space key if this window is active
        if not self.isActiveWindow():
            event.ignore()
            return
            
        if self.stack.currentIndex() == 3:  # Gallery view
            # Let the gallery view handle arrow keys
            if event.key() in (Qt.Key_Left, Qt.Key_Right):
                self.gallery_view.keyPressEvent(event)
                return

        if self.current_path == "This PC":  # Special handling for "This PC" view
            if event.key() in (Qt.Key_Return, Qt.Key_Enter):
                self.handle_enter_key()
                event.accept()
                return
            elif event.key() == Qt.Key_Space:
                event.accept()  # Ignore space in "This PC" view
                return
                
        current_index = self.stack.currentIndex()
        if current_index == 3:  # GalleryView
            super().keyPressEvent(event)
            return
            
        if event.key() == Qt.Key_Backspace:
            event.accept()
            self.move_to_trash_selected()
        elif event.key() in (Qt.Key_Return, Qt.Key_Enter):
            event.accept()
            self.handle_enter_key()
        elif event.key() == Qt.Key_Space:
            event.accept()
            self.handle_space_key()
        else:
            super().keyPressEvent(event)
        
    def rename_selected(self):
        """Allows renaming of the selected file or folder"""
        view = self.stack.currentWidget()
        indexes = view.selectedIndexes()
        
        if indexes:
            file_path = self.model.filePath(indexes[0])
            self.rename_item(file_path)
                    
    def copy_selected(self):
        """Copies selected files to clipboard"""
        view = self.stack.currentWidget()
        indexes = view.selectedIndexes()
        
        # For list view which has multiple columns, we need unique rows
        if isinstance(view, QTreeView):
            # Get unique rows (since each column in a row is a separate index)
            rows = set()
            unique_indexes = []
            for index in indexes:
                if index.row() not in rows:
                    rows.add(index.row())
                    unique_indexes.append(index)
            indexes = unique_indexes
        
        if indexes:
            self.clipboard.files = [self.model.filePath(index) for index in indexes]
            self.clipboard.operation = 'copy'

    def cut_selected(self):
        """Cuts selected files to clipboard"""
        view = self.stack.currentWidget()
        indexes = view.selectedIndexes()
        
        # Handle multiple columns in list view
        if isinstance(view, QTreeView):
            rows = set()
            unique_indexes = []
            for index in indexes:
                if index.row() not in rows:
                    rows.add(index.row())
                    unique_indexes.append(index)
            indexes = unique_indexes
        
        if indexes:
            self.clipboard.files = [self.model.filePath(index) for index in indexes]
            self.clipboard.operation = 'cut'

    def paste_files(self):
        from PyQt5.QtMultimedia import QSoundEffect
        from PyQt5.QtCore import QUrl
        import shutil
        """Paste files from clipboard to current directory with conflict handling"""
        if not self.clipboard.files or not self.clipboard.operation:
            return

        dest_dir = self.current_path if os.path.isdir(self.current_path) else QDir.homePath()
        icon_provider = QFileIconProvider()

        # Wykryj konflikty
        conflicts = []
        for src_path in self.clipboard.files:
            if not os.path.exists(src_path):
                continue
            base_name = os.path.basename(src_path)
            dest_path = os.path.join(dest_dir, base_name)
            if os.path.exists(dest_path):
                conflicts.append((src_path, dest_path))

        # Handle conflicts if any
        if conflicts:
            continue_paste, updated_files = self._handle_copy_conflicts(conflicts)
            if not continue_paste:
                return
            self.clipboard.files = updated_files

        # Perform the copy/move operation
        for src_path in self.clipboard.files:
            if not os.path.exists(src_path):
                continue

            base_name = os.path.basename(src_path)
            dest_path = os.path.join(dest_dir, base_name)

            if os.path.abspath(src_path) == os.path.abspath(dest_path):
                print(f"Pomijam kopiowanie samego siebie: {src_path}")
                continue

            file_info = QFileInfo(src_path)
            icon = icon_provider.icon(file_info)
            pixmap = icon.pixmap(64, 64)

            progress_dialog = CopyProgressDialog(base_name, pixmap, self)
            worker = CopyWorker(src_path, dest_path, self.clipboard.operation)
            progress_dialog.set_worker(worker)

            worker.signals.progress.connect(progress_dialog.progress.setValue)
            worker.signals.finished.connect(progress_dialog.accept)
            worker.signals.error.connect(self.show_error_dialog)                
            worker.start()

            if progress_dialog.exec_() == QDialog.Rejected:
                try:
                    if os.path.exists(dest_path):
                        if os.path.isdir(dest_path):
                            shutil.rmtree(dest_path)
                        else:
                            os.remove(dest_path)
                except Exception as e:                   
                    msg = self.create_message_box()
                    msg.setIcon(QMessageBox.Warning)
                    msg.setWindowTitle("Error")
                      
                    msg.setText(f"Could not remove partially copied file:\n{e}")
                    ok_button = msg.addButton("OK", QMessageBox.AcceptRole)
                    ok_button.setMinimumWidth(120)
                     
                    msg.exec_()  


                break

            if self.clipboard.operation == 'cut' and worker._is_running:
                try:
                    if os.path.isdir(worker.src):
                        shutil.rmtree(worker.src)
                    else:
                        os.remove(worker.src)
                except Exception as e:
                    msg = self.create_message_box()
                    msg.setIcon(QMessageBox.Warning)
                    msg.setWindowTitle("Error")
                      
                    msg.setText(f"Could not remove original file:\n{e}")
                    ok_button = msg.addButton("OK", QMessageBox.AcceptRole)
                    ok_button.setMinimumWidth(120)
                     
                    msg.exec_() 
        if self.clipboard.operation == 'cut':
            self.clipboard.files = []
            self.clipboard.operation = None
        self.sound = QSoundEffect()
        self.sound.setSource(QUrl.fromLocalFile("sounds\complete.wav"))
        self.sound.setVolume(0.8)
        self.sound.play()

       
    def open_new_window(self, path=None):
        """Opens a new instance of the file manager in a separate window."""
        new_file_manager = file_manager(path)
        new_file_manager.show()
        
    def open_in_new_window(self, path):
        """Opens the selected folder in a new window."""
        self.open_new_window(path)
        
    def toggle_search_bar(self):
        if self.search_bar.isVisible():
            self.search_bar.hide()
        else:
            self.search_bar.show()
            self.search_bar.setFocus()
 
        
        
    def sort_by(self, mode):
        if mode == "name":
            column = 0  # Nazwa pliku
        elif mode == "size":
            column = 1  # Rozmiar
        elif mode == "type":
            column = 2  # Typ
        elif mode == "date":
            column = 3  # Data modyfikacji
        else:
            return

        self.current_sort_key = mode

        for action in self.sort_actions.values():
            action.setChecked(False)

        if mode in self.sort_actions:
            self.sort_actions[mode].setChecked(True)

        if mode == "date":
            self.list_view.sortByColumn(column, Qt.DescendingOrder)
        else:
            self.list_view.sortByColumn(column, Qt.AscendingOrder)

        if mode == "name":
            self.icon_view.model().sort(0, Qt.AscendingOrder)
        elif mode == "date":
            self.icon_view.model().sort(3, Qt.DescendingOrder)
        elif mode == "size":
            self.icon_view.model().sort(1, Qt.AscendingOrder)
        elif mode == "type":
            self.icon_view.model().sort(2, Qt.AscendingOrder)


    def on_locations_click(self, item):
        # Cancel any active search
        if self.search_active:
            self.search_bar.clear()
            self.model.setNameFilters([])  # Reset filters
            self.model.setNameFilterDisables(False)
            self.search_active = False
            
        self.clear_selections_in_other_lists(self.locations)
        path = item.data(Qt.UserRole)
        if path == "thispc":
            self.show_this_pc()
        elif os.path.exists(path):
            # Check if the path is a drive (root of a filesystem)
            if self.is_drive_path(path):
                self.switch_to_normal_view(path)
            else:
                self.set_directory(path)
                
    def is_drive_path(self, path):
        # Normalize the path for comparison
        norm_path = os.path.normpath(path)
        # On Windows, drive roots are like "C:\", "D:\"
        if os.name == 'nt':
            # Check if the path is exactly a drive root (e.g., "C:\")
            if len(norm_path) == 3 and norm_path[1:3] == ':\\':
                return True
            # Also check for network drives? (e.g., "\\server\share")
            # But for simplicity, we assume local drives only
        else:
            # For Unix-like systems, check if the path is in the list of mount points
            for part in psutil.disk_partitions():
                if part.mountpoint == norm_path:
                    return True
        return False        


            
    def refresh_drives_list(self):
        """Refreshes the list of drives in the Locations section with Eject buttons"""
        self.locations.clear()
        self.locations.setIconSize(QSize(32, 32)) 
        # Add "This PC"
        pc_item = QListWidgetItem(QIcon("icons\pc.png"), "This PC")
        pc_item.setData(Qt.UserRole, "thispc")
        self.locations.addItem(pc_item)
        
        # Add all available drives
        for drive in self.list_drives():
            # Create a widget for each drive
            item_widget = QWidget()
            item_layout = QHBoxLayout(item_widget)
            item_layout.setContentsMargins(5, 2, 5, 2)
            
            # Add drive icon and name
            if drive['path'].startswith('C:\\'):
                icon = QLabel()
                icon.setPixmap(QIcon("icons\internal.png").pixmap(64, 32))
                item_layout.addWidget(icon)
                
                name_label = QLabel("Macintosh HD")
            else:
                icon = QLabel()
                icon.setPixmap(QIcon("icons\external.png").pixmap(64, 32))
                item_layout.addWidget(icon)
                
                name_label = QLabel(drive['name'] if drive['name'] else os.path.basename(drive['path'].rstrip("\\/")))
            
            item_layout.addWidget(name_label)
            item_layout.addStretch()
                        
            # Create QListWidgetItem and set widget
            item = QListWidgetItem()
            item.setData(Qt.UserRole, drive['path'])
            item.setSizeHint(item_widget.sizeHint())
            
            self.locations.addItem(item)
            self.locations.setItemWidget(item, item_widget)
            self.locations.setStyleSheet('font-size: 26px; font-family: "Segoe UI", "Helvetica Neue", sans-serif; background: transparent')
            self.locations.setObjectName("sidebar")
                        
    def show_sidebar_context_menu(self, pos):
        sender = self.sender()
        menu = CustomMenu(parent=self)
        
        item = sender.itemAt(pos)
        if item:
            path = item.data(Qt.UserRole)
            if path:
                if path != "thispc":
                    open_new_action = menu.addAction("Open in New Window")
                    open_new_action.triggered.connect(lambda: self.open_in_new_window(path))
                    # Add Get Info for both drives and folders
                    file_info_action = menu.addAction("Get Info")
                    file_info_action.setShortcut('CTRL+I')
                    file_info_action.triggered.connect(lambda _=None: self.show_file_or_drive_info(path))                
                else:
                    return

        else:
            new_window_action = menu.addAction("New file_manager Window")
            new_window_action.triggered.connect(self.open_new_window)
            
        menu.exec_(sender.viewport().mapToGlobal(pos))

        
    def show_context_menu(self, pos):
        view = self.sender()

        if hasattr(self, 'current_tag') and self.current_tag:
            # Obsługa menu kontekstowego dla widoku tagów
            self.show_tag_context_menu(pos, view)
            return
        
        if self.current_path == "This PC":
            menu = CustomMenu(parent=self)
            
            index = view.indexAt(pos)
            if index.isValid():
                # This is a drive item
                drive_path = index.data(Qt.UserRole)
                if drive_path:
                    # Add drive-specific actions
                    get_info_action = menu.addAction("Get Info")
                    get_info_action.setShortcut('CTRL+I')
                    get_info_action.triggered.connect(lambda: self.show_drive_info(index))
                    
                    menu.addSeparator()
            
            # Add basic options
            new_window_action = menu.addAction("New file_manager Window")
            new_window_action.triggered.connect(self.open_new_window)
            
            refresh_action = menu.addAction("Refresh")
            refresh_action.triggered.connect(self.refresh_drives_list)
            
            menu.exec_(view.viewport().mapToGlobal(pos))
            return
            
        elif self.current_path == "Trash":
            menu = CustomMenu(parent=self)

            restore_action = menu.addAction("Restore")
            restore_action.triggered.connect(self.restore_from_recycle_bin)
            empty_action = menu.addAction("Empty Trash")
            empty_action.triggered.connect(self.empty_trash)
            menu.addSeparator()
            refresh_action2 = menu.addAction("Refresh")
            refresh_action2.triggered.connect(self.show_recycle_bin)
            menu.addSeparator()
            new_window_action = menu.addAction("New file_manager Window")
            new_window_action.triggered.connect(self.open_new_window)



            



            menu.exec_(view.viewport().mapToGlobal(pos))
            return
           
        
        index = view.indexAt(pos)
        menu = CustomMenu(parent=self)       
        indexes = view.selectedIndexes()
        
        # For list view with multiple columns, get unique rows
        if isinstance(view, QTreeView):
            rows = set()
            unique_indexes = []
            for index in indexes:
                if index.row() not in rows:
                    rows.add(index.row())
                    unique_indexes.append(index)
            indexes = unique_indexes
        
        has_selection = len(indexes) > 0
        single_selection = len(indexes) == 1

        if has_selection:
            # Get all selected paths
            paths = [self.model.filePath(index) for index in indexes]
            
            # Add "Open" actions only for single selection
            if single_selection:
                file_path = paths[0]
                current_tags = [tag for tag, tagged_paths in self.tags.items() if file_path in tagged_paths]
                
                if os.path.isdir(file_path):
                    open_new_window_action = menu.addAction("Open in New Window")
                    open_new_window_action.triggered.connect(lambda: self.open_in_new_window(file_path))

                else:
                    open_action = menu.addAction("Open")
                    open_action.triggered.connect(lambda: os.startfile(file_path))
                    
                    # Add "Open as Admin" for files
                    open_admin_action = menu.addAction("Open as Admin")
                    open_admin_action.triggered.connect(lambda: self.open_as_admin(file_path))
                    
                    open_with_action = menu.addAction("Open with...")
                    open_with_action.triggered.connect(lambda: open_with_dialog(file_path))
                    
                
                menu.addSeparator()
            
            # These actions work with any number of selections
            
            menu.addSeparator()
            
            move_to_trash_action = menu.addAction("Move to Trash")
            move_to_trash_action.setShortcut('⭠')
            move_to_trash_action.triggered.connect(self.move_to_trash_selected)
            
            menu.addSeparator()

            copy_action = menu.addAction("Copy")
            copy_action.setShortcut('CTRL+C')
            copy_action.triggered.connect(self.copy_selected)
            
            cut_action = menu.addAction("Cut")
            cut_action.setShortcut('CTRL+X')
            cut_action.triggered.connect(self.cut_selected)
            
            menu.addSeparator()
            
            compress_action = menu.addAction("Compress...")
            compress_action.triggered.connect(self.compress)            
                                    
            # These actions work only with single selection
            if single_selection:
                file_path = paths[0]
                
                file_info_action = menu.addAction("Get Info")
                file_info_action.setShortcut('CTRL+I')
                file_info_action.triggered.connect(lambda: self.show_file_info(file_path))
                

                
                rename_action = menu.addAction("Rename")
                rename_action.setShortcut('Enter')
                rename_action.triggered.connect(lambda: self.rename_item(file_path))
                
                copy_path_action = menu.addAction("Copy as Path")
                copy_path_action.triggered.connect(lambda: QApplication.clipboard().setText(os.path.normpath(file_path)))
                
                create_shortcut_action = menu.addAction("Create Shortcut")
                create_shortcut_action.triggered.connect(lambda: self.create_shortcut(file_path))
                

                
                menu.addSeparator()
                quick_look_action = menu.addAction("Quick Look")
                quick_look_action.setShortcut('Space')
                quick_look_action.triggered.connect(lambda: self.show_media_preview(file_path))
                
                
                menu.addSeparator()
                
                
                current_tags = [tag for tag in self.available_tags if file_path in self.tags.get(tag, [])]
                
                from PyQt5.QtWidgets import QWidgetAction
                tag_widget = self.create_tags_widget(file_path)
                tag_action = QWidgetAction(menu)
                tag_action.setDefaultWidget(tag_widget)
                menu.addAction(tag_action)

        else:
            # Clicked on empty space
            new_window_action = menu.addAction("New file_manager Window")
            new_window_action.triggered.connect(self.open_new_window)

            new_folder_action = menu.addAction("New Folder")
            new_folder_action.triggered.connect(self.create_new_folder)
            menu.addSeparator()
            paste_action = menu.addAction("Paste")
            paste_action.setShortcut('CTRL+V')
            paste_action.triggered.connect(self.paste_files)
            menu.addSeparator()
            terminal_action = menu.addAction("Open in Terminal")
            terminal_action.triggered.connect(self.open_in_terminal)
            copy_path_action = menu.addAction("Copy as Path")
            copy_path_action.triggered.connect(lambda: QApplication.clipboard().setText(os.path.normpath(self.current_path)))              
           

        menu.exec_(view.viewport().mapToGlobal(pos))
        
        
        

    def show_tag_context_menu(self, pos, view):
        menu = CustomMenu(parent=self)
        
        index = view.indexAt(pos)
        indexes = view.selectedIndexes()
        
        if isinstance(view, QTreeView):
            rows = set()
            unique_indexes = []
            for index in indexes:
                if index.row() not in rows:
                    rows.add(index.row())
                    unique_indexes.append(index)
            indexes = unique_indexes
        
        has_selection = len(indexes) > 0
        single_selection = len(indexes) == 1

        if has_selection:
            paths = []
            for index in indexes:
                if hasattr(view.model(), 'data'):
                    file_path = index.data(Qt.UserRole)
                    if file_path:
                        paths.append(file_path)
            
            if single_selection:
                file_path = paths[0]
                
                if os.path.isdir(file_path):
                    open_new_window_action = menu.addAction("Open in New Window")
                    open_new_window_action.triggered.connect(lambda: self.open_in_new_window(file_path))
                else:
                    open_action = menu.addAction("Open")
                    open_action.triggered.connect(lambda: os.startfile(file_path))
                    
                menu.addSeparator()
            
            tag_menu = menu.addMenu("Edit Tags...")
            current_tags = [tag for tag, tagged_paths in self.tags.items() if file_path in tagged_paths]
            
            for tag in self.available_tags:
                action = tag_menu.addAction(tag)
                action.setCheckable(True)
                action.setChecked(tag in current_tags)
                action.triggered.connect(lambda checked, t=tag, p=file_path: self.toggle_single_tag(p, t))
            
            menu.addSeparator()
            
            move_to_trash_action = menu.addAction("Move to Trash")
            move_to_trash_action.setShortcut('⭠')
            move_to_trash_action.triggered.connect(self.move_to_trash_selected)
            
            menu.addSeparator()

            copy_action = menu.addAction("Copy")
            copy_action.setShortcut('CTRL+C')
            copy_action.triggered.connect(self.copy_selected)
            
            cut_action = menu.addAction("Cut")
            cut_action.setShortcut('CTRL+X')
            cut_action.triggered.connect(self.cut_selected)
            
            menu.addSeparator()
            
            compress_action = menu.addAction("Compress...")
            compress_action.triggered.connect(self.compress)            
                                    
            if single_selection:
                file_info_action = menu.addAction("Get Info")
                file_info_action.setShortcut('CTRL+I')
                file_info_action.triggered.connect(lambda: self.show_file_info(file_path))
                
                rename_action = menu.addAction("Rename")
                rename_action.setShortcut('Enter')
                rename_action.triggered.connect(lambda: self.rename_item(file_path))
                
                copy_path_action = menu.addAction("Copy as Path")
                copy_path_action.triggered.connect(lambda: QApplication.clipboard().setText(os.path.normpath(file_path)))
                
                create_shortcut_action = menu.addAction("Create Shortcut")
                create_shortcut_action.triggered.connect(lambda: self.create_shortcut(file_path))
                
                menu.addSeparator()
                quick_look_action = menu.addAction("Quick Look")
                quick_look_action.setShortcut('Space')
                quick_look_action.triggered.connect(lambda: self.show_media_preview(file_path))
        else:
            new_window_action = menu.addAction("New file_manager Window")
            new_window_action.triggered.connect(self.open_new_window)

        menu.exec_(view.viewport().mapToGlobal(pos))

        
        
    def open_as_admin(self, path):
        """Open file as administrator"""
        try:
            if not os.path.exists(path):
                raise FileNotFoundError(f"Path not found: {path}")

            ctypes.windll.shell32.ShellExecuteW(
                None,
                "runas",            # As administrator
                path,               # File to open
                None,               # Arguments (None if not applicable)
                None,               # Working directory
                1                   # Window display mode
            )

        except Exception as e:
            msg = self.create_message_box()
            msg.setIcon(QMessageBox.Warning)
            msg.setWindowTitle("Error")
              
            msg.setText(f"Could not open as administrator:\n{e}")
            ok_button = msg.addButton("OK", QMessageBox.AcceptRole)
            ok_button.setMinimumWidth(120)
            msg.exec_()

        
    def create_shortcut(self, target_path):
        import winshell
        """Create a shortcut (.lnk file) for the selected file/folder"""
        if not target_path or not os.path.exists(target_path):
            return
        
        # Get the parent directory of the target
        parent_dir = os.path.dirname(target_path)
        base_name = os.path.basename(target_path)
        
        # Default shortcut name
        shortcut_name = f"Shortcut to {base_name}"
        shortcut_path = os.path.join(parent_dir, shortcut_name + ".lnk")
        
        # Check if shortcut already exists
        counter = 1
        while os.path.exists(shortcut_path):
            shortcut_name = f"Shortcut to {base_name} ({counter})"
            shortcut_path = os.path.join(parent_dir, shortcut_name + ".lnk")
            counter += 1
        
        try:
            # Create the shortcut
            shortcut = winshell.shortcut(shortcut_path)
            shortcut.path = target_path
            shortcut.working_directory = os.path.dirname(target_path)
            shortcut.description = f"Shortcut to {base_name}"
            shortcut.write()
            
            # Refresh the view to show the new shortcut
            self.model.setRootPath(self.model.rootPath())
        except Exception as e:
            msg = self.create_message_box()
            msg.setIcon(QMessageBox.Warning)
            msg.setWindowTitle("Error")
              
            msg.setText(f"Could not create shortcut:\n{str(e)}")
            ok_button = msg.addButton("OK", QMessageBox.AcceptRole)
            ok_button.setMinimumWidth(120)
             
            msg.exec_() 
                       

    def toggle_hidden_files(self):
        """Toggle showing hidden files"""
        self.show_hidden_files = not self.show_hidden_files
        self.update_hidden_files_filter()
        
        # Refresh the view
        if self.current_path and self.current_path not in ["This PC", "Trash"]:
            self.model.setRootPath(self.current_path)      

    def compress(self):        
        import zipfile
        import datetime
        
        view = self.stack.currentWidget()
        indexes = view.selectedIndexes()
        if not indexes:
            return

        if isinstance(self.sender().parent().parent(), QTreeView):
            rows = set()
            indexes = [i for i in indexes if not (i.row() in rows or rows.add(i.row()))]

        file_paths = [self.model.filePath(index) for index in indexes]

        folder = os.path.dirname(file_paths[0])
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H-%M-%S")
        zip_name = os.path.join(folder, f"Archive {timestamp}.zip")

        try:
            with zipfile.ZipFile(zip_name, 'w', zipfile.ZIP_DEFLATED) as archive:
                for path in file_paths:
                    if os.path.isdir(path):
                        for root, dirs, files in os.walk(path):
                            for file in files:
                                full_path = os.path.join(root, file)
                                rel_path = os.path.relpath(full_path, folder)
                                archive.write(full_path, rel_path)
                    else:
                        rel_path = os.path.relpath(path, folder)
                        archive.write(path, rel_path)
            
            msg = self.create_message_box()
            msg.setIcon(QMessageBox.Information)
            msg.setWindowTitle("Compression complete")
              
            msg.setText(f"Created archive:\n{zip_name}")
            ok_button = msg.addButton("OK", QMessageBox.AcceptRole)
            ok_button.setMinimumWidth(120)
             
            msg.exec_() 
            
        except Exception as e:
            msg = self.create_message_box()
            msg.setIcon(QMessageBox.Critical)
            msg.setWindowTitle("Error")
              
            msg.setText(f"Failed to compress files:\n{str(e)}")
            ok_button = msg.addButton("OK", QMessageBox.AcceptRole)
            ok_button.setMinimumWidth(120)
             
            msg.exec_() 

        
    def open_in_terminal(self):
        """Open custom terminal in current directory"""
        if not self.current_path or self.current_path in ["This PC", "Trash"]:
            return

        if not os.path.isdir(self.current_path):
            return

        try:
            import subprocess
            import sys
            path = os.path.normpath(self.current_path)

            terminal_exe = r"Terminal.exe" # add the path to your terminal excutable!!
            subprocess.Popen([terminal_exe, "--cd", path], cwd=path)
        except Exception as e:
            print(f"Could not open terminal:\n{e}")

    def rename_item(self, file_path):
        """Starts renaming the selected file or folder in the current view."""
        if not file_path or not os.path.exists(file_path):
            return
            
        view = self.stack.currentWidget()
        index = self.model.index(file_path)
        
        if not index.isValid():
            return
            
        if isinstance(view, QListView):
            self.icon_view.edit(index)
        elif isinstance(view, QTreeView):
            self.list_view.edit(index)
        elif isinstance(view, QColumnView):
            self.column_view.edit(index)


            


            
    def can_rename_selected(self):
        """Checks if the selected item can be renamed"""
        view = self.stack.currentWidget()
        indexes = view.selectedIndexes()
        return len(indexes) == 1

    def show_file_info(self, file_path=None):
        from PyQt5.QtWidgets import QLayout
        """Show detailed file information with file icon, rename capability and preview"""
        if file_path is None or isinstance(file_path, bool):
            view = self.stack.currentWidget()
            indexes = view.selectedIndexes()
            
            if isinstance(view, QTreeView):
                rows = set()
                unique_indexes = []
                for index in indexes:
                    if index.row() not in rows:
                        rows.add(index.row())
                        unique_indexes.append(index)
                indexes = unique_indexes
            
            if len(indexes) != 1:
                return
                
            # Handle "This PC" view differently
            if self.current_path == "This PC":
                self.show_drive_info(indexes[0])
                return
                
            file_path = self.model.filePath(indexes[0])
        
        # Validate file_path
        if not file_path or not isinstance(file_path, str) or not os.path.exists(file_path):
            print(f"Invalid file path: {file_path}")
            return

        dialog = QDialog(self)
        dialog.setWindowTitle("File Info")
        dialog.setWindowIcon(QIcon("icons\file_manager.ico"))
        dialog.setFixedSize(500, 1100)
        
        # Apply theme settings immediately
        set_dark_title_bar(dialog, not self.theme)
        if self.theme:
            dialog.setStyleSheet(open("style.qss").read())
        else:
            dialog.setStyleSheet(open("style_dark.qss").read())
        
        main_layout = QVBoxLayout(dialog)
        
        # Top section with icon and preview
        top_section = QHBoxLayout()
        
        file_info = QFileInfo(file_path)
        icon_provider = QFileIconProvider()
        file_icon = icon_provider.icon(file_info)
        
        icon_label = QLabel()
        icon_label.setPixmap(file_icon.pixmap(64, 64))
        icon_label.setAlignment(Qt.AlignLeft)
        
        icon_name_label = QLabel(file_info.fileName())
        icon_name_label.setStyleSheet("font-weight: bold")
        icon_name_label.setWordWrap(True)
        icon_name_label.setFixedWidth(380)
        

        
        name_line2 = QLabel(f"Modified: {file_info.lastModified().toString('yyyy-MM-dd hh:mm:ss')}")
        name_line2.setStyleSheet("color: #666; font-size: 20px")
        
        top_section.addWidget(icon_label)
        
        icon_text = QVBoxLayout()
        icon_text.addWidget(icon_name_label)
        icon_text.addWidget(name_line2)           


        
        # Preview area
        preview_label = QLabel()
        preview_label.setAlignment(Qt.AlignCenter)
        preview_label.setFixedSize(456, 300)
        preview_label.setStyleSheet("border: 1px solid rgba(125, 125, 125, 0.3)")
        preview_label.setObjectName("preview")
        
        # Image preview logic
        is_image = file_path.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp', '.webp'))
        if is_image and os.path.isfile(file_path):
            try:
                pixmap = QPixmap(file_path)
                if not pixmap.isNull():
                    scaled = pixmap.scaled(preview_label.width(), preview_label.height(), 
                                         Qt.KeepAspectRatio, Qt.SmoothTransformation)
                    preview_label.setPixmap(scaled)
            except:
                preview_label.setText("Preview not available")
        else:
            preview_label.setText("No preview available")
        
        top_section.addLayout(icon_text)
        top_section.addStretch()
        main_layout.addLayout(top_section)

        # Separator with theme-aware color
        separator = QWidget()
        separator.setFixedHeight(2)
        separator.setStyleSheet("background-color: %s;" % ("#dddddd" if self.theme else "#444444"))
        main_layout.addWidget(separator)


        # Get image resolution
        pixmap = QPixmap(file_path)
        if not pixmap.isNull():
            width = pixmap.width()
            height = pixmap.height()
            resolution = f"{width} × {height}"
        else:
            resolution = "n/a"
            
        dimensions = QLabel("")
        dimensions.setStyleSheet("padding-top: 18px; padding-left: 13px; padding-bottom: 16px")
        dimensions.setText(f"Dimensions: {resolution}")
        main_layout.addWidget(dimensions)

        
        # Info sections
        sections = [
            ("General:", self.create_info_section(file_info)),
            ("More Info:", dimensions),
            ("Name & Extension:", self.create_rename_section(file_path, dialog)),
            ("Preview:", preview_label)
        ]
        
        for i, (header_text, content) in enumerate(sections):
            header = QLabel(header_text)
            main_layout.addWidget(header)
            
            if isinstance(content, QWidget):
                main_layout.addWidget(content)
            elif isinstance(content, QLayout):
                main_layout.addLayout(content)

            # Add separator unless it's the last section (i.e., "Preview")
            if i < len(sections) - 1:
                sep = QWidget()
                sep.setFixedHeight(2)
                sep.setStyleSheet("margin-top: 0px; margin-bottom: 0px; background-color: %s;" % ("#dddddd" if self.theme else "#444444"))
                main_layout.addWidget(sep)
        
        # Done button
        button_box = QHBoxLayout()
        cancel_btn = QPushButton("Done")
        cancel_btn.setFixedWidth(110)
        cancel_btn.setStyleSheet("""
            QPushButton {
                background-color: #0263e4; 
                border: 1px solid #0b4ca5; 
                border-radius: 8px; 
                color: white;
                padding: 6px;
            }
            QPushButton:pressed {
                background-color: #0385ff;
            }
        """)
        cancel_btn.clicked.connect(dialog.close)
        button_box.addStretch()
        button_box.addWidget(cancel_btn)
        main_layout.addStretch()
        main_layout.addLayout(button_box)
        
        # Ensure title bar updates after show
        dialog.showEvent = lambda e: set_dark_title_bar(dialog, not self.theme)
        
        dialog.exec_()

    def create_info_section(self, file_info):
        """Helper method to create file info section"""
        from PyQt5.QtWidgets import QFormLayout
        info_widget = QWidget()
        info_layout = QFormLayout(info_widget)
        info_layout.setVerticalSpacing(10)
        
        info_items = [
            ("Kind:", file_info.suffix().upper() if not file_info.isDir() else "Folder"),
            ("Size:", format_size(file_info.size())),
            ("Where:", file_info.path()),
            ("Created:", file_info.birthTime().toString('yyyy-MM-dd hh:mm:ss')),
            ("Modified:", file_info.lastModified().toString('yyyy-MM-dd hh:mm:ss'))
        ]
        
        for label, value in info_items:
            value_label = QLabel(value)
            value_label.setWordWrap(True)
            info_layout.addRow(label, value_label)
        
        return info_widget

    def create_rename_section(self, file_path, dialog):
        """Helper method to create rename section"""
        rename_box = QHBoxLayout()

        self.name_edit = QLineEdit()
        file_info = QFileInfo(file_path)
        full_name = file_info.fileName()  # pełna nazwa z rozszerzeniem
        self.name_edit.setText(full_name)

        rename_btn = QPushButton("Rename")
        rename_btn.setFixedWidth(110)
        rename_btn.setStyleSheet("""
            QPushButton {
                padding: 6px;
                border-radius: 5px;
            }
        """)
        rename_btn.clicked.connect(lambda: self.rename_file(file_path))

        rename_box.addWidget(self.name_edit)
        rename_box.addWidget(rename_btn)

        return rename_box

        

    def rename_file(self, old_path):
        """Rename the file based on user input"""
        new_name = self.name_edit.text().strip()
        if not new_name:
            return

        file_info = QFileInfo(old_path)
        dir_path = file_info.path()
        new_path = os.path.join(dir_path, new_name)

        if os.path.exists(new_path):
            msg = self.create_message_box()
            msg.setIcon(QMessageBox.Warning)
            msg.setWindowTitle("Rename")
            msg.setWindowIcon(QIcon(r\icons\file_manager.ico"))
            msg.setText("A file with that name already exists.")
            ok_button = msg.addButton("OK", QMessageBox.AcceptRole)
            ok_button.setMinimumWidth(120)
            msg.exec_()
            return

        try:
            os.rename(old_path, new_path)
            self.model.setRootPath(self.model.rootPath())  # refresh
        except Exception as e:
            msg = self.create_message_box()
            msg.setIcon(QMessageBox.Warning)
            msg.setWindowTitle("Error")
             
            msg.setText(f"Could not rename file:\n{str(e)}")
            ok_button = msg.addButton("OK", QMessageBox.AcceptRole)
            ok_button.setMinimumWidth(120)
            msg.exec_()
                     
    def show_file_or_drive_info(self, path=None):
        """Show info for either a file/folder or a drive"""
        from PyQt5.QtCore import QModelIndex

        if path is None:
            # Handle selection from main view
            if self.current_path == "This PC":
                view = self.stack.currentWidget()
                indexes = view.selectedIndexes()
                if indexes:
                    self.show_drive_info(indexes[0])
            else:
                self.show_file_info()
        else:
            # Handle path passed from sidebar
            if path == "thispc":
                self.show_drive_info()  # Show This PC info
            elif os.path.exists(path):
                if os.path.ismount(path):  # It's a drive
                    # Create a proper dummy index
                    class DummyIndex(QModelIndex):
                        def __init__(self, path):
                            super().__init__()
                            self._path = path
                            self._name = os.path.basename(path.rstrip("\\/"))
                        
                        def data(self, role=Qt.DisplayRole):
                            if role == Qt.DisplayRole:
                                return self._name
                            elif role == Qt.UserRole:
                                return self._path
                            return None
                    
                    self.show_drive_info(DummyIndex(path))
                else:  # Regular file/folder
                    self.show_file_info(path)
               

    def show_drive_info(self, index=None):
        from PyQt5.QtWidgets import QProgressBar, QFormLayout
        
        # Handle both direct calls and context menu calls
        if index is None:
            view = self.stack.currentWidget()
            indexes = view.selectedIndexes()
            if not indexes:
                return
            index = indexes[0]
        
        # Get drive path - handle both model items and direct path strings
        if hasattr(index, 'data'):
            drive_path = index.data(Qt.UserRole)
        else:
            drive_path = index  # Assume it's already the path
        
        if not drive_path:
            return
            
        dialog = QDialog(self)
        dialog.setWindowTitle("Drive Info")
        dialog.setFixedSize(500, 1100)
        
        set_dark_title_bar(dialog, not self.theme)
        if self.theme:
            dialog.setStyleSheet(open("style.qss").read())
        else:
            dialog.setStyleSheet(open("style_dark.qss").read())
        
        main_layout = QVBoxLayout(dialog)
        
        # Top section with icon and name
        top_section = QHBoxLayout()
        
        # Set appropriate icon based on drive type
        if drive_path.startswith('C:\\'):
            icon = QIcon("icons\hdd.png")
            name = "Macintosh HD"
        else:
            icon = QIcon("icons\usb.png")
            
            if hasattr(index, 'data'):
                name = index.data()  # Get the display name from the model
            else:
                name = os.path.basename(drive_path.rstrip("\\/"))
        
        icon_label = QLabel()
        icon_label.setPixmap(icon.pixmap(64, 64))
        icon_label.setAlignment(Qt.AlignLeft)
        
        name_label = QLabel(name)
        name_label.setStyleSheet("font-weight: bold")
        
        top_section.addWidget(icon_label)
        top_section.addWidget(name_label)
        top_section.addStretch()
        main_layout.addLayout(top_section)

        # Separator
        separator = QWidget()
        separator.setFixedHeight(2)
        separator.setStyleSheet("background-color: %s;" % ("#dddddd" if self.theme else "#444444"))
        main_layout.addWidget(separator)
        
        header = QLabel("General:")
        main_layout.addWidget(header)

        # Drive information section
        try:
            usage = psutil.disk_usage(drive_path)
            total_size = format_size(usage.total)
            free_space = format_size(usage.free)
            used_space = format_size(usage.used)
            percent_used = usage.percent
            file_system = get_file_system(drive_path)
        except:
            total_size = free_space = used_space = "Unknown"
            percent_used = 0

        info_widget = QWidget()
        info_layout = QFormLayout(info_widget)
        info_layout.setVerticalSpacing(10)
        
        info_items = [
            ("Total Size:", total_size),
            ("Used Space:", f"{used_space} ({percent_used}%)"),
            ("Free Space:", free_space),
            ("File System:", file_system),  # You could get this from psutil or other methods
            ("Path:", drive_path)
        ]
        
        for label, value in info_items:
            value_label = QLabel(value)
            value_label.setWordWrap(True)
            info_layout.addRow(label, value_label)
        
        main_layout.addWidget(info_widget)
        
        # Separator
        separator2 = QWidget()
        separator2.setFixedHeight(2)
        separator2.setStyleSheet("background-color: %s;" % ("#dddddd" if self.theme else "#444444"))
        main_layout.addWidget(separator2)
        
        header2 = QLabel("More Info:")
        main_layout.addWidget(header2)
        
        # Add a visual representation of disk usage
        usage_widget = QWidget()
        usage_layout = QVBoxLayout(usage_widget)
        
        usage_label = QLabel("Disk Usage:")
        usage_layout.addWidget(usage_label)
        
        # Progress bar showing used space
        progress = QProgressBar()
        progress.setValue(int(percent_used))
        progress.setTextVisible(False)
        progress.setFixedHeight(20)
        progress.setStyleSheet("""
            QProgressBar {
                background: rgba(124, 124, 124, 0.2);
                border-radius: 5px;
            }
            QProgressBar::chunk {
                background-color: #50aaf1;
                border-radius: 4px;
            }
        """)
        usage_layout.addWidget(progress)
        
        # Legend
        legend = QHBoxLayout()
        used_color = QLabel()
        used_color.setFixedSize(16, 16)
        used_color.setStyleSheet("background-color: #50aaf1; border-radius: 4px;")
        legend.addWidget(used_color)
        legend.addWidget(QLabel("Used space"))
        legend.addSpacing(20)
        
        free_color = QLabel()
        free_color.setFixedSize(16, 16)
        free_color.setStyleSheet("background-color: rgba(124, 124, 124, 0.3); border-radius: 4px;")
        legend.addWidget(free_color)
        legend.addWidget(QLabel("Free space"))
        legend.addStretch()
        
        usage_layout.addLayout(legend)
        main_layout.addWidget(usage_widget)
        
        
        
        # Separator
        separator3 = QWidget()
        separator3.setFixedHeight(2)
        separator3.setStyleSheet("background-color: %s;" % ("#dddddd" if self.theme else "#444444"))
        main_layout.addWidget(separator3)
        
        header3 = QLabel("Preview:")
        main_layout.addWidget(header3)
        
        icon_label2 = QLabel()
        scaled_pixmap = icon.pixmap(512, 512).scaled(256, 256, Qt.KeepAspectRatio, Qt.SmoothTransformation)
        icon_label2.setPixmap(scaled_pixmap)
        icon_label2.setAlignment(Qt.AlignCenter)
        icon_label2.setStyleSheet("border: 1px solid rgba(125, 125, 125, 0.3)")
        icon_label2.setObjectName("preview")
        icon_label2.setFixedSize(456, 300)
        main_layout.addWidget(icon_label2)
        main_layout.addStretch()
        
        # Done button
        button_box = QHBoxLayout()
        done_btn = QPushButton("Done")
        done_btn.setFixedWidth(110)
        done_btn.setStyleSheet("""
            QPushButton {
                background-color: #0263e4; 
                border: 1px solid #0b4ca5; 
                border-radius: 8px; 
                color: white;
                padding: 6px;
            }
            QPushButton:pressed {
                background-color: #0385ff;
            }
        """)
        done_btn.clicked.connect(dialog.close)
        button_box.addStretch()
        button_box.addWidget(done_btn)
        main_layout.addLayout(button_box)
        
        dialog.exec_()




    def create_new_folder(self):
        """Create a new folder and immediately start renaming it"""
        base_path = self.current_path if os.path.isdir(self.current_path) else QDir.homePath()
        
        # Create a default folder name
        counter = 1
        new_folder_name = "Untitled Folder"
        while os.path.exists(os.path.join(base_path, new_folder_name)):
            counter += 1
            new_folder_name = f"Untitled Folder {counter}"
        
        try:
            folder_path = os.path.join(base_path, new_folder_name)
            os.makedirs(folder_path)
            
            # Refresh the view to show the new folder
            self.model.setRootPath(self.model.rootPath())
            
            # Find the index of the newly created folder
            new_folder_index = self.model.index(folder_path)
            
            # Select the new folder
            view = self.stack.currentWidget()
            if isinstance(view, QListView):
                self.icon_view.setCurrentIndex(new_folder_index)
                self.icon_view.scrollTo(new_folder_index)
                QTimer.singleShot(100, lambda: self.icon_view.edit(new_folder_index))
            elif isinstance(view, QTreeView):
                self.list_view.setCurrentIndex(new_folder_index)
                self.list_view.scrollTo(new_folder_index)
                QTimer.singleShot(100, lambda: self.list_view.edit(new_folder_index))
            elif isinstance(view, QColumnView):
                self.column_view.setCurrentIndex(new_folder_index)
                self.column_view.scrollTo(new_folder_index)
                QTimer.singleShot(100, lambda: self.column_view.edit(new_folder_index))
                
        except Exception as e:
            # Create themed error message
            msg = self.create_message_box()
            msg.setIcon(QMessageBox.Critical)
            msg.setWindowTitle("Error")
             
            msg.setText(f"Could not create folder:\n{e}")
            
            # Create styled OK button
            ok_button = msg.addButton(" OK ", QMessageBox.AcceptRole)               
            set_dark_title_bar(msg, not self.theme)
            msg.exec_()
        
        
    def search_files(self, text):
        if not self.current_path or self.current_path == "This PC":
            return
        self.search_active = bool(text)    
        if not text:
            self.model.setNameFilters([])
            self.model.setNameFilterDisables(False)

            index = self.model.index(self.current_path)

            if self.stack.currentIndex() == 0:
                self.icon_view.setRootIndex(index)
            elif self.stack.currentIndex() == 1:
                self.list_view.setRootIndex(index)
            elif self.stack.currentIndex() == 2:
                self.column_view.setRootIndex(index)
            elif self.stack.currentIndex() == 3:
                self.load_gallery_view()

            return

        self.model.setFilter(QDir.AllEntries | QDir.NoDotAndDotDot)
        self.model.setNameFilters([f"*{text}*"])
        self.model.setNameFilterDisables(False)

        index = self.model.index(self.current_path)

        if self.stack.currentIndex() == 0:
            self.icon_view.setRootIndex(index)
        elif self.stack.currentIndex() == 1:
            self.list_view.setRootIndex(index)
        elif self.stack.currentIndex() == 2:
            self.column_view.setRootIndex(index)
        elif self.stack.currentIndex() == 3:
            self.load_gallery_view()


    def go_back(self):
        if self.history_index > 0:
            # Save the current scroll position before navigating back
            if self.current_path and self.current_path not in ["This PC", "Trash"]:
                current_view = self.stack.currentWidget()
                if isinstance(current_view, (QListView, QTreeView, QColumnView)):
                    scroll_position = current_view.verticalScrollBar().value()
                    self.scroll_positions[self.current_path] = scroll_position
            
            # Navigate back in history
            self.history_index -= 1
            previous_path = self.history[self.history_index]
            self.set_directory(previous_path, update_history=False)
            
            # Restore the saved scroll position for the previous path
            if previous_path in self.scroll_positions:
                QTimer.singleShot(100, lambda: self.restore_scroll_position(previous_path))

    def restore_scroll_position(self, path):
        """Restore the scroll position for the given path if it was saved previously."""
        if path in self.scroll_positions:
            current_view = self.stack.currentWidget()
            if isinstance(current_view, (QListView, QTreeView, QColumnView)):
                scroll_bar = current_view.verticalScrollBar()
                scroll_bar.setValue(self.scroll_positions[path])

    def go_forward(self):
        if self.history_index < len(self.history) - 1:
            # Save the current scroll position before navigating forward
            if self.current_path and self.current_path not in ["This PC", "Trash"]:
                current_view = self.stack.currentWidget()
                if isinstance(current_view, (QListView, QTreeView, QColumnView)):
                    scroll_position = current_view.verticalScrollBar().value()
                    self.scroll_positions[self.current_path] = scroll_position
            
            # Navigate forward in history
            self.history_index += 1
            next_path = self.history[self.history_index]
            self.set_directory(next_path, update_history=False)
            
            # Restore the saved scroll position for the next path
            if next_path in self.scroll_positions:
                QTimer.singleShot(100, lambda: self.restore_scroll_position(next_path))

        
    def show_this_pc(self):
        self.current_path = "This PC"
        self.update_path_bar("This PC")
        self.current_folder_label.setText("This PC")
        self.trash_bar.hide()

        drives = self.list_drives()
        self.create_this_pc_models(drives)
        
        current_index = self.stack.currentIndex()
        
        if current_index == 0:  # Icon view
            self.icon_view.setModel(self.this_pc_icon_model)
            self.icon_view.setIconSize(QSize(128, 128))  # Set appropriate icon size
            self.icon_view.setGridSize(QSize(150, 180))  # Adjust grid size to fit icons
            self.icon_view.doubleClicked.disconnect()
            self.icon_view.doubleClicked.connect(self.open_drive_from_this_pc)
            self.icon_view_layout.setContentsMargins(10, 10, 0, 0)
        elif current_index == 1:  # List view
            self.list_view.setModel(self.this_pc_list_model)
            self.list_view.setIconSize(QSize(32, 32))  # Set icon size for list view
            self.list_view.doubleClicked.disconnect()
            self.list_view.doubleClicked.connect(self.open_drive_from_this_pc_list)
        elif current_index == 2:  # Column view
            self.column_view.setModel(None)
            self.column_view.setDragEnabled(False)  # Disable dragging in column view for "This PC"
            self.column_view.setAcceptDrops(False)   
        elif current_index == 3:  # Gallery view
            self.gallery_view.clear()
            self.gallery_view.metadata_label.setText("Gallery view not available for the This PC tab.")

    def create_this_pc_models(self, drives):
        from PyQt5.QtGui import QStandardItemModel, QStandardItem
        """Creates models for the 'This PC' view"""
        # Icon view model
        self.this_pc_icon_model = QStandardItemModel()
        
        # List view model
        self.this_pc_list_model = QStandardItemModel()
        self.this_pc_list_model.setHorizontalHeaderLabels(["Name", "Type", "Total Size", "Free Space"])

        for drive in drives:
            # Get drive information
            drive_path = drive['path']
            drive_name = drive['name'] if 'name' in drive else os.path.basename(drive_path.rstrip("\\/"))
            
            # Set appropriate icon based on drive type
            if drive_path.startswith('C:\\'):
                icon = QIcon("icons\hdd.png")  # Use your internal drive icon
            else:
                icon = QIcon("icons\usb.png")  # Use your external drive icon
                
            # Add to icon view
            item = QStandardItem(icon, drive_name)
            item.setData(drive_path, Qt.UserRole)  # Store the full path as user data
            item.setEditable(False)
            item.setDropEnabled(False)
            item.setDragEnabled(False)
            self.this_pc_icon_model.appendRow(item)

            # Add to list view
            try:
                usage = psutil.disk_usage(drive_path)
                total_size = format_size(usage.total)
                free_space = format_size(usage.free)
            except:
                total_size = "Unknown"
                free_space = "Unknown"

            row = [
                QStandardItem(icon, drive_name),
                QStandardItem("Local Disk"),
                QStandardItem(total_size),
                QStandardItem(free_space)
            ]
            for item in row:
                item.setEditable(False)
                item.setData(drive_path, Qt.UserRole)
            
            self.this_pc_list_model.appendRow(row)


    def open_drive_from_this_pc(self, index):
        drive = index.data(Qt.UserRole)
        self.switch_to_normal_view(drive)

    def open_drive_from_this_pc_list(self, index):
        drive = index.data(Qt.UserRole)
        self.switch_to_normal_view(drive)

    def switch_to_normal_view(self, drive):
        """Switch from 'This PC' view to normal file system view for a specific drive"""
        if drive and os.path.exists(drive):
            # Restore the original model for the current view
            current_index = self.stack.currentIndex()
            
            if current_index == 0:  # Icon view
                self.icon_view.setModel(self.model)
                self.icon_view.doubleClicked.disconnect()
                self.icon_view.doubleClicked.connect(self.on_item_double_click)
                self.icon_view.setRootIndex(self.model.index(drive))
            elif current_index == 1:  # List view
                self.list_view.setModel(self.model)
                self.list_view.doubleClicked.disconnect()
                self.list_view.doubleClicked.connect(self.on_item_double_click)
                self.list_view.setRootIndex(self.model.index(drive))
            elif current_index == 2:  # Column view
                self.column_view.setModel(self.model)
                self.column_view.doubleClicked.disconnect()
                self.column_view.doubleClicked.connect(self.on_item_double_click)
                self.column_view.setRootIndex(self.model.index(drive))
            elif current_index == 3:  # Gallery view
                self.gallery_view.setModel(self.model)
                self.gallery_view.doubleClicked.disconnect()
                self.gallery_view.doubleClicked.connect(self.on_item_double_click)
                self.load_gallery_view()
            
            self.current_path = drive
            self.update_path_bar(drive)
            folder_name = os.path.basename(drive.rstrip("\\/")) if drive else "This PC"
            self.current_folder_label.setText(folder_name or drive)

    def list_drives(self):
        """Returns list of available drives with their names"""
        drives = []
        try:
            for part in psutil.disk_partitions():
                try:
                    if os.path.exists(part.mountpoint):
                        # Normalize path (removes trailing backslash on Windows)
                        normalized = os.path.normpath(part.mountpoint)
                        # For Windows ensure it ends with backslash
                        if os.name == 'nt' and not normalized.endswith('\\'):
                            normalized += '\\'
                        
                        # Get drive name if available
                        drive_name = ""
                        if os.name == 'nt':
                            try:
                                import ctypes
                                kernel32 = ctypes.windll.kernel32
                                volume_name = ctypes.create_unicode_buffer(1024)
                                file_system = ctypes.create_unicode_buffer(1024)
                                kernel32.GetVolumeInformationW(
                                    ctypes.c_wchar_p(normalized),
                                    volume_name,
                                    ctypes.sizeof(volume_name),
                                    None,
                                    None,
                                    None,
                                    file_system,
                                    ctypes.sizeof(file_system)
                                )
                                drive_name = volume_name.value
                            except:
                                pass
                        
                        drives.append({
                            'path': normalized,
                            'name': drive_name if drive_name else f"Local Disk ({normalized[:2]})",
                            'type': part.fstype
                        })
                except:
                    continue
        except:
            pass
        
        # Sort drives (C:, D:, etc.)
        if os.name == 'nt':
            drives.sort(key=lambda x: x['path'][0].upper())
        
        return drives

    def load_gallery_view(self):
        """Load images into gallery view from current directory"""
        if not self.current_path or not os.path.isdir(self.current_path):
            self.gallery_view.clear()
            self.gallery_view.metadata_label.setText("No directory selected.")
            return
            
        # Clear existing thumbnails safely
        try:
            self.gallery_view.clear()
        except RuntimeError:
            pass
            
        # Get all image files in directory
        image_extensions = ('.png', '.jpg', '.jpeg', '.gif', '.bmp', '.webp')
        image_files = []
        
        try:
            for filename in sorted(os.listdir(self.current_path)):
                if filename.lower().endswith(image_extensions):
                    file_path = os.path.join(self.current_path, filename)
                    image_files.append(file_path)
        except Exception as e:
            print(f"Error listing directory: {e}")
            return
        
        # Load images in a way that's interruptible
        for i, file_path in enumerate(image_files):
            try:
                pixmap = QPixmap(file_path)
                if not pixmap.isNull():
                    self.gallery_view.add_thumbnail(pixmap, file_path, i)
            except Exception as e:
                print(f"Error loading image {file_path}: {e}")
        
        # Update the current path display
        self.update_path_bar(self.current_path)
        folder_name = os.path.basename(self.current_path.rstrip("\\/")) if self.current_path else "This PC"
        self.current_folder_label.setText(folder_name or self.current_path)
        
        # Only try to select if we have thumbnails
        if hasattr(self.gallery_view, 'thumbnails') and self.gallery_view.thumbnails:
            try:
                self.gallery_view.select_thumbnail(0)
            except RuntimeError:
                pass
        else:
            self.gallery_view.metadata_label.setText("No images found in this folder.")

    def change_view(self, id):
        if self.current_path == "This PC":
            # Handle "This PC" view
            if id == 0:  # Icon view
                self.icon_view.setModel(self.this_pc_icon_model)
                self.icon_view.doubleClicked.disconnect()
                self.icon_view.doubleClicked.connect(self.open_drive_from_this_pc)
                self.icon_view_layout.setContentsMargins(10, 10, 0, 0)
            elif id == 1:  # List view
                self.list_view.setModel(self.this_pc_list_model)
                self.list_view.doubleClicked.disconnect()
                self.list_view.doubleClicked.connect(self.open_drive_from_this_pc_list)
                self.icon_view_layout.setContentsMargins(0, 0, 0, 0)
            elif id == 2:  # Column view
                self.icon_view.setModel(self.this_pc_list_model)
                self.icon_view_layout.setContentsMargins(0, 0, 0, 0)
                self.ensure_column_view_initialized()
            elif id == 3:  # Gallery view
                self.gallery_view.clear()
                self.gallery_view.metadata_label.setText("Gallery view not available for the This PC tab.")
                self.icon_view_layout.setContentsMargins(0, 0, 0, 0)
        elif self.current_path == "Trash":
            if id == 0:  # Icon view
                self.trash_bar.show()
                self.icon_view.setModel(self.recycle_bin_model)
                self.icon_view.setIconSize(QSize(128, 128))
                self.icon_view.setGridSize(QSize(150, 180))
                self.icon_view_layout.setContentsMargins(10, 0, 0, 0)
            elif id == 1:  # List view
                self.trash_bar.show()
                self.list_view.setModel(self.recycle_bin_model)
                self.icon_view_layout.setContentsMargins(0, 0, 0, 0)
            elif id == 2:  # Column view
                self.trash_bar.hide()
                self.column_view.setModel(None)
                self.icon_view_layout.setContentsMargins(0, 0, 0, 0)
                self.ensure_column_view_initialized()
            elif id == 3:  # Gallery view
                self.trash_bar.show()
                self.gallery_view.clear()
                self.gallery_view.metadata_label.setText("Gallery view not available for Trash.")
                self.icon_view_layout.setContentsMargins(0, 0, 0, 0)
        else:
            # Handle normal directory view
            if id == 0:  # Icon view
                self.icon_view.setModel(self.model)
                self.icon_view.doubleClicked.disconnect()
                self.icon_view.doubleClicked.connect(self.on_item_double_click)
                self.icon_view.setRootIndex(self.model.index(self.current_path))
                self.icon_view_layout.setContentsMargins(10, 0, 0, 0)
            elif id == 1:  # List view
                self.list_view.setModel(self.model)
                self.list_view.doubleClicked.disconnect()
                self.list_view.doubleClicked.connect(self.on_item_double_click)
                self.list_view.setRootIndex(self.model.index(self.current_path))
                self.icon_view_layout.setContentsMargins(0, 0, 0, 0)
            elif id == 2:  # Column view
                self.column_view.setModel(self.model)
                self.column_view.doubleClicked.disconnect()
                self.column_view.doubleClicked.connect(self.on_item_double_click)
                self.column_view.setRootIndex(self.model.index(self.current_path))
                self.icon_view_layout.setContentsMargins(0, 0, 0, 0)
                self.ensure_column_view_initialized()
            elif id == 3:  # Gallery view
                if self.current_path and os.path.isdir(self.current_path):
                    self.load_gallery_view()
                self.icon_view_layout.setContentsMargins(0, 0, 0, 0)
        
        self.stack.setCurrentIndex(id)
        # Force refresh of the view
        if self.current_path and self.current_path not in ["This PC", "Trash"]:
            self.model.setRootPath(self.current_path)



    def update_metadata(self, index=None):
        """Update metadata based on selected thumbnail"""
        if self.stack.currentIndex() == 3:  # Gallery view
            if index is None:
                if self.gallery_view.current_index >= 0:
                    selected_file = self.gallery_view.thumbnails[self.gallery_view.current_index].file_path
                    self.update_metadata_for_file(selected_file)
            else:
                file_path = self.model.filePath(index)
                self.update_metadata_for_file(file_path)

    def update_metadata_for_file(self, file_path):
        """Update metadata panel with information about specific file"""
        if not file_path:
            self.gallery_view.metadata_label.setText("No selection")
            return
            
        file_info = QFileInfo(file_path)
        
        # Basic info
        name = file_info.fileName()
        size = format_size(file_info.size())
        modified = file_info.lastModified().toString("yyyy-MM-dd hh:mm:ss")
        is_dir = "Yes" if file_info.isDir() else "No"
        
        # Image specific metadata (if available)
        img_metadata = ""
        if file_path.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp')):
            try:
                from PIL import Image
                from PIL.ExifTags import TAGS
                img = Image.open(file_path)
                width, height = img.size
                img_metadata = f"\nDimensions: {width}×{height} px"
                
                # Try to get EXIF data
                try:
                    exif_data = img._getexif()
                    if exif_data:
                        for tag_id, value in exif_data.items():
                            tag = TAGS.get(tag_id, tag_id)
                            if tag in ['DateTime', 'Model', 'Make', 'ExposureTime', 'FNumber', 'ISOSpeedRatings']:
                                if tag == 'ExposureTime':
                                    value = f"1/{int(1/value)}" if value < 1 else str(value)
                                elif tag == 'FNumber':
                                    value = f"f/{value}"
                                elif tag == 'ISOSpeedRatings':
                                    tag = 'ISO'
                                img_metadata += f"\n{tag}: {value}"
                except:
                    pass
            except ImportError:
                pass
        
        text = f"""
        <b>{name}</b><br>
        Size: {size} | Modified: {modified} | Folder: {is_dir}
        {img_metadata}
        """
        self.gallery_view.metadata_label.setText(text)
        

    def update_path_bar(self, path, tag=None):
        # Clear existing widgets in the path bar
        for i in reversed(range(self.path_bar.count())):
            widget = self.path_bar.itemAt(i).widget()
            if widget:
                widget.deleteLater()

        icon_provider = QFileIconProvider()
        
        if path == "This PC":
            btn = QPushButton("This PC")
            btn.setIcon(QIcon("icons\pc.png"))
            btn.setIconSize(QSize(28, 28))
            btn.setObjectName("btm")
            btn.setStyleSheet("""
                QPushButton {
                    border: none;
                    padding: 4px 8px;
                    background: transparent;
                }
                QPushButton:hover {
                    background: rgba(0, 0, 0, 0);
                }
            """)
            self.path_bar.addWidget(btn)
            return


        elif path == "Trash":
            
            btn = QPushButton("This PC")
            btn.setIcon(QIcon("icons\pc.png"))
            btn.setIconSize(QSize(28, 28))
            btn.setObjectName("btm")
            btn.setStyleSheet("""
                QPushButton {
                    border: none;
                    padding: 4px 8px;
                    background: transparent;
                }
                QPushButton:hover {
                    background: rgba(125, 125, 125, 0);
                }
            """)
            btn.clicked.connect(lambda: self.show_this_pc())
            self.path_bar.addWidget(btn)

            separator = QPushButton(">")
            separator.setObjectName("btm")
            separator.setStyleSheet("""
                QPushButton {
                    border: none;
                    padding: 4px 2px;
                    background: transparent;
                    font-family: "MS UI Gothic";
                    font-weight: bold;
                }
            """)
            separator.setEnabled(False)
            self.path_bar.addWidget(separator)

            btn = QPushButton("Trash")
            btn.setIcon(QIcon("icons\trash.png"))
            btn.setIconSize(QSize(24, 24))
            btn.setObjectName("btm")
            btn.setStyleSheet("""
                QPushButton {
                    border: none;
                    padding: 4px 8px;
                    background: transparent;
                }
                QPushButton:hover {
                    background: rgba(0, 0, 0, 0);
                }
            """)
            self.path_bar.addWidget(btn)
            return


        btn = QPushButton("This PC")
        btn.setIcon(QIcon("icons\pc.png"))
        btn.setIconSize(QSize(28, 28))
        btn.setObjectName("btm")
        btn.setStyleSheet("""
            QPushButton {
                border: none;
                padding: 4px 8px;
                background: transparent;
            }
            QPushButton:hover {
                background: rgba(0, 0, 0, 0);
            }
        """)
        btn.clicked.connect(lambda: self.show_this_pc())
        self.path_bar.addWidget(btn)

        separator = QPushButton(">")
        separator.setObjectName("btm")
        separator.setStyleSheet("""
            QPushButton {
                border: none;
                padding: 4px 2px;
                background: transparent;
                font-family: "MS UI Gothic";
                font-weight: bold;
            }
        """)
        separator.setEnabled(False)
        self.path_bar.addWidget(separator)

        # normalize path 
        norm_path = os.path.normpath(path)
        if os.name == 'nt':  # Windows
            drive, tail = os.path.splitdrive(norm_path)
            parts = []
            if drive:
                parts.append(drive)
            parts.extend(filter(None, tail.split(os.sep)))
        else:  # Unix
            parts = list(filter(None, norm_path.split(os.sep)))

        current_path = ""
        for i, part in enumerate(parts):
            if os.name == 'nt' and i == 0 and part.endswith(":"):
                current_path = part + "\\"
            else:
                current_path = os.path.join(current_path, part)

            file_info = QFileInfo(current_path)
            icon = icon_provider.icon(file_info)

            btn = QPushButton(part)
            btn.setObjectName("btm")
            btn.setIcon(icon)
            btn.setIconSize(QSize(16, 16))
            btn.setStyleSheet("""
                QPushButton {
                    border: none;
                    padding: 4px 8px;
                    background: transparent;
                    text-align: left;
                }
                QPushButton:hover {
                    background: rgba(0, 0, 0, 0);
                }
            """)
            btn.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)  
            btn.adjustSize()
            min_width = btn.sizeHint().width() + 20 
            btn.setFixedWidth(min_width)
            btn.clicked.connect(lambda checked, p=current_path: self.set_directory(p))
            self.path_bar.addWidget(btn)

            if i < len(parts) - 1:
                separator = QPushButton(">")
                separator.setObjectName("btm")
                separator.setStyleSheet("""
                    QPushButton {
                        border: none;
                        padding: 4px 2px;
                        background: transparent;
                        font-family: "MS UI Gothic";
                        font-weight: bold;
                    }
                """)
                separator.setEnabled(False)
                self.path_bar.addWidget(separator)

    def on_sidebar_click(self, item):        
        # Cancel any active search
        if self.search_active:
            self.search_bar.clear()
            self.model.setNameFilters([])  # Reset filters
            self.model.setNameFilterDisables(False)
            self.search_active = False
            
        self.clear_selections_in_other_lists(self.sidebar)
        path = item.data(Qt.UserRole)
        if path == "thispc":
            self.show_this_pc()
        elif os.path.exists(path):
            # Make sure we're using the correct model for the current view
            current_index = self.stack.currentIndex()
            
            if current_index == 0:  # Icon view
                self.icon_view.setModel(self.model)
                self.icon_view.doubleClicked.disconnect()
                self.icon_view.doubleClicked.connect(self.on_item_double_click)
            elif current_index == 1:  # List view
                self.list_view.setModel(self.model)
                self.list_view.doubleClicked.disconnect()
                self.list_view.doubleClicked.connect(self.on_item_double_click)
            elif current_index == 2:  # Column view
                self.column_view.setModel(self.model)
                self.column_view.doubleClicked.disconnect()
                self.column_view.doubleClicked.connect(self.on_item_double_click)
                
            self.set_directory(path)

    def on_item_double_click(self, index: QModelIndex):
        if not index.isValid():
            return
            
        """Handle double click - opens files/folders without triggering rename"""
        if QApplication.keyboardModifiers() == Qt.NoModifier:
            # Handle different model types
            model = index.model()
            
            # For tagged items (QStandardItemModel)
            if hasattr(model, 'item') and not hasattr(model, 'filePath'):
                file_path = index.data(Qt.UserRole)  # Get path from UserRole
                if not file_path:
                    # Fallback to sibling column if UserRole not set
                    file_path = index.sibling(index.row(), 1).data()
            else:
                # For regular filesystem items
                file_path = model.filePath(index)
                
            if file_path and os.path.exists(file_path):
                if os.path.isdir(file_path):
                    self.set_directory(file_path)
                else:
                    try:
                        os.startfile(file_path)
                    except Exception as e:
                        print(f"Cannot open: {e}")
                                                                       


# TAGS

    def load_tags(self):
        """Load tags from JSON file"""
        try:
            with open(self.tags_file, 'r') as f:
                self.tags = json.load(f)
        except (FileNotFoundError, json.JSONDecodeError):
            self.tags = defaultdict(list)

    def save_tags(self):
        """Save tags to JSON file"""
        with open(self.tags_file, 'w') as f:
            json.dump(self.tags, f)


    def get_tag_icon(self, tag_name):
        """Get icon for a tag based on its name (colored circle)"""
        color = self.available_tags.get(tag_name, "#808080")  # Default to gray if not found

        pixmap = QPixmap(24, 24)
        pixmap.fill(Qt.transparent)  # przezroczyste tło

        painter = QPainter(pixmap)
        painter.setRenderHint(QPainter.Antialiasing)
        painter.setBrush(QBrush(QColor(color)))
        painter.setPen(Qt.NoPen)
        painter.drawEllipse(0, 0, 24, 24)
        painter.end()

        return QIcon(pixmap)

    def update_tags_sidebar(self):
        """Update the tags section in sidebar"""
        # Clear existing tags
        for i in reversed(range(self.sidebar_layout.count())): 
            widget = self.sidebar_layout.itemAt(i).widget()
            if widget and widget.objectName() == "tag_item":
                widget.deleteLater()
        
        # Add current tags (only those that have files assigned)
        for tag in sorted(self.available_tags.keys()):
            if tag in self.tags and self.tags[tag]:  # Only show tags with files
                btn = QPushButton(tag)
                btn.setObjectName("tag_item")
                btn.setIcon(self.get_tag_icon(tag))
                btn.setStyleSheet("""
                    QPushButton {
                        text-align: left;
                        padding: 6px;
                    }
                """)
                btn.setFixedHeight(42)
                btn.clicked.connect(lambda checked, t=tag: self.show_tagged_files(t))
                self.sidebar_layout.addWidget(btn)
        self.sidebar_layout.addStretch()

    def show_tagged_files(self, tag):
        from PyQt5.QtGui import QStandardItemModel, QStandardItem
        """Show all files with the given tag"""
        self.trash_bar.hide()
        self.clear_selections_in_other_lists(None)
        
        # Store the current tag context
        self.current_tag = tag
        tagged_files = self.tags.get(tag, [])
        
        # Create a filtered model that works with the existing file system model
        self.tagged_files_model = QStandardItemModel()
        self.tagged_files_model.setHorizontalHeaderLabels(["Name", "Path"])
        
        for file_path in tagged_files:
            if os.path.exists(file_path):
                file_name = os.path.basename(file_path)
                item = QStandardItem(file_name)
                item.setData(file_path, Qt.UserRole)  # Store full path as user data
                item.setIcon(self.icon_provider.icon(QFileInfo(file_path)))
                path_item = QStandardItem(file_path)
                self.tagged_files_model.appendRow([item, path_item])
        
        # Set the model based on current view
        current_index = self.stack.currentIndex()
        
        if current_index == 0:  # Icon view
            self.icon_view.setModel(self.tagged_files_model)
            self.icon_view.doubleClicked.disconnect()
            self.icon_view.doubleClicked.connect(self.on_item_double_click)
        elif current_index == 1:  # List view
            self.list_view.setModel(self.tagged_files_model)
        elif current_index == 2:  # Column view
            # Column view doesn't work well with standard item model
            self.column_view.setModel(None)
        elif current_index == 3:  # Gallery view
            self.gallery_view.clear()
            self.gallery_view.metadata_label.setText("Gallery view not available for Tags.")
        
        self.current_path = f"Tag: {tag}"
        self.current_folder_label.setText(f"Tag: {tag}")
        self.update_path_bar(f"Tag: {tag}")

    def on_tagged_item_double_click(self, index):
        """Handle double click on items in tagged files view"""
        if QApplication.keyboardModifiers() == Qt.NoModifier:
            # Pobierz ścieżkę z modelu tagów
            file_path = index.data(Qt.UserRole)
            if not file_path:
                # Fallback do drugiej kolumny
                file_path = index.sibling(index.row(), 1).data()
                
            if file_path and os.path.exists(file_path):
                if os.path.isdir(file_path):
                    # Przełącz na normalny widok systemu plików
                    self.current_tag = None
                    self.set_directory(file_path)
                else:
                    try:
                        os.startfile(file_path)
                    except Exception as e:
                        msg = self.create_message_box()
                        msg.setIcon(QMessageBox.Warning)
                        msg.setWindowTitle("Error")
                         
                        msg.setText(f"Cannot open file:\n{e}")
                        ok_button = msg.addButton("OK", QMessageBox.AcceptRole)
                        ok_button.setMinimumWidth(120)
                        msg.exec_()

    def set_directory(self, path, update_history=True):
        # Cancel any active search
        if self.search_active:
            self.search_bar.clear()
            self.model.setNameFilters([])  # Reset filters
            self.model.setNameFilterDisables(False)
            self.search_active = False
        
        if not hasattr(self, 'scroll_positions'):
            self.scroll_positions = {}

        # Save current scroll position
        if self.current_path and self.current_path not in ["This PC", "Trash"]:
            current_view = self.stack.currentWidget()
            if isinstance(current_view, (QListView, QTreeView, QColumnView)):
                scroll_position = current_view.verticalScrollBar().value()
                self.scroll_positions[self.current_path] = scroll_position
        
        # Clear tag context when navigating to a normal directory
        self.current_tag = None
        
        if path:
            path = os.path.normpath(path)
            if not os.path.exists(path):
                return
            if not os.path.isdir(path):
                path = os.path.dirname(path)
        
        self.current_path = path or ""
        
        if update_history:
            if self.history_index < len(self.history) - 1:
                self.history = self.history[:self.history_index + 1]
            self.history.append(self.current_path)
            self.history_index += 1
        
        # Get current view index
        current_index = self.stack.currentIndex()
        
        # Set appropriate model and root index based on path
        if self.current_path == "This PC":
            self.show_this_pc()
            return
        elif self.current_path == "Trash":
            self.show_recycle_bin()
            return
        else:
            self.trash_bar.hide()
        
        # Set model and root index for the current view
        if current_index == 0:  # Icon view
            self.icon_view.setModel(self.model)
            self.icon_view.setRootIndex(self.model.index(self.current_path))
            
            if self.current_path == "This PC":
                self.icon_view_layout.setContentsMargins(10, 10, 0, 0)
            else: 
                self.icon_view_layout.setContentsMargins(10, 0, 0, 0)
                        
        elif current_index == 1:  # List view
            self.list_view.setModel(self.model)
            self.list_view.setRootIndex(self.model.index(self.current_path))
            self.icon_view_layout.setContentsMargins(0, 0, 0, 0)
        elif current_index == 2:  # Column view
            self.column_view.setModel(self.model)
            self.column_view.setRootIndex(self.model.index(self.current_path))
            self.icon_view_layout.setContentsMargins(0, 0, 0, 0)
        elif current_index == 3:  # Gallery view
            self.load_gallery_view()
            self.icon_view_layout.setContentsMargins(0, 0, 0, 0)
        
        # Update UI elements
        self.update_path_bar(self.current_path if self.current_path else "This PC")
        folder_name = os.path.basename(self.current_path.rstrip("\\/")) if self.current_path and self.current_path != "This PC" else "This PC"
        self.current_folder_label.setText(folder_name or self.current_path)
        
        # Restore scroll position after a short delay
        if self.current_path in self.scroll_positions:
            QTimer.singleShot(100, lambda: self.restore_scroll_position(self.current_path))

    def toggle_single_tag(self, path, tag):
        """Add or remove single tag from file"""
        if tag not in self.tags:
            self.tags[tag] = []
            
        if path in self.tags[tag]:
            self.tags[tag].remove(path)
        else:
            self.tags[tag].append(path)
        
        self.save_tags()
        self.update_tags_sidebar()

    def create_tags_widget(self, file_path):
        """Zwraca widget z kolorowymi kółkami tagów w poziomie"""

        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(8, 6, 8, 6)
        layout.setSpacing(6)

        title = QLabel("Tags...")
        title.setStyleSheet('margin-bottom: 4px; font-size: 25px; font-family: "Segoe UI", "Helvetica Neue", sans-serif;')
        layout.addWidget(title)

        row = QHBoxLayout()
        row.setSpacing(8)

        current_tags = [tag for tag in self.available_tags if file_path in self.tags.get(tag, [])]

        for tag in self.available_tags:
            icon = self.get_tag_icon(tag)
            btn = QPushButton()
            btn.setIcon(icon)
            btn.setIconSize(QSize(32, 32))
            btn.setFixedSize(32, 32)
            btn.setCheckable(True)
            btn.setChecked(tag in current_tags)
            btn.setStyleSheet("""
                QPushButton {
                    border: none;
                    background: transparent;
                }
                QPushButton:checked {
                    background: rgba(0, 0, 0, 0.1);
                    border-radius: 10px;
                }
            """)
            btn.clicked.connect(lambda checked, t=tag, p=file_path: self.toggle_single_tag(p, t))
            row.addWidget(btn)

        layout.addLayout(row)
        return widget
        
    def clear_selections_in_other_lists(self, current_list):
        """Deselect items in all lists except the current one"""
        if current_list != self.sidebar:
            self.sidebar.clearSelection()
        if current_list != self.locations:
            self.locations.clearSelection()
                    
                    
# TRASH

    def show_recycle_bin(self):
        from PyQt5.QtGui import QStandardItemModel, QStandardItem
        # Zmień model na specjalny model dla kosza
        self.current_path = "Trash"
        self.current_folder_label.setText("Trash")       
        
        # Utwórz model dla kosza
        self.recycle_bin_model = QStandardItemModel()
        self.recycle_bin_model.clear()
        self.recycle_bin_model.setHorizontalHeaderLabels(["Name", "Original Location", "Size", "Date Deleted"])
        
        # Pobierz zawartość kosza
        shell = win32com.Dispatch("Shell.Application")
        folder = shell.NameSpace(10)  # 10 = Recycle Bin
        items = folder.Items()
        
        for item in items:
            name = item.Name
            row_item = QStandardItem(name)
            row_item.setData(name, Qt.DisplayRole)  # Set display text
            original_path = folder.GetDetailsOf(item, 1)  # Original location
            date_deleted = folder.GetDetailsOf(item, 2)   # Date deleted
            size = folder.GetDetailsOf(item, 3)           # Size

            # Dodaj rozszerzenie na podstawie ścieżki tymczasowej (item.Path)
            try:
                path = item.Path
                extension = os.path.splitext(path)[1]
                if extension and not name.lower().endswith(extension.lower()):
                    name += extension
            except:
                path = ""

            # Ustaw ikonę
            icon = QIcon.fromTheme("unknown")  # Domyślna ikona
            
            try:
                # Próbuj uzyskać pełną ścieżkę do pliku (jeśli istnieje)
                path = item.Path
                if os.path.exists(path):
                    if os.path.isdir(path):
                        icon = self.icon_provider.icon(QFileIconProvider.Folder)
                    else:
                        icon = self.icon_provider.icon(QFileIconProvider.File)
            except:
                pass  # item.Path może nie być dostępne — np. skrót/element tymczasowy

            # Twórz elementy wiersza
            row = [
                QStandardItem(name),
                QStandardItem(original_path),
                QStandardItem(size),
                QStandardItem(date_deleted)
            ]
            pixmap = icon.pixmap(128, 128)
            row[0].setIcon(QIcon(pixmap))
            
            for col_item in row:
                col_item.setEditable(False)
                col_item.setData(item, Qt.UserRole)  # Przechowuj COM obiekt
            
            self.recycle_bin_model.appendRow(row)
        
        # Ustaw odpowiedni model w zależności od aktualnego widoku
        current_index = self.stack.currentIndex()
        
        if current_index == 0:  # Icon view
            self.icon_view.setModel(self.recycle_bin_model)
            self.trash_bar.show()
            self.icon_view.setIconSize(QSize(128, 128))
            self.icon_view.setGridSize(QSize(150, 180))
            self.icon_view.setDragEnabled(False)  # Wyłącz przeciąganie z kosza
            self.icon_view.setAcceptDrops(True)   # Pozwól na upuszczanie do kosza
        elif current_index == 1:  # List view
            self.list_view.setModel(self.recycle_bin_model)
            self.trash_bar.show()
            self.list_view.setDragEnabled(False)  # Wyłącz przeciąganie z kosza
            self.list_view.setAcceptDrops(True)   # Pozwól na upuszczanie do kosza
        elif current_index == 2:  # Column view
            self.column_view.setModel(None)
            self.trash_bar.hide()
            self.column_view.setDragEnabled(False)  # Wyłącz przeciąganie z kosza
            self.column_view.setAcceptDrops(True)   # Pozwól na upuszczanie do kosza
        elif current_index == 3:  # Gallery view
            self.trash_bar.show()
            self.gallery_view.clear()
            self.gallery_view.metadata_label.setText("Gallery view not available for Trash.")
        
        # Aktualizuj pasek ścieżki
        self.update_path_bar("Trash")
        

    def restore_from_recycle_bin(self):
        import shutil
        """Restores selected files from recycle bin"""
        if self.current_path != "Trash":
            return

        view = self.stack.currentWidget()
        indexes = view.selectedIndexes()
        
        # For list view, get unique rows
        if isinstance(view, QTreeView):
            rows = set()
            unique_indexes = []
            for index in indexes:
                if index.row() not in rows:
                    rows.add(index.row())
                    unique_indexes.append(index)
            indexes = unique_indexes

        if not indexes:
            return
            
        # Get all selected items
        items_to_restore = []
        for index in indexes:
            item = self.recycle_bin_model.item(index.row(), 0)
            com_item = item.data(Qt.UserRole)
            original_path = self.recycle_bin_model.item(index.row(), 1).text()
            items_to_restore.append((com_item, original_path))

        # Try to restore each item
        restored_count = 0
        for com_item, original_path in items_to_restore:
            try:
                # Use Verbs().DoIt() like in PowerShell
                verbs = com_item.Verbs()
                restored = False
                for i in range(verbs.Count):
                    verb = verbs.Item(i)
                    name = verb.Name.replace("&", "").strip().lower()
                    if "przywróć" in name or "restore" in name:
                        verb.DoIt()
                        restored_count += 1
                        restored = True
                        break
                
                if not restored:
                    raise Exception("Restore verb not found")

            except Exception as e:
                # Fallback: manually move file
                try:
                    temp_path = com_item.Path
                    if os.path.exists(temp_path):
                        parent_dir = os.path.dirname(original_path)
                        if not os.path.exists(parent_dir):
                            os.makedirs(parent_dir)
                        shutil.move(temp_path, original_path)
                        restored_count += 1
                except Exception as fallback_error:
                    msg = self.create_message_box()
                    msg.setIcon(QMessageBox.Warning)
                    msg.setWindowTitle("Error")
                     
                    msg.setText(f"Could not restore {original_path}:\n{str(fallback_error)}")
                    ok_button = msg.addButton("OK", QMessageBox.AcceptRole)
                    ok_button.setMinimumWidth(120)
                     
                    msg.exec_()

                    continue

        if restored_count > 0:
            self.show_recycle_bin()
        else:
            msg = self.create_message_box()
            msg.setIcon(QMessageBox.Warning)
            msg.setWindowTitle("Error")
             
            msg.setText("Could not restore any items.")
            ok_button = msg.addButton("OK", QMessageBox.AcceptRole)
            ok_button.setMinimumWidth(120)
             
            msg.exec_()



    def empty_trash(self):
        """Empties the recycle bin after user confirmation"""
        if self.current_path != "Trash":
            return
      
        msg = self.create_message_box()
        msg.setIcon(QMessageBox.Question)
        msg.setWindowTitle("Confirm")
         
        msg.setText('Are you sure you want to empty the Trash?')

        yes_button = msg.addButton("Yes", QMessageBox.YesRole)
        no_button = msg.addButton("No", QMessageBox.NoRole)

        for btn in msg.buttons():
            btn.setMinimumWidth(120)

         

        msg.exec_()

        if msg.clickedButton() == yes_button:             
            try:
                flags = 0x00000001 | 0x00000002  # No confirmation + No progress UI
                result = ctypes.windll.shell32.SHEmptyRecycleBinW(None, None, flags)

                if result == 0:
                    self.show_recycle_bin()  # ← ODŚWIEŻ widok tylko po udanym czyszczeniu
                    
                    
                    from PyQt5.QtMultimedia import QSoundEffect
                    from PyQt5.QtCore import QUrl


                    self.sound = QSoundEffect()
                    self.sound.setSource(QUrl.fromLocalFile("icons\empty.wav"))
                    self.sound.setVolume(1)
                    self.sound.play()
                    
                else:                    
                    msg = self.create_message_box()
                    msg.setIcon(QMessageBox.Warning)
                    msg.setWindowTitle("Warning")
                     
                    msg.setText("Trash could not be emptied.")
                    ok_button = msg.addButton("OK", QMessageBox.AcceptRole)
                    ok_button.setMinimumWidth(120)
                     
                    msg.exec_()
                    
            except Exception as e:
                msg = self.create_message_box()
                msg.setIcon(QMessageBox.Critical)
                msg.setWindowTitle("Error")
                 
                msg.setText(f"Could not empty Trash:\n{str(e)}")
                ok_button = msg.addButton("OK", QMessageBox.AcceptRole)
                ok_button.setMinimumWidth(120)
                 
                msg.exec_()                


    def get_original_path_from_trash(self, trash_path):
        """Get original path of a file in Trash"""
        if not hasattr(self, 'recycle_bin_model'):
            return None
            
        for row in range(self.recycle_bin_model.rowCount()):
            path_item = self.recycle_bin_model.item(row, 0)
            if path_item and path_item.data(Qt.UserRole) and hasattr(path_item.data(Qt.UserRole), 'Path'):
                try:
                    if path_item.data(Qt.UserRole).Path == trash_path:
                        return self.recycle_bin_model.item(row, 1).text()
                except:
                    continue
        return None
                

                
# Quick file preview 

    def show_media_preview(self, file_path=None):        
        """Show media_preview preview for selected files"""
        
        from media_preview.media_preview import MediaPreview
        
        # If no file_path provided, get it from selection
        if file_path is None:
            view = self.stack.currentWidget()
            indexes = view.selectedIndexes()
            
            if isinstance(view, QTreeView):
                rows = set()
                unique_indexes = []
                for index in indexes:
                    if index.row() not in rows:
                        rows.add(index.row())
                        unique_indexes.append(index)
                indexes = unique_indexes
            
            if len(indexes) == 1:
                file_path = self.model.filePath(indexes[0])
                self.last_previewed_file = file_path
        
        if file_path and os.path.exists(file_path):
            # Create new preview window
            self.preview_window = MediaPreview(file_path)
            self.preview_window.show()
            self.preview_window.raise_()
            self.preview_window.activateWindow()
       
    def is_supported_file(self, path):
        """All files should show preview, even unsupported ones"""
        return True  # Always return True to show preview for all files

    def closeEvent(self, event):
        # Close any open preview window
        if hasattr(self, 'preview_window') and self.preview_window:
            self.preview_window.close()
        
        # Stop the registry monitor
        if hasattr(self, 'registry_monitor'):
            self.registry_monitor.stop()
            
        super().closeEvent(event)
        
    def show_error_dialog(self, e):
        msg = self.create_message_box()
        msg.setIcon(QMessageBox.Warning)
        msg.setWindowTitle("Error")
         
        msg.setText(f"Operation failed:\n{e}")
        button = msg.addButton("OK", QMessageBox.AcceptRole)
        button.setMinimumWidth(120)         
        msg.exec_()
    

                
class CustomMenu(QMenu):
    def __init__(self, title="", parent=None):
        super().__init__(title, parent)
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.NoDropShadowWindowHint | Qt.Popup)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setStyleSheet('''
            QMenu {
                border-radius: 10px;
                padding: 6px;
                border: 1px solid transparent;
                font-size: 25px; 
                font-family: "Segoe UI", "Helvetica Neue", sans-serif;
            }
            QMenu::item {
                padding: 6px 12px;
                border-radius: 6px;
            }
            QMenu::item:selected {
                background-color: rgba(135, 135, 135, 100)
            }
            QMenu::separator {
                margin-left: 20px;
                margin-right: 20px;
                background-color: rgba(135, 135, 135, 0.2);
                height: 2px;
            }
        ''')
        


        
class ColumnViewDelegate(QStyledItemDelegate):
    def __init__(self, parent=None):
        super().__init__(parent)
        
    def paint(self, painter, option, index):
        from PyQt5.QtGui import QPalette
        painter.save()
        
        if option.state & QStyle.State_Selected:
            highlight_color = QColor("#0064e1")
            painter.fillRect(option.rect, highlight_color)
            
            option.palette.setColor(QPalette.HighlightedText, Qt.white)
        else:
            option.palette.setColor(QPalette.Highlight, option.palette.color(QPalette.Active, QPalette.Highlight))
            option.palette.setColor(QPalette.HighlightedText, option.palette.color(QPalette.Active, QPalette.HighlightedText))
        
        super().paint(painter, option, index)
        
        painter.restore()
        
        
class CopyProgressDialog(QDialog):
    def __init__(self, file_name, icon_pixmap, parent=None):
        super().__init__(parent)
        self.setWindowTitle(" ")
        self.setFixedSize(900, 100)
        self.setWindowModality(Qt.WindowModal)
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)
        self.worker = None  # Reference to worker

        # Apply theme settings
        self.settings = QSettings(
            "HKEY_CURRENT_USER\\Software\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize",
            QSettings.NativeFormat
        )
        self.theme = self.settings.value("AppsUseLightTheme", True)
        set_dark_title_bar(self, not self.theme)

        layout = QHBoxLayout()
        self.setLayout(layout)

        self.icon = QLabel()
        self.icon.setFixedSize(64, 64)
        self.icon.setPixmap(icon_pixmap)
        layout.addWidget(self.icon, alignment=Qt.AlignHCenter)

        vbox = QVBoxLayout()
        vbox.addStretch()

        self.status_label = QLabel(f"Copying: {file_name}")
        vbox.addWidget(self.status_label)

        from PyQt5.QtWidgets import QProgressBar

        self.progress = QProgressBar()
        self.progress.setTextVisible(False)
        self.progress.setFixedHeight(14)
        vbox.addWidget(self.progress)
        vbox.addStretch()
        layout.addLayout(vbox)

        self.cancel_button = QPushButton()
        self.cancel_button.setFixedSize(45, 45)
        self.cancel_button.setIcon(QIcon("close.png"))
        self.cancel_button.setIconSize(QSize(40, 40))
        self.cancel_button.setStyleSheet("background: none; border: none")
        layout.addWidget(self.cancel_button)

        if self.theme:
            self.setStyleSheet('''
                QProgressBar {
                    background-color: #dbdbdb;
                    border: 1px solid #c0c0c0;
                    border-radius: 7px;
                    height: 14px;
                    text-align: center;
                }
                QProgressBar::chunk {
                    background-color: #419bf9;
                    border-top-left-radius: 6px;
                    border-bottom-left-radius: 6px;
                }
            ''')
        else:
            self.setStyleSheet('''
                QProgressBar {
                    background-color: #4b4b4b;
                    border: 1px solid #666666;
                    border-radius: 7px;
                    height: 14px;
                    text-align: center;
                }                
                QProgressBar::chunk {
                    background-color: #3a79de;
                    border-top-left-radius: 6px;
                    border-bottom-left-radius: 6px;
                }
            ''')

    def showEvent(self, event):
        """Ensure title bar theme is applied when shown"""
        super().showEvent(event)
        set_dark_title_bar(self, not self.theme)

    def set_worker(self, worker):
        """Sets the worker for this dialog"""
        self.worker = worker
        self.cancel_button.clicked.connect(self.cancel_operation)

    def cancel_operation(self):
        """Cancels the copy operation"""
        if self.worker:
            self.worker.stop()  # Stop the worker
        self.reject()  # Close the dialog





def open_with_dialog(file_path):
    import ctypes.wintypes
    SEE_MASK_INVOKEIDLIST = 0x0000000C

    class SHELLEXECUTEINFO(ctypes.Structure):
        _fields_ = [
            ('cbSize', ctypes.wintypes.DWORD),
            ('fMask', ctypes.wintypes.ULONG),
            ('hwnd', ctypes.wintypes.HWND),
            ('lpVerb', ctypes.wintypes.LPCWSTR),
            ('lpFile', ctypes.wintypes.LPCWSTR),
            ('lpParameters', ctypes.wintypes.LPCWSTR),
            ('lpDirectory', ctypes.wintypes.LPCWSTR),
            ('nShow', ctypes.c_int),
            ('hInstApp', ctypes.wintypes.HINSTANCE),
            ('lpIDList', ctypes.c_void_p),
            ('lpClass', ctypes.wintypes.LPCWSTR),
            ('hkeyClass', ctypes.wintypes.HKEY),
            ('dwHotKey', ctypes.wintypes.DWORD),
            ('hIcon', ctypes.wintypes.HANDLE),
            ('hProcess', ctypes.wintypes.HANDLE),
        ]

    ShellExecuteEx = ctypes.windll.shell32.ShellExecuteExW

    sei = SHELLEXECUTEINFO()
    sei.cbSize = ctypes.sizeof(sei)
    sei.fMask = SEE_MASK_INVOKEIDLIST
    sei.hwnd = None
    sei.lpVerb = "openas"  # to wymusza "Otwórz za pomocą"
    sei.lpFile = file_path
    sei.lpParameters = None
    sei.lpDirectory = None
    sei.nShow = 1  # SW_SHOWNORMAL

    ShellExecuteEx(ctypes.byref(sei))
    

class CustomFileSystemModel(QFileSystemModel):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setIconProvider(CustomIconProvider())
        
    def data(self, index, role=Qt.DisplayRole):
        if role == Qt.DisplayRole and index.column() == 0:
            file_info = self.fileInfo(index)
            file_path = file_info.filePath()
            
            if file_path.lower().endswith('.lnk'):
                return f"↗ {file_info.fileName()}"
                
        return super().data(index, role)

class CustomIconProvider(QFileIconProvider):
    def icon(self, info):
        if info.isFile() and info.filePath().lower().endswith('.lnk'):
            try:
                shell = win32com.Dispatch("WScript.Shell")
                shortcut = shell.CreateShortCut(info.filePath())
                target_path = shortcut.Targetpath
                if os.path.exists(target_path):
                    target_info = QFileInfo(target_path)
                    return super().icon(target_info)
            except:
                pass
        return super().icon(info)


class RegistryMonitorWorker(QObject):
    theme_changed = pyqtSignal()

    def __init__(self):
        super().__init__()
        self._running = True
        self.last_value = None

    def run(self):
        win32con = lazy_import('win32con')
        win32api = lazy_import('win32api')
        self.last_value = self._get_current_theme_value()
        while self._running:
            try:
                current_value = self._get_current_theme_value()
                if current_value is not None and current_value != self.last_value:
                    self.last_value = current_value
                    self.theme_changed.emit()
                time.sleep(1)  # Check every second
            except Exception as e:
                print(f"Registry monitor error: {e}")
                time.sleep(5)  # Wait longer after errors

    def stop(self):
        self._running = False

    def _get_current_theme_value(self):
        try:
            win32con = lazy_import('win32con')
            win32api = lazy_import('win32api')

            key = win32api.RegOpenKeyEx(
                win32con.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize",
                0,
                win32con.KEY_READ
            )
            value, _ = win32api.RegQueryValueEx(key, "AppsUseLightTheme")
            win32api.RegCloseKey(key)
            return value
        except Exception as e:
            print(f"Error reading registry: {e}")
            return None

class RegistryMonitor:
    def __init__(self):
        self.thread = QThread()
        self.worker = RegistryMonitorWorker()
        self.worker.moveToThread(self.thread)
        self.thread.started.connect(self.worker.run)

    def start(self):
        self.thread.start()

    def stop(self):
        self.worker.stop()
        self.thread.quit()
        self.thread.wait()

    def connect_theme_changed(self, callback):
        self.worker.theme_changed.connect(callback)

def format_size(bytes):
    """Format file size in human-readable format"""
    for unit in ['B', 'KB', 'MB', 'GB', 'TB']:
        if bytes < 1024.0:
            return f"{bytes:.1f} {unit}"
        bytes /= 1024.0
    return f"{bytes:.1f} PB"

def set_dark_title_bar(widget, enable: bool):
    try:
        if not widget.winId():  # Sprawdź czy widget ma już przypisane window handle
            # Jeśli nie, odłóż ustawienie na później
            def delayed_set():
                try:
                    hwnd = int(widget.winId())
                    _set_dark_mode(hwnd, enable)
                except:
                    pass
            
            QTimer.singleShot(100, delayed_set)
            return
            
        hwnd = int(widget.winId())
        _set_dark_mode(hwnd, enable)
        
    except Exception as e:
        print(f"[Title Bar] Failed to set dark mode: {e}")

def _set_dark_mode(hwnd, enable):
    platform = lazy_import('platform')
    """Wewnętrzna funkcja ustawiająca tryb dark"""
    build_number = int(platform.version().split('.')[2])
    DWMWA_USE_IMMERSIVE_DARK_MODE = 20 if build_number >= 22000 else 19

    value = ctypes.c_int(1 if enable else 0)
    ctypes.windll.dwmapi.DwmSetWindowAttribute(
        ctypes.c_void_p(hwnd),
        ctypes.c_int(DWMWA_USE_IMMERSIVE_DARK_MODE),
        ctypes.byref(value),
        ctypes.sizeof(value)
    )

class ThemedMessageBox(QMessageBox):
    def __init__(self, parent=None):
        super().__init__(parent)
        settings = QSettings(
            "HKEY_CURRENT_USER\\Software\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize",
            QSettings.NativeFormat
        )
        self.theme = settings.value("AppsUseLightTheme", True)
        set_dark_title_bar(self, not self.theme)
        
    def showEvent(self, event):
        """Zawsze aktualizuj motyw przy pokazywaniu"""
        super().showEvent(event)
        set_dark_title_bar(self, not self.theme)


class GalleryView(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.main_layout = QVBoxLayout(self)
        self.main_layout.setContentsMargins(0, 0, 0, 0)
        self.main_layout.setSpacing(0)
        
        # Large preview area
        self.preview_area = QLabel()
        self.preview_area.setAlignment(Qt.AlignCenter)
        self.preview_area.setMinimumHeight(400)
        self.preview_area.setObjectName("preview")
        
        # Thumbnail scroll area
        self.thumbnail_scroll = QScrollArea()
        self.thumbnail_scroll.setFixedHeight(200)
        self.thumbnail_scroll.setObjectName("scroll")
        self.thumbnail_scroll.setWidgetResizable(True)
        self.thumbnail_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.thumbnail_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)

        separator = QWidget()
        separator.setObjectName("seplight")
        separator.setFixedHeight(2)
        
        separator2 = QWidget()
        separator2.setObjectName("seplight")
        separator2.setFixedHeight(2)       
        
        
        # Thumbnail container
        self.thumbnail_container = QWidget()
        self.thumbnail_layout = QHBoxLayout(self.thumbnail_container)
        self.thumbnail_layout.setContentsMargins(5, 5, 5, 5)
        self.thumbnail_layout.setSpacing(10)
        self.thumbnail_scroll.setWidget(self.thumbnail_container)
        
        # Metadata panel
        self.metadata_panel = QWidget()
        self.metadata_panel.setFixedHeight(100)
        self.metadata_layout = QVBoxLayout(self.metadata_panel)
        self.metadata_label = QLabel()
        self.metadata_label.setWordWrap(True)
        self.metadata_layout.addWidget(self.metadata_label)
        
        # Add widgets to main layout
        self.main_layout.addWidget(self.preview_area, 2)
        self.main_layout.addWidget(separator)
        
        self.main_layout.addWidget(self.thumbnail_scroll, 1)
        
        self.main_layout.addWidget(separator2)
        
        self.main_layout.addWidget(self.metadata_panel, 0)
        
        # Current image index
        self.current_index = -1
        self.thumbnails = []

        # Enable keyboard focus
        self.setFocusPolicy(Qt.StrongFocus)
        self.selected_thumbnail = None
        
    def keyPressEvent(self, event):
        """Handle keyboard events for navigation"""
        if event.key() == Qt.Key_Left:
            self.select_previous()
        elif event.key() == Qt.Key_Right:
            self.select_next()
        else:
            super().keyPressEvent(event)
            
    def select_previous(self):
        """Select the previous thumbnail"""
        if self.current_index > 0:
            self.select_thumbnail(self.current_index - 1)
            
    def select_next(self):
        """Select the next thumbnail"""
        if self.current_index < len(self.thumbnails) - 1:
            self.select_thumbnail(self.current_index + 1)


    def set_directory(self, path):
        """Set the current directory and load images"""
        self.current_directory = path
        self.load_images(path)
        
    def clear(self):
        """Clear all thumbnails and preview"""
        for i in reversed(range(self.thumbnail_layout.count())): 
            widget = self.thumbnail_layout.itemAt(i).widget()
            if widget:
                widget.deleteLater()
        self.thumbnails = []
        self.preview_area.clear()
        self.metadata_label.clear()
        self.current_index = -1
        
    def load_images(self, directory):
        """Load images from directory"""
        self.clear()
        if not os.path.isdir(directory):
            return
            
        image_extensions = ('.png', '.jpg', '.jpeg', '.gif', '.bmp')
        for filename in os.listdir(directory):
            if filename.lower().endswith(image_extensions):
                file_path = os.path.join(directory, filename)
                try:
                    pixmap = QPixmap(file_path)
                    if not pixmap.isNull():
                        self.add_thumbnail(pixmap, file_path, len(self.thumbnails))
                except Exception as e:
                    print(f"Error loading image {filename}: {e}")
                    
    def add_thumbnail(self, pixmap, file_path, index):
        """Add a thumbnail to the scroll area"""
        thumbnail = QLabel()
        thumbnail.setFixedSize(100, 100)
        thumbnail.setAlignment(Qt.AlignCenter)
        thumbnail.setPixmap(pixmap.scaled(96, 96, Qt.KeepAspectRatio, Qt.SmoothTransformation))
        
        # Store reference to original pixmap
        thumbnail.full_pixmap = pixmap
        thumbnail.file_path = file_path
        thumbnail.index = index
        
        # Connect click event
        thumbnail.mousePressEvent = lambda e, idx=index: self.select_thumbnail(idx)
        
        self.thumbnail_layout.addWidget(thumbnail)
        self.thumbnails.append(thumbnail)
        
    def select_thumbnail(self, index):
        """Select thumbnail and show full preview"""
        if not hasattr(self, 'thumbnails') or index < 0 or index >= len(self.thumbnails):
            return
            
        if not self.thumbnails:  # Check if thumbnails list is empty
            return
            
        thumbnail = self.thumbnails[index]
        if not thumbnail:  # Check if thumbnail exists
            return
            
        try:
            self.current_index = index
            
            # Show full preview
            preview_pixmap = thumbnail.full_pixmap.scaled(
                self.preview_area.width(), 
                self.preview_area.height(),
                Qt.KeepAspectRatio,
                Qt.SmoothTransformation
            )
            self.preview_area.setPixmap(preview_pixmap)

            # Update metadata
            self.update_metadata(thumbnail.file_path)

            # Ensure thumbnail is visible
            self.thumbnail_scroll.ensureWidgetVisible(thumbnail)

            # Reset previous selection style if it exists
            if hasattr(self, 'selected_thumbnail') and self.selected_thumbnail:
                try:
                    self.selected_thumbnail.setStyleSheet("")
                except RuntimeError:
                    pass  # Widget was deleted, ignore

            # Highlight selected thumbnail
            try:
                thumbnail.setStyleSheet("""
                    background-color: #0064e1;
                    border-radius: 10px;
                    padding: 4px;
                """)
                self.selected_thumbnail = thumbnail
            except RuntimeError:
                pass  # Widget was deleted, ignore

            # Set focus for keyboard navigation
            self.setFocus()
        except RuntimeError:
            pass  # Handle case where widgets are deleted during operation

            
    def update_metadata(self, file_path):
        """Update metadata panel with file information"""
        if not file_path:
            self.metadata_label.clear()
            return

        file_info = QFileInfo(file_path)
        name = file_info.fileName()
        size = format_size(file_info.size())
        modified = file_info.lastModified().toString("yyyy-MM-dd hh:mm:ss")

        # Get image resolution
        pixmap = QPixmap(file_path)
        if not pixmap.isNull():
            width = pixmap.width()
            height = pixmap.height()
            resolution = f"{width} × {height}"
        else:
            resolution = "Unknown"

        text = f"""
        <b>{name}</b><br>
        Size: {size} | Resolution: {resolution} | Modified: {modified}
        
        """
        self.metadata_label.setText(text)

        



class GalleryItemDelegate(QStyledItemDelegate):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.pixmap_cache = {}
        self.thumbnail_size = QSize(64, 64) 
        self.cell_size = QSize(128, 116)
        self.current_editor = None 

    def paint(self, painter, option, index):
        from PyQt5.QtGui import QPainterPath
        from PyQt5.QtCore import QRectF


        # Get file path - handle both QFileSystemModel and QStandardItemModel
        model = index.model()
        if hasattr(model, 'filePath'):  # QFileSystemModel
            file_path = model.filePath(index)
        else:  # QStandardItemModel
            file_path = index.data(Qt.UserRole)  # Get path from UserRole
            
        # Handle COM objects from Recycle Bin
        if hasattr(file_path, 'Path'):  # This is a COM object from Recycle Bin
            try:
                file_path = file_path.Path  # Get actual path from COM object
            except:
                file_path = ""

        # Get icon - handle both model types
        if hasattr(model, 'fileIcon'):  # QFileSystemModel
            icon = model.fileIcon(index)
        else:  # QStandardItemModel
            icon = index.data(Qt.DecorationRole)  # Get icon from DecorationRole
            if not icon or icon.isNull():
                # Fallback to default file icon if no icon is set
                icon_provider = QFileIconProvider()
                file_info = QFileInfo(file_path) if file_path else QFileInfo()
                icon = icon_provider.icon(file_info)


        view = self.parent()
        if hasattr(view, 'viewport'):
            viewport = view.viewport()
        else:
            viewport = view

        # Set up painter and options
        painter.save()
        painter.setRenderHint(QPainter.Antialiasing)
        painter.setRenderHint(QPainter.SmoothPixmapTransform)
        
        # Draw selection highlight with rounded corners
        if option.state & QStyle.State_Selected:
            path = QPainterPath()
            radius = 12  # Adjust this value to change corner rounding
            rect = option.rect.adjusted(2, 2, -2, -2)  # Slightly smaller than the item rect
            path.addRoundedRect(QRectF(rect), radius, radius)
            painter.fillPath(path, QColor("#0064e1"))
        
        # Calculate content area (centered in cell)
        cell_width = self.cell_size.width()
        cell_height = self.cell_size.height()
        content_rect = QRect(
            option.rect.x() + (option.rect.width() - cell_width) // 2,
            option.rect.bottom() - cell_height,
            cell_width,
            cell_height
        )
        
        # Check if it's an image file
        is_image = file_path.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp', '.webp'))
        
        if is_image:
            # Try to load image
            pixmap = self.pixmap_cache.get(file_path)
            if pixmap is None:
                try:
                    # Load image and create fixed-size thumbnail
                    pixmap = QPixmap(file_path)
                    if not pixmap.isNull():
                        # Scale to fit 128x128 while keeping aspect ratio
                        pixmap = pixmap.scaled(
                            self.thumbnail_size, 
                            Qt.KeepAspectRatio, 
                            Qt.SmoothTransformation,
                        )
                        self.pixmap_cache[file_path] = pixmap
                except:
                    pixmap = None
            
            if pixmap and not pixmap.isNull():
                # Calculate position to center the image in cell
                img_x = content_rect.x() + (content_rect.width() - pixmap.width()) // 2
                
                                
                image_height = pixmap.height()
                text_height = 30
                spacing = 10  # między obrazkiem a tekstem

                total_height = image_height + spacing + text_height
                bottom_margin = 0
                start_y = content_rect.bottom() - total_height - bottom_margin

                img_y = start_y
                img_rect = QRect(img_x, img_y, pixmap.width(), pixmap.height())

                text_rect = QRect(
                    content_rect.x(),
                    img_rect.bottom() + spacing,
                    content_rect.width(),
                    text_height
                )
                
                
                img_rect = QRect(img_x, img_y, pixmap.width(), pixmap.height())
                
                # Draw image
                painter.drawPixmap(img_rect, pixmap)
                
                # Draw filename below image (centered)
                text_rect = QRect(
                    content_rect.x(),
                    img_rect.bottom() + 10,  # 10px margin below image
                    content_rect.width(),
                    30  # Fixed height for text
                )
                
                # Set text color based on selection
                if option.state & QStyle.State_Selected:
                    painter.setPen(option.palette.highlightedText().color())
                else:
                    painter.setPen(option.palette.text().color())
                
                # Draw filename (elided if too long)
                text = index.data(Qt.DisplayRole)
                painter.drawText(text_rect, Qt.AlignBottom | Qt.AlignHCenter | Qt.TextWordWrap, 
                               option.fontMetrics.elidedText(text, Qt.ElideRight, text_rect.width()))
                painter.restore()
                return
        
        # Fallback to default icon if not an image or if image loading failed
        self.draw_default_icon(painter, option, index, content_rect, icon)
        painter.restore()
    
    def draw_default_icon(self, painter, option, index, content_rect, icon):
        # Get icon from parameter (already determined in paint method)
        if icon is None:
            # Fallback if no icon was provided
            icon_provider = QFileIconProvider()
            icon = icon_provider.icon(QFileIconProvider.File)
        
        pixmap = icon.pixmap(64, 64)  # Standard icon size
        
        # Calculate position to center the icon
        icon_x = content_rect.x() + (content_rect.width() - pixmap.width()) // 2
        icon_y = content_rect.y() + (content_rect.height() - pixmap.height()) // 2 - 15
        icon_rect = QRect(icon_x, icon_y, pixmap.width(), pixmap.height())
        
        # Draw icon
        painter.drawPixmap(icon_rect, pixmap)
        
        # Draw filename below icon
        text_rect = QRect(
            content_rect.x(),
            icon_rect.bottom() + 10,
            content_rect.width(),
            30
        )
        
        # Set text color based on selection
        if option.state & QStyle.State_Selected:
            painter.setPen(option.palette.highlightedText().color())
        else:
            painter.setPen(option.palette.text().color())
        
        # Draw filename (elided if too long)
        text = index.data(Qt.DisplayRole)  # Get display text directly
        painter.drawText(text_rect, Qt.AlignBottom | Qt.AlignHCenter | Qt.TextWordWrap, 
                       option.fontMetrics.elidedText(text, Qt.ElideRight, text_rect.width()))
    
    def sizeHint(self, option, index):
        # Fixed size for all cells in grid
        return self.cell_size

    def createEditor(self, parent, option, index):
        # Create editor as a child widget
        self.current_editor = QLineEdit(parent)
        self.current_editor.setFixedWidth(128)
        self.current_editor.setFixedHeight(36)
        self.current_editor.setStyleSheet("""
            QLineEdit {
                padding: 0px 4px;
                padding-bottom: 4px;
                margin-left: 1px;
                margin-right: 3px;
                font-size: 26px;
            }
        """)
        self.current_editor.setObjectName("rename")
        return self.current_editor

    def updateEditorGeometry(self, editor, option, index):
        # Calculate position - align left in the text label area
        cell_width = self.cell_size.width()
        cell_height = self.cell_size.height()
        
        # Calculate text area position (below image)
        text_x = option.rect.x() + (option.rect.width() - cell_width) // 2
        text_y = option.rect.y() + 10 + 64 + 10  # Top margin + image height + bottom margin
        
        editor_width = cell_width - 20  # Leave some padding
        editor_height = 30
        
        # Align editor to the left (with small left margin)
        editor_x = text_x + 1  # 10px left margin instead of centering
        editor_y = text_y
        
        # Get the viewport from the parent QListView
        list_view = self.parent()  # This should be the QListView
        if hasattr(list_view, 'viewport'):
            viewport = list_view.viewport()
            # Ensure it stays within viewport bounds
            editor_x = max(0, min(editor_x, viewport.width() - editor_width))
            editor_y = max(0, min(editor_y, viewport.height() - editor_height))
        
        editor.setGeometry(editor_x, editor_y, editor_width, editor_height)
        editor.raise_()

    def setEditorData(self, editor, index):
        full_name = index.data()
        model = index.model()
        file_path = model.filePath(index)
        
        if full_name.startswith("↗ "):
            display_name = full_name[2:]
        else:
            display_name = full_name
            
        if '.' in display_name and not os.path.isdir(file_path):
            base_name = display_name.rsplit('.', 1)[0]
            editor.setText(base_name)
        else:
            editor.setText(display_name)
        editor.selectAll()

    def setModelData(self, editor, model, index):
        old_full_name = index.data()
        new_base_name = editor.text().strip()
        
        if not new_base_name or new_base_name == old_full_name:
            return
            
        file_path = model.filePath(index)
        if not file_path:
            return
            
        if '.' in old_full_name and not os.path.isdir(file_path):
            extension = old_full_name.rsplit('.', 1)[1]
            new_name = f"{new_base_name}.{extension}"
        else:
            new_name = f"{new_base_name}"
            
        base_path = os.path.dirname(file_path)
        new_path = os.path.join(base_path, new_name)
        
        try:
            os.rename(file_path, new_path)
        except Exception as e:            
            msg = self.create_message_box()
            msg.setIcon(QMessageBox.Warning)
            msg.setWindowTitle("Error")
             
            msg.setText(f"Could not rename file:\n{e}")
            ok_button = msg.addButton("OK", QMessageBox.AcceptRole)
            ok_button.setMinimumWidth(120)
            msg.exec_()
            return False

        model.setRootPath(model.rootPath())
        editor.close()  # Zamknij edytor po zapisaniu
        
    def destroyEditor(self, editor, index):
        if self.current_editor == editor:
            self.current_editor = None
        super().destroyEditor(editor, index)


class CopySignals(QObject):
    progress = pyqtSignal(int)
    finished = pyqtSignal()
    error = pyqtSignal(str)

class CopyWorker(QThread):
    def __init__(self, src, dst, operation):
        super().__init__()
        self.src = src
        self.dst = dst
        self.operation = operation
        self.signals = CopySignals()
        self._is_running = True

    def run(self):
        try:
            if os.path.isdir(self.src):
                self.copy_folder(self.src, self.dst)
            else:
                self.copy_file(self.src, self.dst)
            self.signals.finished.emit()
        except Exception as e:
            self.signals.error.emit(str(e))

    def copy_file(self, src, dst):
        import shutil
        # Check if destination exists and handle accordingly
        if os.path.exists(dst):
            # Remove existing file if it exists (user already confirmed)
            try:
                if os.path.isdir(dst):
                    shutil.rmtree(dst)
                else:
                    os.remove(dst)
            except Exception as e:
                self.signals.error.emit(f"Could not remove existing file: {str(e)}")
                return
        
        total_size = os.path.getsize(src)
        copied = 0
        
        with open(src, 'rb') as fsrc:
            with open(dst, 'wb') as fdst:
                while self._is_running:
                    buf = fsrc.read(1024*1024)  # 1MB buffer
                    if not buf:
                        break
                    fdst.write(buf)
                    copied += len(buf)
                    progress = int((copied / total_size) * 100)
                    self.signals.progress.emit(progress)

    def copy_folder(self, src, dst):
        if not os.path.exists(dst):
            os.makedirs(dst)
        
        items = []
        for root, dirs, files in os.walk(src):
            for name in files:
                items.append(os.path.join(root, name))
            for name in dirs:
                items.append(os.path.join(root, name))
        
        total = len(items)
        for i, item in enumerate(items):
            if not self._is_running:
                break
                
            rel_path = os.path.relpath(item, src)
            dest_path = os.path.join(dst, rel_path)
            
            if os.path.isdir(item):
                if not os.path.exists(dest_path):
                    os.makedirs(dest_path)
            else:
                self.copy_file(item, dest_path)
            
            progress = int((i + 1) / total * 100)
            self.signals.progress.emit(progress)

    def stop(self):
        self._is_running = False


class GlobalClipboard:
    _instance = None
    
    def __new__(cls):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
            cls._instance.files = []
            cls._instance.operation = None
        return cls._instance
        
class RenameDelegate(QStyledItemDelegate):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.current_editor = None
        self.parent_widget = parent  # Store reference to parent widget

    def createEditor(self, parent, option, index):
        self.current_editor = QLineEdit(parent)
        self.current_editor.setObjectName("rename")
        
        if hasattr(self.parent_widget, 'stack'):
            current_index = self.parent_widget.stack.currentIndex()
        
        if current_index == 0:
            self.current_editor.setStyleSheet("border-radius: 0px; border-bottom-left-radius: 12px; border-bottom-right-radius: 12px; padding: 0px 4px;")
        if current_index == 2:
            self.current_editor.setStyleSheet("border-radius: 0px; margin-left: 38px; padding: 0px 1px;") 
            self.current_editor.setMaximumWidth(303)
        else:    
            self.current_editor.setStyleSheet("border-radius: 0px; padding: 0px 1px;")  
            
        self.current_editor.installEventFilter(self)
        return self.current_editor

    def eventFilter(self, editor, event):
        from PyQt5.QtCore import QEvent
        if event.type() == QEvent.KeyPress and event.key() in (Qt.Key_Return, Qt.Key_Enter):
            self.commitData.emit(editor)
            self.closeEditor.emit(editor, QStyledItemDelegate.SubmitModelCache)
            return True
        return super().eventFilter(editor, event)

    def setEditorData(self, editor, index):
        full_name = index.data()
        model = index.model()
        
        # Handle QFileSystemModel
        if hasattr(model, 'filePath'):
            file_path = model.filePath(index)
        # Handle QStandardItemModel
        else:
            file_path = index.data(Qt.UserRole)  # Get path from UserRole
        
        if not file_path:
            editor.setText(full_name)
            editor.selectAll()
            return
            
        # Remove arrow "↗" from the beginning of the filename (for .lnk shortcuts)
        display_name = full_name
        if full_name.startswith("↗ "):
            display_name = full_name[2:]
            
        # Edit only the name without extension for files
        if '.' in display_name and not os.path.isdir(file_path):
            base_name = display_name.rsplit('.', 1)[0]
            editor.setText(base_name)
        else:
            editor.setText(display_name)
        editor.selectAll()

    def setModelData(self, editor, model, index):
        old_full_name = index.data()
        new_base_name = editor.text().strip()
        
        if not new_base_name or new_base_name == old_full_name:
            return None
            
        # Get file path based on model type
        if hasattr(model, 'filePath'):
            file_path = model.filePath(index)
        else:
            file_path = index.data(Qt.UserRole)
            
        if not file_path:
            return None
            
        # Preserve original extension for files
        if '.' in old_full_name and not os.path.isdir(file_path):
            extension = old_full_name.rsplit('.', 1)[1]
            new_name = f"{new_base_name}.{extension}"
        else:
            new_name = f"{new_base_name}"
            
        base_path = os.path.dirname(file_path)
        new_path = os.path.join(base_path, new_name)
        
        try:
            # Check if destination already exists
            if os.path.exists(new_path):
                if hasattr(self.parent_widget, 'create_message_box'):
                    msg = self.parent_widget.create_message_box()
                    msg.setIcon(QMessageBox.Warning)
                    msg.setWindowTitle("Error")
                     
                    msg.setText(f"A file or folder with the name '{new_name}' already exists.")
                    ok_button = msg.addButton("OK", QMessageBox.AcceptRole)
                    ok_button.setMinimumWidth(120)
                    msg.exec_()
                return None
                
            os.rename(file_path, new_path)
        except Exception as e:
            if hasattr(self.parent_widget, 'create_message_box'):
                msg = self.parent_widget.create_message_box()
                msg.setIcon(QMessageBox.Warning)
                msg.setWindowTitle("Error")
                 
                msg.setText(f"Could not rename file:\n{e}")
                ok_button = msg.addButton("OK", QMessageBox.AcceptRole)
                ok_button.setMinimumWidth(120)
                msg.exec_()
            return None

        # Refresh the model
        if hasattr(model, 'setRootPath'):
            model.setRootPath(model.rootPath())
        
        return None

    def updateEditorGeometry(self, editor, option, index):
        # Special handling for column view
        if hasattr(self.parent_widget, 'stack') and self.parent_widget.stack.currentIndex() == 2:
            # For column view, align editor with the text
            rect = option.rect
            editor.setGeometry(rect.x(), rect.y(), rect.width(), rect.height())
        else:
            # Default behavior for other views
            super().updateEditorGeometry(editor, option, index)

def get_file_system(drive_path):
    for part in psutil.disk_partitions(all=False):
        if part.mountpoint == drive_path or drive_path.startswith(part.mountpoint):
            return part.fstype
    return "Unknown"



if __name__ == "__main__":    
    if sys.platform == 'win32':
        # For Windows, set DPI awareness and process priority
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
            ctypes.windll.kernel32.SetPriorityClass(ctypes.windll.kernel32.GetCurrentProcess(), 0x00008000)  # HIGH_PRIORITY_CLASS
        except:
            pass        
    app = QApplication(sys.argv)

    path = None
    show_recycle_bin_directly = False
    show_thispc_directly = False

    if len(sys.argv) > 1:
        combined_args = ' '.join(sys.argv[1:]).strip()

        # Check if the argument starts with '/O ' and extract the path accordingly
        if combined_args.startswith('/O '):
            path_candidate = combined_args[3:].strip('"')
        else:
            path_candidate = combined_args.strip('"')

        lower_candidate = path_candidate.lower()

        if lower_candidate == 'trash':
            show_recycle_bin_directly = True
        elif lower_candidate == 'thispc':
            show_thispc_directly = True
        else:
            path = os.path.normpath(path_candidate)

    # CHeck if an instance of file_manager is already running
    existing_file_manager = next((w for w in app.topLevelWidgets() if isinstance(w, file_manager)), None)

    if existing_file_manager is None:
        # New instance
        file_manager = file_manager(path if not show_recycle_bin_directly and not show_thispc_directly else None)
        file_manager.show()

        # Handle command-line arguments for special cases -> shortcuts like "trash" or "thispc"
        if show_recycle_bin_directly:
            file_manager.show_recycle_bin()
        elif show_thispc_directly:
            file_manager.show_thispc()
    else:
        if show_recycle_bin_directly:
            existing_file_manager.show_recycle_bin()
        elif show_thispc_directly:
            existing_file_manager.show_thispc()
        elif path and os.path.exists(path):
            existing_file_manager.set_directory(path)

        existing_file_manager.raise_()
        existing_file_manager.activateWindow()

    sys.exit(app.exec_())