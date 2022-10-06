''' NVIDIA Instant Replay auto-cutter 1/11/22 '''
from __future__ import annotations  # 0.10mb  / 13.45       https://stackoverflow.com/questions/33533148/how-do-i-type-hint-a-method-with-the-type-of-the-enclosing-class
import tracemalloc                  # 0.125mb / 13.475mb
import gc                           # 0.275mb / 13.625mb    <- Heavy, but probably worth it
import os                           # 0.10mb  / 13.45mb
import sys                          # 0.125mb / 13.475mb
import time                         # 0.125mb / 13.475mb
import subprocess                   # 0.125mb / 13.475mb
import shutil                       # 0.125mb / 13.475mb
import logging                      # 0.65mb  / 14.00mb

import keyboard                     # 2.05mb  / 15.40mb     <- TODO Find lighter alternative?
import pymediainfo                  # 3.75mb  / 17.1mb      https://stackoverflow.com/questions/15041103/get-total-length-of-videos-in-a-particular-directory-in-python
import winsound                     # 0.21mb  / 13.56mb     ^ https://stackoverflow.com/questions/3844430/how-to-get-the-duration-of-a-video-in-python
import pystray                      # 3.29mb  / 16.64mb
from PIL import Image               # 2.13mb  / 15.48mb
from datetime import datetime       # 0.125mb / 13.475mb
from traceback import format_exc    # 0.35mb  / 13.70mb     <- Heavy, but probably worth it
from threading import Thread        # 0.125mb / 13.475mb

import ctypes
from pystray._util import win32
from win32_setctime import setctime
tracemalloc.start()                 # start recording memory usage AFTER libraries have been imported

# Starts with roughly ~36.7mb of memory usage. Roughly 9.78mb combined from imports alone, without psutil and cv2/pymediainfo (9.63mb w/o tracemalloc).
# Detecting shadowplay videos via their encoding is possible (but useless) https://github.com/rebane2001/NvidiaInstantRename/blob/mane/InstantRenameCLI.py

'''
TODO extended backup system with more than 1 undo possible at a time
TODO add deletion "confirmation"? (press delete hotkey twice? add delete submenu for the tray?)
TODO show what "action" you can undo in the menu in some way (as a submenu?)
TODO add dedicated config file, possibly separate file for defining tray menu
TODO ability to auto-rename folders using same system as aliases
TODO cropping ability -> pick crop region after saving instant replay OR before starting recording
        - have ability to pick region before (?) and after recording
        - this should be very lightweight and simple -> possible idea:
            - use ffmpeg to extract a frame from a given clip (if done after recording)
            - display frame fullscreen, then allow user to draw a region over the frame
TODO pystray subclass improvements
        - add dynamic tooltips? (when hovering over icon)
        - add double-click support? https://learn.microsoft.com/en-us/windows/win32/inputdev/wm-lbuttondblclk
            - left-clicks still get registered
            - pystray's _mainloop is blocking
            - an entire thread dedicated to handling double-clicks would likely be needed (unreasonable)
FIXME: `INSTANT_REPLAY_HOTKEY` suddenly stops working with no error (has only happened once)
FIXME: Alt + Arrow Keys emits number shortcuts (is it just my keyboard?)
        - Alt + Up    = Alt + 8
        - Alt + Down  = Alt + 2
        - Alt + Left  = Alt + 4
        - Alt + Right = Alt + 6
'''

# ---------------------
# Settings
# ---------------------
AUDIO = True
RENAME = True
RENAME_FORMAT = '?game ?date #?count'
RENAME_DATE_FORMAT = '%y.%m.%d'     # https://strftime.org/
RENAME_COUNT_START_NUMBER = 1
RENAME_COUNT_PADDED_ZEROS = 0

CHECK_FOR_NEW_CLIPS_ON_LAUNCH = True
SEND_DELETED_FILES_TO_RECYCLE_BIN = True
MAX_BACKUPS = 5


# --- Paths ---
LOG_PATH = None
BACKUP_DIR = 'Backups'
RESOURCE_DIR = 'resources'
HISTORY_PATH = 'history.txt'
UNDO_LIST_PATH = 'undo.txt'
ICON_PATH = 'icon.ico'

# NOTE: These only apply if the associated path above is not absolute.
SAVE_LOG_FILE_TO_APPDATA_FOLDER = False
SAVE_HISTORY_TO_APPDATA_FOLDER = False
SAVE_UNDO_LIST_TO_APPDATA_FOLDER = False
SAVE_BACKUPS_TO_APPDATA_FOLDER = False
SAVE_BACKUPS_TO_VIDEO_FOLDER = False

# NOTE: Subfolders only, i.e. ['Other', 'Movies']. Not case-sensitive.
IGNORE_VIDEOS_IN_THESE_FOLDERS = []


# --- Hotkeys ---
CONCATENATE_HOTKEY = 'alt + c'
DELETE_HOTKEY = 'ctrl + alt + d'
UNDO_HOTKEY = 'alt + u'
LENGTH_HOTKEY = 'alt'
LENGTH_DICTIONARY = {
    '1': 10,
    '2': 20,
    '3': 30,
    '4': 40,
    '5': 50,
    '6': 60,
    '7': 70,
    '8': 80,
    '9': 90
}


# --- Tray Menu ---
'''
--- CUSTOM TRAY MENU TUTORIAL ---
`TRAY_ADVANCED_MODE_MENU` defines the custom tray you wish to use. It's a list
(or tuple) of items. Items may be strings which directly insert an item into
the menu, or a single-key dictionary with a custom name as the key (string),
and the item you wish to insert and rename as the value (string). A dictionary
with a list as the value will become a submenu, with the key being the title.
Submenus work just like the base menu, and can be nested indefinitely.

`TRAY_LEFT_CLICK_ACTION` and `TRAY_MIDDLE_CLICK_ACTION` define
a single normal tray item to activate upon that specific click.

Normal tray items:
    'open_log':             Opens this program's log file.
    'open_video_folder':    Opens the currently defined "Videos" folder.
    'open_install_folder':  Opens this program's root folder.
    'play_most_recent':     Plays your most recent clip.
    'explore_most_recent':  Opens your most recent clip in Explorer.
    'delete_most_recent':   Deletes your most recent clip.
    'concatenate_last_two': Concatenates your two most recent clips.
    'undo':                 Undoes the last trim or concatenation.
    'clear_history':        Clears your clip history.
    'update':               Manually checks for new clips and refreshes existing ones.
    'quit':                 Exits this program.

Special tray items:
    'separator':            Adds a separator in the menu.
                                -Cannot be renamed.
    'memory':               Displays current RAM usage.
                                -This is somewhat misleading and not worth using.
                                -Use '?memory' in the title to represent where the number will be:
                                    {'RAM: ?memory': 'memory'}
                                -This item will be greyed out and is informational only.
    'recent_clips':         Displays your most recent clips.
                                -Cannot be renamed.
                                -Renaming this item will simply place it within a submenu:
                                    {'Recent clips': 'recent_clips'}
                                -See "Recent clip menu settings" below for lots of customization.

---

Submenu example:
    {'Quick actions': [
        {'Play most recent clip': 'play_most_recent'},
        {'View last clip in explorer': 'explore_most_recent'},
        {'Concatenate last two clips': 'concatenate_last_two'},
        {'Delete most recent clip': 'delete_most_recent'},
    ]}
'''

TRAY_LEFT_CLICK_ACTION = 'open_video_folder'
TRAY_MIDDLE_CLICK_ACTION = 'play_most_recent'
TRAY_ADVANCED_MODE = True
TRAY_ADVANCED_MODE_MENU = (
    {'View log': 'open_log'},
    {'View videos': 'open_video_folder'},
    {'View root': 'open_install_folder'},
    'separator',
    {'Play last clip': 'play_most_recent'},
    {'Explore last clip': 'explore_most_recent'},
    {'Splice last clips': 'concatenate_last_two'},
    {'Delete last clip': 'delete_most_recent'},
    {'Undo last action': 'undo'},
    'separator',
    'recent_clips',
    'separator',
    {'Update clips': 'update'},
    {'Clear history': 'clear_history'},
    {'Exit': 'quit'},
)


# --- Basic mode (TRAY_ADVANCED_MODE = False) only ---
TRAY_SHOW_QUICK_ACTIONS = True
TRAY_RECENT_CLIPS_IN_SUBMENU = False
TRAY_QUICK_ACTIONS_IN_SUBMENU = True


# --- Recent clip menu settings ---
TRAY_RECENT_CLIP_COUNT = 5
TRAY_RECENT_CLIPS_HAVE_UNIQUE_SUBMENUS = True
TRAY_RECENT_CLIPS_SUBMENU_EXTRA_INFO = True    # TODO have auto_update turn on or off based on these settings
TRAY_EXTRA_INFO_DATE_FORMAT = '%a %#D %#I:%M:%S%p'
TRAY_CLIPS_PLAY_ON_CLICK = True
''' ?date - "1/17/22 12:09am" (see TRAY_RECENT_CLIP_DATE_FORMAT)
    ?recency - "2 days ago"
    ?recencyshort - "2d"
    ?size - "244.1mb"
    ?length - "1:30" <- 90 seconds
    ?clip - Name of a clip only.
    ?clipdir - Name and immediate parent directory only.
    ?clippath - Full path to a clip. '''
TRAY_RECENT_CLIP_NAME_FORMAT = '(?recencyshort) - ?clip'
TRAY_RECENT_CLIP_DATE_FORMAT = '%#I:%M%p'
TRAY_RECENT_CLIP_DEFAULT_TEXT = ' --'


# --- Game aliases ---
GAME_ALIASES = {    # NOTE: the game titles must be lowercase and have no double-spaces
    "left 4 dead": "L4D1",
    "left 4 dead 2": "L4D2",
    "battlefield 4": "BF4",
    "dead by daylight": "DBD",
    "counter-strike global offensive": "CSGO",
    "the binding of isaac rebirth": "TBOI",
    "team fortress 2": "TF2",
    "tom clancy's rainbow six siege": "R6"
}


# --- Registry setting overrides ---
VIDEO_PATH_OVERRIDE = ''
INSTANT_REPLAY_HOTKEY_OVERRIDE = ''
TRAY_ALIGN_CENTER = False


# ---------------------
# Aliases
# ---------------------
pjoin = os.path.join
exists = os.path.exists
getstat = os.stat
getsize = os.path.getsize
basename = os.path.basename
dirname = os.path.dirname
abspath = os.path.abspath
splitext = os.path.splitext
splitdrive = os.path.splitdrive


# ---------------------
# Constants & paths
# ---------------------
CLIP_BUFFER = max(5, TRAY_RECENT_CLIP_COUNT)
CWD = dirname(os.path.realpath(__file__))
os.chdir(CWD)

APPDATA_PATH = pjoin(os.path.expandvars('%LOCALAPPDATA%'), 'Instant Replay Suite')
RESOURCE_DIR = abspath(RESOURCE_DIR)

if splitdrive(ICON_PATH)[0]: ICON_PATH = abspath(ICON_PATH)
else: ICON_PATH = pjoin(RESOURCE_DIR if exists(RESOURCE_DIR) else CWD, 'icon.ico')
if splitdrive(HISTORY_PATH)[0]: HISTORY_PATH = abspath(HISTORY_PATH)
else: HISTORY_PATH = pjoin(APPDATA_PATH if SAVE_HISTORY_TO_APPDATA_FOLDER else CWD, HISTORY_PATH)
if splitdrive(UNDO_LIST_PATH)[0]: UNDO_LIST_PATH = abspath(UNDO_LIST_PATH)
else: UNDO_LIST_PATH = pjoin(APPDATA_PATH if SAVE_UNDO_LIST_TO_APPDATA_FOLDER else CWD, UNDO_LIST_PATH)
if not LOG_PATH: LOG_PATH = basename(__file__.replace('.pyw', '.log').replace('.py', '.log'))
if splitdrive(LOG_PATH)[0]: LOG_PATH = abspath(LOG_PATH)
else: LOG_PATH = pjoin(APPDATA_PATH if SAVE_LOG_FILE_TO_APPDATA_FOLDER else CWD, LOG_PATH)

assert exists(ICON_PATH), f'No icon exists at {ICON_PATH}!'
if not exists(dirname(HISTORY_PATH)): os.makedirs(dirname(HISTORY_PATH))
if not exists(dirname(UNDO_LIST_PATH)): os.makedirs(dirname(UNDO_LIST_PATH))
if not exists(dirname(LOG_PATH)): os.makedirs(dirname(LOG_PATH))

if isinstance(IGNORE_VIDEOS_IN_THESE_FOLDERS, str): IGNORE_VIDEOS_IN_THESE_FOLDERS = (IGNORE_VIDEOS_IN_THESE_FOLDERS,)
IGNORE_VIDEOS_IN_THESE_FOLDERS = tuple(path.strip().lower() for path in IGNORE_VIDEOS_IN_THESE_FOLDERS)


# ---------------------
# Logging
# ---------------------
logging.basicConfig(
    level=logging.INFO,
    encoding='utf-16',
    format='{asctime} {lineno:<3} {levelname} {funcName}: {message}',
    datefmt='%I:%M:%S%p',
    style='{',
    handlers=(logging.FileHandler(LOG_PATH, 'w', delay=False),
              logging.StreamHandler()))


# ---------------------
# Registry Settings
# ---------------------
# get ShadowPlay video path from registry
if not VIDEO_PATH_OVERRIDE:
    try:
        import winreg   # NOTE: ShadowPlay settings are encoded in utf-16 and have a NULL character at the end
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r'SOFTWARE\NVIDIA Corporation\Global\ShadowPlay\NVSPCAPS')
        VIDEO_PATH = winreg.QueryValueEx(key, 'DefaultPathW')[0].decode('utf-16')[:-1]
    except:
        logging.critical(f'Could not find video path from registry: {format_exc()}\n\nPlease set VIDEO_PATH_OVERRIDE.')
        sys.exit(2)
else: VIDEO_PATH = VIDEO_PATH_OVERRIDE.strip()
logging.info('Video path: ' + VIDEO_PATH)

# get Instant Replay hotkey from registry (each key is a separate value)
if not INSTANT_REPLAY_HOTKEY_OVERRIDE:
    try:
        import winreg   # NOTE: ShadowPlay settings are encoded in utf-16 and have a NULL character at the end
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r'SOFTWARE\NVIDIA Corporation\Global\ShadowPlay\NVSPCAPS')

        # I have no idea how to actually decode numbers, modifiers, and function keys, so I cheated for all 3
        # NOTE: ctrl and alt do not have left/right counterparts
        # NOTE: number is plainly visible inside the encoded string, so we can just pluck it out
        modifier_keys = {'\t': 'tab', '': 'shift', '': 'ctrl', '': 'alt', '': 'capslock'}
        f_keys = {'P': 'f1', 'Q': 'f2', 'R': 'f3', 'S': 'f4', 'T': 'f5', 'U': 'f6',
                  'V': 'f7', 'W': 'f8', 'X': 'f9', 'Y': 'f10', 'Z': 'f11', '{': 'f12'}  # TODO: Function keys beyond F12
        total_keys_encoded_string = winreg.QueryValueEx(key, 'DVRHKeyCount')[0]
        total_keys = int(str(total_keys_encoded_string)[5])

        hotkey = []
        for key_number in range(total_keys):
            hotkey_part = winreg.QueryValueEx(key, f'DVRHKey{key_number}')[0].decode('utf-16')[:-1]
            if hotkey_part in modifier_keys: hotkey.append(modifier_keys[hotkey_part])
            elif not hotkey_part.isupper(): hotkey.append(f_keys[hotkey_part])
            else: hotkey.append(hotkey_part)

        INSTANT_REPLAY_HOTKEY = ' + '.join(hotkey)
    except:
        logging.critical(f'Could not detect Instant-Replay hotkey from registry: {format_exc()}\n\nPlease set INSTANT_REPLAY_HOTKEY_OVERRIDE.')
        sys.exit(3)
else: INSTANT_REPLAY_HOTKEY = INSTANT_REPLAY_HOTKEY_OVERRIDE.strip().lower()
logging.info(f'Instant replay hotkey: "{INSTANT_REPLAY_HOTKEY}"')

# get taskbar position from registry (NOTE: 0 = left, 1 = top, 2 = right, 3 = bottom)
if TRAY_ALIGN_CENTER: MENU_ALIGNMENT = win32.TPM_CENTERALIGN | win32.TPM_TOPALIGN
else:
    try:    # NOTE: this value takes a few moments to update after moving the taskbar (if you're testing this)
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\StuckRects3')
        taskbar_position = winreg.QueryValueEx(key, 'Settings')[0][12]

        # top-right alignment adjusts itself automatically for all EXCEPT left taskbars
        if taskbar_position == 0: MENU_ALIGNMENT = win32.TPM_LEFTALIGN | win32.TPM_BOTTOMALIGN
        else: MENU_ALIGNMENT = win32.TPM_RIGHTALIGN | win32.TPM_TOPALIGN
    except: logging.warning(f'Could not detect taskbar position for menu-alignment: {format_exc()}')
logging.info(f'Menu alignment: {MENU_ALIGNMENT}')


# ---------------------
# Backup dir cleanup
# ---------------------
if exists(BACKUP_DIR): BACKUP_DIR = abspath(BACKUP_DIR)
elif SAVE_BACKUPS_TO_VIDEO_FOLDER: BACKUP_DIR = pjoin(VIDEO_PATH, BACKUP_DIR)
elif SAVE_BACKUPS_TO_APPDATA_FOLDER: BACKUP_DIR = pjoin(APPDATA_PATH, BACKUP_DIR)
else: BACKUP_DIR = pjoin(CWD, BACKUP_DIR)

# VIDEO_PATH and BACKUP_DIR must be on the same drive or we'll get OSError 17
if (splitdrive(VIDEO_PATH)[0] != splitdrive(BACKUP_DIR)[0] or
    os.path.ismount(VIDEO_PATH) != os.path.ismount(BACKUP_DIR)):
    msg = ("Your video folder and the path for saving temporary "
           "backups are not on the same drive. Instant Replay Suite "
           "cannot backup and restore videos across drives without "
           "copying them back and forth.\n\nVideo folder: "
           f"{VIDEO_PATH}\nBackup folder: {BACKUP_DIR}\n\nPlease "
           "set `SAVE_BACKUPS_TO_VIDEO_FOLDER` or specify an "
           "absolute path for `BACKUP_DIR` on the same drive.")
    MessageBox = ctypes.windll.user32.MessageBoxW   # flags are !-symbol + stay on top
    MessageBox(None, msg, 'Invalid Backup Directory', 0x00040030)
    logging.error(msg)
    exit(17)

if not exists(BACKUP_DIR): os.makedirs(BACKUP_DIR)


# ---------------------
# Utility functions
# ---------------------
#get_memory = lambda: psutil.Process().memory_info().rss / (1024 * 1024)
get_memory = lambda: tracemalloc.get_traced_memory()[0] / 1048576

def verify_ffmpeg():
    if exists('ffmpeg.exe'): return
    else:
        for path in os.environ.get('PATH', '').split(';'):
            try:
                if 'ffmpeg.exe' in os.listdir(path):
                    return
            except: pass

    msg = ("FFmpeg was not detected. FFmpeg is required for all of this "
           "program's editing features. Please ensure `ffmpeg.exe` is "
           "either in your PATH or in this program's install folder.\n\n"
           "You can download FFmpeg for Windows here (not clickable, sorry): "
           "https://www.gyan.dev/ffmpeg/builds/")
    MessageBox = ctypes.windll.user32.MessageBoxW   # flags are warning symbol + stay on top
    MessageBox(None, msg, 'Instant Replay Suite — FFmpeg not detected', 0x00040010)
    raise FileNotFoundError('''\n\n---\n
FFmpeg was not detected. FFmpeg is required for all of Instant Replay Suite's editing features.

Please ensure `ffmpeg.exe` is either in your PATH or in this program's install folder.
You can download FFmpeg for Windows from here: https://www.gyan.dev/ffmpeg/builds/\n\n---''')


def quit_tray(icon):
    logging.info('Closing system tray icon and exiting suite.')
    icon.visible = False
    icon.stop()
    tracemalloc.stop()
    logging.info(f'Clip history: {cutter.last_clips}')
    with open(HISTORY_PATH, 'w') as history: history.write('\n'.join(c.path if isinstance(c, Clip) else c for c in cutter.last_clips))
    try: sys.exit(0)
    except SystemExit: pass         # avoid harmless yet annoying SystemExit error


def delete(path: str):
    ''' Deletes a clip at the given `path`. '''
    logging.info('Deleting clip: ' + path)
    try:
        if SEND_DELETED_FILES_TO_RECYCLE_BIN: send2trash.send2trash(path)
        else: os.remove(path)
    except:
        logging.error(f'(!) Error while deleting file {path}: {format_exc()}')
        play_alert('error')


def play_alert(sound: str) -> bool:
    if AUDIO:
        path = pjoin(RESOURCE_DIR, f'{sound}.wav')
        logging.info('Playing alert: ' + path)
        try: winsound.PlaySound(path, winsound.SND_ASYNC)
        except:
            winsound.MessageBeep(winsound.MB_ICONHAND)      # play OS error sound
            if sound != 'error':    # intentional error sound (and there's no custom error.wav) -> ignore missing file error
                logging.error(f'(!) Error while playing sound {path}: {format_exc()}')
                return False
    return True


def refresh_backups(*paths):
    ''' Deletes outdated backup files. Ignores `paths` while counting,
        even if `MAX_BACKUPS` is 0. Assumes all backups use the format
        "{time.time_ns()}*.mp4". '''
    old_backups = 0
    for filename in reversed(os.listdir(BACKUP_DIR)):
        if filename[-4:] != '.txt' and filename[:19].isnumeric():
            path = pjoin(BACKUP_DIR, filename)
            if path not in paths:
                old_backups += 1
                if old_backups >= MAX_BACKUPS:
                    try: os.remove(path)
                    except: logging.warning(f'Failed to delete outdated backup "{path}": {format_exc()}')


def ffmpeg(out, cmd):
    temp_path = f'{out[:-4]}_temp.mp4'
    shutil.copy2(out, temp_path)
    os.remove(out)
    cmd = f'ffmpeg -y {cmd.replace("%tp", temp_path)} -hide_banner -loglevel warning'
    logging.info(f'Performing ffmpeg operation: {cmd}')
    process = subprocess.Popen(cmd, shell=True)
    process.wait()
    os.remove(temp_path)
    logging.info('ffmpeg operation successful.')


def trim_off_start_in_place(clip: Clip, length: float):
    clip_path = clip.path
    assert exists(clip_path), f'Path {clip_path} does not exist!'
    logging.info(f'Trimming clip {clip.name} to {length} seconds')
    logging.info(f'Clip is {clip.length:.2f} seconds long.')
    if clip.length <= length: return logging.info(f'(?) Video is only {clip.length:.2f} seconds long and cannot be trimmed to {length} seconds.')

    ext = splitext(clip_path)[-1]
    temp_path = pjoin(BACKUP_DIR, f'{time.time_ns()}{ext}')
    os.renames(clip_path, temp_path)

    try:
        #ffmpeg(clip_path, f'-i "%tp" -ss {clip.length - length} -c:v copy -c:a copy "{clip_path}"')
        cmd = f'ffmpeg -y -i "{temp_path}" -ss {clip.length - length} -c:v copy -c:a copy "{clip_path}" -hide_banner -loglevel warning'
        logging.info(cmd)
        process = subprocess.Popen(cmd, shell=True)
        process.wait()
        with open(UNDO_LIST_PATH, 'w') as undo: undo.write(f'{basename(temp_path)} -> {clip_path} -> Trimmed to {length:g} seconds\n')
        refresh_backups(temp_path)
        setctime(clip_path, getstat(temp_path).st_ctime)    # ensure edited clip retains original creation time
        logging.info(f'Trim to {length} seconds successful.\n')
    except:
        logging.error(f'(!) Error while trimming clip: {format_exc()}')
        if exists(clip_path): os.remove(clip_path)
        os.rename(temp_path, clip_path)
        logging.info('Successfully restored clip after error.')


def get_video_duration(path: str) -> float:     # ? -> https://stackoverflow.com/questions/10075176/python-native-library-to-read-metadata-from-videos
    for track in pymediainfo.MediaInfo.parse(path).tracks:
        if track.track_type == "Video":
            return track.duration / 1000
    return 0


# ---------------------
# Custom Pystray class
# ---------------------
WM_MBUTTONUP = 0x0208
class Icon(pystray._win32.Icon):
    ''' This subclass auto-updates the menu before opening,
        allowing dynamic titles/actions to always be up-to-date,
        and adds support for middle-clicks and custom menu alignments.
        Full comments can be found in the original _win32.Icon class. '''
    def _on_notify(self, wparam, lparam):
        if lparam == win32.WM_LBUTTONUP:
            self()

        elif lparam == WM_MBUTTONUP:
            if MIDDLE_CLICK_ACTION:
                MIDDLE_CLICK_ACTION()

        elif self._menu_handle and lparam == win32.WM_RBUTTONUP:
            self._update_menu()

            win32.SetForegroundWindow(self._hwnd)
            point = ctypes.wintypes.POINT()
            win32.GetCursorPos(ctypes.byref(point))
            hmenu, descriptors = self._menu_handle

            index = win32.TrackPopupMenuEx(
                hmenu,
                MENU_ALIGNMENT | win32.TPM_RETURNCMD,
                point.x,
                point.y,
                self._menu_hwnd,
                None
            )

            if index > 0:
                descriptors[index - 1](self)


# ---------------------
# Clip class
# ---------------------
class Clip:
    __slots__ = ('working', 'path', 'name', 'time', 'size', 'date', 'full_date', 'length', 'length_string', 'length_size_string')
    def __repr__(self): return self.name

    def __init__(self, path, stat, rename=False):
        self.working = False
        self.path, self.name = self.rename(path) if rename else (abspath(path), basename(path))   # abspath for consistent formatting
        self.time = stat.st_ctime
        self.size = f'{(stat.st_size / 1048576):.1f}mb'
        self.date = datetime.strftime(datetime.fromtimestamp(stat.st_ctime), TRAY_RECENT_CLIP_DATE_FORMAT)
        self.full_date = datetime.strftime(datetime.fromtimestamp(stat.st_ctime), TRAY_EXTRA_INFO_DATE_FORMAT)
        self.length = get_video_duration(self.path)     # NOTE: this could be faster, but uglier
        length_int = int(self.length)
        self.length_string = f'{length_int // 60}:{length_int % 60:02}'
        self.length_size_string = f'Length: {self.length_string} ({self.size})'


    def refresh(self):
        ''' Refreshes various statistics for the clip, including creation-time, size, and length. '''
        stat = getstat(self.path)
        self.time = stat.st_ctime
        self.size = f'{(stat.st_size / 1048576):.1f}mb'
        self.length = get_video_duration(self.path)
        length_int = int(self.length)
        self.length_string = f'{length_int // 60}:{length_int % 60:02}'
        self.length_size_string = f'Length: {self.length_string} ({self.size})'


    def rename(self, path, name_format=RENAME_FORMAT, date_format=RENAME_DATE_FORMAT):
        ''' Renames the clip according to specified `name_format` and
            `date_format` based on ShadowPlay's default name formatting. '''
        try:
            parts = basename(path).split()
            parts[-1] = '.'.join(parts[-1].split('.')[:-3])

            date_string = ' '.join(parts[-3:])
            date = datetime.strptime(date_string, '%Y.%m.%d - %H.%M.%S')
            game = ' '.join(parts[:-3])
            if game.lower() in GAME_ALIASES: game = GAME_ALIASES[game.lower()]

            renamed_base_no_ext = name_format.replace('?game', game).replace('?date', date.strftime(date_format))
            renamed_path_no_ext = pjoin(dirname(path), renamed_base_no_ext)
            renamed_path = f'{renamed_path_no_ext}.mp4'

            count_detected = '?count' in renamed_path_no_ext
            if count_detected or exists(renamed_path):
                count = RENAME_COUNT_START_NUMBER
                if not count_detected:      # if forced to add a number, use windows-style count: start from (2)
                    count = 2
                    renamed_path_no_ext = f'{renamed_path_no_ext} (?count)'
                while True:
                    count_string = str(count).zfill(RENAME_COUNT_PADDED_ZEROS if count >= 0 else RENAME_COUNT_PADDED_ZEROS + 1)
                    renamed_path = f'{renamed_path_no_ext.replace("?count", count_string)}.mp4'
                    if not exists(renamed_path): break
                    count += 1
            renamed_base = basename(renamed_path)
            logging.info(f'Renaming video to: {renamed_base}')
            os.rename(path, renamed_path)
            logging.info('Rename successful.')
            return abspath(renamed_path), renamed_base      # use abspath to ensure consistent path formatting later on
        except Exception as error:
            logging.warning(f'(!) Clip at {path} could not be renamed (maybe it was already renamed?): "{error}"')
            return path, basename(path)


    def is_working(self, verb):
        if self.working:
            logging.info(f'Busy -- cannot {verb.lower()}. Clip {self.name} is being worked on.')
            play_alert('busy')
            return True
        return False


# ---------------------
# Main class
# ---------------------
class AutoCutter:
    __slots__ = ('waiting_for_clip', 'last_clip_time', 'last_clips')

    def __init__(self):
        start = time.time()
        self.waiting_for_clip = False
        if exists(HISTORY_PATH):
            with open(HISTORY_PATH, 'r') as history:
                # get all valid paths from history file, create a buffer of clip objects, then start caching paths as strings outside buffer
                lines = (path for path in reversed(history.read().splitlines()) if path and exists(path))   # reversed() is an iterable, not an actual list
                logging.info(f'History file parsed in {time.time() - start:.3f} seconds.')

                last_clips = [Clip(path, getstat(path)) if index <= CLIP_BUFFER else path for index, path in enumerate(lines)]
                last_clips.reverse()                        # .reverse() is a very fast operation
                if last_clips: logging.info(f'Previous {len(last_clips)} clip{"s" if len(last_clips) != 1 else ""} loaded: {last_clips}')
                else: logging.info('No previous clips detected.')
                logging.info(f'Previous clips loaded in {time.time() - start:.2f} seconds.')

                self.last_clips = last_clips
                if CHECK_FOR_NEW_CLIPS_ON_LAUNCH:
                    self.last_clip_time = getstat(last_clips[-1].path).st_ctime
                    self.check_for_clips(manual_update=True)
                del lines
        else:
            self.last_clips = []
            msg = ("It appears to be your first time running Instant Replay Suite "
                   "(or you deleted your history file). Please review your hotkey "
                   "and renaming settings and restart if necessary.\n\nWould you "
                   "like to organize, rename, and add all existing clips in "
                   f"{VIDEO_PATH}? Click cancel to exit Instant Replay Suite.")
            MessageBox = ctypes.windll.user32.MessageBoxW   # flags are ?-symbol, stay on top, Yes/No/Cancel
            response = MessageBox(None, msg, 'Welcome to Instant Replay Suite', 0x00040023)
            if response == 2:               # Cancel/X
                logging.info('Cancel selected on welcome dialog, closing...')
                exit(2)
            elif response == 7:             # No
                logging.info('No selected on welcome dialog, not retroactively adding clips.')
                self.last_clip_time = time.time()
            elif response == 6:             # Yes
                logging.info('Yes selected on welcome dialog, looking for pre-existing clips...')
                self.last_clip_time = 0     # set last_clip_time to 0 so all .mp4 files are valid
                self.check_for_clips(manual_update=True)

        keyboard.add_hotkey(INSTANT_REPLAY_HOTKEY, self.check_for_clips)
        keyboard.add_hotkey(CONCATENATE_HOTKEY, self.concatenate_last_clips)
        keyboard.add_hotkey(DELETE_HOTKEY, self.delete_clip)
        keyboard.add_hotkey(UNDO_HOTKEY, self.undo)
        for key, length in LENGTH_DICTIONARY.items():
            keyboard.add_hotkey(f'{LENGTH_HOTKEY} + {key}', self.trim_clip, args=(length,))
        logging.info(f'Auto-cutter initialized in {time.time() - start:.2f} seconds.')

    # ---------------------
    # Helper methods
    # ---------------------
    def pop(self, index=-1):
        ''' Wrapper for popping from the last_clips list that converts cached paths from strings to
            Clip objects, if necessary, with a failsafe for non-existent cached paths included.
            Attempts to return popped value on error, if possible -- otherwise returns None. '''
        last_clips = self.last_clips
        popped = None
        try:
            popped = last_clips.pop(index)
            cached_clip_index = -(CLIP_BUFFER + 1)      # +1 to check first clip outside buffer
            while isinstance(cached_clip := last_clips[cached_clip_index], str):    # convert clip to actual Clip object if it's just a string
                if exists(cached_clip):                 # if cached clip exists, convert to Clip object and break loop
                    last_clips[cached_clip_index] = Clip(cached_clip, getstat(cached_clip), rename=False)
                    break
                last_clips.pop(cached_clip_index)       # if cached clip doesn't exist, pop and try next clip
        except IndexError: pass                         # IndexError -> pop was out of range, pass and return None
        except: logging.error(f'(!) Error while popping clip at index {index} <cached_clip_index={cached_clip_index}, len(self.last_clips)={len(self.last_clips)}>: {format_exc()}')
        return popped


    def wait(self, verb=None, alert=None, min_clips=1):
        ''' Checks if we have a clip queued up in another thread and then waits for it. Plays an associated sound
            effect with the action we're going to perform, if specified with `alert`. Checks if we have enough
            min_clips before returning: returns False if the action cannot continue, and True otherwise.
            Waiting process is logged using `verb` to describe the action. '''
        log_verb = verb if verb else alert
        if alert is not None: play_alert(str(alert).lower())
        if self.waiting_for_clip:
            logging.info(f'{log_verb} detected, waiting for instant replay to finish saving...')
            while self.waiting_for_clip: time.sleep(0.2)
        if len(self.last_clips) < min_clips:
            logging.warning(f'(!) Cannot {log_verb.lower()}: Not enough clips have been created since the auto-cutter was started')
            return False
        return True


    def get_clip(self, index=None, path=None, verb=None, alert=None, min_clips=1, patient=True, _recursive=False):
        ''' Safely gets a clip at a specified `index` or `path`. Can wait for the clip if `patient` is True, otherwise pulls a clip immediately, which
            is unlikely to return the wrong clip outside of intentional misuse. Plays an associated sound effect with the action we're going to
            perform, if specified with `alert`, and avoids passing this to the wait() method. Plays an  Uses the `verb`, `alert`, and `min_clips`
            parameters for waiting and checking if the acquired clip is working. Pops clips if they don't exist anymore. Recursively calls itself
            until a valid clip is acquired if `patient` is True. `_recursive` used internally to avoid unnecessary waiting/alerts. '''
        if not _recursive:
            logging.info(f'Getting clip at index {index} (verb={verb} alert={alert} min_clips={min_clips} patient={patient})')
            if alert is not None: play_alert(str(alert).lower())
            if patient and not self.wait(verb=verb if verb else alert, min_clips=min_clips): return

        if index is not None: clip = self.last_clips[index]
        elif path is not None:
            path = abspath(path)
            for last_clip_index, last_clip in enumerate(self.last_clips):
                if isinstance(last_clip, Clip):
                    if last_clip.path == path:
                        clip = last_clip
                        index = last_clip_index
                        break
                elif last_clip == path:
                    return path     # `path` in last_clips but as cached string -> return path immediately
            else: return None       # `path` specified but not present in last_clips -> return None

        if not exists(clip.path):
            self.pop(index)
            if patient: return self.get_clip(index, path, verb, alert, min_clips, patient, _recursive=True)
            else:
                logging.warning(f'Clip at index {index} does not actually exist: {clip.path}')
                play_alert('error')
                raise AssertionError("Clip does not exist")
        if clip.is_working(verb if verb else alert): raise AssertionError("Clip is being worked on")
        return clip

    # ---------------------
    # Acquiring clips
    # ---------------------
    def check_for_clips(self, manual_update=False):
        self.waiting_for_clip = True
        Thread(target=self.check_for_clips_thread, args=(manual_update,), daemon=True).start()


    def check_for_clips_thread(self, manual_update=False):
        try:
            if not manual_update:
                logging.info('Instant replay detected! Waiting 3 seconds...')
                self.last_clip_time = time.time() - 1
                time.sleep(3)
            logging.info(f'Scanning {VIDEO_PATH} for videos')
            last_clip_time = self.last_clip_time    # alias for optimization
            for root, _, files in os.walk(VIDEO_PATH):
                if basename(root).lower() in IGNORE_VIDEOS_IN_THESE_FOLDERS: continue
                if root == BACKUP_DIR: continue
                for file in files:
                    path = pjoin(root, file)
                    stat = getstat(path)            # ↓ skip non-mp4 files ↓
                    if last_clip_time < stat.st_ctime and file[-4:] == '.mp4':
                        logging.info(f'New video detected: {file}')
                        while get_video_duration(path) == 0:
                            logging.debug(f'VIDEO DURATION IS 0 ({path})')
                            time.sleep(0.2)

                        clip = Clip(path, stat, rename=RENAME)
                        self.last_clips.append(clip)
                        self.last_clip_time = stat.st_ctime
                        logging.info(f'Memory usage after adding clip: {get_memory():.2f}mb\n')
                        if manual_update: continue
                        else: return
        except:
            logging.error(f'(!) Error while checking for new clips: {format_exc()}')
            play_alert('error')
        finally:
            if manual_update:
                self.last_clips.sort(key=lambda clip: clip.time if isinstance(clip, Clip) else getstat(clip).st_ctime)
                logging.info('Manual scan complete.')
            self.waiting_for_clip = False
            gc.collect(generation=2)

    # ---------------------
    # Clip actions
    # ---------------------
    def trim_clip(self, length, index=-1, patient=True):
        try:    # recusively trim until the desired clip exists or no clips remain
            clip = self.get_clip(index, verb='Trim', alert=length, min_clips=1, patient=patient)
            trim_off_start_in_place(clip, length)
            clip.refresh()

        except (IndexError, AssertionError): return
        except:
            logging.error(f'(!) Error while cutting last clip: {format_exc()}')
            play_alert('error')


    def concatenate_last_clips(self, index=-1, patient=True):
        ''' 1. Get clip at `index` and clip just before that
            2. Write clip paths to a text file
            3. Concat files to a third temporary file using FFmpeg
            4. Delete text file and move original clips to backup folder
            5. Rename temporary file to the second clip's name
            6. Remove clip at `index` from recent clips and refresh '''
        try:
            clip1 = self.get_clip(index=index - 1, alert='Concatenate', min_clips=1, patient=patient)
            clip2 = self.get_clip(index=index, min_clips=2, patient=patient)
            clip_path1 = clip1.path
            clip_path2 = clip2.path

            base, ext = splitext(clip_path1)
            temp_path = f'{base}_concat{ext}'
            text_path = f'{base}_concatlist.txt'    # write list of clips to text file (with / as separator and single quotes to avoid ffmpeg errors)

            with open(text_path, 'w') as txt:
                txt.write(f"file '{clip_path1.replace(os.sep, '/')}'\nfile '{clip_path2.replace(os.sep, '/')}'")

            cmd = f'ffmpeg -y -f concat -safe 0 -i "{text_path}" -c copy "{temp_path}" -hide_banner -loglevel warning'
            logging.info(cmd)
            process = subprocess.Popen(cmd, shell=True)
            process.wait()

            delete(text_path)
            ext = splitext(clip_path1)[-1]
            temp_path1 = pjoin(BACKUP_DIR, f'{time.time_ns()}_1{ext}')
            os.renames(clip_path1, temp_path1)
            ext = splitext(clip_path2)[-1]
            temp_path2 = pjoin(BACKUP_DIR, f'{time.time_ns()}_2{ext}')
            os.renames(clip_path2, temp_path2)

            with open(UNDO_LIST_PATH, 'w') as undo:
                undo.write(f'{basename(temp_path1)} -> {basename(temp_path2)} -> {clip_path1} -> {clip_path2} -> Concatenated\n')
            refresh_backups(temp_path1, temp_path2)

            os.rename(temp_path, clip_path1)
            setctime(clip_path1, getstat(temp_path1).st_ctime)  # ensure edited clip retains original creation time
            self.pop(index)
            clip1.refresh()
            logging.info('Clips concatenated, renamed, popped, refreshed, and cleaned up successfully.')
        except AssertionError: return
        except:
            logging.error(f'(!) Error while concatenating last two clips: {format_exc()}')
            play_alert('error')


    def delete_clip(self, index=-1, patient=True):
        ''' Deletes a clip at the given `index`. '''
        if patient and not self.wait(alert='Delete'): return    # not a part of get_clip(), to simplify things
        try:
            logging.info(f'Deleting clip at index {index}: {self.last_clips[index].path}')
            delete(self.pop(index).path)                        # pop and remove directly
            logging.info('Deletion successful.\n')

        except IndexError: return
        except:
            logging.error(f'(!) Error while deleting last clip: {format_exc()}')
            play_alert('error')


    def open_clip(self, index=-1, play=TRAY_CLIPS_PLAY_ON_CLICK, patient=True):
        ''' Opens a recent clip by its index. If play is True, this function
            plays the video directly. Otherwise, it's opened in explorer. '''
        try:
            clip_path = self.get_clip(index, verb='Open', patient=patient).path
            os.startfile(clip_path) if play else subprocess.Popen(f'explorer /select,"{clip_path}"', shell=True)
            return clip_path

        except (IndexError, AssertionError): return
        except:
            logging.error(f'(!) Error while cutting last clip: {format_exc()}')
            play_alert('error')


    def compress_clip(self, index=-1, patient=True):
        try:    # rename clip to include (comressing...), compress, then rename back and refresh
            clip = self.get_clip(index, verb='Compress', patient=patient)
            clip.working = True
            try: Thread(target=self.compress_clip_thread, args=(clip,), daemon=True).start()
            except:
                logging.error(f'(!) Error while creating compression thread: {format_exc()}')
                play_alert('error')
            finally: clip.working = False

        except (IndexError, AssertionError): pass
        except:
            logging.error(f'(!) Error while cutting last clip: {format_exc()}')
            play_alert('error')


    def compress_clip_thread(self, clip: Clip):
        try:
            clip.working = True
            base, ext = splitext(clip.path)
            old_path = clip.path
            new_path = f'{base} (compressing...){ext}'
            clip.path = new_path
            os.rename(old_path, new_path)
            logging.info(f'Video size is {clip.size}. Compressing...')
            ffmpeg(clip.path, f'-i "%tp" -vcodec libx265 -crf 28 "{clip.path}"')
            logging.info(f'New compressed size is: {clip.size}')
            os.rename(new_path, old_path)
            clip.path = old_path
            clip.refresh()
        except:
            logging.error(f'(!) Error while compressing clip: {format_exc()}')
            play_alert('error')
        finally:
            clip.working = False


    # TODO what should this do when you undo actions on deleted clips? restore the clip anyways?
    def undo(self, *args, patient=True):    # NOTE: *args used to capture pystray's unused args
        try:
            with open(UNDO_LIST_PATH, 'r') as undo:
                line = undo.readline().strip().split(' -> ')
                if len(line) == 3:      # trim
                    old, new, action = line

                    alert = 'Undoing Trim' if 'trim' in action.lower() else 'Undo'
                    if patient and not self.wait(verb=f'Undo "{action}"', alert=alert): return

                    os.remove(new)
                    os.rename(pjoin(BACKUP_DIR, old), new)
                    logging.info(f'Undo completed for "{action}" on clip "{new}"')
                    clip = self.get_clip(path=new)
                    if isinstance(clip, Clip): clip.refresh()       # refresh clip

                elif len(line) == 5:    # concatenation
                    old, old2, new, new2, action = line

                    alert = 'Undoing Concatenation' if 'concat' in action.lower() else 'Undo'
                    if patient and not self.wait(verb=f'Undo "{action}"', alert=alert): return

                    os.remove(new)
                    os.rename(pjoin(BACKUP_DIR, old), new)
                    os.rename(pjoin(BACKUP_DIR, old2), new2)

                    index = self.last_clips.index(new)              # concat removes new2 from the list...
                    buffer = len(self.last_clips) - CLIP_BUFFER     # ...so new2 needs to be re-added
                    if index < buffer: self.last_clips.insert(index + 1, new2)  # TODO make sure this math is right
                    else: self.last_clips.insert(index + 1, Clip(new2, getstat(new2), rename=False))

                    logging.info(f'Undo completed for "{action}" on clips "{new}" and "{new2}"')
                    clip = self.get_clip(path=new)
                    if isinstance(clip, Clip): clip.refresh()       # refresh original clip
            os.remove(UNDO_LIST_PATH)
        except:
            logging.error(f'(!) Error while undoing last action: {format_exc()}')
            play_alert('error')


###########################################################
if __name__ == '__main__':
    try:
        verify_ffmpeg()
        logging.info('FFmpeg installation verified.')

        logging.info(f'Memory usage before initializing AutoCutter class: {get_memory():.2f}mb')
        cutter = AutoCutter()
        logging.info(f'Memory usage after initializing AutoCutter class: {get_memory():.2f}mb')

        if SEND_DELETED_FILES_TO_RECYCLE_BIN: import send2trash

        # ---------------------
        # Tray-icon functions
        # ---------------------
        def get_clip_tray_title(index: int, format: str = TRAY_RECENT_CLIP_NAME_FORMAT, default: str = TRAY_RECENT_CLIP_DEFAULT_TEXT) -> str:
            ''' Returns the basename of a recent clip by its index. If no clip exists at that index, the default parameter is returned. '''
            try:
                clip = cutter.last_clips[index]
                if not exists(clip.path):
                    cutter.pop(index)
                    return get_clip_tray_title(index)

                # get loose estimate of how long ago it was, if necessary
                if '?recency' in format:
                    time_delta = time.time() - clip.time
                    d = int(time_delta // 86400)
                    h = int(time_delta // 3600)
                    m = int(time_delta // 60)
                    if '?recencyshort' in format:
                        if d:   time_delta_string = f'{d}d'
                        elif h: time_delta_string = f'{h}h'
                        elif m: time_delta_string = f'{m}m'
                        else: time_delta_string = '0m'
                    else:
                        if d:   time_delta_string = f'{d} day{ "s" if d > 1 else ""} ago'
                        elif h: time_delta_string = f'{h} hour{"s" if h > 1 else ""} ago'
                        elif m: time_delta_string = f'{m} min{ "s" if m > 1 else ""} ago'
                        else: time_delta_string = 'just now'
                else: time_delta_string = ''

                return (format
                        .replace('?date', clip.date)
                        .replace('?recencyshort', time_delta_string)
                        .replace('?recency', time_delta_string)
                        .replace('?size', clip.size)
                        .replace('?length', clip.length_string)
                        .replace('?clippath', clip.path)
                        .replace('?clipdir', os.sep.join(clip.path.split(os.sep)[-2:]) if '?clipdir' in format else '')
                        .replace('?clip', clip.name))
            except IndexError: return default
            except: logging.error(f'(!) Error while generating title for system tray clip {clip} at index {index}: {cutter.last_clips} {format_exc()}')


        def get_clip_tray_action(index: int):
            ''' Generates the action associated with each recent clip item in the tray-icon's menu.
                If a submenu is desired, one is created and returned. Otherwise, a callable lambda is returned.
                `index` will be a negative number, starting with -1 for the topmost item. '''
            if TRAY_RECENT_CLIPS_HAVE_UNIQUE_SUBMENUS:
                if TRAY_RECENT_CLIPS_SUBMENU_EXTRA_INFO:
                    last_clips = cutter.last_clips
                    extra_info_items = (
                        pystray.MenuItem(pystray.MenuItem(None, None), None),
                        pystray.MenuItem(lambda _: last_clips[index].length_size_string if len(last_clips) >= -index else TRAY_RECENT_CLIP_DEFAULT_TEXT, None, enabled=False),
                        pystray.MenuItem(lambda _: last_clips[index].full_date if len(last_clips) >= -index else TRAY_RECENT_CLIP_DEFAULT_TEXT, None, enabled=False)
                    )
                else: extra_info_items = tuple()

                get_trim_action = lambda length, index: lambda: cutter.trim_clip(length, index, patient=False)   # workaround for python bug/oddity involving creating lambdas in iterables
                return pystray.Menu(
                    pystray.MenuItem('Trim...', pystray.Menu(*(pystray.MenuItem(f'{length} seconds', get_trim_action(length, index)) for length in LENGTH_DICTIONARY.values()))),
                    pystray.MenuItem('Play...', lambda: cutter.open_clip(index, play=True, patient=False)),
                    pystray.MenuItem('Explore...', lambda: cutter.open_clip(index, play=False, patient=False)),
                    pystray.MenuItem('Splice...', lambda: cutter.concatenate_last_clips(index, patient=False)),
                    #pystray.MenuItem('Compress...', lambda: cutter.compress_clip(index, patient=False)),
                    #pystray.MenuItem('Audio only...', lambda: cutter.???(index, patient=False)),
                    #pystray.MenuItem('Video only...', lambda: cutter.???(index, patient=False)),
                    pystray.MenuItem('Delete...', lambda: cutter.delete_clip(index, patient=False)),
                    *extra_info_items
                )
            return lambda: cutter.open_clip(index)

        # ---------------------
        # Tray-icon setup
        # ---------------------
        SEPARATOR = pystray.MenuItem(pystray.MenuItem(None, None), None)

        # creating the base recent-clip menu
        title_callback = lambda index: lambda _: get_clip_tray_title(index)  # workaround for python bug/oddity involving creating lambdas in iterables
        RECENT_CLIPS_BASE = tuple(pystray.MenuItem(title_callback(i), get_clip_tray_action(i)) for i in range(-1, (TRAY_RECENT_CLIP_COUNT * -1) - 1, -1))

        # action dictionary
        TRAY_ACTIONS = {
            'open_log':             lambda: os.startfile(LOG_PATH),
            'open_video_folder':    lambda: os.startfile(VIDEO_PATH),
            'open_install_folder':  lambda: os.startfile(CWD),
            'play_most_recent':     lambda: cutter.open_clip(play=True),
            'explore_most_recent':  lambda: cutter.open_clip(play=False),
            'delete_most_recent':   lambda: cutter.delete_clip(),               # small RAM drop by making these not lambdas
            'concatenate_last_two': lambda: cutter.concatenate_last_clips(),    # 26.0mb -> 25.8mb on average
            'clear_history':        lambda: cutter.last_clips.clear(),
            'update':               lambda: cutter.check_for_clips(manual_update=True),
            'undo':                 cutter.undo,
            'quit':                 quit_tray,
        }

        # setting special-click actions
        LEFT_CLICK_ACTION = TRAY_ACTIONS.get(TRAY_LEFT_CLICK_ACTION)
        MIDDLE_CLICK_ACTION = TRAY_ACTIONS.get(TRAY_MIDDLE_CLICK_ACTION)
        if LEFT_CLICK_ACTION is None: logging.warning(f'(X) Left click action "{TRAY_LEFT_CLICK_ACTION}" does not exist')
        if MIDDLE_CLICK_ACTION is None: logging.warning(f'(X) Middle click action "{TRAY_MIDDLE_CLICK_ACTION}" does not exist')

        # setting the left-click action in pystray has an unusual implementation
        LEFT_CLICK_ACTION = pystray.MenuItem(None, action=LEFT_CLICK_ACTION, default=True, visible=False)

        # creating menu -- advanced mode
        if TRAY_ADVANCED_MODE:
            exit_item_exists = False            # variable for making sure an exit item is included
            def evaluate_menu(items, menu):     # function for recursively solving menus/submenus and exporting them to a list
                global exit_item_exists
                for item_dict in items:
                    assert not isinstance(item_dict, set), (f'Tray item {item_dict} is improperly written.\n\n'
                                                            'Usage: {"title of item": "action_of_item"}\n       '
                                                            'OR\n       {"title of submenu": {nested_menu}}\n\n'
                                                            'Note the colon (:) between the title and it\'s action.')
                    if isinstance(item_dict, str):
                        item = item_dict.strip().lower()
                        if item == 'separator': menu.append(SEPARATOR)
                        elif item == 'recent_clips': menu.extend(RECENT_CLIPS_BASE)
                        elif item == 'memory': menu.append(pystray.MenuItem(lambda _: f'Memory usage: {get_memory():.2f}mb', None, enabled=False))
                        continue

                    for name, action in item_dict.items():
                        try:                            # create and append menu item
                            action = action.strip().lower()
                            if action == 'quit': exit_item_exists = True    # mark that an exit item is in the menu
                            elif action == 'memory':    # get_mem_title -> workaround for python bug/oddity involving creating lambdas in iterables
                                get_mem_title = lambda name: lambda _: name.replace('?memory', f'{get_memory():.2f}mb')
                                menu.append(pystray.MenuItem(get_mem_title(name), None, enabled=False))
                                continue
                            elif action == 'recent_clips':
                                action = (action,)      # set action as tuple and raise AttributeError to read it as a submenu
                                raise AttributeError
                            menu.append(pystray.MenuItem(name, action=TRAY_ACTIONS[action]))
                        except AttributeError:          # AttributeError -> item is a submenu
                            if isinstance(action, list) or isinstance(action, tuple):
                                submenu = []
                                evaluate_menu(action, submenu)
                                menu.append(pystray.MenuItem(name, pystray.Menu(*submenu)))
                        except KeyError: logging.warning(f'(X) The following menu item does not exist: "{name}": "{action}"')

            tray_menu = [LEFT_CLICK_ACTION] if LEFT_CLICK_ACTION else []        # start with hidden left-click action included, if present
            evaluate_menu(TRAY_ADVANCED_MODE_MENU, tray_menu)
            if not exit_item_exists: tray_menu.append(pystray.MenuItem('Exit', quit_tray))  # add exit item if needed

        # creating menu -- "basic" mode
        else:   # create the base quick-actions menu
            if not TRAY_SHOW_QUICK_ACTIONS: QUICK_ACTIONS_BASE = tuple()
            else: QUICK_ACTIONS_BASE = (
                pystray.MenuItem('Play most recent clip', lambda: cutter.open_clip(play=True)),
                pystray.MenuItem('View last clip in explorer', lambda: cutter.open_clip(play=False)),
                pystray.MenuItem('Concatenate two last clips', lambda: cutter.concatenate_last_clips()),
                pystray.MenuItem('Delete most recent clip', lambda: cutter.delete_clip()),
                pystray.MenuItem('Undo most recent action', cutter.undo)
            )

            # set up final quick-action and recent-clip menus + setting their location/organization within the full menu
            RECENT_CLIPS_SEPARATOR = pystray.MenuItem(pystray.MenuItem(None, None), None, visible=True if TRAY_RECENT_CLIP_COUNT and TRAY_SHOW_QUICK_ACTIONS else False)
            if TRAY_RECENT_CLIPS_IN_SUBMENU and TRAY_QUICK_ACTIONS_IN_SUBMENU:
                RECENT_CLIPS_MENU = (pystray.MenuItem('Recent clips', pystray.Menu(*QUICK_ACTIONS_BASE, RECENT_CLIPS_SEPARATOR, *RECENT_CLIPS_BASE)), )
            elif TRAY_QUICK_ACTIONS_IN_SUBMENU:
                RECENT_CLIPS_MENU = (pystray.MenuItem('Quick actions', pystray.Menu(*QUICK_ACTIONS_BASE)), *RECENT_CLIPS_BASE)
            elif TRAY_RECENT_CLIPS_IN_SUBMENU:
                RECENT_CLIPS_MENU = (*QUICK_ACTIONS_BASE, RECENT_CLIPS_SEPARATOR, pystray.MenuItem('Recent clips', pystray.Menu(*RECENT_CLIPS_BASE)))
            else:
                RECENT_CLIPS_MENU = (*QUICK_ACTIONS_BASE, RECENT_CLIPS_SEPARATOR, *RECENT_CLIPS_BASE)
            del RECENT_CLIPS_SEPARATOR
            del QUICK_ACTIONS_BASE

            # create menu
            tray_menu = (
                LEFT_CLICK_ACTION,
                pystray.MenuItem('View log',    action=lambda: os.startfile(LOG_PATH)),
                pystray.MenuItem('View videos', action=lambda: os.startfile(VIDEO_PATH)),
                SEPARATOR,
                *RECENT_CLIPS_MENU,
                SEPARATOR,
                pystray.MenuItem('Check for clips', action=lambda: cutter.check_for_clips(manual_update=True)),
                pystray.MenuItem('Clear history',   action=lambda: cutter.last_clips.clear()),
                pystray.MenuItem('Exit', quit_tray)
            )

        # create system tray icon
        tray_icon = Icon(None, Image.open(ICON_PATH), 'Instant Replay Suite', tray_menu)

        # cleanup *some* extraneous dictionaries/collections/functions
        del verify_ffmpeg
        del get_clip_tray_action
        del title_callback
        del tray_menu
        del SEPARATOR
        del RECENT_CLIPS_BASE
        del TRAY_ADVANCED_MODE_MENU
        del LENGTH_DICTIONARY
        del INSTANT_REPLAY_HOTKEY
        del CONCATENATE_HOTKEY
        del DELETE_HOTKEY
        del LENGTH_HOTKEY
        del TRAY_ACTIONS
        del TRAY_LEFT_CLICK_ACTION
        del TRAY_MIDDLE_CLICK_ACTION

        # final garbage collection to reduce memory usage
        gc.collect(generation=2)
        logging.info(f'Memory usage before initializing system tray icon: {get_memory():.2f}mb')

        # finally, run system tray icon
        logging.info('Running.')
        tray_icon.run()
    except SystemExit: pass
    except:
        logging.critical(f'(!) Error while initalizing Instant Replay Suite: {format_exc()}')
        play_alert('error')
        time.sleep(2.5)   # sleep to give error sound time to play
