''' NVIDIA Instant Replay auto-cutter 1/11/22 '''
from __future__ import annotations  # 0.10mb  / 13.45       https://stackoverflow.com/questions/33533148/how-do-i-type-hint-a-method-with-the-type-of-the-enclosing-class
import gc                           # 0.275mb / 13.625mb    <- Heavy, but probably worth it
import os                           # 0.10mb  / 13.45mb
import sys                          # 0.125mb / 13.475mb
import json
import time                         # 0.125mb / 13.475mb
import atexit
import ctypes
import logging                      # 0.65mb  / 14.00mb
import subprocess                   # 0.125mb / 13.475mb
import tracemalloc                  # 0.125mb / 13.475mb
from threading import Thread        # 0.125mb / 13.475mb
from datetime import datetime       # 0.125mb / 13.475mb
from traceback import format_exc    # 0.35mb  / 13.70mb     <- Heavy, but probably worth it

import pystray                      # 3.29mb  / 16.64mb
import keyboard                     # 2.05mb  / 15.40mb
import winsound                     # 0.21mb  / 13.56mb     ↓ https://stackoverflow.com/questions/3844430/how-to-get-the-duration-of-a-video-in-python
import pymediainfo                  # 3.75mb  / 17.1mb      https://stackoverflow.com/questions/15041103/get-total-length-of-videos-in-a-particular-directory-in-python
from pystray._util import win32
from win32_setctime import setctime
from configparsebetter import ConfigParseBetter

# Starts with roughly ~36.7mb of memory usage. Roughly 9.78mb combined from imports alone, without psutil and cv2/pymediainfo (9.63mb w/o tracemalloc).
tracemalloc.start()                 # start recording memory usage AFTER libraries have been imported

'''
TODO extended backup system with more than 1 undo possible at a time
TODO add stuff for multi-track recordings
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
'''

# ---------------------------------
#     --- Table of contents ---
#  Aliases
#  Base constants
#  Logging
#  Utility functions
#  Temporary functions
#  Video path
#  Settings
#  Registry settings
#  Other constants & paths
#  Custom Pystray class
#  Custom keyboard listener
#  Clip class
#  Main class
#      Helper methods
#      Acquiring clips
#      Clip actions
#  Tray-icon functions
#  Tray-icon setup
# ---------------------------------


# ---------------------
# Aliases
# ---------------------
parsemedia = pymediainfo.MediaInfo.parse
sep = os.sep
sepjoin = sep.join
pjoin = os.path.join
exists = os.path.exists
getstat = os.stat
getsize = os.path.getsize
basename = os.path.basename
dirname = os.path.dirname
abspath = os.path.abspath
splitext = os.path.splitext
splitpath = os.path.split
splitdrive = os.path.splitdrive
ismount = os.path.ismount
isdir = os.path.isdir
listdir = os.listdir
makedirs = os.makedirs
rename = os.rename
remove = os.remove


# ---------------------
# Base constants
# ---------------------
TITLE = 'Instant Replay Suite'
VERSION = '1.0.0'
REPOSITORY_URL = 'https://github.com/thisismy-github/instant-replay-suite'

# ---

IS_COMPILED = getattr(sys, 'frozen', False)
SCRIPT_START_TIME = time.time()

# current working directory
SCRIPT_PATH = sys.executable if IS_COMPILED else os.path.realpath(__file__)
CWD = dirname(SCRIPT_PATH)
os.chdir(CWD)

# other paths that will always be the same no matter what
RESOURCE_FOLDER = pjoin(CWD, 'resources')
BIN_FOLDER = pjoin(CWD, 'bin')
APPDATA_FOLDER = pjoin(os.path.expandvars('%LOCALAPPDATA%'), TITLE)
CONFIG_PATH = pjoin(CWD, 'config.settings.ini')
CUSTOM_MENU_PATH = pjoin(CWD, 'config.menu.ini')
SHADOWPLAY_REGISTRY_PATH = r'SOFTWARE\NVIDIA Corporation\Global\ShadowPlay\NVSPCAPS'

LOG_PATH = splitext(basename(SCRIPT_PATH))[0] + '.log'
MEDIAINFO_DLL_PATH = pjoin(BIN_FOLDER, 'MediaInfo.dll') if IS_COMPILED else None


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
# Utility functions
# ---------------------
#get_memory = lambda: psutil.Process().memory_info().rss / (1024 * 1024)
get_memory = lambda: tracemalloc.get_traced_memory()[0] / 1048576


def show_message(title: str, msg: str, flags: int = 0x00040030):
    ''' Displays a MessageBoxW on the screen with a `title` and
        `msg`. Default `flags` are <!-symbol + stay on top>.
        https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-messageboxw '''
    logging.info(f'Showing message box "{title}":\n\n---\n{msg}\n---\n')
    return ctypes.windll.user32.MessageBoxW(None, msg, title, flags)


def delete(path: str) -> None:
    ''' Robustly deletes a given `path`. '''
    logging.info('Deleting: ' + path)
    try:
        if SEND_DELETED_FILES_TO_RECYCLE_BIN: send2trash.send2trash(path)
        else: remove(path)
    except:
        logging.error(f'(!) Error while deleting file {path}: {format_exc()}')
        play_alert('error')


def renames(old: str, new: str) -> None:
    ''' `os.py`'s super-rename, but without deleting empty directories. '''
    head, tail = splitpath(new)
    if head and tail and not exists(head): makedirs(head)
    rename(old, new)


def get_video_duration(path: str) -> float:     # ? -> https://stackoverflow.com/questions/10075176/python-native-library-to-read-metadata-from-videos
    ''' Returns a precise duration for the video at `path`.
        Returns 0 if `path` is corrupt or an invalid format. '''
    for track in parsemedia(path, library_file=MEDIAINFO_DLL_PATH).tracks:
        if track.track_type == "Video":
            return track.duration / 1000
    return 0


def auto_rename_clip(path: str) -> None:
    ''' Renames `path` according to `RENAME_FORMAT` and `RENAME_DATE_FORMAT`,
        so long as `path` ends with a date formatted as `%Y.%m.%d - %H.%M.%S`,
        which is found at the end of all ShadowPlay clip names. '''
    try:
        parts = basename(path).split()
        parts[-1] = '.'.join(parts[-1].split('.')[:-3])

        date_string = ' '.join(parts[-3:])
        date = datetime.strptime(date_string, '%Y.%m.%d - %H.%M.%S')
        game = ' '.join(parts[:-3])
        if game.lower() in GAME_ALIASES: game = GAME_ALIASES[game.lower()]

        renamed_base_no_ext = RENAME_FORMAT.replace('?game', game).replace('?date', date.strftime(RENAME_DATE_FORMAT))
        renamed_path_no_ext = pjoin(dirname(path), renamed_base_no_ext)
        renamed_path = renamed_path_no_ext + '.mp4'

        count_detected = '?count' in renamed_path_no_ext
        protected_paths = cutter.protected_paths    # this is exactly what i want to avoid but i'm leaving it for now
        if count_detected or exists(renamed_path) or renamed_path in protected_paths:
            count = RENAME_COUNT_START_NUMBER
            if not count_detected:                  # if forced to add number, use windows-style count: start from (2)
                count = 2
                renamed_path_no_ext += ' (?count)'
            while True:
                count_string = str(count).zfill(RENAME_COUNT_PADDED_ZEROS + (1 if count >= 0 else 2))
                renamed_path = renamed_path_no_ext.replace("?count", count_string) + '.mp4'
                if not exists(renamed_path) and renamed_path not in protected_paths: break
                count += 1

        renamed_base = basename(renamed_path)
        logging.info(f'Renaming video to: {renamed_base}')
        rename(path, renamed_path)                  # super-rename not needed
        logging.info('Rename successful.')
        return abspath(renamed_path), renamed_base  # use abspath to ensure consistent path formatting later on
    except Exception as error:
        logging.warning(f'(!) Clip at {path} could not be renamed (maybe it was already renamed?): "{error}"')
        return path, basename(path)


def play_alert(sound: str) -> None:
    ''' Plays a system-wide audio alert. `sound` is the filename of a WAV
        file located within `RESOURCE_FOLDER`. Plays a generic OS alert if
        `sound` doesn't exist, or a generic OS error sound if `sound` is
        "error" and "error.wav" doesn't exist. '''
    if not AUDIO: return
    path = pjoin(RESOURCE_FOLDER, f'{sound}.wav')
    logging.info('Playing alert: ' + path)

    if exists(path):
        try: winsound.PlaySound(path, winsound.SND_ASYNC)
        except:
            winsound.MessageBeep(winsound.MB_ICONHAND)      # play OS error sound
            logging.error(f'(!) Error while playing alert {path}: {format_exc()}')
    else:       # generic OS alert for missing file, OS error for actual errors
        if sound == 'error': winsound.MessageBeep(winsound.MB_ICONHAND)
        else: winsound.MessageBeep(winsound.MB_ICONEXCLAMATION)
        logging.warning('(!) Alert doesn\'t exist at path ' + path)


def check_for_updates(manual: bool = True) -> None:
    ''' Checks for updates and notifies user if one is found. If compiled,
        user may download and install the update automatically, prompting
        us to exit while it occurs. Also checks for `update_report.txt`,
        left by a previous update. If `manual` is False, less cleanup is
        done and if the `CHECK_FOR_UPDATES_ON_LAUNCH` setting is also
        False, only the report is checked. '''

    update_report = pjoin(CWD, 'update_report.txt')
    update_report_exists = exists(update_report)
    if manual or CHECK_FOR_UPDATES_ON_LAUNCH or update_report_exists:
        import update

        # set update constants here to avoid circular import
        update.VERSION = VERSION
        update.REPOSITORY_URL = REPOSITORY_URL
        update.IS_COMPILED = IS_COMPILED
        update.SCRIPT_PATH = SCRIPT_PATH
        update.CWD = CWD
        update.RESOURCE_FOLDER = RESOURCE_FOLDER
        update.BIN_FOLDER = BIN_FOLDER
        update.show_message = show_message
        update.HYPERLINK = f'latest release on GitHub here:\n{REPOSITORY_URL}/releases/latest'

        # validate previous update first if needed
        if update_report_exists:
            update.validate_update(update_report)

        # check for updates and exit if we're installing a new version
        if manual or CHECK_FOR_UPDATES_ON_LAUNCH:
            if IS_COMPILED:             # if compiled, override cacert.pem path...
                import certifi.core     # ...to get rid of a pointless folder
                cacert_override_path = pjoin(BIN_FOLDER, 'cacert.pem')
                os.environ["REQUESTS_CA_BUNDLE"] = cacert_override_path
                certifi.core.where = lambda: cacert_override_path
            exit_code = update.check_for_update()
            if exit_code is not None:
                sys.exit(exit_code)


def about() -> None:
    ''' Displays an "About" window with various information/statistics. '''
    seconds_running = time.time() - SCRIPT_START_TIME
    clip_count = len(cutter.last_clips)

    if seconds_running < 60:
        time_delta_string = 'Script has been running for less than a minute'
    else:
        d = seconds_running // 86400
        h = seconds_running // 3600
        m = seconds_running // 60

        suffix = f'{m:g} minute{"s" if m != 1 else ""}'
        if h: suffix = f'{h:g} hour{"s" if m != 1 else ""}, {suffix}'
        if d: suffix = f'{d:g} day{"s" if m != 1 else ""}, {suffix}'
        time_delta_string = 'Script has been running for ' + suffix

    msg = (f'» {TITLE} v{VERSION} «\n{REPOSITORY_URL}\n\n---\n'
           f'{time_delta_string}\nScript is tracking '
           f'{clip_count} clip{"s" if clip_count != 1 else ""}'
           '\n---\n\n© thisismy-github 2022-2023')
    show_message('About ' + TITLE, msg, 0x00010040)  # i-symbol, set foreground


# ---------------------
# Temporary functions
# ---------------------
def abort_launch(code: int, title: str, msg: str, flags: int = 0x00040010) -> None:
    ''' Checks for updates, then displays/logs a `msg` with `title` using
        `flags`, then exits with exit code `code`. Default `flags` are
        <X-symbol + stay on top>. Only to be used during launch. '''
    check_for_updates(manual=False)
    show_message(title, msg, flags)
    sys.exit(code)


def verify_ffmpeg() -> None:
    ''' Checks if FFmpeg exists. If it isn't in the script's folder,
        the user's PATH system variable is checked. If still not found,
        a message box is displayed and the script exits. '''
    logging.info('Verifying FFmpeg installation...')
    if exists('ffmpeg.exe'): return
    else:
        for path in os.environ.get('PATH', '').split(';'):
            try:
                if 'ffmpeg.exe' in listdir(path):
                    return
            except: pass

    msg = ("FFmpeg was not detected. FFmpeg is required for all of this "
           "program's editing features. Please ensure `ffmpeg.exe` is "
           "either in your PATH or in this program's install folder.\n\n"
           "You can download FFmpeg for Windows here (not clickable, sorry): "
           "https://www.gyan.dev/ffmpeg/builds/")
    show_message('FFmpeg not detected', msg)
    sys.exit(3)


def verify_config_files() -> None:
    ''' Displays a message if config and/or menu file is missing, then gives
        user the option to create them immediately and quit or to continue
        with default settings and a default menu. '''
    logging.info('Verifying config.settings and config.menu...')
    if NO_CONFIG or NO_MENU:
        if NO_CONFIG and NO_MENU: parts = ('config file or a menu file', 'them', 'files')
        elif NO_CONFIG: parts = ('config file', 'it', 'file')
        else: parts = ('menu file', 'it', 'file')
        string1, string2, string3 = parts

        msg = (f"You do not have a {string1}. Would you like to exit {TITLE} "
               f"to create {string2} and review them now?\n\n"
               "Press 'No' if you want to continue with the default settings "
               f"(the necessary {string3} will be created on exit).")

        # ?-symbol, stay on top, Yes/No
        response = show_message('Missing config/menu files', msg, 0x00040024)
        if response == 6:       # Yes
            logging.info('Yes selected on missing config/menu dialog, closing...')
            if NO_MENU:         # create AFTER dialog is closed to avoid confusion
                restore_menu_file()
            sys.exit(1)
        elif response == 7:     # No
            logging.info('No selected on missing config/menu dialog, using defaults.')
            if NO_MENU:         # create AFTER dialog is closed to avoid confusion
                restore_menu_file()


def sanitize_json(path, comment_prefix='//'):
    ''' Reads a JSON file at `path`, but fixes common errors users may
        make, while allowing comments and value-only lines. Lines with
        `comment_prefix` are ignored. Designed for reading JSON files
        that are meant to be edited by users. '''
    with open(path, 'r') as file:
        striped = (line.strip() for line in file.readlines())
        raw_json_lines = (line for line in striped if line and line[:2] != comment_prefix)

    json_lines = []
    for line in raw_json_lines:
        # allow value-only lines
        if ':' not in line and line not in ('{', '}', '},'):
            line = '"": ' + line

        # ensure all nested dictionaries have exactly one trailing comma
        line = line.replace('}', '},')
        while ' ,' in line: line = line.replace(' ,', ',')
        while '},,' in line: line = line.replace('},,', '},')
        json_lines.append(line)

    # ensure final bracket exists and does not have a trailing comma
    json_string = '\n'.join(json_lines).rstrip().rstrip(',').rstrip('}') + '}'

    # remove trailing comma from final element of every dictionary
    comma_index = json_string.find(',')     # string ends with }, so we'll...
    bracket_index = json_string.find('}')   # ...run out of commas first
    while comma_index != -1:
        next_comma_index = json_string.find(',', comma_index + 1)
        if next_comma_index > bracket_index or next_comma_index == -1:
            quote_index = json_string.find('"', comma_index)
            if quote_index > bracket_index or quote_index == -1:
                start = json_string[:comma_index]
                end = json_string[comma_index + 1:]
                json_string = start + end
            bracket_index = json_string.find('}', bracket_index + 1)
        comma_index = next_comma_index

    # our `object_pairs_hook` immediately returns the raw list of pairs
    # instead of creating a dictionary, allowing us to use duplicate keys
    return json.loads(json_string, object_pairs_hook=lambda pairs: pairs)


def load_menu() -> list:
    ''' Parse menu at `CUSTOM_MENU_PATH` and warn/exit if parsing fails. '''
    try: return sanitize_json(CUSTOM_MENU_PATH, '//')
    except json.decoder.JSONDecodeError as error:
        msg = ("Error while reading your custom menu file "
               f"({CUSTOM_MENU_PATH}):\n\nJSONDecodeError - {error}"
               "\n\nThe custom menu file follows JSON syntax. If you "
               "need to reset your menu file, delete your existing "
               f"one and restart {TITLE} to generate a fresh copy.")
        show_message('Invalid Menu File', msg, 0x00040010)
        sys.exit(2)                         # ^ X-symbol, stay on top


def restore_menu_file() -> None:
    ''' Creates a fresh menu file at `CUSTOM_MENU_PATH`. '''
    logging.info(f'Creating fresh menu file at {CUSTOM_MENU_PATH}...')
    with open(CUSTOM_MENU_PATH, 'w') as file:
        file.write('''//                     --- CUSTOM TRAY MENU TUTORIAL ---
//
// This file defines a custom menu dictionary for your tray icon. It's JSON format,
// with some leeway in the formatting. Each item consists of a name-action pair,
// and actions may be named however you please. To create a submenu, add a nested
// dictionary as an action (see example below). Submenus work just like the base
// menu, and can be nested indefinitely. Actions without names will still be parsed,
// and will default to a blank name (see "Special tray actions" for exceptions).
//
// Normal tray actions:
//    "open_log":             Opens this program's log file.
//    "open_video_folder":    Opens the currently defined "Videos" folder.
//    "open_install_folder":  Opens this program's root folder.
//    "open_backup_folder":   Opens the currently defined backup folder.
//    "play_most_recent":     Plays your most recent clip.
//    "explore_most_recent":  Opens your most recent clip in Explorer.
//    "delete_most_recent":   Deletes your most recent clip.
//    "concatenate_last_two": Concatenates your two most recent clips.
//    "undo":                 Undoes the last trim or concatenation.
//    "clear_history":        Clears your clip history.
//    "refresh":              Manually checks for new clips and refreshes existing ones.
//    "check_for_updates":    Checks for a new release on GitHub to install.
//    'about':                Shows an "About" window.
//    "quit":                 Exits this program.
//
// Special tray actions:
//    "separator":            Adds a separator in the menu.
//                                - Cannot be named.
//                                    ("separator") OR ("": "separator")
//    "recent_clips":         Displays your most recent clips.
//                                - Naming this item will place it within a submenu:
//                                    ("Recent clips": "recent_clips")
//                                - Not naming this item will display all clips in the base menu:
//                                    ("recent_clips") OR ("": "recent_clips")
//    "memory":               Displays current RAM usage.
//                                - This is somewhat misleading and not worth using.
//                                - Use "?memory" in the title to represent where the number will be:
//                                    ("RAM: ?memory": "memory")
//                                - Not naming this item will default to "Memory usage: ?memorymb":
//                                    ("memory") OR ("": "memory")
//                                - This item will be greyed out and is informational only.
//
// Submenu example:
//    {
//        "Quick actions": {
//            "Play most recent clip": "play_most_recent",
//            "View last clip in explorer": "explore_most_recent",
//            "Concatenate last two clips": "concatenate_last_two",
//            "Delete most recent clip": "delete_most_recent"
//        },
//    }
// ---------------------------------------------------------------------------

{
\t"Open...": {
\t\t"Open root": "open_install_folder",
\t\t"Open videos": "open_video_folder",
\t\t"Open backups": "open_backup_folder",
\t\t"separator",
\t\t"Update check": "check_for_updates",
\t\t"About...": "about",
\t},
\t"View log": "open_log",
\t"separator",
\t"Play last clip": "play_most_recent",
\t"Explore last clip": "explore_most_recent",
\t"Undo last action": "undo",
\t"separator",
\t"recent_clips",
\t"separator",
\t"Refresh clips": "refresh",
\t"Clear history": "clear_history",
\t"Exit": "quit",
}''')


# ---------------------
# Video path
# ---------------------
# We do this before reading config file to set a hint for whether or not
# `SAVE_BACKUPS_TO_VIDEO_FOLDER` should default to True or False later.
try:                # gets ShadowPlay's video path from the registry
    import winreg   # NOTE: ShadowPlay settings are encoded in utf-16 and have a NULL character at the end
    key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, SHADOWPLAY_REGISTRY_PATH)
    VIDEO_FOLDER = winreg.QueryValueEx(key, 'DefaultPathW')[0].decode('utf-16')[:-1]
except:             # don't show message box until later
    logging.warning('(!) Could not read video folder from registry: ' + format_exc())
    VIDEO_FOLDER = None

# True if `VIDEO_FOLDER` and `CWD` are on different drives
BACKUP_FOLDER_HINT = (VIDEO_FOLDER and (splitdrive(VIDEO_FOLDER)[0] != splitdrive(CWD)[0]
                                        or ismount(VIDEO_FOLDER) != ismount(CWD)))


# ---------------------
# Settings
# ---------------------
cfg = ConfigParseBetter(
    CONFIG_PATH,
    caseSensitive=True,
    autosaveOnlyWhenFileDoesNotExist=True,
    comment_prefixes=('//',)
)

# --- Hotkeys ---
cfg.setSection(' --- Trim Hotkeys --- ')
cfg.comment('Usage: <hotkey> = <trim length>')
LENGTH_DICTIONARY = {}
for key, length in cfg.loadAllFromSection():
    try: LENGTH_DICTIONARY[key] = int(float(length.strip()))
    except: logging.warning(f'(!) Could not add trim length "{length}"')
if not LENGTH_DICTIONARY:
    LENGTH_DICTIONARY = {
        'alt + 1': 10,
        'alt + 2': 20,
        'alt + 3': 30,
        'alt + 4': 40,
        'alt + 5': 50,
        'alt + 6': 60,
        'alt + 7': 70,
        'alt + 8': 80,
        'alt + 9': 90
    }
    for name, alias in LENGTH_DICTIONARY.items():
        cfg.load(name, alias)

cfg.setSection(' --- Other Hotkeys --- ')
CONCATENATE_HOTKEY = cfg.load('CONCATENATE', 'alt + c')
UNDO_HOTKEY = cfg.load('UNDO', 'alt + u')
DELETE_HOTKEY = cfg.load('DELETE', 'ctrl + alt + d')

# --- Misc settings ---
cfg.setSection(' --- General --- ')
CHECK_FOR_UPDATES_ON_LAUNCH = cfg.load('CHECK_FOR_UPDATES_ON_LAUNCH', True)
AUDIO = cfg.load('AUDIO', True)
CHECK_FOR_NEW_CLIPS_ON_LAUNCH = cfg.load('CHECK_FOR_NEW_CLIPS_ON_LAUNCH', True)
SEND_DELETED_FILES_TO_RECYCLE_BIN = cfg.load('SEND_DELETED_FILES_TO_RECYCLE_BIN', True)
MAX_BACKUPS = cfg.load('MAX_BACKUPS', 10)

# --- Rename formatting ---
cfg.setSection(' --- Renaming Clips --- ')
cfg.comment('''NAME_FORMAT variables:
    ?game - The game being played. The game's name will be swapped
            for an available alias if `USE_GAME_ALIASES` is True.
    ?date - The clip's timestamp. The timestamp's format is specified by
            DATE_FORMAT - see https://strftime.org/ for date formatting.
    ?count - A count given to create unique clip names. Only increments
             when the clip's name already exists. Best used when ?date
             is absent or isn't very specific (by default, DATE_FORMAT
             only saves the day, not the exact time of a clip).''')
RENAME = cfg.load('AUTO_RENAME_CLIPS', True)
USE_GAME_ALIASES = cfg.load('USE_GAME_ALIASES', True)
RENAME_FORMAT = cfg.load('NAME_FORMAT', '?game ?date #?count')
RENAME_DATE_FORMAT = cfg.load('DATE_FORMAT', '%y.%m.%d')
RENAME_COUNT_START_NUMBER = cfg.load('COUNT_START_NUMBER', 1)
RENAME_COUNT_PADDED_ZEROS = cfg.load('COUNT_PADDED_ZEROS', 1)

# --- Game aliases ---
cfg.setSection(' --- Game Aliases --- ')
cfg.comment('''This section defines aliases to use for ?game in `NAME_FORMAT` for renaming
clips. Not case-sensitive. || Usage: <ShadowPlay\'s name> = <custom name>''')
if USE_GAME_ALIASES:    # lowercase and remove double-spaces from names
    GAME_ALIASES = {' '.join(name.lower().split()): alias for name, alias in cfg.loadAllFromSection()}
    if not GAME_ALIASES:
        GAME_ALIASES = {
            "Left 4 Dead": "L4D1",
            "Left 4 Dead 2": "L4D2",
            "Battlefield 4": "BF4",
            "Dead by Daylight": "DBD",
            "Counter-Strike Global Offensive": "CSGO",
            "The Binding of Isaac Rebirth": "TBOI",
            "Team Fortress 2": "TF2",
            "Tom Clancy's Rainbow Six Siege": "R6"
        }
        for name, alias in GAME_ALIASES.items():
            cfg.load(name, alias)
else: GAME_ALIASES = {}

# --- Paths ---
cfg.setSection(' --- Paths --- ')
ICON_PATH = cfg.load('CUSTOM_ICON')
BACKUP_FOLDER = cfg.load('BACKUP_FOLDER', 'Backups')
HISTORY_PATH = cfg.load('HISTORY', 'history.txt')
UNDO_LIST_PATH = cfg.load('UNDO_LIST', 'undo.txt')

cfg.setSection(' --- Special Folders --- ')
cfg.comment('''These only apply if the associated path
in [Paths] is not an absolute path.''')
SAVE_HISTORY_TO_APPDATA_FOLDER = cfg.load('SAVE_HISTORY_TO_APPDATA_FOLDER', False)
SAVE_UNDO_LIST_TO_APPDATA_FOLDER = cfg.load('SAVE_UNDO_LIST_TO_APPDATA_FOLDER', False)
SAVE_BACKUPS_TO_APPDATA_FOLDER = cfg.load('SAVE_BACKUPS_TO_APPDATA_FOLDER', False)
SAVE_BACKUPS_TO_VIDEO_FOLDER = cfg.load('SAVE_BACKUPS_TO_VIDEO_FOLDER', BACKUP_FOLDER_HINT)

cfg.setSection(' --- Ignored Folders --- ')
cfg.comment('''Subfolders in the video folder that will be ignored during scans.
Names must be enclosed in quotes and comma-separated. Base names
only, i.e. '"Other", "Movies"'. Not case-sensitive.''')
lines = cfg.load('IGNORED_FOLDERS').split(',')
cleaned = (folder.strip().strip('"').strip().lower() for folder in lines)
IGNORE_VIDEOS_IN_THESE_FOLDERS = tuple(folder for folder in cleaned if folder)
del lines
del cleaned

# --- Tray menu ---
cfg.setSection(' --- Tray Menu --- ')
cfg.comment(f'''If `USE_CUSTOM_MENU` is True, {basename(CUSTOM_MENU_PATH)} is used to
create a custom menu. See {basename(CUSTOM_MENU_PATH)} for more information.
If deleted, set this to True and restart {TITLE}.''')
TRAY_ADVANCED_MODE = cfg.load('USE_CUSTOM_MENU', True)

# --- Basic mode (TRAY_ADVANCED_MODE = False) only ---
cfg.comment('Only used if `USE_CUSTOM_MENU` is False:', before='\n')
TRAY_SHOW_QUICK_ACTIONS = cfg.load('SHOW_QUICK_ACTIONS', True)
TRAY_RECENT_CLIPS_IN_SUBMENU = cfg.load('PUT_RECENT_CLIPS_IN_SUBMENU', False)
TRAY_QUICK_ACTIONS_IN_SUBMENU = cfg.load('PUT_QUICK_ACTIONS_IN_SUBMENU', True)

cfg.comment('''Valid left-click and middle-click actions:
    'open_log':             Opens this program's log file.
    'open_video_folder':    Opens the currently defined "Videos" folder.
    'open_install_folder':  Opens this program's root folder.
    'open_backup_folder':   Opens the currently defined backup folder.
    'play_most_recent':     Plays your most recent clip.
    'explore_most_recent':  Opens your most recent clip in Explorer.
    'delete_most_recent':   Deletes your most recent clip.
    'concatenate_last_two': Concatenates your two most recent clips.
    'undo':                 Undoes the last trim or concatenation.
    'clear_history':        Clears your clip history.
    'refresh':              Manually checks for new clips/refreshes existing ones.
    'check_for_updates':    Checks for a new release on GitHub to install.
    'about':                Shows an "About" window.
    'quit':                 Exits this program.''', before='\n')
TRAY_LEFT_CLICK_ACTION = cfg.load('LEFT_CLICK_ACTION', 'open_video_folder')
TRAY_MIDDLE_CLICK_ACTION = cfg.load('MIDDLE_CLICK_ACTION', 'play_most_recent')

# --- Recent clip menu settings ---
cfg.setSection(' --- Tray Menu Recent Clips --- ')
cfg.comment('''MAX_RECENT_CLIPS             - Total number of recent clips to display in the menu
                                (NOT the total number of clips saved in general).
PLAY_RECENT_CLIPS_ON_CLICK   - Play clips on click instead of opening them in explorer.
                                Only used if `EACH_RECENT_CLIP_HAS_SUBMENU` is False.
EACH_RECENT_CLIP_HAS_SUBMENU - If True, each clip will have a dedicated
                                submenu full of editing actions.
SUBMENUS_DISPLAY_EXTRA_INFO  - If True, a separator and two extra lines of info
                                will appear at the bottom of each clip's submenu.
                                Only used if `EACH_RECENT_CLIP_HAS_SUBMENU` is True.
EXTRA_INFO_DATE_FORMAT       - The date format used if `SUBMENUS_DISPLAY_EXTRA_INFO`
                                is True. See https://strftime.org/ for date formatting.''')
TRAY_RECENT_CLIP_COUNT = cfg.load('MAX_RECENT_CLIPS', 10)
TRAY_CLIPS_PLAY_ON_CLICK = cfg.load('PLAY_RECENT_CLIPS_ON_CLICK', True)
TRAY_RECENT_CLIPS_HAVE_UNIQUE_SUBMENUS = cfg.load('EACH_RECENT_CLIP_HAS_SUBMENU', True)
TRAY_RECENT_CLIPS_SUBMENU_EXTRA_INFO = cfg.load('SUBMENUS_DISPLAY_EXTRA_INFO', True)
TRAY_EXTRA_INFO_DATE_FORMAT = cfg.load('EXTRA_INFO_DATE_FORMAT', '%a %#D %#I:%M:%S%p')
cfg.comment('''RECENT_CLIP_NAME_FORMAT variables:
    ?date         - "1/17/22 12:09am" (see https://strftime.org/ for date formatting)
    ?recency      - "2 days ago"
    ?recencyshort - "2d"
    ?size         - "244.1mb"
    ?length       - "1:30" <- 90 seconds
    ?clip         - Name of a clip only.
    ?clipdir      - Name and immediate parent directory only.
    ?clippath     - Full path to a clip.''', before='\n')
TRAY_RECENT_CLIP_NAME_FORMAT = cfg.load('RECENT_CLIP_NAME_FORMAT', '(?recencyshort) - ?clip')
TRAY_RECENT_CLIP_DATE_FORMAT = cfg.load('RECENT_CLIP_DATE_FORMAT', '%#I:%M%p')
TRAY_RECENT_CLIP_DEFAULT_TEXT = cfg.load('EMPTY_SLOT_TEXT', ' --')

# --- Registry setting overrides ---
cfg.setSection(' --- Registry Overrides --- ')
cfg.comment('Used for overriding values obtained from the registry.')
VIDEO_FOLDER_OVERRIDE = cfg.load('VIDEO_FOLDER_OVERRIDE')
INSTANT_REPLAY_HOTKEY_OVERRIDE = cfg.load('INSTANT_REPLAY_HOTKEY_OVERRIDE')
TRAY_ALIGN_CENTER = cfg.load('ALWAYS_CENTER_ALIGN_TRAY_MENU_ON_OPEN', False)


# -----------------------
# Registry settings
# -----------------------
# confirm video folder has been set
if VIDEO_FOLDER_OVERRIDE:
    VIDEO_FOLDER = VIDEO_FOLDER_OVERRIDE.strip()
    logging.info('Overridden video directory: ' + VIDEO_FOLDER)
elif VIDEO_FOLDER is None:
    msg = ("ShadowPlay video path could not be read from your registry:\n\n"
           f"HKEY_CURRENT_USER\\{SHADOWPLAY_REGISTRY_PATH}\\DefaultPathW."
           "\n\nPlease set `VIDEO_FOLDER_OVERRIDE` in your config file.")
    abort_launch(2, 'No Video Folder Detected', msg)
else: logging.info('Video directory: ' + VIDEO_FOLDER)


# get Instant Replay hotkey from registry (each key is a separate value)
if not INSTANT_REPLAY_HOTKEY_OVERRIDE:
    try:
        import winreg
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, SHADOWPLAY_REGISTRY_PATH)

        # I have no idea how to actually decode numbers, modifiers, and function keys, so I cheated for all 3
        # NOTE: ctrl and alt do not have left/right counterparts
        # NOTE: number is plainly visible inside the encoded string, so we can just pluck it out
        modifier_keys = {'\t': 'tab', '': 'shift', '': 'ctrl', '': 'alt', '': 'capslock'}
        f_keys = {'P': 'f1', 'Q': 'f2', 'R': 'f3', 'S': 'f4', 'T': 'f5', 'U': 'f6',
                  'V': 'f7', 'W': 'f8', 'X': 'f9', 'Y': 'f10', 'Z': 'f11', '{': 'f12'}  # TODO: Function keys beyond F12
        total_keys_encoded_string = winreg.QueryValueEx(key, 'DVRHKeyCount')[0]
        total_keys = int(str(total_keys_encoded_string)[5])

        # NOTE: ShadowPlay settings are encoded in utf-16 and have a NULL character at the end
        hotkey = []
        for key_number in range(total_keys):
            hotkey_part = winreg.QueryValueEx(key, f'DVRHKey{key_number}')[0].decode('utf-16')[:-1]
            if hotkey_part in modifier_keys: hotkey.append(modifier_keys[hotkey_part])
            elif not hotkey_part.isupper(): hotkey.append(f_keys[hotkey_part])
            else: hotkey.append(hotkey_part)
        INSTANT_REPLAY_HOTKEY = ' + '.join(hotkey)

    except:
        msg = ("ShadowPlay Instant-Replay hotkey could not be read from your "
               f"registry:\n\nHKEY_CURRENT_USER\\{SHADOWPLAY_REGISTRY_PATH}\\"
               "DVRHKey___\n\nPlease set `INSTANT_REPLAY_HOTKEY_OVERRIDE` in "
               "your config file.\n\nFull error traceback: " + format_exc())
        abort_launch(2, 'No Instant-Replay Hotkey Detected', msg)
else: INSTANT_REPLAY_HOTKEY = INSTANT_REPLAY_HOTKEY_OVERRIDE.strip().lower()
logging.info(f'Instant replay hotkey: "{INSTANT_REPLAY_HOTKEY}"')


# get taskbar position from registry (NOTE: 0 = left, 1 = top, 2 = right, 3 = bottom)
if TRAY_ALIGN_CENTER: MENU_ALIGNMENT = win32.TPM_CENTERALIGN | win32.TPM_TOPALIGN
else:
    try:    # NOTE: this value takes a few moments to update after moving the taskbar
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\StuckRects3')
        taskbar_position = winreg.QueryValueEx(key, 'Settings')[0][12]

        # top-right alignment adjusts itself automatically for all EXCEPT left taskbars
        if taskbar_position == 0: MENU_ALIGNMENT = win32.TPM_LEFTALIGN | win32.TPM_BOTTOMALIGN
        else: MENU_ALIGNMENT = win32.TPM_RIGHTALIGN | win32.TPM_TOPALIGN
    except: logging.warning(f'Could not detect taskbar position for menu-alignment: {format_exc()}')
logging.info(f'Menu alignment: {MENU_ALIGNMENT}')


# ------------------------
# Other constants & paths
# ------------------------
# `CLIP_BUFFER` is how many recent clips should be `Clip` objects instead of just strings
# every clip in the menu + the first one outside the menu should be a `Clip` object
CLIP_BUFFER = max(2, TRAY_RECENT_CLIP_COUNT) + 1

# missing config/menu file (menu file only "missing" when custom menu enabled)
NO_CONFIG = not exists(CONFIG_PATH)
NO_MENU = TRAY_ADVANCED_MODE and not exists(CUSTOM_MENU_PATH)

# misc constants
TRAY_RECENT_CLIP_NAME_FORMAT_HAS_RECENCY = '?recency' in TRAY_RECENT_CLIP_NAME_FORMAT
TRAY_RECENT_CLIP_NAME_FORMAT_HAS_RECENCYSHORT = '?recencyshort' in TRAY_RECENT_CLIP_NAME_FORMAT
TRAY_RECENT_CLIP_NAME_FORMAT_HAS_CLIPDIR = '?clipdir' in TRAY_RECENT_CLIP_NAME_FORMAT

# constructing paths for various files/folders
if splitdrive(ICON_PATH)[0]: ICON_PATH = abspath(ICON_PATH)
else: ICON_PATH = pjoin(RESOURCE_FOLDER if exists(RESOURCE_FOLDER) else CWD, ICON_PATH)
if splitdrive(HISTORY_PATH)[0]: HISTORY_PATH = abspath(HISTORY_PATH)
else: HISTORY_PATH = pjoin(APPDATA_FOLDER if SAVE_HISTORY_TO_APPDATA_FOLDER else CWD, HISTORY_PATH)
if splitdrive(UNDO_LIST_PATH)[0]: UNDO_LIST_PATH = abspath(UNDO_LIST_PATH)
else: UNDO_LIST_PATH = pjoin(APPDATA_FOLDER if SAVE_UNDO_LIST_TO_APPDATA_FOLDER else CWD, UNDO_LIST_PATH)

# ensuring above paths are valid
if not exists(dirname(HISTORY_PATH)): makedirs(dirname(HISTORY_PATH))
if not exists(dirname(UNDO_LIST_PATH)): makedirs(dirname(UNDO_LIST_PATH))

# if icon isn't valid and we're running from the script -> warn and exit
# (when compiled, the .exe's icon is used as the backup)
if not IS_COMPILED and (not exists(ICON_PATH) or not os.path.isfile(ICON_PATH)):
    msg = 'No icon detected at `CUSTOM_ICON`: ' + ICON_PATH
    abort_launch(3, 'No icon detected', msg)

# ensuring backup folder is valid
if exists(BACKUP_FOLDER): BACKUP_FOLDER = abspath(BACKUP_FOLDER)
elif SAVE_BACKUPS_TO_VIDEO_FOLDER: BACKUP_FOLDER = pjoin(VIDEO_FOLDER, BACKUP_FOLDER)
elif SAVE_BACKUPS_TO_APPDATA_FOLDER: BACKUP_FOLDER = pjoin(APPDATA_FOLDER, BACKUP_FOLDER)
else: BACKUP_FOLDER = pjoin(CWD, BACKUP_FOLDER)

# `VIDEO_FOLDER` and `BACKUP_FOLDER` must be on the same drive or we'll get OSError 17
# if they are, warn user -> explain how to fix it -> abort launch
if (splitdrive(VIDEO_FOLDER)[0] != splitdrive(BACKUP_FOLDER)[0]
    or ismount(VIDEO_FOLDER) != ismount(BACKUP_FOLDER)):

    # quietly restore config/menu files if needed so user can reference them
    if NO_CONFIG: cfg.write()
    if NO_MENU: restore_menu_file()

    drive = splitdrive(VIDEO_FOLDER)[0]
    msg = ("Your video folder and the path for saving temporary backups "
           f"are on different drives. {TITLE} cannot backup and restore "
           "videos across drives without copying them back and forth.\n\n"
           f"Video folder: {VIDEO_FOLDER}\nBackup folder: {BACKUP_FOLDER}"
           f"\n\nTo resolve this conflict, open \"{basename(CONFIG_PATH)}"
           "\" and set `SAVE_BACKUPS_TO_VIDEO_FOLDER` to True or set "
           f"`BACKUP_FOLDER` to an absolute path on the {drive} drive.")
    abort_launch(2, 'Invalid Backup Directory', msg, 0x00040030)  # !-symbol, stay on top

# ensure `BACKUP_FOLDER` exists, but only once we've dealt with drive-conflict
if not exists(BACKUP_FOLDER): makedirs(BACKUP_FOLDER)


# ---------------------
# Custom Pystray class
# ---------------------
WM_LBUTTONUP = 0x0202
WM_RBUTTONUP = 0x0205
WM_MBUTTONUP = 0x0208
class Icon(pystray._win32.Icon):
    ''' This subclass auto-updates the menu before opening to allow dynamic
        titles/actions to always be up-to-date, adds support for middle-clicks
        and custom menu alignments, and allows icons to be passed as strings
        rather than using `PIL.Image.Image` (removing it as a dependency by
        just assuming a given .ICO is valid, and using the .exe's icon if it
        isn't). See original _win32.Icon class for original comments. '''

    def _on_notify(self, wparam, lparam):
        ''' Adds auto-updating, middle-click, and menu alignment support. '''
        if lparam == WM_LBUTTONUP:
            self()

        elif lparam == WM_MBUTTONUP:
            if MIDDLE_CLICK_ACTION:
                MIDDLE_CLICK_ACTION()

        elif self._menu_handle and lparam == WM_RBUTTONUP:
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


    def _assert_icon_handle(self):
        ''' Removes usage of `serialized_image` context manager and thus the
            the dependency on `PIL.Image.Image` by assuming `self._icon`
            is the path to a valid .ICO file. Uses .exe's icon if needed. '''
        if self._icon_handle: return
        args = (win32.IMAGE_ICON, 0, 0, win32.LR_DEFAULTSIZE | win32.LR_LOADFROMFILE)

        try:
            handle = win32.LoadImage(None, self._icon, *args)
            if handle is None: raise
            self._icon_handle = handle
            return
        except:
            if IS_COMPILED:             # if we're compiled, take the .exe's icon
                try:
                    # https://stackoverflow.com/questions/90775/how-do-you-load-an-embedded-icon-from-an-exe-file-with-pywin32
                    import win32api     # these libraries cost almost nothing to import...
                    import win32gui     # ...and don't add any files to our compilation

                    # NOTE: for our current icon, index 4 is the best icon, even at different scales
                    RT_ICON = 3         # this is so we don't need `win32con.RT_ICON`
                    icon_index = 4

                    resource = win32api.LoadResource(None, RT_ICON, icon_index)
                    handle = win32gui.CreateIconFromResource(resource, True)
                    if handle is None: raise
                    self._icon_handle = handle
                    return logging.warning(f'Custom icon at {self._icon} was invalid. Using .exe\'s icon.')
                except: logging.warning(f'.exe\'s icon at index {icon_index} was not valid.')

        # warn and exit. use f-string for warning in case `self._icon` isn't a string
        msg = f'The icon at `CUSTOM_ICON` is not a valid .ICO file: {self._icon}'
        show_message('Invalid icon', msg, 0x00040010)
        sys.exit(3)


# -------------------------
# Custom keyboard listener
# -------------------------
INSTANT_REPLAY_HOTKEY_SCANCODES = tuple(sorted(keyboard.key_to_scan_codes(key.strip()) for key in INSTANT_REPLAY_HOTKEY.split('+')))
ALL_INSTANT_REPLAY_HOTKEY_SCANCODES = tuple(code for code_tuple in INSTANT_REPLAY_HOTKEY_SCANCODES for code in code_tuple)
ACTUAL_INSTANT_REPLAY_HOTKEY = tuple(sorted(code_tuple[0] for code_tuple in INSTANT_REPLAY_HOTKEY_SCANCODES))
KEYPAD_DUPLICATES = (71, 72, 73, 75, 77, 79, 80, 81, 82, 83)   # 7, 8, 9, 4, 6, 1, 2, 3, 0, 'decimal'
def pre_process_event(self, event):
    ''' This is an *extremely* convulted way of dealing with
        two major shortcomings with the keyboard library:
            A. Using hotkeys while other keys are held down (like ShadowPlay)
            B. Preventing erronous hotkey triggers when pressing
               buttons that share scancodes with the number pad
               (see AutoCutter.__init__.key_to_scan_codes_no_keypad for more)

        Without this fix, you can accidentally save clips without the script
        detecting it and common shortcuts will trigger Alt + Number hotkeys.

        These shortcomings actually appear to be intentional design choices in
        the keyboard library, as they allow the code to be MUCH faster and far
        more elegant. Elegance which I have bludgeoned with a sledgehammer.

        This took several hours to figure out. It is horrible. Ironically,
        however, this is probably still better than other keyboard hooking
        libraries. This could be much better and much more generalized, but
        I will not be the one to figure it out. I can't believe I even made
        it this far in the first place. Still worth it, though. '''

    scan_code = event.scan_code
    for key_hook in self.nonblocking_keys[scan_code]:
        key_hook(event)

    # A. Allow INSTANT_REPLAY_HOTKEY to be detected even while other keys are held down (like ShadowPlay does)
    if scan_code in ALL_INSTANT_REPLAY_HOTKEY_SCANCODES:        # ALL_INSTANT_REPLAY_HOTKEY_SCANCODES is flattened
        with keyboard._pressed_events_lock:                     # NOTE: ONLY this hotkey works like this because this...
            hotkey = tuple(sorted(keyboard._pressed_events))    # ...implementation is too slow to do every hotkey this way
        for valid_keys in INSTANT_REPLAY_HOTKEY_SCANCODES:      # each "key" is a tuple of possible scan codes that key uses
            if not any(key in hotkey for key in valid_keys):    # see if at least one scan code in each tuple is being pressed
                break
        else: hotkey = ACTUAL_INSTANT_REPLAY_HOTKEY             # nobreak -> set hotkey to something `keyboard` will recognize

    # B. Preventing erronous hotkey triggers when pressing buttons that share scancodes with the number pad (Alt + Arrows, etc.)
    #    I have spent many days trying to figure out a simple, general purpose solution better than this one. I don't think
    #    one exists with the limited information we have. Improving this will require a much more complicated system.
    elif event.is_keypad and scan_code in KEYPAD_DUPLICATES and event.event_type == 'down':
        with keyboard._pressed_events_lock:
            del keyboard._pressed_events[scan_code]
            hotkey = tuple(sorted(keyboard._pressed_events))

    # The default keyboard library code, as seen in keyboard\__init__.py
    else:
        with keyboard._pressed_events_lock:
            hotkey = tuple(sorted(keyboard._pressed_events))    # "hotkey" is a tuple of scan codes being pressed right now
    for callback in self.nonblocking_hotkeys[hotkey]:
        callback(event)
    return scan_code or (event.name and event.name != 'unknown')


# replace the keyboard libraries event processor with our own
keyboard._KeyboardListener.pre_process_event = pre_process_event


# ---------------------
# Clip class
# ---------------------
class Clip:
    __slots__ = ('working', 'path', 'name', 'game', 'time', 'raw_size', 'size', 'date', 'full_date', 'length', 'length_string', 'length_size_string')
    def __repr__(self): return self.name

    def __init__(self, path, stat, rename=False):
        self.working = False

        path, self.name = auto_rename_clip(path) if rename else (abspath(path), basename(path))   # abspath for consistent formatting
        size = f'{(stat.st_size / 1048576):.1f}mb'

        self.path = path
        self.game = basename(dirname(path))
        self.time = stat.st_ctime
        self.raw_size = stat.st_size
        self.size = size
        self.date = datetime.strftime(datetime.fromtimestamp(stat.st_ctime), TRAY_RECENT_CLIP_DATE_FORMAT)
        self.full_date = datetime.strftime(datetime.fromtimestamp(stat.st_ctime), TRAY_EXTRA_INFO_DATE_FORMAT)

        length = get_video_duration(path)   # NOTE: this could be faster, but uglier
        length_int = int(length)
        length_string = f'{length_int // 60}:{length_int % 60:02}'
        self.length = length
        self.length_string = length_string
        self.length_size_string = f'Length: {length_string} ({size})'


    def refresh(self, path=None, stat=None):
        ''' Refreshes various statistics for the clip, including creation
            time, size, and length. Pass `path` and `stat` as minor
            optimizations if you already have direct access to them. '''
        path = path or self.path
        stat = stat or getstat(path)
        size = f'{(stat.st_size / 1048576):.1f}mb'
        self.time = stat.st_ctime
        self.raw_size = stat.st_size
        self.size = size

        length = get_video_duration(path)
        length_int = int(length)
        length_string = f'{length_int // 60}:{length_int % 60:02}'
        self.length = length
        self.length_string = length_string
        self.length_size_string = f'Length: {length_string} ({size})'


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
    __slots__ = ('waiting_for_clip', 'protected_paths', 'last_clips')

    def __init__(self):
        start = time.time()
        self.waiting_for_clip = False

    # --- protecting backup paths ---
        protected_paths = []
        for path in self.get_all_backups():
            path_dirname, path_basename = splitpath(path)
            if '.' not in path_basename[20:]: continue                  # skip files with empty basenames
            protected_path = pjoin(VIDEO_FOLDER, basename(path_dirname), path_basename[20:])
            protected_paths.append(protected_path)
        logging.info(f'Current protected paths: {protected_paths}')
        self.protected_paths = protected_paths

    # --- importing history ---
        # if history file exists, import clips
        if exists(HISTORY_PATH):
            with open(HISTORY_PATH, 'r') as history:
                # get all valid unique paths from history file, create a buffer of clip objects, then start caching paths as strings outside buffer
                lines = []
                addpath = lines.append
                for path in reversed(history.read().splitlines()):      # reversed() is an iterable, not an actual list (no performance loss)
                    if path and exists(path) and path not in lines:     # faster than using a set
                        addpath(path)

                logging.info(f'History file parsed in {time.time() - start:.3f} seconds.')

                # sort history by creation date to resolve most issues before they arise
                lines.sort(key=lambda clip: getstat(clip).st_ctime, reverse=True)
                last_clips = [Clip(path, getstat(path), rename=False)
                              if index < CLIP_BUFFER
                              else path
                              for index, path in enumerate(lines)]
                last_clips.reverse()                                    # .reverse() is a very fast operation
                if last_clips: logging.info(f'Previous {len(last_clips)} clip{"s" if len(last_clips) != 1 else ""} loaded: {last_clips}')
                else: logging.info('No previous clips detected.')
                logging.info(f'Previous clips loaded in {time.time() - start:.2f} seconds.')

                self.last_clips = last_clips
                if CHECK_FOR_NEW_CLIPS_ON_LAUNCH: self.check_for_clips(manual_update=True)
                del lines

        # no history file -- run first-time setup
        else:
            self.last_clips = []

            # put together simple, dynamic message for message box
            msg = ("It appears you do not have a history file. Would you "
                   f"like to add all existing clips in {VIDEO_FOLDER}? ")
            if IGNORE_VIDEOS_IN_THESE_FOLDERS:
                msg += '\n\nThe following subfolders will be ignored:\n\t"'
                msg += '"\n\t"'.join(IGNORE_VIDEOS_IN_THESE_FOLDERS) + '"'
            else:
                msg += ("Your `IGNORED_FOLDERS` setting is empty, so "
                        "all subfolders will be scanned.")
            msg += (f"\n\nAny clips matching ShadowPlay's naming format will "
                    "be renamed to match your own `NAME_FORMAT` setting: "
                    f"\n\t\"{RENAME_FORMAT}\"")
            msg += f'\n\nClick cancel to exit {TITLE}.'

            # ?-symbol, stay on top, Yes/No/Cancel
            response = show_message('Welcome to ' + TITLE, msg, 0x00040023)
            if response == 2:       # Cancel/X
                logging.info('Cancel selected on welcome dialog, closing...')
                sys.exit(1)
            elif response == 7:     # No
                logging.info('No selected on welcome dialog, not retroactively adding clips.')
            elif response == 6:     # Yes
                logging.info('Yes selected on welcome dialog, looking for pre-existing clips...')
                self.check_for_clips(manual_update=True, from_time=0)

    # --- hotkeys ---
        # define instant replay hotkey BEFORE we do the dumb garbage below it
        keyboard.add_hotkey(INSTANT_REPLAY_HOTKEY, self.check_for_clips)

        # This edits `keyboard.key_to_scan_codes` so that it removes any scan codes from
        # `KEYPAD_DUPLICATES`, unless it's the primary (smallest) scan code for that key.
        # This is simpler than what I was doing before - editing `keyboard.add_hotkey`.
        _key_to_scan_codes = keyboard.key_to_scan_codes
        def key_to_scan_codes_no_keypad(*args, **kwargs):
            codes = _key_to_scan_codes(*args, **kwargs)
            return (codes[0], *(code for code in codes[1:] if code not in KEYPAD_DUPLICATES))
        keyboard.key_to_scan_codes = key_to_scan_codes_no_keypad

        keyboard.add_hotkey(CONCATENATE_HOTKEY, self.concatenate_last_clips)
        keyboard.add_hotkey(DELETE_HOTKEY, self.delete_clip)
        keyboard.add_hotkey(UNDO_HOTKEY, self.undo)
        for key, length in LENGTH_DICTIONARY.items():
            keyboard.add_hotkey(key, self.trim_clip, args=(length,))
        logging.info(f'Auto-cutter initialized in {time.time() - start:.2f} seconds.')

    # ---------------------
    # Helper methods
    # ---------------------
    def wait(self, verb=None, alert=None, min_clips=1):
        ''' Checks if a clip is queued up in another thread and waits for it.
            Returns `False` If `self.last_clips` is smaller than `min_clips`.
            Plays sound effect specified by `alert`. Describes action in log
            message with `verb` if specified, otherwise `alert` is used. '''
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
        ''' Safely gets a clip at a specified `index` or `path`. Recursively
            calls itself until a valid clip is obtained if `patient` is True,
            otherwise raises an error if desired clip cannot be obtained.
            `alert` plays the given sound effect, `verb` describes the action
            for log messages, and `min_clips` is the number of clips needed
            to exist if `patient` is True. '''
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


    def cache_clip(self):
        ''' Converts first `Clip` object outside `CLIP_BUFFER`
            to a string. Removes outdated/invalid clips. '''
        try:
            last_clips = self.last_clips
            cache_index = -(CLIP_BUFFER + 1)    # +1 to get first clip outside buffer
            while isinstance(clip := last_clips[cache_index], Clip):
                if exists(clip.path):           # if cached Clip object exists, convert to string and break loop
                    logging.info(f'Caching {clip} at index {cache_index}.')
                    last_clips[cache_index] = clip.path
                    break
                last_clips.pop(cache_index)     # if cached clip doesn't exist, pop and try next clip
        except IndexError: pass                 # IndexError -> pop was out of range, pass
        except: logging.error(f'(!) Error while caching clip at index {cache_index} <len(last_clips)={len(last_clips)}>: {format_exc()}')


    def insert_clip(self, path, index=None, return_index=True):
        ''' Inserts `path` within `self.last_clips` based on its creation date
            (unless `index` is specified). Meant for retroactively adding old
            clips, not appending new ones. Returns the clip's index if
            `return_index` is True, else the final clip object/string. '''
        stat = None
        path = abspath(path)
        cache_index = len(self.last_clips) - CLIP_BUFFER
        logging.info(f'Inserting "{path}" {f"at index {index}" if index else "based on creation date"}.')

        if index is None:
            stat = getstat(path)
            index = self.index_for_time(stat.st_ctime)

        # clips below the cache_index will be strings instead of full `Clip` objects
        if index >= cache_index:
            stat = stat or getstat(path)
            path = Clip(path, stat, rename=False)
            self.last_clips.insert(index, path)
            self.cache_clip()
        else: self.last_clips.insert(index, path)

        return index if return_index else path


    def update_clip(self, path, return_index=False):
        ''' Refreshes a clip at `path`. If no clip object/string exists
            for `path`, one is inserted using `self.insert_clip`. Returns
            the clip's index if `return_index` is True, else the clip
            object/string itself. `path` is assumed to exist. '''
        logging.info(f'Updating clip for path "{path}".')
        try:                # self.get_clip() is not needed here
            index = self.index(path)
            clip = self.last_clips[index]
            if isinstance(clip, Clip): clip.refresh(path)
            return index if return_index else clip
        except ValueError:  # insert clip if self.index() raised a ValueError
            return self.insert_clip(path, return_index=return_index)


    def pop(self, index=-1):
        ''' Wrapper for popping from the last_clips list that converts cached paths from strings to
            Clip objects, if necessary, with a failsafe for non-existent cached paths included.
            Attempts to return popped value on error, if possible -- otherwise returns None. '''
        last_clips = self.last_clips
        popped = None
        try:
            popped = last_clips.pop(index)

            # convert string at end of clip buffer to a Clip object
            # this could be uncache_clip(), but we only use this here
            cache_index = -CLIP_BUFFER          # index of the very last clip in buffer
            while isinstance(clip := last_clips[cache_index], str):
                if exists(clip):                # if cached string exists, convert to Clip object and break loop
                    logging.info(f'Uncaching {clip} at index {cache_index}.')
                    last_clips[cache_index] = Clip(clip, getstat(clip), rename=False)
                    break
                last_clips.pop(cache_index)     # if cached clip doesn't exist, pop and try next clip

        except IndexError: pass                 # IndexError -> pop was out of range, pass and return None
        except: logging.error(f'(!) Error while popping clip at index {index} <cache_index={cache_index}, len(last_clips)={len(last_clips)}>: {format_exc()}')
        return popped


    def index(self, path):
        ''' Returns the index of a given `path` in `self.last_clips`. '''
        path = abspath(path)
        for index, clip in enumerate(self.last_clips):
            if isinstance(clip, Clip):
                if clip.path == path: return index
            elif clip == path: return index
        raise ValueError(f'{path} not in last_clips')


    def index_for_time(self, time):
        ''' Returns the index in `self.last_clips` that `time`
            would occupy if a clip created at that time existed. '''
        # starts with most recent clips to check `Clip` objects first
        for index, clip in enumerate(reversed(self.last_clips)):
            if isinstance(clip, Clip):
                if clip.time <= time: return len(self.last_clips) - index
            elif getstat(clip).st_ctime <= time: return len(self.last_clips) - index
        return 0


    def get_all_backups(self):
        ''' Returns all backups in `BACKUP_FOLDER` as a flattened list.
            Assumes all backups start with `time.time_ns()`. '''
        all_backups = []
        for folder in listdir(BACKUP_FOLDER):
            try:                                    # get all backup .mp4s and delete empty subfolders
                subfolder = pjoin(BACKUP_FOLDER, folder)
                files = listdir(subfolder)
                if files: all_backups.extend(pjoin(subfolder, file) for file in files if file[:19].isnumeric())
                else: os.rmdir(subfolder)
            except: pass
        return all_backups


    def refresh_backups(self, *protected_paths):
        ''' Deletes outdated backup files and empty backup folders. Ignores
            `protected_paths` while counting, even if `MAX_BACKUPS` is 0.
            Outdated files are removed from `self.protected_paths`.
            Assumes all backups start with `time.time_ns()`. '''
        old_backups = 0
        for path in sorted(self.get_all_backups(), reverse=True):
            if path not in protected_paths:
                old_backups += 1
                if old_backups >= MAX_BACKUPS:      # backup is too old
                    try:    # remove backup, remove its folder if empty, and remove its protected status
                        remove(path)
                        path_dirname, path_basename = splitpath(path)
                        if not listdir(dirname(path)): os.rmdir(path_dirname)

                        protected_path = pjoin(VIDEO_FOLDER, basename(path_dirname), path_basename[20:])
                        try: self.protected_paths.remove(protected_path)
                        except ValueError: pass
                    except: logging.warning(f'Failed to delete outdated backup "{path}": {format_exc()}')

        # add new protected paths
        for path in protected_paths:
            path_dirname, path_basename = splitpath(path)
            protected_path = pjoin(VIDEO_FOLDER, basename(path_dirname), path_basename[20:])
            logging.info(f'Adding protected path: {protected_path}')
            self.protected_paths.append(protected_path)

    # ---------------------
    # Acquiring clips
    # ---------------------
    def check_for_clips(self, manual_update=False, from_time=None):
        self.waiting_for_clip = True
        Thread(target=self.check_for_clips_thread, args=(manual_update, from_time), daemon=True).start()


    def check_for_clips_thread(self, manual_update=False, from_time=None):
        ''' Scans `VIDEO_FOLDER` for new .mp4 files to add as clips to
            `self.last_clips`. Runs automatically after every instant-replay.
            Only checks the base-level files inside `VIDEO_FOLDER`'s
            base-level subfolders, as these are the only places ShadowPlay
            will save a recording. Files within nested subfolders or within
            `VIDEO_FOLDER`'s base directory will be ignored. Ignores the
            backup folder (if present) and `IGNORE_VIDEOS_IN_THESE_FOLDERS`.

            If `manual_update` is False:
                - A 3 second delay is used to wait for the last instant-replay
                  to finish saving.
                - The minimum timestamp for new clips is the current time
                  minus one second.
                - The scan ends after the first "new" clip is discovered.

            If `manual_update` is True:
                - The minimum timestamp is the creation date of the latest
                  clip in `self.last_clips`, unless `from_time` is specified.
                - All "new" clips discovered are added.
                - After the scan, `self.last_clips` is sorted and the cache is
                  verified to ensure `Clip` objects outside `CLIP_BUFFER` have
                  become strings, and vice versa.

            NOTE: Detecting ShadowPlay videos via their encoding is possible.
            This would prevent non-ShadowPlay clips from being detected for
            good, but I'm not decided as to whether detecting non-ShadowPlay
            clips is a good thing or not. The limited scanning range mentioned
            above is a good middle-ground for now:
            https://github.com/rebane2001/NvidiaInstantRename/blob/mane/InstantRenameCLI.py '''
        try:

            # determine the timestamp that all new videos should be after
            last_clips = self.last_clips
            if not manual_update:   # auto-update -> wait for instant replay to finish saving
                logging.info('Instant replay detected! Waiting 3 seconds...')
                last_clip_time = time.time() - 1
                time.sleep(3)
            elif from_time is not None: last_clip_time = from_time
            else:                   # if no clips in last_clips -> use history file's creation date
                try: last_clip_time = last_clips[-1].time
                except IndexError:  # if no history file -> use script's start time
                    try: last_clip_time = getstat(HISTORY_PATH).st_ctime
                    except: last_clip_time = SCRIPT_START_TIME
            logging.info(f'Scanning {VIDEO_FOLDER} for videos')

            # ONLY look at the base files of the base subfolders
            for filename in listdir(VIDEO_FOLDER):
                if filename in IGNORE_VIDEOS_IN_THESE_FOLDERS: continue
                folder = pjoin(VIDEO_FOLDER, filename)
                if folder == BACKUP_FOLDER: continue    # user might use absolute path for backup folder

                if isdir(folder):
                    for file in listdir(folder):
                        path = pjoin(folder, file)
                        stat = getstat(path)            # skip non-mp4 files ↓
                        if last_clip_time < stat.st_ctime and file[-4:] == '.mp4':
                            logging.info(f'New video detected: {file}')
                            while get_video_duration(path) == 0:
                                logging.info('Video duration is still 0, retrying...')
                                time.sleep(0.2)         # if duration is 0, the video is still being saved

                            clip = Clip(path, stat, rename=RENAME)
                            last_clips.append(clip)
                            logging.info(f'Memory usage after adding clip: {get_memory():.2f}mb\n')
                            if manual_update: continue  # don't stop after first video on manual scans
                            else: return self.cache_clip()

        except:
            logging.error(f'(!) Error while checking for new clips: {format_exc()}')
            play_alert('error')

        finally:
            if manual_update:       # sort last_clips and then verify the cache
                last_clips.sort(key=lambda clip: clip.time if isinstance(clip, Clip) else getstat(clip).st_ctime)
                for index, clip in enumerate(reversed(last_clips)):
                    if isinstance(clip, Clip) and index >= CLIP_BUFFER:
                        last_clips[-index - 1] = clip.path
                    elif isinstance(clip, str) and index < CLIP_BUFFER:
                        last_clips[-index - 1] = Clip(path, stat, rename=RENAME)
                logging.info('Manual scan complete.')
            self.waiting_for_clip = False
            gc.collect(generation=2)

    # ---------------------
    # Clip actions
    # ---------------------
    def trim_clip(self, length, index=-1, patient=True):
        ''' Trims the clip at `index` down to the last `length` seconds. Clip
            is edited in-place (the original is moved to `BACKUP_FOLDER`). '''
        try:
            clip = self.get_clip(index, verb='Trim', alert=length, min_clips=1, patient=patient)
            clip_path = clip.path
            clip_length = clip.length
            logging.info(f'Trimming clip {clip.name} to {length} seconds')
            logging.info(f'Clip is {clip_length:.2f} seconds long.')
            logging.info(f'{clip_length}: {type(clip_length)} | {length}: {type(length)}')
            if clip_length <= length: return logging.info(f'(?) Video is only {clip_length:.2f} seconds long and cannot be trimmed to {length} seconds.')

            relative_temp_path = pjoin(clip.game, f'{time.time_ns()}.{clip.name}')
            temp_path = pjoin(BACKUP_FOLDER, relative_temp_path)
            renames(clip_path, temp_path)

            try:
                cmd = f'ffmpeg -y -i "{temp_path}" -ss {clip_length - length} -c:v copy -c:a copy "{clip_path}" -hide_banner -loglevel warning'
                logging.info(cmd)
                process = subprocess.Popen(cmd, shell=True)
                process.wait()

                with open(UNDO_LIST_PATH, 'w') as undo:
                    undo.write(f'{relative_temp_path} -> {clip_path} -> Trimmed to {length:g} seconds\n')
                self.refresh_backups(temp_path)

                setctime(clip_path, clip.time)          # ensure edited clip retains original creation time
                clip.refresh()
                logging.info(f'Trim to {length} seconds successful.\n')
            except:
                logging.error(f'(!) Error WHILE trimming clip: {format_exc()}')
                if exists(clip_path): remove(clip_path)
                renames(temp_path, clip_path)
                logging.info('Successfully restored clip after error.')

        except (IndexError, AssertionError): return
        except:
            logging.error(f'(!) Error BEFORE trimming clip: {format_exc()}')
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
            logging.info(f'Concatenating clips "{clip_path1}" and "{clip_path2}"')

            base, ext = splitext(clip_path1)
            temp_path = f'{base}_concat{ext}'       # the temporary name for our final .mp4 file
            text_path = f'{base}_concatlist.txt'    # write list of clips to text file (with / as separator and single quotes to avoid ffmpeg errors)
            with open(text_path, 'w') as txt:
                txt.write(f"file '{clip_path1.replace(sep, '/')}'\nfile '{clip_path2.replace(sep, '/')}'")

            cmd = f'ffmpeg -y -f concat -safe 0 -i "{text_path}" -c copy "{temp_path}" -hide_banner -loglevel warning'
            logging.info(cmd)
            process = subprocess.Popen(cmd, shell=True)
            process.wait()
            delete(text_path)

            # attempt to move first clip to backup folder
            relative_temp_path1 = pjoin(clip1.game, f'{time.time_ns()}.{clip1.name}')
            temp_path1 = pjoin(BACKUP_FOLDER, relative_temp_path1)
            try: renames(clip_path1, temp_path1)
            except:
                delete(temp_path)                               # delete concat clip to avoid confusing user
                logging.error(f'(!) Error while concatenating last two clips: {format_exc()}')
                return play_alert('error')

            # attempt to move second clip to backup folder
            relative_temp_path2 = pjoin(clip2.game, f'{time.time_ns()}.{clip2.name}')
            temp_path2 = pjoin(BACKUP_FOLDER, relative_temp_path2)
            try: renames(clip_path2, temp_path2)
            except:
                delete(temp_path)                               # delete concat clip to avoid confusing user
                renames(temp_path1, clip_path1)                 # restore first clip if second clip failed
                logging.error(f'(!) Error while concatenating last two clips: {format_exc()}')
                return play_alert('error')

            with open(UNDO_LIST_PATH, 'w') as undo:
                undo.write(f'{relative_temp_path1} -> {relative_temp_path2} -> {clip_path1} -> {clip_path2} -> Concatenated\n')
            self.refresh_backups(temp_path1, temp_path2)

            renames(temp_path, clip_path1)
            setctime(clip_path1, getstat(temp_path1).st_ctime)  # ensure edited clip retains original creation time
            self.pop(index)
            clip1.refresh(clip_path1)                           # pass `clip_path1` as slight optimization
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


    # TODO unused and not fully implemented. should not happen automatically.
    # TODO also, these two methods are the only usage of `Clip.working`.
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
            renames(old_path, new_path)
            logging.info(f'Video size is {clip.size}. Compressing...')
            #ffmpeg(clip.path, f'-i "%tp" -vcodec libx265 -crf 28 "{clip.path}"')
            logging.info(f'New compressed size is: {clip.size}')
            renames(new_path, old_path)
            clip.path = old_path
            clip.refresh(old_path)
        except:
            logging.error(f'(!) Error while compressing clip: {format_exc()}')
            play_alert('error')
        finally:
            clip.working = False


    def undo(self, *args, patient=True):    # NOTE: *args used to capture pystray's unused args
        ''' Undoes an action described in `UNDO_LIST_PATH`. Each line
            represents one action that can be undone, and consists of
            multiple strings required for the undo, separated by "-->". '''
        try:
            with open(UNDO_LIST_PATH, 'r') as undo:
                line = undo.readline().strip().split(' -> ')

                # trimming
                # - old -> unedited video's backup name
                # - new -> new, edited video's full path
                if len(line) == 3:
                    old, new, action = line

                    alert = 'Undoing Trim' if 'trim' in action.lower() else 'Undo'
                    if patient and not self.wait(verb=f'Undo "{action}"', alert=alert): return

                    if exists(new): remove(new)                 # delete edited video (if it still exists)
                    renames(pjoin(BACKUP_FOLDER, old), new)     # super-rename in case folder was deleted

                    self.update_clip(path=new)                  # verify and refresh `new`'s clip
                    logging.info(f'Undo completed for "{action}" on clip "{new}"')

                # concatenation
                # - old, old2 -> unedited videos' backup names
                # - new       -> new, edited video's full path
                # - new2      -> the video that was deleted when the original concat happened
                elif len(line) == 5:
                    old, old2, new, new2, action = line

                    alert = 'Undoing Concatenation' if 'concat' in action.lower() else 'Undo'
                    if patient and not self.wait(verb=f'Undo "{action}"', alert=alert): return

                    if exists(new): remove(new)                 # delete edited video (if it still exists)
                    if exists(new2):                            # `new2` has been replaced? (this should NOT happen)
                        logging.warn('(!) Clip shares name with concatenated clip that should have been protected: ' + new2)
                        base, ext = splitext(new2)
                        new2 = f'{base} (pre-concat){ext}'      # add " (pre-concat)" marker to `new2`
                    renames(pjoin(BACKUP_FOLDER, old), new)     # super-rename in case folder was deleted
                    rename(pjoin(BACKUP_FOLDER, old2), new2)

                    # verify/refresh `new` and re-add `new2` (concat removes `new2` from the list)
                    index = self.update_clip(path=new, return_index=True)
                    self.insert_clip(new2, index=index + 1)
                    logging.info(f'Undo completed for "{action}" on clips "{new}" and "{new2}"')

            remove(UNDO_LIST_PATH)
        except:
            logging.error(f'(!) Error while undoing last action: {format_exc()}')
            play_alert('error')


###########################################################
if __name__ == '__main__':
    try:
        check_for_updates(manual=False)
        verify_ffmpeg()
        verify_config_files()

        logging.info(f'Memory usage before initializing AutoCutter class: {get_memory():.2f}mb')
        cutter = AutoCutter()
        logging.info(f'Memory usage after initializing AutoCutter class: {get_memory():.2f}mb')

        if SEND_DELETED_FILES_TO_RECYCLE_BIN: import send2trash

        # ---------------------
        # Tray-icon functions
        # ---------------------
        def quit_tray(icon):
            ''' Quits pystray `icon`, saves history,
                does final cleanup, and exits script. '''
            try:                                # close icon and save history if icon exists
                logging.info('Closing system tray icon and exiting.')
                tracemalloc.stop()
                icon.visible = False
                icon.stop()

                logging.info(f'Clip history: {cutter.last_clips}')
                with open(HISTORY_PATH, 'w') as history:
                    history.write('\n'.join(c.path if isinstance(c, Clip) else c for c in cutter.last_clips))
            except: pass

            try: atexit.unregister(quit_tray)   # unregister quit_tray so we don't run it twice
            except: pass

            try: sys.exit(0)
            except SystemExit: pass             # avoid harmless yet annoying SystemExit error


        def get_clip_tray_title(index: int, default: str = TRAY_RECENT_CLIP_DEFAULT_TEXT) -> str:
            ''' Returns the title of a recent clip item by its `index`.
                If no clip exists at that index, `default` is returned.
                Refreshes and removes edited/deleted clips. '''
            try:
                clip = cutter.last_clips[index]
                path = clip.path
                try:        # assume all paths exist and react accordingly
                    stat = getstat(path)
                    if stat.st_size != clip.raw_size: clip.refresh(path, stat)
                except FileNotFoundError:
                    cutter.pop(index)
                    return get_clip_tray_title(index)

                # get loose estimate of how long ago it was, if necessary
                if TRAY_RECENT_CLIP_NAME_FORMAT_HAS_RECENCY:            # ?recency
                    time_delta = time.time() - clip.time
                    d = int(time_delta // 86400)
                    h = int(time_delta // 3600)
                    m = int(time_delta // 60)
                    if TRAY_RECENT_CLIP_NAME_FORMAT_HAS_RECENCYSHORT:   # ?recencyshort
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

                return (TRAY_RECENT_CLIP_NAME_FORMAT
                        .replace('?date', clip.date)
                        .replace('?recencyshort', time_delta_string)    # ?recencyshort first so ?recency is replaced
                        .replace('?recency', time_delta_string)
                        .replace('?size', clip.size)
                        .replace('?length', clip.length_string)
                        .replace('?clippath', path)
                        .replace('?clipdir', sepjoin(path.split(sep)[-2:]) if TRAY_RECENT_CLIP_NAME_FORMAT_HAS_CLIPDIR else '')
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
            'open_video_folder':    lambda: os.startfile(VIDEO_FOLDER),
            'open_install_folder':  lambda: os.startfile(CWD),
            'open_backup_folder':   lambda: os.startfile(BACKUP_FOLDER),
            'play_most_recent':     lambda: cutter.open_clip(play=True),
            'explore_most_recent':  lambda: cutter.open_clip(play=False),
            'delete_most_recent':   lambda: cutter.delete_clip(),               # small RAM drop by making these not lambdas
            'concatenate_last_two': lambda: cutter.concatenate_last_clips(),    # 26.0mb -> 25.8mb on average
            'clear_history':        lambda: cutter.last_clips.clear(),
            'refresh':              lambda: cutter.check_for_clips(manual_update=True),
            'check_for_updates':    check_for_updates,
            'about':                about,
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
            exit_item_exists = False                # variable for making sure an exit item is included
            def evaluate_menu(item_pairs, menu):    # function for recursively solving menus/submenus and exporting them to a list
                global exit_item_exists
                for name, action in item_pairs:
                    try:
                        action = action.strip().lower()

                        # special actions
                        if action == 'separator':
                            menu.append(SEPARATOR)
                            continue
                        elif action == 'recent_clips':
                            if not name.strip(): menu.extend(RECENT_CLIPS_BASE)
                            else: menu.append(pystray.MenuItem(name, pystray.Menu(*RECENT_CLIPS_BASE)))
                            continue
                        elif action == 'memory':
                            if name:                # get_mem_title -> workaround for python bug/oddity involving creating lambdas in iterables
                                get_mem_title = lambda name: lambda _: name.replace('?memory', f'{get_memory():.2f}mb')
                                menu.append(pystray.MenuItem(get_mem_title(name), None, enabled=False))
                            else:
                                menu.append(pystray.MenuItem(lambda _: f'Memory usage: {get_memory():.2f}mb', None, enabled=False))
                            continue

                        # normal actions -> create and append menu item
                        elif action == 'quit': exit_item_exists = True  # confirm that an exit item is in the menu
                        menu.append(pystray.MenuItem(name, action=TRAY_ACTIONS[action]))

                    # AttributeError means item is (probably) a submenu
                    except AttributeError:
                        if isinstance(action, list):
                            submenu = []
                            evaluate_menu(action, submenu)
                            menu.append(pystray.MenuItem(name, pystray.Menu(*submenu)))
                    except KeyError: logging.warning(f'(X) The following menu item does not exist: "{name}": "{action}"')

            tray_menu = [LEFT_CLICK_ACTION] if LEFT_CLICK_ACTION else []    # start with left-click action included, if present
            evaluate_menu(load_menu(), tray_menu)                           # recursively evaluate custom menu
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
                pystray.MenuItem('View videos', action=lambda: os.startfile(VIDEO_FOLDER)),
                SEPARATOR,
                *RECENT_CLIPS_MENU,
                SEPARATOR,
                pystray.MenuItem('Check for clips', action=lambda: cutter.check_for_clips(manual_update=True)),
                pystray.MenuItem('Clear history',   action=lambda: cutter.last_clips.clear()),
                pystray.MenuItem('Exit', quit_tray)
            )

        # create system tray icon and manually assert that `ICON_PATH` is valid
        tray_icon = Icon(None, ICON_PATH, f'{TITLE} {VERSION}', tray_menu)
        tray_icon._assert_icon_handle()

        # use atexit to register quit function so we always quit
        atexit.register(quit_tray, tray_icon)

        # cleanup *some* extraneous dictionaries/collections/functions
        del abort_launch
        del verify_ffmpeg
        del verify_config_files
        del sanitize_json
        del load_menu
        del restore_menu_file
        del get_clip_tray_action
        del title_callback
        del tray_menu
        del SEPARATOR
        del RECENT_CLIPS_BASE
        del BACKUP_FOLDER_HINT
        del LENGTH_DICTIONARY
        del INSTANT_REPLAY_HOTKEY
        del CONCATENATE_HOTKEY
        del DELETE_HOTKEY
        del TRAY_ACTIONS
        del TRAY_LEFT_CLICK_ACTION
        del TRAY_MIDDLE_CLICK_ACTION
        del CONFIG_PATH
        del CUSTOM_MENU_PATH
        del SHADOWPLAY_REGISTRY_PATH
        del NO_CONFIG
        del NO_MENU

        # final garbage collection to reduce memory usage
        gc.collect(generation=2)
        logging.info(f'Memory usage before initializing system tray icon: {get_memory():.2f}mb')

        # finally, run system tray icon
        logging.info('Running.')
        tray_icon.run()
    except SystemExit: pass
    except:
        logging.critical(f'(!) Error while initalizing {TITLE}: {format_exc()}')
        play_alert('error')
        time.sleep(2.5)   # sleep to give error sound time to play
