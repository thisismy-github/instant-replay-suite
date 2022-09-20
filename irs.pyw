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
#from ctypes import *                # 0.27mb / 1.15mb / 14.5mb
#from ctypes.wintypes import *
tracemalloc.start()                 # start recording memory usage AFTER libraries have been imported

# Starts with roughly ~36.7mb of memory usage. Roughly 9.78mb combined from imports alone, without psutil and cv2/pymediainfo (9.63mb w/o tracemalloc).
# detecting shadowplay videos via their encoding is possible (but useless) https://github.com/rebane2001/NvidiaInstantRename/blob/mane/InstantRenameCLI.py
# NOTE add dynamic tooltips to pystray library
# NOTE add mute audio/audio only options(?), maybe simplify "trim..." submenu if performance is bad
# TODO option to hide system tray icon until specific hotkey is pressed/shortcut is opened
# TODO finish and upload fork of pystray to github
#           - add dynamic tooltips (when hovering over icon)
#           - why does exiting cause it to loop around to .run()?
#           - actions should work with class methods (self makes them break)
#           - expanded default/left-click action functionality (bold option, explicit left-click declaration, etc.)
#           - DEBUG init: Image: failed to import FpxImagePlugin: No module named 'olefile'

# ---------------------
# Settings
# ---------------------
AUDIO = True
RENAME = True
RENAME_FORMAT = '?game ?date #?count'
RENAME_DATE_FORMAT = '%y.%m.%d'         # https://strftime.org/
RENAME_COUNT_START_NUMBER = 1
RENAME_COUNT_PADDED_ZEROS = 0

TRAY_ADVANCED_MODE = True
TRAY_ADVANCED_MODE_MENU = (
    {'View log': 'open_log'},
    {'View videos': 'open_video_folder'},
    {'View root': 'open_install_folder'},
    'separator',
    {'Quick actions': [
        {'Play most recent clip': 'play_most_recent'},
        {'View last clip in explorer': 'explore_most_recent'},
        {'Concatenate last two clips': 'concatenate_last_two'},
        {'Delete most recent clip': 'delete_most_recent'},
    ]},
    #{'Play last clip': 'play_most_recent'},
    #{'Explore last clip': 'explore_most_recent'},
    #{'Concat last clips': 'concatenate_last_two'},
    #{'Recent clips': 'recent_clips'},
    'recent_clips',
    'separator',
    {'RAM: ?memory': 'memory'},
    {'Update clips': 'update'},
    {'Clear history': 'clear_history'},
    {'Exit': 'quit'},
)

# Basic mode (TRAY_ADVANCED_MODE = False) only
TRAY_SHOW_QUICK_ACTIONS = True
TRAY_RECENT_CLIPS_IN_SUBMENU = False
TRAY_QUICK_ACTIONS_IN_SUBMENU = True

''' log      - Open log file on left-click.
    videos   - Open VIDEO_PATH on left-click.
    play     - Play most recent clip on left-click.
    explore  - Open most recent clip in explorer on left-click.
    root     - Open root directory of program on left-click.
    quit     - Exit program on left-click. '''
TRAY_LEFT_CLICK_ACTION = 'videos'

# Recent clip menu settings
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

CONCATENATE_HOTKEY = 'alt + c'
DELETE_HOTKEY = 'ctrl + alt + d'
LENGTH_HOTKEY = 'alt'
LENGTH_DICTIONARY = {
    '1': 10,
    '2': 20,
    '3': 30,
    '4': 40,
    '5': 50,
    '6': 60
}

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

VIDEO_PATH_OVERRIDE = ''
INSTANT_REPLAY_HOTKEY_OVERRIDE = ''

# ---------------------
# Logging
# ---------------------
LOG_PATH = __file__.replace('.pyw', '.log').replace('.py', '.log')
logging.basicConfig(
    level=logging.INFO,
    encoding='utf-16',
    format='{asctime} {lineno:<3} {levelname} {funcName}: {message}',
    datefmt='%I:%M:%S%p',
    style='{',
    handlers=(logging.FileHandler(LOG_PATH, 'w', delay=False),
              logging.StreamHandler()))


# ---------------------
# Aliases & constants
# ---------------------
pjoin = os.path.join
exists = os.path.exists
getstat = os.stat
getsize = os.path.getsize
basename = os.path.basename
dirname = os.path.dirname
abspath = os.path.abspath

CWD = dirname(os.path.realpath(__file__))
HISTORY_PATH = pjoin(CWD, 'recent_clips.txt')
RESOURCE_DIR = pjoin(CWD, 'resources')
ICON_PATH = pjoin(RESOURCE_DIR, 'icon.ico')

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

USAGE_BAD_TRAY_ITEM_ERROR = '\n\nUsage: {"title of item": "action_of_item"}\n       OR\n       {"title of submenu": {nested_menu}}\n\nNote the colon (:) between the title and it\'s action.'
DEV_MIN_CLIP_OBJECTS = 5
CLIP_BUFFER = max(DEV_MIN_CLIP_OBJECTS, TRAY_RECENT_CLIP_COUNT)


# ---------------------
# Utility functions
# ---------------------
#get_memory = lambda: psutil.Process().memory_info().rss / (1024 * 1024)
get_memory = lambda: tracemalloc.get_traced_memory()[0] / 1048576

def quit_tray():
    logging.info('Closing system tray icon and exiting suite.')
    tray_icon.visible = False
    tray_icon.stop()
    tracemalloc.stop()
    logging.info(f'Clip history: {cutter.last_clips}')
    with open(HISTORY_PATH, 'w') as history: history.write('\n'.join(c.path if isinstance(c, Clip) else c for c in cutter.last_clips))
    try: sys.exit(0)
    except SystemExit: pass         # avoid harmless yet annoying SystemExit error


def delete(path: str):
    logging.info('Deleting clip: ' + path)
    try: os.remove(path)
    except:
        logging.error(f'Error while deleting file {path}: {format_exc()}')
        play_alert('error')


def play_alert(sound: str) -> bool:
    if AUDIO:
        path = pjoin(RESOURCE_DIR, f'{sound}.wav')
        logging.info('Playing alert: ' + path)
        try: winsound.PlaySound(path, winsound.SND_ASYNC)
        except:
            winsound.MessageBeep(winsound.MB_ICONHAND)      # play OS error sound
            if sound != 'error':    # intentional error sound (and there's no custom error.wav) -> ignore missing file error
                logging.error(f'Error while playing sound {path}: {format_exc()}')
                return False
    return True


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


def trim_off_start_in_place(clip: Clip, length: float):     # see old_instant_replay_suite.py for path-only version
    assert exists(clip.path), f'Path {clip.path} does not exist!'
    logging.info(f'Trimming clip {clip.name} to {length} seconds')
    if clip.length <= length: return logging.info(f'(?) Video is only {clip.length:.2f} seconds long and cannot be trimmed to {length} seconds.')
    logging.info(f'Clip is {clip.length:.2f} seconds long.')

    base, ext = os.path.splitext(clip.path)
    temp_path = f'{base}_temp{ext}'
    shutil.copy2(clip.path, temp_path)

    #ffmpeg(clip.path, f'-i "%tp" -ss {clip.length - length} -c:v copy -c:a copy "{clip.path}"')
    cmd = f'ffmpeg -y -i "{temp_path}" -ss {clip.length - length} -c:v copy -c:a copy "{clip.path}" -hide_banner -loglevel warning'
    logging.info(cmd)
    process = subprocess.Popen(cmd, shell=True)
    process.wait()
    try: os.remove(temp_path)
    except: logging.error(f'Error while deleting temporary file {temp_path}: {format_exc()}')
    logging.info(f'Trim to {length} seconds successful.\n')


def get_video_duration(path: str) -> float:     # ? -> https://stackoverflow.com/questions/10075176/python-native-library-to-read-metadata-from-videos
    for track in pymediainfo.MediaInfo.parse(path).tracks:
        if track.track_type == "Video":
            return track.duration / 1000
    return 0


# ---------------------
# Custom Pystray class
# ---------------------
class Icon(pystray._win32.Icon):
    ''' This subclass auto-updates the menu before opening,
        allowing dynamic titles/actions to always be up-to-date. '''
    def _on_notify(self, wparam, lparam):
        if lparam == 0x0205: self._update_menu()   # "win32.WM_RBUTTONUP" in pystray
        super()._on_notify(wparam, lparam)


# ---------------------
# Clip class
# ---------------------
class Clip:
    __slots__ = ('working', 'path', 'name', 'time', 'size', 'date', 'full_date', 'length', 'length_string', 'length_size_string')
    def __repr__(self): return self.name

    def __init__(self, path, stat, rename=False):
        self.working = False
        self.path, self.name = self.rename(path) if rename else (abspath(path), basename(path))   # os.path.abspath for consistent formatting
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
        stat = os.stat(self.path)
        self.time = stat.st_ctime
        self.size = f'{(stat.st_size / 1048576):.1f}mb'
        self.length = get_video_duration(self.path)
        length_int = int(self.length)
        self.length_string = f'{length_int // 60}:{length_int % 60:02}'
        self.length_size_string = f'Length: {self.length_string} ({self.size})'
        gc.collect(generation=2)


    def rename(self, path, name_format=RENAME_FORMAT, date_format=RENAME_DATE_FORMAT):
        ''' Renames clip according to specified `name_format` and `date_format` based on Shadow Play's default name formatting. '''
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
        return abspath(renamed_path), renamed_base      # use os.path.abspath to ensure consistent path formatting later on


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
        self.waiting_for_clip = False
        self.last_clip_time = time.time()
        if exists(HISTORY_PATH):
            with open(HISTORY_PATH, 'r') as history:
                # get all valid paths from history file, create a buffer of clip objects, then start caching paths as strings outside buffer
                lines = (path for path in reversed(history.read().splitlines()) if path and exists(path))   # reversed() is an iterable, not an actual list
                self.last_clips = [Clip(p, getstat(p)) if i <= CLIP_BUFFER else p for i, p in enumerate(lines)]
                self.last_clips.reverse()                   # .reverse() is a very fast operation
                if self.last_clips: logging.info(f'Previous {len(self.last_clips)} clip{"s" if len(self.last_clips) != 1 else ""} loaded: {self.last_clips}')
        else: self.last_clips = []

        keyboard.add_hotkey(INSTANT_REPLAY_HOTKEY, self.set_last_clip)
        keyboard.add_hotkey(CONCATENATE_HOTKEY, self.concatenate_last_clips)
        keyboard.add_hotkey(DELETE_HOTKEY, self.delete_clip)
        for number_key, length in LENGTH_DICTIONARY.items():
            keyboard.add_hotkey(f'{LENGTH_HOTKEY} + {number_key}', self.trim_clip, args=(length,))
        logging.info('Auto-cutter initialized.')

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
            cached_clip_index = -1 * (CLIP_BUFFER + 1)  # +1 to check first clip outside buffer
            while isinstance(cached_clip := last_clips[cached_clip_index], str):    # convert clip to actual Clip object if it's just a string
                if exists(cached_clip):                 # if cached clip exists, convert to Clip object and break loop
                    last_clips[cached_clip_index] = Clip(cached_clip, getstat(cached_clip))
                    break
                last_clips.pop(cached_clip_index)       # if cached clip doesn't exist, pop and try next clip
        except IndexError: pass                         # IndexError -> pop was out of range, pass and return None
        except: logging.error(f'Error while popping clip at index {index} <cached_clip_index={cached_clip_index}, len(self.last_clips)={len(self.last_clips)}>: {format_exc()}')
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


    def get_clip(self, index, verb=None, alert=None, min_clips=1, patient=True, _recursive=False):    # this function gets rid of ~60 lines of repetitive code
        ''' Safely gets a clip at a specified `index`. Can wait for the clip if `patient` is True, otherwise pulls a clip immediately, which
            is unlikely to return the wrong clip outside of intentional misuse. Plays an associated sound effect with the action we're going to
            perform, if specified with `alert`, and avoids passing this to the wait() method. Plays an  Uses the `verb`, `alert`, and `min_clips`
            parameters for waiting and checking if the acquired clip is working. Pops clips if they don't exist anymore. Recursively calls itself
            until a valid clip is acquired if `patient` is True. `_recursive` used internally to avoid unnecessary waiting/alerts. '''
        if not _recursive:
            logging.info(f'Getting clip at index {index} (verb={verb} alert={alert} min_clips={min_clips} patient={patient})')
            if alert is not None: play_alert(str(alert).lower())
            if patient and not self.wait(verb=verb if verb else alert, min_clips=min_clips): return

        clip = self.last_clips[index]
        if not exists(clip.path):
            self.pop(index)
            if patient: return self.get_clip(index, verb, alert, min_clips, patient, _recursive=True)
            else:
                logging.warning(f'Clip at index {index} does not actually exist: {clip.path}')
                play_alert('error')
                raise AssertionError("Clip does not exist")
        if clip.is_working(verb if verb else alert): raise AssertionError("Clip is being worked on")
        return clip

    # ---------------------
    # Acquiring clips
    # ---------------------
    def set_last_clip(self, manual_update=False):
        self.waiting_for_clip = True
        Thread(target=self.set_last_clip_thread, args=(manual_update,), daemon=True).start()


    def set_last_clip_thread(self, manual_update=False):
        try:
            if not manual_update:
                logging.info('Instant replay detected! Waiting 3 seconds...')
                self.last_clip_time = time.time() - 1
                time.sleep(3)
            logging.info(f'Scanning {VIDEO_PATH} for videos')
            for root, _, files in os.walk(VIDEO_PATH):
                for file in files:
                    path = pjoin(root, file)
                    stat = getstat(path)
                    if self.last_clip_time < stat.st_ctime:
                        logging.info(f'New video detected: {file}')
                        while get_video_duration(path) == 0:
                            print('VIDEO IS 0')
                            time.sleep(0.2)

                        clip = Clip(path, stat, rename=RENAME)
                        self.last_clips.append(clip)
                        logging.info(f'Memory usage after adding clip: {get_memory():.2f}mb\n')
                        if manual_update: continue
                        else: return
        except:
            logging.error(f'Error while setting last clip: {format_exc()}')
            play_alert('error')
        finally:
            if manual_update: self.last_clips.sort(key=lambda clip: clip.time)
            self.waiting_for_clip = False

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
            logging.error(f'Error while cutting last clip: {format_exc()}')
            play_alert('error')


    def concatenate_last_clips(self, patient=True):
        try:
            clip1 = self.get_clip(index=-2, alert='Concatenate', min_clips=1, patient=patient)
            clip2 = self.get_clip(index=-1, alert='Concatenate', min_clips=2, patient=patient)
            clip1path = clip1.path         # these could be popped here, but it's safer and simpler to do it this way
            clip2path = clip2.path

            base, ext = os.path.splitext(clip1path)
            temp_path = f'{base}_concat{ext}'
            text_path = f'{base}_concatlist.txt'    # write list of clips to text file (with / as separator and single quotes to avoid ffmpeg errors)

            with open(text_path, 'w') as txt: txt.write(f"file '{clip1path.replace(os.sep, '/')}'\nfile '{clip2path.replace(os.sep, '/')}'")
            cmd = f'ffmpeg -y -f concat -safe 0 -i "{text_path}" -c copy "{temp_path}" -hide_banner -loglevel warning'
            logging.info(cmd)
            process = subprocess.Popen(cmd, shell=True)
            process.wait()

            for file in (clip1path, clip2path, text_path): delete(file)
            os.rename(temp_path, clip1path)
            self.pop()
            clip1.refresh()
            logging.info('Clips concatenated, renamed, popped, refreshed, and cleaned up successfully.')
        except AssertionError: return
        except:
            logging.error(f'Error while concatenating last two clips: {format_exc()}')
            play_alert('error')


    def delete_clip(self, index=-1, patient=True):
        if patient and not self.wait(alert='Delete'): return    # not a part of get_clip(), to simplify things
        try:
            logging.info(f'Deleting clip at index {index}: {self.last_clips[index].path}')
            os.remove(self.pop(index).path)                     # pop and remove directly
            logging.info('Deletion successful.\n')

        except IndexError: return
        except:
            logging.error(f'Error while deleting last clip: {format_exc()}')
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
            logging.error(f'Error while cutting last clip: {format_exc()}')
            play_alert('error')


    def compress_clip(self, index=-1, patient=True):
        try:    # rename clip to include (comressing...), compress, then rename back and refresh
            clip = self.get_clip(index, verb='Compress', patient=patient)
            clip.working = True
            try: Thread(target=self.compress_clip_thread, args=(clip,), daemon=True).start()
            except:
                logging.error(f'Error while creating compression thread: {format_exc()}')
                play_alert('error')
            finally: clip.working = False

        except (IndexError, AssertionError): pass
        except:
            logging.error(f'Error while cutting last clip: {format_exc()}')
            play_alert('error')


    def compress_clip_thread(self, clip: Clip):
        try:
            clip.working = True
            base, ext = os.path.splitext(clip.path)
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
            logging.error(f'Error while compressing clip: {format_exc()}')
            play_alert('error')
        finally:
            clip.working = False


###########################################################
if __name__ == '__main__':
    try:
        logging.info(f'Memory usage before initializing AutoCutter class: {get_memory():.2f}mb')
        cutter = AutoCutter()
        logging.info(f'Memory usage after initializing AutoCutter class: {get_memory():.2f}mb')

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
            except: logging.error(f'Error while generating title for system tray clip {clip} at index {index}: {cutter.last_clips} {format_exc()}')


        def get_clip_tray_action(index: int):
            ''' Generates the action associated with each recent clip item in the tray-icon's menu.
                If a submenu is desired, one is created and returned. Otherwise, a callable lambda is returned. '''
            if TRAY_RECENT_CLIPS_HAVE_UNIQUE_SUBMENUS:
                if TRAY_RECENT_CLIPS_SUBMENU_EXTRA_INFO:
                    extra_info_items = (
                        pystray.MenuItem(pystray.MenuItem(None, None), None),
                        pystray.MenuItem(lambda _: cutter.last_clips[index].length_size_string if len(cutter.last_clips) * -1 <= index else TRAY_RECENT_CLIP_DEFAULT_TEXT, None, enabled=False),
                        pystray.MenuItem(lambda _: cutter.last_clips[index].full_date if len(cutter.last_clips) * -1 <= index else TRAY_RECENT_CLIP_DEFAULT_TEXT, None, enabled=False)
                    )
                else: extra_info_items = tuple()

                get_trim_lambda = lambda length, index: lambda: cutter.trim_clip(length, index, patient=False)   # workaround for python bug/oddity involving creating lambdas in iterables
                return pystray.Menu(
                    pystray.MenuItem('Trim...', pystray.Menu(*(pystray.MenuItem(f'{length} seconds', get_trim_lambda(length, index)) for length in LENGTH_DICTIONARY.values()))),
                    pystray.MenuItem('Play...', lambda: cutter.open_clip(index, play=True, patient=False)),
                    pystray.MenuItem('Explore...', lambda: cutter.open_clip(index, play=False, patient=False)),
                    #pystray.MenuItem('Concatenate with prior clip', cutter.concatenate_last_clips),
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
        get_title_function = lambda index: lambda _: get_clip_tray_title(index)  # workaround for python bug/oddity involving creating lambdas in iterables
        RECENT_CLIPS_BASE = tuple(pystray.MenuItem(get_title_function(i), get_clip_tray_action(i)) for i in range(-1, (TRAY_RECENT_CLIP_COUNT * -1) - 1, -1))

        # action dictionaries
        LEFT_CLICK_ACTIONS = {
            'log':     lambda: pystray.MenuItem(None, action=lambda: os.startfile(LOG_PATH), default=True, visible=False),
            'videos':  lambda: pystray.MenuItem(None, action=lambda: os.startfile(VIDEO_PATH), default=True, visible=False),
            'play':    lambda: pystray.MenuItem(None, action=lambda: cutter.open_clip(play=True), default=True, visible=False),
            'explore': lambda: pystray.MenuItem(None, action=lambda: cutter.open_clip(play=False), default=True, visible=False),
            'root':    lambda: pystray.MenuItem(None, action=lambda: os.startfile(CWD), default=True, visible=False),
            'quit':    lambda: pystray.MenuItem(None, action=quit_tray, default=True, visible=False),
        }
        TRAY_ADVANCED_MODE_ACTIONS = {
            'open_log':             lambda name: pystray.MenuItem(name, action=lambda: os.startfile(LOG_PATH)),
            'open_video_folder':    lambda name: pystray.MenuItem(name, action=lambda: os.startfile(VIDEO_PATH)),
            'open_install_folder':  lambda name: pystray.MenuItem(name, action=lambda: os.startfile(CWD)),
            'play_most_recent':     lambda name: pystray.MenuItem(name, action=lambda: cutter.open_clip(play=True)),
            'explore_most_recent':  lambda name: pystray.MenuItem(name, action=lambda: cutter.open_clip(play=False)),
            'delete_most_recent':   lambda name: pystray.MenuItem(name, action=cutter.delete_clip),             # small RAM drop by making these not lambdas
            'concatenate_last_two': lambda name: pystray.MenuItem(name, action=cutter.concatenate_last_clips),  # 26.0mb -> 25.8mb on average
            'clear_history':        lambda name: pystray.MenuItem(name, action=cutter.last_clips.clear),
            'update':               lambda name: pystray.MenuItem(name, action=lambda: cutter.set_last_clip(manual_update=True)),
            'quit':                 lambda name: pystray.MenuItem(name, action=quit_tray),
        }

        # setting left-click action
        if TRAY_LEFT_CLICK_ACTION in LEFT_CLICK_ACTIONS: LEFT_CLICK_ACTION = LEFT_CLICK_ACTIONS[TRAY_LEFT_CLICK_ACTION]()
        elif TRAY_LEFT_CLICK_ACTION in TRAY_ADVANCED_MODE_ACTIONS: LEFT_CLICK_ACTION = TRAY_ADVANCED_MODE_ACTIONS[TRAY_LEFT_CLICK_ACTION]()
        else:
            LEFT_CLICK_ACTION = pystray.MenuItem(pystray.MenuItem(None, None), None, visible=False)
            logging.warning(f'(X) Left click action "{TRAY_LEFT_CLICK_ACTION}" does not exist')

        # creating menu -- advanced mode
        if TRAY_ADVANCED_MODE:
            exit_item_exists = False            # variable for making sure an exit item is included
            def evaluate_menu(items, menu):     # function for recursively solving menus/submenus and exporting them to a list
                global exit_item_exists
                for item_dict in items:
                    assert not isinstance(item_dict, set), f'Tray item {item_dict} is improperly written. {USAGE_BAD_TRAY_ITEM_ERROR}'
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
                                action = (action,)  # set action as tuple and raise AttributeError to read it as a submenu
                                raise AttributeError
                            menu.append(TRAY_ADVANCED_MODE_ACTIONS[action](name))
                        except AttributeError:      # AttributeError -> item is a submenu
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
                pystray.MenuItem('Concatenate two last clips', cutter.concatenate_last_clips),
                pystray.MenuItem('Delete most recent clip', cutter.delete_clip)
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
                pystray.MenuItem('Check for clips', action=lambda: cutter.set_last_clip(manual_update=True)),
                pystray.MenuItem('Clear history',   action=cutter.last_clips.clear),
                pystray.MenuItem('Exit', quit_tray)
            )

        #CreateIconFromResourceEx = windll.user32.CreateIconFromResourceEx
        #size_x, size_y = 32, 32
        #LR_DEFAULTCOLOR = 0
        #with open(ICON_PATH, "rb") as f:
        #    png = f.read()
        #hicon = CreateIconFromResourceEx(png, len(png), 1, 0x30000, size_x, size_y, LR_DEFAULTCOLOR)

        # create system tray icon
        tray_icon = Icon(None, Image.open(ICON_PATH), 'Instant Replay Suite', tray_menu)

        # cleanup *some* extraneous dictionaries/collections/functions
        del get_clip_tray_action
        del get_title_function
        del tray_menu
        del SEPARATOR
        del RECENT_CLIPS_BASE
        del TRAY_ADVANCED_MODE_MENU
        del LENGTH_DICTIONARY
        del INSTANT_REPLAY_HOTKEY
        del CONCATENATE_HOTKEY
        del DELETE_HOTKEY
        del LENGTH_HOTKEY

        # final garbage collection to reduce memory usage
        gc.collect(generation=2)
        logging.info(f'Memory usage before initializing system tray icon: {get_memory():.2f}mb')

        # finally, run system tray icon
        tray_icon.run()
    except:
        logging.error(f'Error while initalizing Instant Replay Suite: {format_exc()}')
        play_alert('error')
        time.sleep(2.5)   # sleep to give error sound time to play
