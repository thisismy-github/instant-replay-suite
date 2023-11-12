''' Contains the necessary functions for update-checking, downloading,
    validating, and migrating. Uses this project's GitHub releases
    page to detect, compare, and download new versions.

    Adapted from my `PyPlayer` project. '''

import os
import time
import logging
from traceback import format_exc

# ---

''' Because we don't use a shared constants file like PyPlayer does, we'll
    set these constants from within IRS itself to avoid circular imports. '''
VERSION = None
REPOSITORY_URL = None
IS_COMPILED = None
SCRIPT_PATH = None
CWD = None
RESOURCE_FOLDER = None
BIN_FOLDER = None
show_message = None
cfg = None

HYPERLINK = None
logger = logging.getLogger('update.py')


# --------------------------
# Utility functions/classes
# --------------------------
class InsufficientSpaceError(Exception): pass


# https://stackoverflow.com/questions/11887762/how-do-i-compare-version-numbers-in-python
def get_later_version(version_a: str, version_b: str) -> str:
    ''' Returns the greater of two version strings, with mild future-proofing.
        Allows for an arbitrary number in each sequence of the version, with
        an arbitrary number of sequences. '''
    atuple = tuple(map(int, (version_a.split('.'))))
    btuple = tuple(map(int, (version_b.split('.'))))
    return version_a if atuple > btuple else version_b


def check_available_space(space_needed: int, path: str,
                          overhead_factor: float = 1.1) -> bool:
    ''' Shows a warning and raises `InsufficientSpaceError` if
        `space_needed` (in bytes) multiplied by `overhead_factor`
        is less than what's available on `path`'s drive. '''
    import shutil
    drive = os.path.splitdrive(path)[0]
    needed = space_needed * overhead_factor
    available = shutil.disk_usage(drive).free
    if needed > available:
        msg = (f"There is not enough space available on your {drive}\\ drive."
               f"\n\nFree space required: {needed / 1048576:.0f}mb"
               f"\nFree space available: {available / 1048576:.0f}mb")
        show_message('Insufficient space remamining', msg, 0x00040010)
        raise InsufficientSpaceError    # X-symbol, stay on top ^


def download(url: str, path: str) -> None:
    ''' Downloads file from `url` in chunks and saves it to `path`. '''
    import requests
    download_response = requests.get(url, stream=True)
    download_response.raise_for_status()

    # check if we have enough space (raises InsufficientSpaceError)
    total_size = int(download_response.headers.get('content-length'))
    check_available_space(total_size, path)

    downloaded = 0
    mb_per_chunk = 4

    # download in chunks (not really necessary since we don't have a GUI)
    with open(path, 'wb') as file:
        logger.info(f'Downloading {total_size / 1048576:.2f}mb')
        chunk_size = mb_per_chunk * (1024 * 1024)
        start_time = time.time()
        for chunk in download_response.iter_content(chunk_size=chunk_size):
            file.write(chunk)
            downloaded += len(chunk)
            percent = (downloaded / total_size) * 100
            logger.info(f'{percent:.0f}% ({downloaded / 1048576:.0f}mb/{total_size / 1048576:.2f}mb)')
        logger.info(f'File downloaded after {time.time() - start_time:.1f} seconds.')


# --------------------------
# Update functions
# --------------------------
def check_for_update(show_message_for_no_update: bool = False, lock_file: str = '') -> None:
    ''' Checks GitHub for an update, and downloads it if requested by the
        user. Returns an exit code if we need to close for an incoming update.
        `lock_file` specifies the path to an open file to give to the updater
        so that it has a way to see when we've closed.

        The following format is assumed:
            > REPOSITORY_URL -- https://github.com/thisismy-github/instant-replay-suite
            > TITLE          -- Instant Replay Suite
            > VERSION        -- 0.1.0 beta
                - version number can be variable length, "beta" modifier is optional

        The following example is expected:
            ->     https://github.com/thisismy-github/instant-replay-suite/releases/latest
              ->   https://github.com/thisismy-github/instant-replay-suite/releases/tags/v1.2.3
                -> https://github.com/thisismy-github/instant-replay-suite/releases/download/v1.2.3/instant-replay-suite_1.2.3.zip

        NOTE: It's possible to directly access the latest version of an asset by doing
              {REPOSITORY_URL}/releases/latest/download/*asset_name*,
              but that requires not including the version in the asset's filename. '''

    import requests
    release_url = f'{REPOSITORY_URL}/releases/latest'
    logger.info(f'Checking {release_url} for updates')

    try:
        response = requests.get(release_url)
        response.raise_for_status()
        latest_version_url = response.url.rstrip('/')
        latest_version = latest_version_url.split('/')[-1].lstrip('v')

        current_version = VERSION.split()[0]
        logger.info(f'Latest version: {latest_version} | Current version: {current_version}')

        # the formats of the current and latest versions are different
        if len(latest_version) != len(current_version):
            logger.error('(!) Github release URL could not be parsed correctly.')
            msg = ("The URL for the latest Github release has an unexpected "
                   f"format.\n\nGithub version: '{latest_version}'\nCurrent "
                   f"version: '{current_version}'\n\nNewer versions may use "
                   "a different naming scheme. You can manually check the ")
            show_message('Update URL mismatch', msg + HYPERLINK)
            return                      # make sure we don't return anything

        # current version is older than latest version -> update available
        if get_later_version(latest_version, current_version) != current_version:
            title = f'Update {latest_version} available'
            intro = f'An update is available on Github ({current_version} -> {latest_version}). '
            outro = '\n\nYou can view the ' + HYPERLINK
            if IS_COMPILED:             # script users cannot auto-update
                msg = 'Would you like to download and install this update automatically?'
                flags = 0x00040044      # i-symbol, Yes/No, stay-on-top
            else:                       # NOTE: auto-updating is Windows-only for now
                msg = 'You cannot auto-update while running directly from the script.'
                flags = 0x00040040      # i-symbol, OK, stay-on-top

            response = show_message(title, intro + msg + outro, flags)
            if response == 6:           # "Yes" button -> begin auto-update
                name = REPOSITORY_URL.split('/')[-1]
                filename = f'{name}_{latest_version}.zip'
                download_url = f'{latest_version_url.replace("/tag/", "/download/")}/{filename}'
                download_path = os.path.join(CWD, filename)
                if download_update(latest_version, download_url, download_path, lock_file):
                    return 99

        # otherwise, we must be up to date
        else:
            msg = 'You\'re up to date!'
            logger.info(msg)
            if show_message_for_no_update:
                show_message('No update found', msg, 0x00040040)  # i-symbol, stay on top

    except requests.exceptions.ConnectionError:
        logger.warning('(!) Update check was unable to reach GitHub (no internet connection?): ' + format_exc())
    except:
        logger.error('(!) UPDATE-CHECK FAILED: ' + format_exc())


def download_update(latest_version: str, download_url: str, download_path: str, lock_file: str) -> None:
    ''' Downloads update from `download_url` to `download_path` and installs
        it using our updater-utility. Passes `latest_version` and `lock_file`
        to the updater-utility. '''

    try:
        logger.info(f'Downloading version {latest_version} from {download_url} to {download_path}')
        download(download_url, download_path)
        logger.info('Update download successful, restarting...')

        # locate updater executable -> check both root folder and bin folder
        original_updater_path = os.path.join(CWD, 'updater.exe')    # IS_COMPILED is assumed here
        if not os.path.exists(original_updater_path):
            original_updater_path = os.path.join(BIN_FOLDER, 'updater.exe')
            if not os.path.exists(original_updater_path):
                raise FileNotFoundError(f'Could not find updater at {original_updater_path}')

        # copy updater utility to temporary path so it can be replaced during the update
        import shutil
        import subprocess
        active_updater_path = os.path.join(CWD, f'{time.time()}_updater.exe')
        logger.info(f'Copying updater-utility to temporary path ({active_updater_path})')
        shutil.copy2(original_updater_path, active_updater_path)

        # mark edited/deleted resouces as ignored so they don't get replaced during the update
        default_file = os.path.join(RESOURCE_FOLDER, '!defaults.txt')
        edited = []
        deleted = []
        folder_name = os.path.basename(RESOURCE_FOLDER)
        if os.path.exists(default_file):
            with open(default_file, 'r') as defaults:
                for line in defaults:   # format is "<filename>: <size>"
                    try:
                        line = line.strip()
                        if line and line[:2] != '//':
                            filename, expected_size = line.split(': ')
                            path = os.path.join(RESOURCE_FOLDER, filename)
                            if not os.path.exists(path):
                                deleted.append(f'"{folder_name}/{filename}"')
                            elif os.path.getsize(path) != int(expected_size):
                                edited.append(f'"{folder_name}/{filename}"')
                    except:
                        pass
        ignored = edited + deleted      # we handle both edits and deletes the same way, but this may change
        logger.info(f'Ignoring edited resources: {ignored}')

        # run updater utility and close ourselves
        logger.info('Update-utility starting, main script closing...')
        add_to_report = f'"{VERSION.split()[0]} -> {latest_version}" "{active_updater_path}"'
        updater_cmd = (f'{active_updater_path} {download_path} '    # the updater and the zip file we want it to unpack
                       f'--destination {CWD} '                      # the destination to unzip the file to
                       f'--cmd "{SCRIPT_PATH}" '                    # the command the updater should run to restart us
                       f'--lock-files "{lock_file}" '               # tell updater to wait for lock-file to be deleted
                       f'--ignore {" ".join(ignored)} '             # tell updater not to extract these resource files
                       f'--add-to-report {add_to_report}')          # write versions and temp-updater's path in report
        logger.info('Update-utility command:\n\n' + updater_cmd.replace('--', '\n--'))
        subprocess.Popen(updater_cmd)
        return True

    except InsufficientSpaceError:
        pass
    except:
        logger.error(f'(!) Could not download latest version. New naming format? Missing updater? {format_exc()}')

        reasons = ''
        if os.path.exists(active_updater_path):
            try: os.remove(active_updater_path)
            except: reasons += f'Additionally, the temporary update-utility file at {active_updater_path} could not be deleted.\n\n'
        if os.path.exists(download_path):
            try: os.remove(download_path)
            except: reasons += f'Additionally, the downloaded .zip file at {download_path} could not be deleted.\n\n'

        msg = (f'Update {latest_version} failed to install.\n\nThere could '
               'have been an error while creating the download link, the '
               'download may have failed, the update utility may be missing, '
               'or newer versions may use a different format for updating.\n'
               f'\n{reasons}You can still manually download the {HYPERLINK}')
        show_message('Update download failed', msg)
    return False


def validate_update(update_report_path: str) -> None:
    ''' Parses update report at `update_report_path`, which tells us if the
        update was successful (and what errors occurred if it wasn't), what
        version we've updated from, and what files we need to clean up. '''

    logger.info(f'Update report detected at {update_report_path}, validating...')
    with open(update_report_path) as report:
        lines = tuple(line.strip() for line in report)              # version_change = "<old_version> -> <new_version>"
        version_change, active_updater_path, download_path, status = lines

        try: os.remove(active_updater_path)
        except: logger.warning(f'Could not clean up temporary updater after update: {download_path}')
        try: os.remove(download_path)
        except: logger.warning(f'Could not clean up .zip file after update: {download_path}')

        if status != 'SUCCESS':
            logger.warning(f'(!) UPDATE FAILED: {status}')
            msg = (f"The update failed while unpacking:\n{status}.\n\n"
                   "If needed, you can manually download the " + HYPERLINK)
            return show_message('Update failed', msg)

    try: os.remove(update_report_path)
    except: logger.warning('Failed to delete update report after validation.')

    logger.info('Update validated.')
    update_migration(version_change.split(' -> ')[0])
    msg = f'Update from {version_change} successful.'
    show_message('Update successful', msg, 0x00040040)              # i-symbol + stay on top


def update_migration(old_version: str) -> None:
    ''' Handles additional work required to migrate
        `old_version` to the latest version, if any. '''
    older_than = lambda v: old_version != v and get_later_version(old_version, v) == v
    if older_than('1.3.0'):
        try:
            import irs
            cfg.load('MAX_RECENT_CLIPS', 10, section=' --- Tray Menu Recent Clips --- ')
            irs.TRAY_RECENT_CLIP_COUNT = cfg.moveSetting(
                oldKey='MAX_RECENT_CLIPS',
                oldSection=' --- Tray Menu Recent Clips --- ',
                newKey='MAX_CLIPS_VISIBLE_IN_MENU',
                newSection=' --- General --- ',
                replace=False
            )
        except:
            pass
