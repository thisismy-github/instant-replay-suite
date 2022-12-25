''' This script compiles Instant Replay Suite to an executable. Make sure all
    appropriate libraries in requirements.txt are installed. View !readme.txt
    for more information. Cross-platform features are holdovers as this script
    was adapted from the version I made for PyPlayer.

    A virtualenv is highly recommended for an accurate release folder.
    If you do not run this script from within it, the script will
    reopen itself within the virtualenv automatically.

    thisismy-github 4/16/22 '''


import os
import sys
import glob
import shutil
import platform
import subprocess


pjoin = os.path.join
PLATFORM = platform.system()

CWD = os.path.dirname(os.path.realpath(__file__))
SCRIPT_DIR = os.path.dirname(CWD)
RELEASE_DIR = pjoin(CWD, 'compiled', 'release')
BIN_DIR = pjoin(RELEASE_DIR, 'bin')


# ensure python is running from a venv, if possible
if 'venv' not in sys.executable.split(os.sep):
    VENV_DIR = pjoin(SCRIPT_DIR, 'venv')
    if not os.path.exists(VENV_DIR): VENV_DIR = pjoin(CWD, 'venv')

    if PLATFORM == 'Windows': venv = pjoin(VENV_DIR, 'Scripts', 'python.exe')
    else: venv = pjoin(VENV_DIR, 'bin', 'python3')
    if os.path.exists(venv):
        print(f'Restarting with virtual environment at "{venv}"')
        process = subprocess.Popen(f'"{venv}" "{sys.argv[0]}"', shell=True)
        process.wait()
        sys.exit(0)


def get_default_resources():
    ''' Generates a new "!defaults.txt" file in the resource folder. '''
    resource_dir = os.path.join(RELEASE_DIR, 'resources')
    output = os.path.join(resource_dir, '!defaults.txt')

    comments = '''
// This lists the current default resources and their sizes. This
// is used to detect which resources should not be replaced while
// updating. Do not edit this file.
    '''

    with open(output, 'w') as out:
        out.write(comments.strip() + '\n\n')

        for filename in os.listdir(resource_dir):
            if filename[-4:] != '.txt':
                path = os.path.join(resource_dir, filename)
                out.write(f'{filename}: {os.path.getsize(path)}\n')


def compile():
    print(f'\nCompiling Instant Replay Suite (sys.executable="{sys.executable}")...\n')
    pyinstaller = f'"{sys.executable}" -m PyInstaller'
    args = f'--distpath "{pjoin(CWD, "compiled")}" --workpath "{pjoin(CWD, "build")}"'
    subprocess.call(f'{pyinstaller} "{pjoin(CWD, "main.spec")}" --noconfirm {args}')
    subprocess.call(f'{pyinstaller} "{pjoin(CWD, "updater.spec")}" --noconfirm {args}')

    if not os.path.isdir(BIN_DIR): os.makedirs(BIN_DIR)

    print('Moving updater to bin folder...')
    name = 'updater' + ('.exe' if PLATFORM == 'Windows' else '')
    shutil.move(pjoin(CWD, 'compiled', name), pjoin(BIN_DIR, name))

    print('Copying cacert.pem...')
    certifi_dir = pjoin(RELEASE_DIR, 'certifi')
    shutil.copy2(pjoin(certifi_dir, 'cacert.pem'), pjoin(BIN_DIR, 'cacert.pem'))
    shutil.rmtree(certifi_dir)
 
    print(f'Deleting files defined in {pjoin(CWD, "exclude.txt")}...')
    with open(pjoin(CWD, 'exclude.txt')) as exclude:
        for line in exclude:
            line = line.strip()
            if not line: continue
            for path in glob.glob(pjoin(RELEASE_DIR, line.strip())):
                print(f'exists={os.path.exists(path)} - {path}')
                if os.path.exists(path):
                    if os.path.isdir(path): shutil.rmtree(path)
                    else: os.remove(path)

    print('Generating "!defaults.txt" file for resources...')
    get_default_resources()


# ---------------------
# Windows
# ---------------------
def compile_windows():
    compile()
    print(f'\nPerforming post-compilation tasks for {PLATFORM}...')

    print('Moving .pyd and .dll files to bin folder...')
    for pattern in ('*.pyd', '*.dll'):
        for path in glob.glob(pjoin(RELEASE_DIR, pattern)):
            print(f'{path} -> {pjoin(BIN_DIR, os.path.basename(path))}')
            shutil.move(path, pjoin(BIN_DIR, os.path.basename(path)))

    print('Moving MediaInfo.dll from pymediainfo folder to bin folder...')
    old = pjoin(RELEASE_DIR, 'pymediainfo')
    shutil.move(pjoin(old, 'MediaInfo.dll'), pjoin(BIN_DIR, 'MediaInfo.dll'))
    os.rmdir(old)

    print('Moving python3*.dll back to root folder...')
    for path in glob.glob(pjoin(BIN_DIR, 'python*.dll')):
        filename = os.path.basename(path)
        print(path, filename, filename != 'python3.dll')
        if filename != 'python3.dll' and 'com' not in filename:
            shutil.move(path, pjoin(RELEASE_DIR, filename))


#######################################
if __name__ == '__main__':
    while True:
        try:
            compile_windows() if PLATFORM == 'Windows' else compile()
            choice = input('\nDone! Type anything to exit, or press enter to recompile... ')
            if choice != '': break
        except:
            import traceback
            input(f'\n(!) Compile failed:\n\n{traceback.format_exc()}')