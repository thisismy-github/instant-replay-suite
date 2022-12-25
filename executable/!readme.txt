---

To compile Instant Replay Suite on Windows, simply run "build.py".

This assumes that you have pip and already have all the appropriate libraries from
requirements.txt installed. If not, you can do "pip -r --upgrade requirements.txt".

For compatibility with Windows 7 and Windows Vista, Python 3.8 must be used.

---

Summary of contents:
    build             -- PyInstaller build files that are created during the first compilation.
                         These are not needed for anything, but speed up future compilations.

    compiled          -- PyInstaller's actual compilation folder. This is where the final products
                         are placed. Normally called "dist", but renamed to "compiled" for clarity.

    compiled/release  -- The compiled files for this program's main script(s). This is where compiled
                         files for the launcher and updater will be merged, the "include" files will
                         be added, and all other misc files will be placed. The launcher and updater
                         folders and include-files will all be merged and deleted automatically.
                         As the name implies, this is the folder that would be released on Github.

    build.py          -- A cross-platform Python script for compiling. This searches for a venv, uses
                         the .spec files to compile both our script and its updater, and then performs
                         several post-compilation activities to clean up and finish the compile.

    exclude.txt       -- A list of likely files and folders to be included with each compilation on
                         Windows that do not appear to actually be necessary and should be deleted
                         automatically after compilation. Based on compilations done through a
                         virtualenv on Windows 10 using only the packages in requirements.txt.

    hook.py           -- A runtime hook added to main.spec which runs at startup. For this script,
                         its only purpose is adding our "bin" folder to sys.path, allowing the
                         executable to look for .dll and .pyd files within it, thus letting us
                         hide many files and cut down on clutter.

    updater.py        -- A cross-platform Python script for safely installing a downloaded update.
                         Waits for script to close, extracts a given .zip file, leaves a report
                         detailing success/failure as well as what files should be cleaned up
                         (without deleting them itself), and restarts the original executable.

    *.spec            -- PyInstaller .spec files specifying how to compile our main script(s) and
                         separate update-utility.
