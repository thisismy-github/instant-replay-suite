---

To compile Instant Replay Suite on Windows, simply run "build.py".

This assumes that you have pip and already have all the appropriate libraries from
requirements.txt installed. If not, you can do "pip -r --upgrade requirements.txt".

For compatibility with Windows 7 and Windows Vista, Python 3.8 must be used.

---

Summary of contents:
    build             -- PyInstaller build files that are created during the first compilation.
                         These are not needed for anything, but speed up future compilations.

    compiled          -- PyInstaller's actual compilation folder, containing the compiled files and
                         executable for Instant Replay Suite's main script, irs.pyw. This is where
                         our resources, the updater's executable, and all of our other misc. files
                         will be placed. This is the folder that would be released on GitHub.

    build.py          -- A cross-platform Python script for compiling. This searches for a venv, uses
                         the .spec files to compile both our script and its updater, and then performs
                         several post-compilation activities to clean up and finish the compile.

    exclude.txt       -- A list of likely files and folders to be included with each compilation on
                         Windows that do not appear to actually be necessary and should be deleted
                         automatically after compilation. Based on compilations done through a
                         virtualenv on Windows 10 using only the packages in requirements.txt.

    updater.py        -- A cross-platform Python script for safely installing a downloaded update.
                         Waits for script to close, extracts a given .zip file, leaves a report
                         detailing success/failure as well as what files should be cleaned up
                         (without deleting them itself), and restarts the original executable.

    *.spec            -- PyInstaller .spec files specifying how to compile our main script(s) and
                         separate update-utility.
