# Instant Replay Suite [![Latest Release](https://img.shields.io/github/release/thisismy-github/instant-replay-suite/all.svg)](https://github.com/thisismy-github/instant-replay-suite/releases) [![Download Count](https://img.shields.io/github/downloads/thisismy-github/instant-replay-suite/total?color=success)](https://github.com/thisismy-github/instant-replay-suite/releases/latest)

A suite of tools for formatting and editing Instant Replay clips, both in-game and out.

![Screenshot #1](https://i.imgur.com/6Q4S8r6.png)

# Core Features

## Trimming

Use customizable hotkeys to instantly trim clips to any length you choose, or even enter the trim length yourself using your number keys (requires [FFmpeg](https://www.gyan.dev/ffmpeg/builds/ffmpeg-release-essentials.zip)).

## Splicing

Use another hotkey to instantly concatenate two consecutive clips together (requires [FFmpeg](https://www.gyan.dev/ffmpeg/builds/ffmpeg-release-essentials.zip)).

## Backups/Undo

Automatically undo your last edit using another hotkey. Backups for your last several edited clips will be saved to give you further peace of mind.

## Automatically Rename Clips

Renames clips as they're saved using a highly customizable naming format. Allows you to set "aliases" for games to shorten their names while renaming. A handful of common aliases are included by default, i.e. "Left 4 Dead 2" becomes "L4D2".

## System Tray Icon

A *highly* customizable tray menu that can be accessed at any time that offers all of the same features listed above. Edit, delete, play, and explore any clip you want, not just the latest one, while also getting to see key information about each clip. Offers customizable left-click and middle-click actions.

# Other features

## Auto-updates

GitHub will be checked for a new release every time you launch Instant Replay Suite. If you're using the [compiled release](https://github.com/thisismy-github/instant-replay-suite/releases/latest), you'll have the option to automatically download/install the update straight from GitHub.

## Audio Alerts

Plays (yes, customizable) .WAV files in real-time to give you audible feedback for various actions being taken, even while in-game.

## Retroactively Add Existing Clips

On first launch, you'll have the option to add all existing clips in your video folder. This includes renaming any clips using ShadowPlay's naming format to match your desired format.

## Anti-Cheat Safe

Instant Replay Suite does **not** simulate any keyboard inputs. It only reads them, making it safe to use online. The cost is that you have to manually save clips yourself and *then* use separate hotkeys to edit them, but this in turn allows you to save clips and worry about editing them later.

# Contributing/Compiling

See the `executable` folder for details on how to compile (it's REALLY easy).

- [If you're new to contributing in general, you can use this guide](https://www.dataschool.io/how-to-contribute-on-github/).
- Follow the [seven rules of a great commit message](https://cbea.ms/git-commit/#seven-rules).
- Try to match the style of the code surrounding your addition. Don't let your code stick out like a sore thumb.
- Code should be as self-documenting as possible, with only minor explanatory comments/docstrings (if any).
- `configparsebetter.py` is currently off-limits (part of a future project).
- Avoid introducing new, heavy dependencies where possible.
- Avoid relative paths where possible (use `CWD` for the root folder and `pjoin` for creating paths).
- Be mindful of `REPOSITORY_URL` when making commits if you changed it while testing.
