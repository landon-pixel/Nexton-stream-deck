# NextDeck

Turn a Nextion HMI panel into a Windows desktop control surface with Stream Deck-style pages, app launch buttons, hotkeys, media controls, and custom artwork.

## Highlights

- Serial bridge for Nextion touch events over COM
- Multi-page deck profiles mapped to Nextion page ids
- Responsive tiles that scale with the window size
- Light and dark themes
- Layout presets including `5 x 3` and `3 x 2`
- App import for `.exe`, `.lnk`, `.url`, `.bat`, `.cmd`, and `.ps1`
- Per-tile custom name, icon/photo, shortcut keys, and label sync target
- Media controls such as `play_pause`, `next_track`, and `previous_track`
- Profile storage in JSON for easy backup and editing

## Requirements

- Windows
- Python 3.13 or newer
- A Nextion HMI connected over serial
- Matching baud settings on both the Nextion side and the app side

## Quick Start

1. Install dependencies:

```powershell
pip install -r requirements.txt
```

2. Start the app:

```powershell
python app.py
```

If Windows starts you in the wrong folder, use:

```powershell
.\run_app.ps1
```

or double-click `run_app.bat`.

To run it without leaving a terminal window open, use:

```powershell
.\run_app_silent.bat
```

or double-click `run_app_silent.vbs`.

## How It Works

The app listens for standard Nextion touch event packets:

```text
0x65 page component event 0xFF 0xFF 0xFF
```

When a touch press arrives, the app matches the `page + component` pair to a configured tile and runs the tile action on your PC.

## Features

### App Tiles

- Launch desktop apps, files, scripts, and URLs
- Send keyboard shortcuts
- Launch an app and then send follow-up shortcut keys
- Import metadata from Windows shortcuts and executables

### Visual Customization

- Custom tile names
- Custom icons or photos per tile
- Automatic square crop/resize for imported artwork
- Automatic tile scaling as the window resizes
- Stable spacing so tiles shrink instead of colliding

### Pages And Layouts

- Multiple named deck pages
- Nextion page id mapping per page
- Layout presets for different panel densities
- Page duplication for quick iteration

### Nextion Integration

- Label sync back to named HMI components
- Manual sync per tile or batch sync for all labels

## Setup Your Nextion Project

For each tappable control on the HMI:

- Enable `Touch Press Event`
- Note the page id and component id
- Map that pair in the app editor

If a tile uses label sync, set a `label_target` like:

```text
page0.b0
```

The app will send commands like:

```text
page0.b0.txt="OBS"
```

## Adding Apps

Use `Import App` in the editor and choose one of:

- `.exe`
- `.lnk`
- `.url`
- `.bat`
- `.cmd`
- `.ps1`

The app will try to collect:

- a display name
- the launch target or URL
- an icon source

You can override all of that afterward with:

- `Custom Name`
- `Choose Photo/Icon`
- `Shortcut Keys`

## Media Controls

For Spotify and other media apps, use a tile with `Action Type` set to `hotkey` and one of these payloads:

- `play_pause`
- `next_track`
- `previous_track`
- `media_stop`

You can also put those values in `Shortcut Keys` if you want them to fire after a launch action.

## Profile Format

Profiles are stored in the app data profile folder, with a default profile created automatically.

In the source workspace you will still see a starter profile at:

```text
profiles/default.json
```

Each profile stores:

- `rows`
- `cols`
- `pages`
- `active_page`
- `theme_mode`

Each tile stores:

- `slot`
- `page_id`
- `component_id`
- `label`
- `label_target`
- `action_type`
- `payload`
- `icon_path`
- `source_path`
- `shortcut_keys`

## Project Layout

```text
app.py
nextion_stream_deck/
  actions.py
  config.py
  metadata.py
  protocol.py
  serial_bridge.py
  ui.py
profiles/
tests/
```

## Tile Editor

- Select any tile to edit its name, action, ids, label target, icon, and shortcut keys
- The editor panel scrolls when needed
- `Apply` and `Test Action` stay pinned at the bottom of the editor

## Development

Run the tests with:

```powershell
python -m unittest discover -s tests -v
```

To build a Windows executable:

```powershell
.\build_app.ps1
```

The packaged app will be created in `dist/NextDeck.exe`.

## Current Notes

- The app expects a real Windows serial `COM` port for the Nextion connection
- If a Nextion is only powered over USB and no `COM` port appears, you may need a USB-to-TTL serial adapter depending on the model
- For best custom artwork results, use square PNG images
  
## Future Updates

- Optimize the app more to reduce storage required and resources being used
- Add a custom theme, and RGB mode
- Create a linux port and possibly an android port
- Broaden to other hmi devices, and create nextion firmware
- Possibly create a wireless esp or raspberry pi version(would be a fork)
