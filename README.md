# Nextion Stream Deck

This desktop app turns a Nextion HMI panel into a Stream Deck-style launcher on Windows.

It listens for touch events from a Nextion display over serial, maps each touched component to an action, and runs that action on the PC.

## What it does

- Connects to a Nextion HMI over a COM port
- Supports multiple deck pages, each mapped to a Nextion page id
- Supports fixed-size deck tiles and layout presets including `3 x 2`
- Maps `page + component` touch events to buttons
- Supports a light and dark editor theme
- Runs actions such as:
  - launching apps or files
  - opening URLs
  - running shell commands
  - sending keyboard shortcuts
- Imports Windows apps and shortcuts, then fills in the label, launch target, and icon
- Saves button profiles as JSON
- Can push labels back to named text/button components on the Nextion

## Quick start

1. Install Python 3.13 or newer.
2. Install dependencies:

```powershell
pip install -r requirements.txt
```

3. Run the app:

```powershell
python app.py
```

If Windows starts in another folder and cannot find `app.py`, use one of these launchers from the project directory instead:

```powershell
.\run_app.ps1
```

or double-click `run_app.bat`.

## Nextion setup

The app expects normal Nextion touch event packets (`0x65 page component event 0xFF 0xFF 0xFF`).

For each tappable control on your HMI:

- make sure the component has `Touch Press Event` enabled in the HMI project
- note the page id and component id
- map that pair in the app

Optional label sync:

- if a mapped button includes a `label_target`, the app will send a command like:

```text
page0.b0.txt="OBS"
```

That means your HMI project should contain a component with that exact object name.

## Adding custom apps

Use the `Import App` button in the editor and choose:

- `.exe` files
- `.lnk` shortcuts
- `.url` internet shortcuts
- `.bat`, `.cmd`, or `.ps1` scripts

The app will try to collect:

- display name
- launch target or URL
- icon image cached into `assets/icons`

You can also customize each tile with:

- a custom display name in `Custom Name`
- custom art with `Choose Photo/Icon`
- optional `Shortcut Keys` like `ctrl+shift+s`

If `Shortcut Keys` is filled in for a launch action, the app launches the target and then sends that shortcut.

Custom art is automatically converted into a square tile image so button sizes stay consistent.

## Pages

The editor supports multiple named deck pages.

- `Page Name` controls the editor tab name
- `Nextion Page ID` determines which Nextion page should trigger those buttons
- `Layout` can switch the whole profile between presets like `5 x 3` and `3 x 2`
- `Duplicate Page` makes it easy to clone a layout and tweak it

Each page still uses the same grid size, but every page has its own buttons, labels, icons, and actions.

## Profile format

Profiles are stored in `profiles/default.json`.

Each profile includes:

- `rows`
- `cols`
- `pages`
- `active_page`
- `theme_mode`

Each page includes:

- `name`
- `nextion_page_id`
- `buttons`

Each button entry includes:

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

## Notes

- `hotkey` actions use native Windows key injection.
- Media controls can use `play_pause`, `next_track`, and `previous_track` as hotkey payloads or shortcut keys.
- `command` actions run through PowerShell.
- If your Nextion is connected through a USB serial adapter, select the matching COM port in the app.
