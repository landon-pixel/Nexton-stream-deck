from __future__ import annotations

import configparser
from dataclasses import dataclass
import hashlib
import json
import os
from pathlib import Path
import subprocess

from nextion_stream_deck.config import ICON_CACHE_DIR


@dataclass
class AppMetadata:
    label: str
    action_type: str
    payload: str
    source_path: str = ""
    icon_path: str = ""


def import_app_metadata(source: str) -> AppMetadata:
    path = Path(source).expanduser().resolve()
    suffix = path.suffix.lower()

    if suffix == ".url":
        metadata = _metadata_from_url(path)
    elif suffix == ".lnk":
        metadata = _metadata_from_shortcut(path)
    else:
        metadata = _metadata_from_path(path)

    metadata.source_path = str(path)
    metadata.icon_path = metadata.icon_path or extract_icon_png(path)
    return metadata


def _metadata_from_url(path: Path) -> AppMetadata:
    parser = configparser.ConfigParser()
    parser.read(path, encoding="utf-8")
    url = parser.get("InternetShortcut", "URL", fallback="").strip()
    icon_file = parser.get("InternetShortcut", "IconFile", fallback="").strip()
    label = path.stem
    icon_path = ""
    if icon_file:
        icon_candidate = Path(os.path.expandvars(icon_file))
        if icon_candidate.exists():
            icon_path = str(icon_candidate)
    return AppMetadata(label=label, action_type="url", payload=url, icon_path=icon_path)


def _metadata_from_shortcut(path: Path) -> AppMetadata:
    payload = _powershell_json(
        f"""
$shortcut = (New-Object -ComObject WScript.Shell).CreateShortcut('{_ps_escape(str(path))}')
$target = $shortcut.TargetPath
$arguments = $shortcut.Arguments
$iconLocation = $shortcut.IconLocation
$label = [System.IO.Path]::GetFileNameWithoutExtension('{_ps_escape(path.name)}')
if ($target -and (Test-Path $target)) {{
  $version = (Get-Item $target).VersionInfo
  if ($version.FileDescription) {{ $label = $version.FileDescription }}
}}
[pscustomobject]@{{
  label = $label
  target = $target
  arguments = $arguments
  iconLocation = $iconLocation
}} | ConvertTo-Json -Compress
"""
    )
    target = str(payload.get("target", "")).strip()
    arguments = str(payload.get("arguments", "")).strip()
    label = str(payload.get("label", "")).strip() or path.stem
    icon_location = str(payload.get("iconLocation", "")).strip()
    icon_path = ""
    if icon_location:
        icon_path = icon_location.split(",")[0].strip()
    launch_payload = target if not arguments else f'"{target}" {arguments}'
    return AppMetadata(label=label, action_type="launch", payload=launch_payload, icon_path=icon_path)


def _metadata_from_path(path: Path) -> AppMetadata:
    if path.suffix.lower() in {".ps1", ".bat", ".cmd"}:
        return AppMetadata(label=path.stem, action_type="command", payload=str(path))

    description = _powershell_description(path)
    label = description or path.stem
    return AppMetadata(label=label, action_type="launch", payload=str(path))


def extract_icon_png(path: Path) -> str:
    if not path.exists():
        return ""
    ICON_CACHE_DIR.mkdir(parents=True, exist_ok=True)
    digest = hashlib.sha1(str(path).encode("utf-8")).hexdigest()[:16]
    destination = ICON_CACHE_DIR / f"{digest}.png"
    if destination.exists():
        return str(destination)

    script = f"""
Add-Type -AssemblyName System.Drawing
$source = '{_ps_escape(str(path))}'
$destination = '{_ps_escape(str(destination))}'
$resolved = $source
if ($source.ToLower().EndsWith('.lnk')) {{
  $shortcut = (New-Object -ComObject WScript.Shell).CreateShortcut($source)
  if ($shortcut.IconLocation) {{
    $resolved = $shortcut.IconLocation.Split(',')[0]
  }} elseif ($shortcut.TargetPath) {{
    $resolved = $shortcut.TargetPath
  }}
}}
if ($resolved -and (Test-Path $resolved)) {{
  $icon = [System.Drawing.Icon]::ExtractAssociatedIcon($resolved)
  if ($icon -ne $null) {{
    $bitmap = $icon.ToBitmap()
    $bitmap.Save($destination, [System.Drawing.Imaging.ImageFormat]::Png)
    $bitmap.Dispose()
    $icon.Dispose()
    'ok'
  }}
}}
"""
    result = subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script],
        capture_output=True,
        text=True,
        check=False,
    )
    if destination.exists() and result.returncode == 0:
        return str(destination)
    return ""


def _powershell_description(path: Path) -> str:
    payload = _powershell_json(
        f"""
$item = Get-Item '{_ps_escape(str(path))}'
$version = $item.VersionInfo
[pscustomobject]@{{
  description = $version.FileDescription
}} | ConvertTo-Json -Compress
"""
    )
    return str(payload.get("description", "")).strip()


def _powershell_json(script: str) -> dict:
    result = subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script],
        capture_output=True,
        text=True,
        check=False,
    )
    if result.returncode != 0:
        return {}
    output = result.stdout.strip()
    if not output:
        return {}
    try:
        return json.loads(output)
    except json.JSONDecodeError:
        return {}


def _ps_escape(value: str) -> str:
    return value.replace("'", "''")
