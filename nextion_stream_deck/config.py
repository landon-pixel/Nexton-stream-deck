from __future__ import annotations

from dataclasses import asdict, dataclass, field
import json
from pathlib import Path

from nextion_stream_deck.paths import app_data_dir


PROFILE_DIR = app_data_dir() / "profiles"
DEFAULT_PROFILE_PATH = PROFILE_DIR / "default.json"
ICON_CACHE_DIR = app_data_dir() / "icons"


@dataclass
class ButtonMapping:
    slot: int
    page_id: int = 0
    component_id: int = 0
    label: str = ""
    label_target: str = ""
    action_type: str = "launch"
    payload: str = ""
    icon_path: str = ""
    source_path: str = ""
    shortcut_keys: str = ""


@dataclass
class DeckPage:
    name: str
    nextion_page_id: int = 0
    buttons: list[ButtonMapping] = field(default_factory=list)


@dataclass
class Profile:
    name: str = "Default"
    baud_rate: int = 9600
    rows: int = 3
    cols: int = 5
    pages: list[DeckPage] = field(default_factory=list)
    active_page: int = 0
    theme_mode: str = "dark"
    style_mode: str = "default"  # "default" shows background, "alternate" uses plain theme colors


def create_default_buttons(rows: int = 3, cols: int = 5, page_id: int = 0) -> list[ButtonMapping]:
    buttons: list[ButtonMapping] = []
    for slot in range(rows * cols):
        buttons.append(
            ButtonMapping(
                slot=slot,
                page_id=page_id,
                component_id=slot + 1,
                label=f"Key {slot + 1}",
                label_target=f"page{page_id}.b{slot}",
            )
        )
    return buttons


def create_default_profile(rows: int = 3, cols: int = 5) -> Profile:
    return Profile(
        rows=rows,
        cols=cols,
        pages=[DeckPage(name="Page 1", nextion_page_id=0, buttons=create_default_buttons(rows, cols, 0))],
        active_page=0,
    )


def ensure_default_profile(path: Path = DEFAULT_PROFILE_PATH) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    if not path.exists():
        save_profile(create_default_profile(), path)
    return path


def load_profile(path: Path = DEFAULT_PROFILE_PATH) -> Profile:
    ensure_default_profile(path)
    data = json.loads(path.read_text(encoding="utf-8"))
    rows = int(data.get("rows", 3))
    cols = int(data.get("cols", 5))
    pages = _load_pages(data, rows, cols)
    profile = Profile(
        name=data.get("name", "Default"),
        baud_rate=int(data.get("baud_rate", 9600)),
        rows=rows,
        cols=cols,
        pages=pages,
        active_page=min(max(int(data.get("active_page", 0)), 0), max(len(pages) - 1, 0)),
        theme_mode=str(data.get("theme_mode", "dark")).lower(),
        style_mode=str(data.get("style_mode", "default")).lower(),
    )
    if not profile.pages:
        profile.pages = create_default_profile(rows, cols).pages
    return profile


def save_profile(profile: Profile, path: Path = DEFAULT_PROFILE_PATH) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    payload = asdict(profile)
    path.write_text(json.dumps(payload, indent=2), encoding="utf-8")


def ensure_page_shape(profile: Profile) -> None:
    if not profile.pages:
        profile.pages = create_default_profile(profile.rows, profile.cols).pages
    for index, page in enumerate(profile.pages):
        if not page.name:
            page.name = f"Page {index + 1}"
        if page.nextion_page_id < 0:
            page.nextion_page_id = index
        _ensure_buttons(page, profile.rows, profile.cols, page.nextion_page_id)
    if profile.active_page >= len(profile.pages):
        profile.active_page = 0


def _load_pages(data: dict, rows: int, cols: int) -> list[DeckPage]:
    raw_pages = data.get("pages")
    if raw_pages:
        pages: list[DeckPage] = []
        for index, raw_page in enumerate(raw_pages):
            nextion_page_id = int(raw_page.get("nextion_page_id", index))
            buttons = [ButtonMapping(**button) for button in raw_page.get("buttons", [])]
            page = DeckPage(
                name=raw_page.get("name", f"Page {index + 1}"),
                nextion_page_id=nextion_page_id,
                buttons=buttons,
            )
            _ensure_buttons(page, rows, cols, nextion_page_id)
            pages.append(page)
        return pages

    # Legacy flat-profile migration.
    legacy_buttons = [ButtonMapping(**button) for button in data.get("buttons", [])]
    legacy_page = DeckPage(name="Page 1", nextion_page_id=0, buttons=legacy_buttons)
    _ensure_buttons(legacy_page, rows, cols, 0)
    return [legacy_page]


def _ensure_buttons(page: DeckPage, rows: int, cols: int, page_id: int) -> None:
    expected = rows * cols
    if not page.buttons:
        page.buttons = create_default_buttons(rows, cols, page_id)
    existing_slots = {button.slot for button in page.buttons}
    defaults = create_default_buttons(rows, cols, page_id)
    for default_button in defaults:
        if default_button.slot not in existing_slots:
            page.buttons.append(default_button)
    page.buttons.sort(key=lambda button: button.slot)
    for button in page.buttons[:expected]:
        if not button.label_target:
            button.label_target = f"page{page_id}.b{button.slot}"
        button.page_id = page_id
    del page.buttons[expected:]
