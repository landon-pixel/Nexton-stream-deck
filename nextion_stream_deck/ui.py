from __future__ import annotations

import math
import queue
from pathlib import Path
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk

from nextion_stream_deck.actions import run_mapping
from nextion_stream_deck.config import DEFAULT_PROFILE_PATH, ICON_CACHE_DIR, PROFILE_DIR, ButtonMapping, DeckPage, ensure_page_shape, load_profile, save_profile
from nextion_stream_deck.metadata import import_app_metadata
from nextion_stream_deck.paths import resource_path
from nextion_stream_deck.protocol import NextionTouchEvent
from nextion_stream_deck.serial_bridge import NextionBridge


ACTION_TYPES = ("launch", "url", "command", "hotkey")
APP_TITLE = "NextDeck"
APP_VERSION = "1.0.0"
LAYOUT_PRESETS = {
    "5 x 3": (5, 3),
    "3 x 2": (3, 2),
}
IMAGE_SIZE = 96
MIN_TILE_IMAGE = 44
PANEL_WIDTH = 440
HEADER_HEIGHT = 108
OUTER_PAD = 24

THEMES = {
    "dark": {
        "window_bg": "#050816",
        "panel_bg": "#0f1728",
        "panel_border": "#22304d",
        "card_idle": "#15233b",
        "card_active": "#3b82f6",
        "card_hover": "#22365e",
        "header_fg": "#f8fafc",
        "text_primary": "#dbe8ff",
        "text_muted": "#94a3b8",
        "field_bg": "#0a1020",
        "field_fg": "#eef4ff",
        "button_fg": "#eef4ff",
        "placeholder_inner": "#7c3aed",
        "accent": "#8b5cf6",
    },
    "light": {
        "window_bg": "#eaf0ff",
        "panel_bg": "#ffffff",
        "panel_border": "#c8d7f0",
        "card_idle": "#dce9ff",
        "card_active": "#2563eb",
        "card_hover": "#c5dcff",
        "header_fg": "#14213d",
        "text_primary": "#1e293b",
        "text_muted": "#5b6b86",
        "field_bg": "#f8fbff",
        "field_fg": "#10213f",
        "button_fg": "#10213f",
        "placeholder_inner": "#93c5fd",
        "accent": "#2563eb",
    },
}


class App:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("NextDeck")
        self.default_geometry = "1460x900"
        self.root.geometry(self.default_geometry)
        self.root.minsize(1220, 760)
        self.profile_path = DEFAULT_PROFILE_PATH
        self.profile = load_profile(self.profile_path)
        ensure_page_shape(self.profile)

        self.selected_slot = 0
        self.message_queue: queue.Queue[tuple[str, object]] = queue.Queue()
        self.bridge = NextionBridge(self._queue_event, self._queue_status)
        self.icon_cache: dict[str, tk.PhotoImage] = {}
        self.grid_buttons: list[tk.Canvas] = []
        self._render_job: str | None = None
        self.port_combo: ttk.Combobox | None = None
        self.settings_port_combo: ttk.Combobox | None = None

        self.port_var = tk.StringVar()
        self.baud_var = tk.StringVar(value=str(self.profile.baud_rate))
        self.status_var = tk.StringVar(value="Disconnected")
        self.last_touch_var = tk.StringVar(value="No touch events yet")
        self.page_var = tk.StringVar()
        self.page_name_var = tk.StringVar()
        self.nextion_page_var = tk.StringVar()
        self.theme_var = tk.StringVar(value=self.profile.theme_mode or "dark")
        self.layout_var = tk.StringVar(value=self._layout_label())

        self.slot_title_var = tk.StringVar()
        self.page_id_var = tk.StringVar()
        self.component_id_var = tk.StringVar()
        self.label_var = tk.StringVar()
        self.label_target_var = tk.StringVar()
        self.action_type_var = tk.StringVar(value=ACTION_TYPES[0])
        self.icon_path_var = tk.StringVar()
        self.source_path_var = tk.StringVar()
        self.shortcut_var = tk.StringVar()

        self._load_brand_assets()
        self._configure_theme()
        self._build_ui()
        self._sync_window_mode_with_layout()
        self._refresh_page_tabs()
        self._render_grid()
        self._load_mapping_into_editor(0)
        self.refresh_ports()
        self.root.after(50, self._process_messages)
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        self.root.bind("<Configure>", self._on_root_resize)

    @property
    def current_page(self) -> DeckPage:
        ensure_page_shape(self.profile)
        return self.profile.pages[self.profile.active_page]

    def _theme(self) -> dict[str, str]:
        return THEMES.get(self.theme_var.get().strip().lower(), THEMES["dark"])

    def _load_brand_assets(self) -> None:
        self.background_image = self._load_photo(resource_path("assets", "background", "cool background.png"))
        self.logo_image = self._load_photo(resource_path("assets", "logo", "Nextdeck logo.png"), subsample=12)
        self.header_logo = self._load_photo(resource_path("assets", "logo", "Nextdeck logo.png"), subsample=18)
        self.icon_logo = self._load_photo(resource_path("assets", "logo", "Nextdeck logo.png"), subsample=14)
        if self.icon_logo:
            self.root.iconphoto(True, self.icon_logo)

    def _load_photo(self, path: Path, subsample: int | None = None) -> tk.PhotoImage | None:
        if not path.exists():
            return None
        try:
            image = tk.PhotoImage(file=str(path))
            if subsample and subsample > 1:
                image = image.subsample(subsample, subsample)
            return image
        except tk.TclError:
            return None

    def _configure_theme(self) -> None:
        colors = self._theme()
        self.root.configure(bg=colors["window_bg"])
        style = ttk.Style()
        try:
            style.theme_use("vista")
        except tk.TclError:
            pass
        style.configure("Deck.TFrame", background=colors["panel_bg"])
        style.configure("Deck.TLabel", background=colors["panel_bg"], foreground=colors["text_primary"], font=("Segoe UI", 10))
        style.configure("Hero.TLabel", background=colors["panel_bg"], foreground=colors["header_fg"], font=("Segoe UI Semibold", 18))
        style.configure("Muted.TLabel", background=colors["panel_bg"], foreground=colors["text_muted"], font=("Segoe UI", 10))
        style.configure("TitleBar.TLabel", background=colors["window_bg"], foreground=colors["header_fg"], font=("Segoe UI Semibold", 18))
        style.configure("TitleBarMuted.TLabel", background=colors["window_bg"], foreground=colors["text_muted"], font=("Segoe UI", 10))
        style.configure("Accent.TButton", font=("Segoe UI Semibold", 10))

    def _build_ui(self) -> None:
        colors = self._theme()

        if hasattr(self, "top_frame") and self.top_frame.winfo_exists():
            self.top_frame.destroy()
        if hasattr(self, "main_frame") and self.main_frame.winfo_exists():
            self.main_frame.destroy()
        if hasattr(self, "background_canvas") and self.background_canvas.winfo_exists():
            self.background_canvas.destroy()
        if hasattr(self, "deck_shell") and self.deck_shell.winfo_exists():
            self.deck_shell.destroy()
        if hasattr(self, "editor_shell") and self.editor_shell.winfo_exists():
            self.editor_shell.destroy()

        self.background_canvas = tk.Canvas(self.root, bg=colors["window_bg"], highlightthickness=0, bd=0)
        self.background_canvas.place(x=0, y=0, relwidth=1, relheight=1)
        self.background_item = self.background_canvas.create_image(0, 0, anchor="center", image=self.background_image)

        self.top_frame = tk.Frame(self.root, bg=colors["window_bg"], padx=26, pady=16)
        self.top_frame.place(x=OUTER_PAD, y=OUTER_PAD - 4, relwidth=1, width=-(OUTER_PAD * 2), height=HEADER_HEIGHT)

        if self.header_logo:
            tk.Label(self.top_frame, image=self.header_logo, bg=colors["window_bg"], bd=0).pack(side="left", padx=(0, 16), pady=(4, 0))

        title_stack = tk.Frame(self.top_frame, bg=colors["window_bg"])
        title_stack.pack(side="left", fill="x", expand=True, pady=(8, 0))
        ttk.Label(title_stack, text=APP_TITLE, style="TitleBar.TLabel").pack(anchor="w")
        ttk.Label(
            title_stack,
            text="Nextion control surface with app tiles, pages, labels, and media keys.",
            style="TitleBarMuted.TLabel",
        ).pack(anchor="w", pady=(2, 0))

        top_controls = tk.Frame(self.top_frame, bg=colors["window_bg"])
        top_controls.pack(side="right", pady=(10, 0))
        ttk.Label(top_controls, textvariable=self.status_var, style="TitleBarMuted.TLabel").pack(side="left", padx=(0, 12))
        ttk.Button(top_controls, text="Settings", command=self.show_settings).pack(side="left")
        ttk.Button(top_controls, text="About", command=self.show_about).pack(side="left", padx=(10, 0))

        self.deck_shell = tk.Frame(
            self.root,
            bg=colors["panel_bg"],
            highlightbackground=colors["panel_border"],
            highlightthickness=1,
            bd=0,
        )
        self.deck_shell.place(
            x=OUTER_PAD,
            y=HEADER_HEIGHT + OUTER_PAD,
            relwidth=1,
            width=-(PANEL_WIDTH + OUTER_PAD * 4),
            relheight=1,
            height=-(HEADER_HEIGHT + OUTER_PAD * 3),
        )

        self.editor_shell = tk.Frame(
            self.root,
            bg=colors["panel_bg"],
            highlightbackground=colors["panel_border"],
            highlightthickness=1,
            bd=0,
            width=PANEL_WIDTH,
        )
        self.editor_shell.place(
            relx=1,
            x=-(PANEL_WIDTH + OUTER_PAD * 2),
            y=HEADER_HEIGHT + OUTER_PAD,
            width=PANEL_WIDTH,
            relheight=1,
            height=-(HEADER_HEIGHT + OUTER_PAD * 3),
        )

        self._build_deck_shell()
        self._build_editor_shell()

        self.deck_holder.bind("<Configure>", self._on_deck_resize)

    def _build_deck_shell(self) -> None:
        colors = self._theme()
        top = tk.Frame(self.deck_shell, bg=colors["panel_bg"], padx=24, pady=22)
        top.pack(fill="x")

        header_text = tk.Frame(top, bg=colors["panel_bg"])
        header_text.pack(side="left", fill="x", expand=True)
        ttk.Label(header_text, text="Deck Pages", style="Hero.TLabel").pack(anchor="w")
        ttk.Label(header_text, textvariable=self.last_touch_var, style="Muted.TLabel").pack(anchor="w", pady=(4, 0))

        self.page_tabs = ttk.Combobox(top, textvariable=self.page_var, state="readonly", width=26)
        self.page_tabs.pack(side="right", padx=(12, 0))
        self.page_tabs.bind("<<ComboboxSelected>>", self._on_page_selected)

        meta = tk.Frame(self.deck_shell, bg=colors["panel_bg"], padx=24)
        meta.pack(fill="x", pady=(0, 12))
        for column in range(8):
            meta.grid_columnconfigure(column, weight=1 if column in (1, 3, 5) else 0)

        self._inline_field(meta, "Page", self.page_name_var, 0, 0, width=24)
        self._inline_field(meta, "Nextion ID", self.nextion_page_var, 0, 2, width=8)

        ttk.Label(meta, text="Layout", style="Deck.TLabel").grid(row=0, column=4, sticky="w", padx=(14, 0))
        self.layout_box = ttk.Combobox(meta, textvariable=self.layout_var, values=tuple(LAYOUT_PRESETS.keys()), state="readonly", width=8)
        self.layout_box.grid(row=0, column=5, sticky="ew", padx=(8, 12))

        controls = tk.Frame(meta, bg=colors["panel_bg"])
        controls.grid(row=0, column=6, columnspan=2, sticky="e")
        ttk.Button(controls, text="Apply Page", command=self.apply_page_settings).pack(side="left")
        ttk.Button(controls, text="Add", command=self.add_page).pack(side="left", padx=(8, 0))
        ttk.Button(controls, text="Copy", command=self.duplicate_page).pack(side="left", padx=(8, 0))
        ttk.Button(controls, text="Delete", command=self.delete_page).pack(side="left", padx=(8, 0))

        self.deck_holder = tk.Frame(self.deck_shell, bg=colors["panel_bg"], padx=24, pady=10)
        self.deck_holder.pack(fill="both", expand=True)

    def _build_editor_shell(self) -> None:
        colors = self._theme()
        shell = tk.Frame(self.editor_shell, bg=colors["panel_bg"])
        shell.pack(fill="both", expand=True)

        canvas = tk.Canvas(shell, bg=colors["panel_bg"], highlightthickness=0, bd=0)
        scrollbar = ttk.Scrollbar(shell, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y", padx=(0, 8), pady=(18, 86))
        canvas.pack(side="top", fill="both", expand=True, padx=(0, 2), pady=(0, 0))

        wrap = tk.Frame(canvas, bg=colors["panel_bg"], padx=22, pady=22)
        canvas_window = canvas.create_window((0, 0), window=wrap, anchor="nw")

        def sync_editor_scroll(_event: object | None = None) -> None:
            canvas.configure(scrollregion=canvas.bbox("all"))
            canvas.itemconfigure(canvas_window, width=max(260, canvas.winfo_width() - 4))

        wrap.bind("<Configure>", sync_editor_scroll)
        canvas.bind("<Configure>", sync_editor_scroll)

        ttk.Label(wrap, textvariable=self.slot_title_var, style="Hero.TLabel").pack(anchor="w")
        ttk.Label(wrap, text="Edit the selected tile, app action, and sync target.", style="Muted.TLabel").pack(anchor="w", pady=(2, 12))

        ids_row = tk.Frame(wrap, bg=colors["panel_bg"])
        ids_row.pack(fill="x")
        self._stack_field(ids_row, "Tile Name", self.label_var)

        pair_row = tk.Frame(wrap, bg=colors["panel_bg"])
        pair_row.pack(fill="x", pady=(12, 0))
        left = tk.Frame(pair_row, bg=colors["panel_bg"])
        left.pack(side="left", fill="x", expand=True)
        right = tk.Frame(pair_row, bg=colors["panel_bg"])
        right.pack(side="left", fill="x", expand=True, padx=(12, 0))
        self._stack_field(left, "Page ID", self.page_id_var, compact=True)
        self._stack_field(right, "Component ID", self.component_id_var, compact=True)

        self._stack_field(wrap, "Label Target", self.label_target_var, compact=True)

        action_row = tk.Frame(wrap, bg=colors["panel_bg"])
        action_row.pack(fill="x", pady=(12, 0))
        action_left = tk.Frame(action_row, bg=colors["panel_bg"])
        action_left.pack(side="left", fill="x", expand=True)
        action_right = tk.Frame(action_row, bg=colors["panel_bg"])
        action_right.pack(side="left", fill="x", expand=True, padx=(12, 0))
        ttk.Label(action_left, text="Action Type", style="Deck.TLabel").pack(anchor="w", pady=(0, 6))
        self.action_box = ttk.Combobox(action_left, textvariable=self.action_type_var, values=ACTION_TYPES, state="readonly")
        self.action_box.pack(fill="x")
        self._stack_field(action_right, "Shortcut Keys", self.shortcut_var, compact=True, top_padding=0)

        ttk.Label(wrap, text="Payload", style="Deck.TLabel").pack(anchor="w", pady=(12, 6))
        self.payload_entry = tk.Text(
            wrap,
            height=4,
            bg=colors["field_bg"],
            fg=colors["field_fg"],
            insertbackground=colors["field_fg"],
            relief="flat",
            font=("Consolas", 10),
            wrap="word",
        )
        self.payload_entry.pack(fill="x")

        self._stack_field(wrap, "Source Path", self.source_path_var, compact=True)
        self._stack_field(wrap, "Custom Photo/Icon", self.icon_path_var, compact=True)

        utility_rows = [
            ("Import App", self.import_app, "Choose Photo/Icon", self.choose_icon),
            ("Clear Art", self.clear_icon, "Use Current Name", self.use_source_name),
            ("Sync Label", self.sync_selected_label, "Sync All Labels", self.sync_all_labels),
            ("Open Profile", self.open_profile, "Save Profile", self.save_profile),
        ]
        for left_text, left_cmd, right_text, right_cmd in utility_rows:
            row = tk.Frame(wrap, bg=colors["panel_bg"])
            row.pack(fill="x", pady=(10, 0))
            ttk.Button(row, text=left_text, command=left_cmd).pack(side="left", fill="x", expand=True)
            ttk.Button(row, text=right_text, command=right_cmd).pack(side="left", fill="x", expand=True, padx=(10, 0))

        footer = tk.Frame(shell, bg=colors["panel_bg"], padx=18, pady=14)
        footer.pack(side="bottom", fill="x")
        ttk.Button(footer, text="Apply", style="Accent.TButton", command=self.apply_current_edits).pack(side="left", fill="x", expand=True)
        ttk.Button(footer, text="Test Action", command=self.test_action).pack(side="left", fill="x", expand=True, padx=(10, 0))

    def _inline_field(
        self,
        parent: tk.Widget,
        label: str,
        variable: tk.StringVar,
        row: int,
        column: int,
        width: int,
        parent_is_grid: bool = True,
    ) -> None:
        colors = self._theme()
        if parent_is_grid:
            ttk.Label(parent, text=label, style="Deck.TLabel").grid(row=row, column=column, sticky="w")
            tk.Entry(
                parent,
                textvariable=variable,
                width=width,
                bg=colors["field_bg"],
                fg=colors["field_fg"],
                insertbackground=colors["field_fg"],
                relief="flat",
            ).grid(row=row, column=column + 1, sticky="ew", padx=(8, 0))
            return
        tk.Label(parent, text=label, bg=colors["panel_bg"], fg=colors["text_muted"], font=("Segoe UI", 10)).pack(side="left")
        tk.Entry(
            parent,
            textvariable=variable,
            width=width,
            bg=colors["field_bg"],
            fg=colors["field_fg"],
            insertbackground=colors["field_fg"],
            relief="flat",
        ).pack(side="left", padx=(8, 0))

    def _stack_field(
        self,
        parent: tk.Widget,
        label: str,
        variable: tk.StringVar,
        compact: bool = False,
        top_padding: int = 12,
    ) -> None:
        colors = self._theme()
        ttk.Label(parent, text=label, style="Deck.TLabel").pack(anchor="w", pady=(top_padding, 6))
        tk.Entry(
            parent,
            textvariable=variable,
            bg=colors["field_bg"],
            fg=colors["field_fg"],
            insertbackground=colors["field_fg"],
            relief="flat",
        ).pack(fill="x", ipady=2 if compact else 0)

    def _refresh_page_tabs(self) -> None:
        ensure_page_shape(self.profile)
        labels = [f"{index + 1}. {page.name}" for index, page in enumerate(self.profile.pages)]
        self.page_tabs["values"] = labels
        self.page_var.set(labels[self.profile.active_page])
        self.page_name_var.set(self.current_page.name)
        self.nextion_page_var.set(str(self.current_page.nextion_page_id))
        self.layout_var.set(self._layout_label())

    def _tile_dimensions(self) -> tuple[int, int]:
        width = self.deck_holder.winfo_width() or self.deck_holder.winfo_reqwidth() or 900
        height = self.deck_holder.winfo_height() or self.deck_holder.winfo_reqheight() or 520
        gutter = 16
        available_width = max(240, width - gutter * (self.profile.cols + 1))
        available_height = max(220, height - gutter * (self.profile.rows + 1))
        tile_width = max(120, min(360, available_width // self.profile.cols))
        tile_height = max(118, min(320, available_height // self.profile.rows))
        return tile_width, tile_height

    def _render_grid(self) -> None:
        colors = self._theme()
        for child in self.deck_holder.winfo_children():
            child.destroy()
        self.grid_buttons = []

        tile_width, tile_height = self._tile_dimensions()
        for row in range(self.profile.rows):
            self.deck_holder.rowconfigure(row, weight=1)
        for col in range(self.profile.cols):
            self.deck_holder.columnconfigure(col, weight=1)

        for mapping in self.current_page.buttons:
            image = self._icon_for_mapping(mapping, tile_width, tile_height)
            tile = tk.Canvas(
                self.deck_holder,
                width=tile_width,
                height=tile_height,
                bg=colors["panel_bg"],
                highlightthickness=0,
                bd=0,
            )
            row = mapping.slot // self.profile.cols
            col = mapping.slot % self.profile.cols
            tile.grid(row=row, column=col, sticky="nsew", padx=8, pady=8)
            tile.image = image
            self._paint_tile(tile, mapping, selected=(mapping.slot == self.selected_slot))
            tile.bind("<Button-1>", lambda _e, slot=mapping.slot: self._load_mapping_into_editor(slot))
            self.grid_buttons.append(tile)

    @staticmethod
    def _button_caption(mapping: ButtonMapping) -> str:
        return f"{mapping.label or f'Slot {mapping.slot + 1}'}\nP{mapping.page_id} C{mapping.component_id}"

    def _icon_for_mapping(self, mapping: ButtonMapping, tile_width: int = 220, tile_height: int = 180) -> tk.PhotoImage:
        target_size = max(MIN_TILE_IMAGE, min(IMAGE_SIZE, tile_width - 42, tile_height - 86))
        if mapping.icon_path:
            path = Path(mapping.icon_path)
            if path.exists():
                key = f"{path.resolve()}::{target_size}"
                if key not in self.icon_cache:
                    try:
                        image = tk.PhotoImage(file=str(path))
                        source_size = max(image.width(), image.height(), 1)
                        if target_size < source_size:
                            factor = max(1, math.ceil(source_size / target_size))
                            image = image.subsample(factor, factor)
                        self.icon_cache[key] = image
                    except tk.TclError:
                        self.icon_cache[key] = self._placeholder_icon(mapping.label, target_size)
                return self.icon_cache[key]
        return self._placeholder_icon(mapping.label, target_size)

    def _placeholder_icon(self, label: str, size: int = IMAGE_SIZE) -> tk.PhotoImage:
        safe_size = max(MIN_TILE_IMAGE, size)
        key = f"placeholder:{(label[:1] or '?').upper()}:{self.theme_var.get()}:{safe_size}"
        if key in self.icon_cache:
            return self.icon_cache[key]
        inset = max(4, safe_size // 12)
        image = tk.PhotoImage(width=safe_size, height=safe_size)
        image.put(self._theme()["accent"], to=(0, 0, safe_size - 1, safe_size - 1))
        image.put(self._theme()["placeholder_inner"], to=(inset, inset, safe_size - inset - 1, safe_size - inset - 1))
        self.icon_cache[key] = image
        return image

    def _paint_tile(self, tile: tk.Canvas, mapping: ButtonMapping, selected: bool) -> None:
        colors = self._theme()
        tile.delete("all")
        width = int(tile.cget("width"))
        height = int(tile.cget("height"))
        radius = max(16, min(34, min(width, height) // 5))
        outer_pad = max(4, min(12, width // 18))
        inner_pad = max(10, min(18, width // 12))
        text_block = max(46, min(76, height // 3))
        meta_offset = max(14, min(22, height // 8))
        title_offset = max(34, min(48, height // 4))
        icon_band_bottom = max(inner_pad + MIN_TILE_IMAGE + 12, height - text_block)
        title_font = max(8, min(11, width // 16))
        meta_font = max(7, min(9, width // 22))
        fill = colors["card_active"] if selected else colors["card_idle"]
        outline = colors["accent"] if selected else colors["panel_border"]
        self._round_rect(tile, outer_pad, outer_pad, width - outer_pad, height - outer_pad, radius, fill=fill, outline=outline, width=2)
        self._round_rect(
            tile,
            inner_pad,
            inner_pad,
            width - inner_pad,
            icon_band_bottom,
            max(12, radius - 8),
            fill="#08101f",
            outline="",
        )
        if getattr(tile, "image", None):
            tile.create_image(width / 2, inner_pad + tile.image.height() / 2 + 4, image=tile.image)
        tile.create_text(
            width / 2,
            height - title_offset,
            text=mapping.label or f"Slot {mapping.slot + 1}",
            fill=colors["button_fg"],
            font=("Segoe UI Semibold", title_font),
            width=max(56, width - inner_pad * 2),
            justify="center",
        )
        tile.create_text(
            width / 2,
            height - meta_offset,
            text=f"P{mapping.page_id} C{mapping.component_id}",
            fill=colors["text_muted"],
            font=("Segoe UI", meta_font),
        )

    @staticmethod
    def _round_rect(canvas: tk.Canvas, x1: int, y1: int, x2: int, y2: int, radius: int, **kwargs: object) -> int:
        points = [
            x1 + radius, y1,
            x2 - radius, y1,
            x2, y1,
            x2, y1 + radius,
            x2, y2 - radius,
            x2, y2,
            x2 - radius, y2,
            x1 + radius, y2,
            x1, y2,
            x1, y2 - radius,
            x1, y1 + radius,
            x1, y1,
        ]
        return canvas.create_polygon(points, smooth=True, splinesteps=36, **kwargs)

    def _layout_label(self) -> str:
        for label, (cols, rows) in LAYOUT_PRESETS.items():
            if self.profile.cols == cols and self.profile.rows == rows:
                return label
        return f"{self.profile.cols} x {self.profile.rows}"

    def _sync_window_mode_with_layout(self) -> None:
        try:
            if self.profile.cols == 3 and self.profile.rows == 2:
                self.root.state("zoomed")
            else:
                if self.root.state() == "zoomed":
                    self.root.state("normal")
                self.root.geometry(self.default_geometry)
        except tk.TclError:
            pass

    def refresh_ports(self) -> None:
        ports = self.bridge.available_ports()
        if self.port_combo and self.port_combo.winfo_exists():
            self.port_combo["values"] = ports
        if self.settings_port_combo and self.settings_port_combo.winfo_exists():
            self.settings_port_combo["values"] = ports
        if ports and self.port_var.get() not in ports:
            self.port_var.set(ports[0])

    def connect(self) -> None:
        try:
            baud_rate = int(self.baud_var.get().strip())
            port = self.port_var.get().strip()
            if not port:
                raise ValueError("Select a COM port.")
            self.bridge.connect(port, baud_rate)
            self.status_var.set(f"Connected to {port} @ {baud_rate}")
        except Exception as exc:
            messagebox.showerror("Connection failed", str(exc))

    def disconnect(self) -> None:
        self.bridge.disconnect()
        self.status_var.set("Disconnected")

    def _queue_event(self, event: NextionTouchEvent) -> None:
        self.message_queue.put(("event", event))

    def _queue_status(self, message: str) -> None:
        self.message_queue.put(("status", message))

    def _process_messages(self) -> None:
        while not self.message_queue.empty():
            kind, payload = self.message_queue.get_nowait()
            if kind == "status":
                self.status_var.set(str(payload))
            elif kind == "event":
                self._handle_touch_event(payload)
        self.root.after(50, self._process_messages)

    def _handle_touch_event(self, event: NextionTouchEvent) -> None:
        state = "press" if event.pressed else "release"
        self.last_touch_var.set(f"Last touch: page {event.page_id}, component {event.component_id}, {state}")
        if not event.pressed:
            return
        mapping = self._find_mapping(event.page_id, event.component_id)
        if not mapping:
            self.status_var.set(f"No mapping for page {event.page_id}, component {event.component_id}")
            return
        page_index, slot = mapping
        try:
            current = self.profile.pages[page_index].buttons[slot]
            result = run_mapping(current.action_type, current.payload, current.shortcut_keys)
            self.status_var.set(result)
            self.profile.active_page = page_index
            self._refresh_page_tabs()
            self._render_grid()
            self._highlight_slot(slot)
        except Exception as exc:
            self.status_var.set(str(exc))

    def _find_mapping(self, page_id: int, component_id: int) -> tuple[int, int] | None:
        for page_index, page in enumerate(self.profile.pages):
            if page.nextion_page_id != page_id:
                continue
            for mapping in page.buttons:
                if mapping.component_id == component_id:
                    return page_index, mapping.slot
        return None

    def _highlight_slot(self, slot: int) -> None:
        self._load_mapping_into_editor(slot)
        if slot < len(self.grid_buttons):
            self.grid_buttons[slot].flash()

    def _load_mapping_into_editor(self, slot: int) -> None:
        self.selected_slot = slot
        mapping = self.current_page.buttons[slot]
        self.slot_title_var.set(f"{self.current_page.name} · Tile {slot + 1}")
        self.page_id_var.set(str(mapping.page_id))
        self.component_id_var.set(str(mapping.component_id))
        self.label_var.set(mapping.label)
        self.label_target_var.set(mapping.label_target)
        self.action_type_var.set(mapping.action_type or ACTION_TYPES[0])
        self.icon_path_var.set(mapping.icon_path)
        self.source_path_var.set(mapping.source_path)
        self.shortcut_var.set(mapping.shortcut_keys)
        self.payload_entry.delete("1.0", tk.END)
        self.payload_entry.insert("1.0", mapping.payload)

        for index, button in enumerate(self.grid_buttons):
            current = self.current_page.buttons[index]
            self._paint_tile(button, current, selected=(index == slot))

    def apply_current_edits(self) -> None:
        try:
            mapping = self.current_page.buttons[self.selected_slot]
            nextion_page_id = int(self.nextion_page_var.get().strip())
            mapping.page_id = int(self.page_id_var.get().strip())
            mapping.component_id = int(self.component_id_var.get().strip())
            mapping.label = self.label_var.get().strip()
            mapping.label_target = self.label_target_var.get().strip()
            mapping.action_type = self.action_type_var.get().strip()
            mapping.payload = self.payload_entry.get("1.0", tk.END).strip()
            mapping.icon_path = self.icon_path_var.get().strip()
            mapping.source_path = self.source_path_var.get().strip()
            mapping.shortcut_keys = self.shortcut_var.get().strip()
            self.current_page.name = self.page_name_var.get().strip() or self.current_page.name
            self.current_page.nextion_page_id = nextion_page_id
            for button in self.current_page.buttons:
                button.page_id = nextion_page_id
            self.page_id_var.set(str(nextion_page_id))
            self._refresh_page_tabs()
            self._render_grid()
            self._load_mapping_into_editor(self.selected_slot)
            self.status_var.set(f"Updated {self.current_page.name} tile {self.selected_slot + 1}")
        except Exception as exc:
            messagebox.showerror("Invalid mapping", str(exc))

    def test_action(self) -> None:
        try:
            self.apply_current_edits()
            mapping = self.current_page.buttons[self.selected_slot]
            result = run_mapping(mapping.action_type, mapping.payload, mapping.shortcut_keys)
            self.status_var.set(result)
        except Exception as exc:
            messagebox.showerror("Action failed", str(exc))

    def import_app(self) -> None:
        chosen = filedialog.askopenfilename(
            title="Import app or shortcut",
            filetypes=[("Apps and shortcuts", "*.exe *.lnk *.url *.bat *.cmd *.ps1"), ("All files", "*.*")],
        )
        if not chosen:
            return
        try:
            metadata = import_app_metadata(chosen)
            processed_icon = ""
            if metadata.icon_path:
                try:
                    processed_icon = self._prepare_custom_art(Path(metadata.icon_path))
                except Exception:
                    processed_icon = metadata.icon_path
            self.label_var.set(metadata.label)
            self.action_type_var.set(metadata.action_type)
            self.source_path_var.set(metadata.source_path)
            self.icon_path_var.set(processed_icon)
            self.shortcut_var.set("")
            self.payload_entry.delete("1.0", tk.END)
            self.payload_entry.insert("1.0", metadata.payload)
            self.apply_current_edits()
            self.status_var.set(f"Imported {Path(chosen).name}")
        except Exception as exc:
            messagebox.showerror("Import failed", str(exc))

    def choose_icon(self) -> None:
        chosen = filedialog.askopenfilename(
            title="Choose icon or photo",
            filetypes=[("Image files", "*.png *.jpg *.jpeg *.bmp *.gif *.ppm *.pgm"), ("All files", "*.*")],
        )
        if not chosen:
            return
        try:
            processed = self._prepare_custom_art(Path(chosen))
            self.icon_path_var.set(processed)
            self.apply_current_edits()
        except Exception as exc:
            messagebox.showerror("Image failed", str(exc))

    def clear_icon(self) -> None:
        self.icon_path_var.set("")
        self.apply_current_edits()

    def use_source_name(self) -> None:
        source = self.source_path_var.get().strip()
        if not source:
            return
        self.label_var.set(Path(source).stem)
        self.apply_current_edits()

    def _command_for_label(self, mapping: ButtonMapping) -> str:
        if not mapping.label_target:
            raise ValueError("Label target is empty.")
        safe_label = mapping.label.replace("\\", "\\\\").replace('"', '\\"')
        return f'{mapping.label_target}.txt="{safe_label}"'

    def sync_selected_label(self) -> None:
        try:
            self.apply_current_edits()
            mapping = self.current_page.buttons[self.selected_slot]
            self.bridge.send_command(self._command_for_label(mapping))
            self.status_var.set(f"Synced label for tile {self.selected_slot + 1}")
        except Exception as exc:
            messagebox.showerror("Sync failed", str(exc))

    def sync_all_labels(self) -> None:
        try:
            for page in self.profile.pages:
                for mapping in page.buttons:
                    if mapping.label_target:
                        self.bridge.send_command(self._command_for_label(mapping))
            self.status_var.set("Queued label sync for all mapped tiles")
        except Exception as exc:
            messagebox.showerror("Sync failed", str(exc))

    def add_page(self) -> None:
        self.apply_current_edits()
        page_number = len(self.profile.pages) + 1
        page = DeckPage(name=f"Page {page_number}", nextion_page_id=len(self.profile.pages), buttons=[])
        self.profile.pages.append(page)
        ensure_page_shape(self.profile)
        self.profile.active_page = len(self.profile.pages) - 1
        self._refresh_page_tabs()
        self._render_grid()
        self._load_mapping_into_editor(0)
        self.status_var.set(f"Added {page.name}")

    def duplicate_page(self) -> None:
        self.apply_current_edits()
        source = self.current_page
        clone_buttons = [ButtonMapping(**mapping.__dict__) for mapping in source.buttons]
        clone = DeckPage(name=f"{source.name} Copy", nextion_page_id=len(self.profile.pages), buttons=clone_buttons)
        for mapping in clone.buttons:
            mapping.page_id = clone.nextion_page_id
            if mapping.label_target.startswith(f"page{source.nextion_page_id}."):
                mapping.label_target = mapping.label_target.replace(f"page{source.nextion_page_id}.", f"page{clone.nextion_page_id}.", 1)
        self.profile.pages.append(clone)
        self.profile.active_page = len(self.profile.pages) - 1
        self._refresh_page_tabs()
        self._render_grid()
        self._load_mapping_into_editor(0)
        self.status_var.set(f"Duplicated {source.name}")

    def delete_page(self) -> None:
        if len(self.profile.pages) == 1:
            messagebox.showinfo("Cannot delete", "At least one page must remain.")
            return
        removed = self.current_page.name
        del self.profile.pages[self.profile.active_page]
        self.profile.active_page = max(0, self.profile.active_page - 1)
        ensure_page_shape(self.profile)
        self._refresh_page_tabs()
        self._render_grid()
        self._load_mapping_into_editor(0)
        self.status_var.set(f"Deleted {removed}")

    def apply_page_settings(self) -> None:
        try:
            if self.layout_var.get() in LAYOUT_PRESETS:
                cols, rows = LAYOUT_PRESETS[self.layout_var.get()]
                self.profile.cols = cols
                self.profile.rows = rows
                ensure_page_shape(self.profile)
                self.selected_slot = min(self.selected_slot, len(self.current_page.buttons) - 1)
                self._sync_window_mode_with_layout()
            self.current_page.name = self.page_name_var.get().strip() or self.current_page.name
            page_id = int(self.nextion_page_var.get().strip())
            self.current_page.nextion_page_id = page_id
            for mapping in self.current_page.buttons:
                mapping.page_id = page_id
                if mapping.label_target.startswith("page"):
                    mapping.label_target = f"page{page_id}.b{mapping.slot}"
            self._refresh_page_tabs()
            self._render_grid()
            self._load_mapping_into_editor(self.selected_slot)
            self.status_var.set(f"Updated {self.current_page.name}")
        except Exception as exc:
            messagebox.showerror("Page settings failed", str(exc))

    def _on_page_selected(self, _event: object) -> None:
        index = self.page_tabs.current()
        if index < 0:
            return
        self.apply_current_edits()
        self.profile.active_page = index
        self._refresh_page_tabs()
        self._render_grid()
        self._load_mapping_into_editor(0)

    def _on_theme_changed(self, _event: object) -> None:
        self.profile.theme_mode = self.theme_var.get().strip().lower()
        self._configure_theme()
        self._build_ui()
        self._refresh_page_tabs()
        self._render_grid()
        self._load_mapping_into_editor(self.selected_slot)

    def open_profile(self) -> None:
        chosen = filedialog.askopenfilename(
            title="Open profile",
            filetypes=[("JSON files", "*.json")],
            initialdir=str(PROFILE_DIR),
        )
        if not chosen:
            return
        self.profile_path = Path(chosen)
        self.profile = load_profile(self.profile_path)
        ensure_page_shape(self.profile)
        self.theme_var.set(self.profile.theme_mode or "dark")
        self._configure_theme()
        self._sync_window_mode_with_layout()
        self.baud_var.set(str(self.profile.baud_rate))
        self._refresh_page_tabs()
        self._render_grid()
        self._load_mapping_into_editor(0)
        self.status_var.set(f"Loaded profile {self.profile_path.name}")

    def save_profile(self) -> None:
        try:
            self.apply_current_edits()
            self.profile.baud_rate = int(self.baud_var.get().strip())
            self.profile.theme_mode = self.theme_var.get().strip().lower()
            save_profile(self.profile, self.profile_path)
            self.status_var.set(f"Saved {self.profile_path.name}")
        except Exception as exc:
            messagebox.showerror("Save failed", str(exc))

    def _prepare_custom_art(self, source: Path) -> str:
        ICON_CACHE_DIR.mkdir(parents=True, exist_ok=True)
        destination = ICON_CACHE_DIR / f"custom_{source.stem}_{IMAGE_SIZE}.png"
        counter = 1
        while destination.exists() and destination.resolve() == source.resolve():
            destination = ICON_CACHE_DIR / f"custom_{source.stem}_{counter}_{IMAGE_SIZE}.png"
            counter += 1
        script = f"""
Add-Type -AssemblyName System.Drawing
$source = '{str(source).replace("'", "''")}'
$destination = '{str(destination).replace("'", "''")}'
$size = {IMAGE_SIZE}
$image = [System.Drawing.Image]::FromFile($source)
$ratio = [Math]::Max($size / $image.Width, $size / $image.Height)
$scaledWidth = [int][Math]::Ceiling($image.Width * $ratio)
$scaledHeight = [int][Math]::Ceiling($image.Height * $ratio)
$x = [int](($size - $scaledWidth) / 2)
$y = [int](($size - $scaledHeight) / 2)
$bitmap = New-Object System.Drawing.Bitmap $size, $size
$graphics = [System.Drawing.Graphics]::FromImage($bitmap)
$graphics.Clear([System.Drawing.Color]::Transparent)
$graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
$graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
$graphics.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
$graphics.DrawImage($image, $x, $y, $scaledWidth, $scaledHeight)
$bitmap.Save($destination, [System.Drawing.Imaging.ImageFormat]::Png)
$graphics.Dispose()
$bitmap.Dispose()
$image.Dispose()
"""
        result = subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script],
            capture_output=True,
            text=True,
            check=False,
        )
        if result.returncode != 0 or not destination.exists():
            raise RuntimeError("Could not process that image into a tile-sized icon.")
        return str(destination)

    def _on_root_resize(self, event: tk.Event) -> None:
        if event.widget is not self.root:
            return
        self.background_canvas.configure(width=event.width, height=event.height)
        self.background_canvas.coords(self.background_item, event.width / 2, event.height / 2)

    def _on_deck_resize(self, _event: tk.Event) -> None:
        if self._render_job:
            self.root.after_cancel(self._render_job)
        self._render_job = self.root.after(70, self._rerender_after_resize)

    def _rerender_after_resize(self) -> None:
        self._render_job = None
        self._render_grid()
        self._load_mapping_into_editor(min(self.selected_slot, len(self.current_page.buttons) - 1))

    def on_close(self) -> None:
        self.bridge.disconnect()
        self.root.destroy()

    def show_about(self) -> None:
        colors = self._theme()
        win = tk.Toplevel(self.root)
        win.title("About NextDeck")
        win.transient(self.root)
        win.configure(bg=colors["panel_bg"])
        win.resizable(False, False)

        body = tk.Frame(win, bg=colors["panel_bg"], padx=22, pady=22)
        body.pack(fill="both", expand=True)
        if self.logo_image:
            tk.Label(body, image=self.logo_image, bg=colors["panel_bg"]).pack(anchor="center", pady=(0, 10))
        tk.Label(body, text=APP_TITLE, bg=colors["panel_bg"], fg=colors["header_fg"], font=("Segoe UI Semibold", 18)).pack(anchor="center")
        tk.Label(body, text=f"Version {APP_VERSION}", bg=colors["panel_bg"], fg=colors["text_muted"], font=("Segoe UI", 10)).pack(anchor="center", pady=(4, 10))
        tk.Label(
            body,
            text="A Nextion-powered desktop deck with pages, branded visuals, media controls, and custom app tiles.",
            bg=colors["panel_bg"],
            fg=colors["text_primary"],
            wraplength=340,
            justify="center",
            font=("Segoe UI", 10),
        ).pack(anchor="center")
        tk.Label(body, text=f"Profiles: {PROFILE_DIR}", bg=colors["panel_bg"], fg=colors["text_muted"], wraplength=360, justify="center").pack(anchor="center", pady=(14, 0))
        ttk.Button(body, text="Close", command=win.destroy).pack(anchor="center", pady=(18, 0))

    def show_settings(self) -> None:
        colors = self._theme()
        win = tk.Toplevel(self.root)
        win.title("Settings")
        win.transient(self.root)
        win.configure(bg=colors["panel_bg"])
        win.resizable(False, False)

        theme_var = tk.StringVar(value=self.theme_var.get())
        baud_var = tk.StringVar(value=self.baud_var.get())
        port_var = tk.StringVar(value=self.port_var.get())

        body = tk.Frame(win, bg=colors["panel_bg"], padx=22, pady=22)
        body.pack(fill="both", expand=True)

        tk.Label(body, text="Theme", bg=colors["panel_bg"], fg=colors["text_muted"], font=("Segoe UI", 10)).grid(row=0, column=0, sticky="w")
        ttk.Combobox(body, textvariable=theme_var, values=("dark", "light"), state="readonly", width=12).grid(row=0, column=1, sticky="ew", padx=(10, 0))
        tk.Label(body, text="COM Port", bg=colors["panel_bg"], fg=colors["text_muted"], font=("Segoe UI", 10)).grid(row=1, column=0, sticky="w", pady=(12, 0))
        self.settings_port_combo = ttk.Combobox(body, textvariable=port_var, state="readonly", width=14)
        self.settings_port_combo.grid(row=1, column=1, sticky="ew", padx=(10, 0), pady=(12, 0))
        ttk.Button(body, text="Refresh Ports", command=self.refresh_ports).grid(row=1, column=2, sticky="ew", padx=(10, 0), pady=(12, 0))
        tk.Label(body, text="Default Baud", bg=colors["panel_bg"], fg=colors["text_muted"], font=("Segoe UI", 10)).grid(row=2, column=0, sticky="w", pady=(12, 0))
        tk.Entry(body, textvariable=baud_var, bg=colors["field_bg"], fg=colors["field_fg"], insertbackground=colors["field_fg"], relief="flat").grid(row=2, column=1, sticky="ew", padx=(10, 0), pady=(12, 0))
        link_row = tk.Frame(body, bg=colors["panel_bg"])
        link_row.grid(row=3, column=0, columnspan=3, sticky="w", pady=(16, 0))
        ttk.Button(link_row, text="Connect", style="Accent.TButton", command=lambda: self._apply_settings_connection(port_var, baud_var, theme_var, win, connect=True)).pack(side="left")
        ttk.Button(link_row, text="Disconnect", command=self.disconnect).pack(side="left", padx=(10, 0))
        tk.Label(body, text=f"Profile file:\n{self.profile_path}", bg=colors["panel_bg"], fg=colors["text_muted"], justify="left", wraplength=360).grid(row=4, column=0, columnspan=3, sticky="w", pady=(16, 0))
        body.grid_columnconfigure(1, weight=1)
        self.refresh_ports()

        def apply_settings() -> None:
            self.theme_var.set(theme_var.get())
            self.baud_var.set(baud_var.get())
            self.port_var.set(port_var.get())
            self.profile.theme_mode = self.theme_var.get().strip().lower()
            self.profile.baud_rate = int(self.baud_var.get().strip())
            self._configure_theme()
            self._build_ui()
            self._refresh_page_tabs()
            self._render_grid()
            self._load_mapping_into_editor(self.selected_slot)
            save_profile(self.profile, self.profile_path)
            self.settings_port_combo = None
            win.destroy()

        controls = tk.Frame(body, bg=colors["panel_bg"])
        controls.grid(row=5, column=0, columnspan=3, sticky="e", pady=(18, 0))
        ttk.Button(controls, text="Cancel", command=win.destroy).pack(side="left")
        ttk.Button(controls, text="Apply", command=apply_settings).pack(side="left", padx=(10, 0))

    def _apply_settings_connection(
        self,
        port_var: tk.StringVar,
        baud_var: tk.StringVar,
        theme_var: tk.StringVar,
        win: tk.Toplevel,
        connect: bool,
    ) -> None:
        self.port_var.set(port_var.get())
        self.baud_var.set(baud_var.get())
        self.theme_var.set(theme_var.get())
        self.profile.theme_mode = self.theme_var.get().strip().lower()
        self.profile.baud_rate = int(self.baud_var.get().strip())
        if connect:
            self.connect()
        save_profile(self.profile, self.profile_path)


def main() -> None:
    root = tk.Tk()
    App(root)
    root.mainloop()
