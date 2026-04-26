from __future__ import annotations

import queue
from pathlib import Path
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk

from nextion_stream_deck.actions import run_mapping
from nextion_stream_deck.config import DEFAULT_PROFILE_PATH, ICON_CACHE_DIR, ButtonMapping, DeckPage, ensure_page_shape, load_profile, save_profile
from nextion_stream_deck.metadata import import_app_metadata
from nextion_stream_deck.protocol import NextionTouchEvent
from nextion_stream_deck.serial_bridge import NextionBridge


ACTION_TYPES = ("launch", "url", "command", "hotkey")
LAYOUT_PRESETS = {
    "5 x 3": (5, 3),
    "3 x 2": (3, 2),
}
# Tile size constants - 3x2 uses dynamic sizing, 5x3 uses fixed
FIXED_TILE_WIDTH = 180
FIXED_TILE_HEIGHT = 160
IMAGE_SIZE = 88
THEMES = {
    "dark": {
        "card_active": "#1d4ed8",
        "card_idle": "#1e293b",
        "window_bg": "#0b1120",
        "panel_bg": "#e5eefc",
        "text_primary": "#e2e8f0",
        "text_muted": "#94a3b8",
        "header_fg": "#e2e8f0",
        "editor_fg": "#172554",
        "editor_title_fg": "#0f172a",
        "tile_fg": "#ffffff",
        "payload_bg": "#ffffff",
        "payload_fg": "#0f172a",
        "placeholder_inner": "#60a5fa",
    },
    "light": {
        "card_active": "#2563eb",
        "card_idle": "#dbeafe",
        "window_bg": "#f8fafc",
        "panel_bg": "#ffffff",
        "text_primary": "#0f172a",
        "text_muted": "#475569",
        "header_fg": "#0f172a",
        "editor_fg": "#1e3a8a",
        "editor_title_fg": "#0f172a",
        "tile_fg": "#0f172a",
        "payload_bg": "#f8fafc",
        "payload_fg": "#0f172a",
        "placeholder_inner": "#bfdbfe",
    },
}


class App:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Nextion Stream Deck")
        self.default_geometry = "1360x860"
        self.root.geometry(self.default_geometry)
        self.root.minsize(1180, 760)
        self.profile_path = DEFAULT_PROFILE_PATH
        self.profile = load_profile(self.profile_path)
        ensure_page_shape(self.profile)
        self.selected_slot = 0
        self.message_queue: queue.Queue[tuple[str, object]] = queue.Queue()
        self.bridge = NextionBridge(self._queue_event, self._queue_status)
        self.icon_cache: dict[str, tk.PhotoImage] = {}

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
        self.action_type_var = tk.StringVar()
        self.icon_path_var = tk.StringVar()
        self.source_path_var = tk.StringVar()
        self.shortcut_var = tk.StringVar()

        self.grid_buttons: list[tk.Button] = []

        self._configure_theme()
        self._build_ui()
        self._sync_window_mode_with_layout()
        self._refresh_page_tabs()
        self._render_grid()
        self._load_mapping_into_editor(0)
        self.refresh_ports()
        self.root.after(50, self._process_messages)
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    @property
    def current_page(self) -> DeckPage:
        ensure_page_shape(self.profile)
        return self.profile.pages[self.profile.active_page]

    def _theme(self) -> dict[str, str]:
        return THEMES.get(self.theme_var.get().strip().lower(), THEMES["dark"])

    def _configure_theme(self) -> None:
        colors = self._theme()
        self.root.configure(bg=colors["window_bg"])
        style = ttk.Style()
        try:
            style.theme_use("vista")
        except tk.TclError:
            pass
        style.configure("Header.TFrame", background=colors["window_bg"])
        style.configure("Panel.TFrame", background=colors["panel_bg"])
        style.configure("Panel.TLabel", background=colors["panel_bg"], foreground=colors["editor_fg"], font=("Segoe UI", 10))
        style.configure("Header.TLabel", background=colors["window_bg"], foreground=colors["header_fg"], font=("Segoe UI", 10))
        style.configure("Title.TLabel", background=colors["window_bg"], foreground=colors["text_primary"], font=("Segoe UI Semibold", 18))
        style.configure("Subtle.TLabel", background=colors["window_bg"], foreground=colors["text_muted"], font=("Segoe UI", 10))
        style.configure("EditorTitle.TLabel", background=colors["panel_bg"], foreground=colors["editor_title_fg"], font=("Segoe UI Semibold", 16))
        style.configure("Accent.TButton", font=("Segoe UI Semibold", 10))

    def _build_ui(self) -> None:
        self.root.columnconfigure(0, weight=5)
        self.root.columnconfigure(1, weight=3)
        self.root.rowconfigure(1, weight=1)

        header = ttk.Frame(self.root, style="Header.TFrame", padding=(20, 18))
        header.grid(row=0, column=0, columnspan=2, sticky="ew")
        header.columnconfigure(10, weight=1)

        ttk.Label(header, text="Nextion Stream Deck", style="Title.TLabel").grid(row=0, column=0, columnspan=3, sticky="w")
        ttk.Label(
            header,
            text="Import apps, set custom names and art, map pages, and push labels back to the HMI.",
            style="Subtle.TLabel",
        ).grid(row=1, column=0, columnspan=8, sticky="w", pady=(4, 14))

        ttk.Label(header, text="COM Port", style="Header.TLabel").grid(row=2, column=0, sticky="w")
        self.port_combo = ttk.Combobox(header, textvariable=self.port_var, state="readonly", width=12)
        self.port_combo.grid(row=2, column=1, padx=(8, 12))
        ttk.Button(header, text="Refresh", command=self.refresh_ports).grid(row=2, column=2, padx=(0, 16))

        ttk.Label(header, text="Baud", style="Header.TLabel").grid(row=2, column=3, sticky="w")
        ttk.Entry(header, textvariable=self.baud_var, width=10).grid(row=2, column=4, padx=(8, 12))
        ttk.Button(header, text="Connect", style="Accent.TButton", command=self.connect).grid(row=2, column=5, padx=(0, 8))
        ttk.Button(header, text="Disconnect", command=self.disconnect).grid(row=2, column=6, padx=(0, 16))
        ttk.Label(header, text="Theme", style="Header.TLabel").grid(row=2, column=7, sticky="w")
        theme_box = ttk.Combobox(header, textvariable=self.theme_var, values=("dark", "light"), state="readonly", width=10)
        theme_box.grid(row=2, column=8, padx=(8, 12))
        theme_box.bind("<<ComboboxSelected>>", self._on_theme_changed)

        ttk.Label(header, textvariable=self.status_var, style="Header.TLabel").grid(row=2, column=10, sticky="e")
        ttk.Label(header, textvariable=self.last_touch_var, style="Subtle.TLabel").grid(row=3, column=0, columnspan=11, sticky="w", pady=(10, 0))

        left = ttk.Frame(self.root, style="Header.TFrame", padding=(20, 0, 10, 20))
        left.grid(row=1, column=0, sticky="nsew")
        left.columnconfigure(0, weight=1)
        left.rowconfigure(2, weight=1)

        page_bar = ttk.Frame(left, style="Header.TFrame")
        page_bar.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        page_bar.columnconfigure(0, weight=1)
        self.page_tabs = ttk.Combobox(page_bar, textvariable=self.page_var, state="readonly")
        self.page_tabs.grid(row=0, column=0, sticky="ew", padx=(0, 10))
        self.page_tabs.bind("<<ComboboxSelected>>", self._on_page_selected)
        self.mode_label = ttk.Label(page_bar, text="", style="Subtle.TLabel")
        self.mode_label.grid(row=1, column=0, columnspan=4, sticky="ew", pady=(8, 0))
        ttk.Button(page_bar, text="Add Page", command=self.add_page).grid(row=0, column=1, padx=(0, 8))
        ttk.Button(page_bar, text="Rename", command=self.rename_page).grid(row=0, column=2, padx=(0, 8))
        ttk.Button(page_bar, text="Delete", command=self.delete_page).grid(row=0, column=3)

        page_meta = ttk.Frame(left, style="Header.TFrame")
        page_meta.grid(row=1, column=0, sticky="ew", pady=(0, 12))
        ttk.Label(page_meta, text="Page Name", style="Header.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Entry(page_meta, textvariable=self.page_name_var, width=28).grid(row=0, column=1, padx=(8, 14))
        ttk.Label(page_meta, text="Nextion Page ID", style="Header.TLabel").grid(row=0, column=2, sticky="w")
        ttk.Entry(page_meta, textvariable=self.nextion_page_var, width=8).grid(row=0, column=3, padx=(8, 10))
        ttk.Label(page_meta, text="Layout", style="Header.TLabel").grid(row=0, column=4, sticky="w")
        layout_box = ttk.Combobox(page_meta, textvariable=self.layout_var, values=tuple(LAYOUT_PRESETS.keys()), state="readonly", width=8)
        layout_box.grid(row=0, column=5, padx=(8, 10))
        ttk.Button(page_meta, text="Apply Page", command=self.apply_page_settings).grid(row=0, column=6)

        self.grid_frame = tk.Frame(left, bg=self._theme()["window_bg"], highlightthickness=0)
        self.grid_frame.grid(row=2, column=0, sticky="nsew")

        editor = ttk.Frame(self.root, style="Panel.TFrame", padding=(18, 18))
        editor.grid(row=1, column=1, sticky="nsew", padx=(10, 20), pady=(0, 20))
        editor.columnconfigure(1, weight=1)
        editor.rowconfigure(9, weight=1)

        ttk.Label(editor, textvariable=self.slot_title_var, style="EditorTitle.TLabel").grid(
            row=0, column=0, columnspan=3, sticky="w", pady=(0, 14)
        )

        self._field(editor, "Page ID", self.page_id_var, 1)
        self._field(editor, "Component ID", self.component_id_var, 2)
        self._field(editor, "Custom Name", self.label_var, 3)
        self._field(editor, "Label Target", self.label_target_var, 4)

        ttk.Label(editor, text="Action Type", style="Panel.TLabel").grid(row=5, column=0, sticky="w", pady=6)
        action_box = ttk.Combobox(editor, textvariable=self.action_type_var, values=ACTION_TYPES, state="readonly")
        action_box.grid(row=5, column=1, columnspan=2, sticky="ew", pady=6)

        ttk.Label(editor, text="Payload", style="Panel.TLabel").grid(row=6, column=0, sticky="nw", pady=6)
        self.payload_entry = tk.Text(
            editor,
            height=5,
            width=36,
            bg=self._theme()["payload_bg"],
            fg=self._theme()["payload_fg"],
            relief="flat",
            insertbackground=self._theme()["payload_fg"],
            font=("Consolas", 10),
        )
        self.payload_entry.grid(row=6, column=1, columnspan=2, sticky="nsew", pady=6)

        self._field(editor, "Source Path", self.source_path_var, 7)
        self._field(editor, "Shortcut Keys", self.shortcut_var, 8)
        self._field(editor, "Custom Photo/Icon", self.icon_path_var, 9)

        buttons = ttk.Frame(editor, style="Panel.TFrame")
        buttons.grid(row=10, column=0, columnspan=3, sticky="ew", pady=(14, 8))
        for column in range(3):
            buttons.columnconfigure(column, weight=1)
        ttk.Button(buttons, text="Apply", command=self.apply_current_edits).grid(row=0, column=0, sticky="ew", padx=(0, 6))
        ttk.Button(buttons, text="Test Action", command=self.test_action).grid(row=0, column=1, sticky="ew", padx=6)
        ttk.Button(buttons, text="Import App", command=self.import_app).grid(row=0, column=2, sticky="ew", padx=(6, 0))

        buttons2 = ttk.Frame(editor, style="Panel.TFrame")
        buttons2.grid(row=11, column=0, columnspan=3, sticky="ew", pady=(0, 8))
        for column in range(3):
            buttons2.columnconfigure(column, weight=1)
        ttk.Button(buttons2, text="Choose Photo/Icon", command=self.choose_icon).grid(row=0, column=0, sticky="ew", padx=(0, 6))
        ttk.Button(buttons2, text="Clear Art", command=self.clear_icon).grid(row=0, column=1, sticky="ew", padx=6)
        ttk.Button(buttons2, text="Sync Label", command=self.sync_selected_label).grid(row=0, column=2, sticky="ew", padx=(6, 0))

        buttons3 = ttk.Frame(editor, style="Panel.TFrame")
        buttons3.grid(row=12, column=0, columnspan=3, sticky="ew", pady=(0, 8))
        for column in range(3):
            buttons3.columnconfigure(column, weight=1)
        ttk.Button(buttons3, text="Open Profile", command=self.open_profile).grid(row=0, column=0, sticky="ew", padx=(0, 6))
        ttk.Button(buttons3, text="Save Profile", command=self.save_profile).grid(row=0, column=1, sticky="ew", padx=6)
        ttk.Button(buttons3, text="Duplicate Page", command=self.duplicate_page).grid(row=0, column=2, sticky="ew", padx=(6, 0))

        buttons4 = ttk.Frame(editor, style="Panel.TFrame")
        buttons4.grid(row=13, column=0, columnspan=3, sticky="ew", pady=(0, 8))
        for column in range(2):
            buttons4.columnconfigure(column, weight=1)
        ttk.Button(buttons4, text="Sync All Labels", command=self.sync_all_labels).grid(row=0, column=0, sticky="ew", padx=(0, 6))
        ttk.Button(buttons4, text="Use Current Name", command=self.use_source_name).grid(row=0, column=1, sticky="ew", padx=(6, 0))

        help_text = (
            "Import an app, then tweak the custom name, shortcut keys, and photo/icon for that tile. "
            "If Shortcut Keys is filled for a launch action, the app launches first and the shortcut fires after."
        )
        ttk.Label(editor, text=help_text, style="Panel.TLabel", wraplength=360, justify="left").grid(
            row=14, column=0, columnspan=3, sticky="w", pady=(10, 0)
        )

    @staticmethod
    def _field(parent: ttk.Frame, label: str, variable: tk.StringVar, row: int) -> None:
        ttk.Label(parent, text=label, style="Panel.TLabel").grid(row=row, column=0, sticky="w", pady=6)
        ttk.Entry(parent, textvariable=variable).grid(row=row, column=1, columnspan=2, sticky="ew", pady=6)

    def _refresh_page_tabs(self) -> None:
        ensure_page_shape(self.profile)
        labels = [f"{index + 1}. {page.name}" for index, page in enumerate(self.profile.pages)]
        self.page_tabs["values"] = labels
        self.page_var.set(labels[self.profile.active_page])
        self.page_name_var.set(self.current_page.name)
        self.nextion_page_var.set(str(self.current_page.nextion_page_id))

    def _render_grid(self) -> None:
        for child in self.grid_frame.winfo_children():
            child.destroy()
        self.grid_buttons = []
        self.grid_frame.configure(bg=self._theme()["window_bg"])
        
        # Determine if we're in 3x2 mode - use same fixed sizing as other layouts
        is_three_by_two = self.profile.cols == 3 and self.profile.rows == 2
        self.mode_label.config(text=f"Debug: cols={self.profile.cols} rows={self.profile.rows} is_3x2={is_three_by_two}")
        
        # Set fixed grid frame size regardless of layout
        grid_width = self.profile.cols * (FIXED_TILE_WIDTH + 16) + 16
        grid_height = self.profile.rows * (FIXED_TILE_HEIGHT + 16) + 16
        self.grid_frame.configure(width=grid_width, height=grid_height)
        
        # All layouts use fixed row/column sizing (same as original 5x3)
        for row in range(self.profile.rows):
            self.grid_frame.rowconfigure(row, weight=0, minsize=FIXED_TILE_HEIGHT + 16)
        for col in range(self.profile.cols):
            self.grid_frame.columnconfigure(col, weight=0, minsize=FIXED_TILE_WIDTH + 16)
        
        for mapping in self.current_page.buttons:
            image = self._icon_for_mapping(mapping)
            row = mapping.slot // self.profile.cols
            col = mapping.slot % self.profile.cols
            
            tile = tk.Frame(
                self.grid_frame,
                bg=self._theme()["card_idle"],
                highlightthickness=2,
                highlightbackground=self._theme()["window_bg"],
                bd=0,
            )
            
            button = tk.Button(
                tile,
                text=self._button_caption(mapping),
                image=image,
                compound="top",
                anchor="center",
                justify="center",
                command=lambda slot=mapping.slot: self._load_mapping_into_editor(slot),
                bg=self._theme()["card_idle"],
                fg=self._theme()["tile_fg"],
                activebackground=self._theme()["card_active"],
                activeforeground=self._theme()["tile_fg"],
                relief="flat",
                bd=0,
                highlightthickness=0,
                font=("Segoe UI Semibold", 10),
            )
            button.image = image
            
            # All layouts use fixed size with place
            tile.grid(row=row, column=col, sticky="nw", padx=8, pady=8)
            tile.configure(width=FIXED_TILE_WIDTH, height=FIXED_TILE_HEIGHT)
            tile.grid_propagate(False)
            button.place(x=0, y=0, width=FIXED_TILE_WIDTH, height=FIXED_TILE_HEIGHT)
            
            self.grid_buttons.append(button)

    @staticmethod
    def _button_caption(mapping: ButtonMapping) -> str:
        footer = f"P{mapping.page_id} C{mapping.component_id}"
        return f"{mapping.label or f'Slot {mapping.slot + 1}'}\n{footer}"

    def _icon_for_mapping(self, mapping: ButtonMapping) -> tk.PhotoImage:
        if mapping.icon_path:
            path = Path(mapping.icon_path)
            if path.exists():
                key = f"{path.resolve()}:{self.theme_var.get()}"
                if key not in self.icon_cache:
                    try:
                        self.icon_cache[key] = tk.PhotoImage(file=str(path))
                    except tk.TclError:
                        self.icon_cache[key] = self._placeholder_icon(mapping.label)
                return self.icon_cache[key]
        return self._placeholder_icon(mapping.label)

    def _placeholder_icon(self, label: str) -> tk.PhotoImage:
        key = f"placeholder:{(label[:1] or '?').upper()}:{self.theme_var.get()}"
        if key in self.icon_cache:
            return self.icon_cache[key]
        image = tk.PhotoImage(width=IMAGE_SIZE, height=IMAGE_SIZE)
        image.put(self._theme()["card_active"], to=(0, 0, IMAGE_SIZE - 1, IMAGE_SIZE - 1))
        image.put(self._theme()["placeholder_inner"], to=(8, 8, IMAGE_SIZE - 9, IMAGE_SIZE - 9))
        self.icon_cache[key] = image
        return image

    def _layout_label(self) -> str:
        for label, (cols, rows) in LAYOUT_PRESETS.items():
            if self.profile.cols == cols and self.profile.rows == rows:
                return label
        return f"{self.profile.cols} x {self.profile.rows}"

    def _sync_window_mode_with_layout(self) -> None:
        # Don't auto-zoom window - just ensure tiles expand to fill available space
        pass

    def refresh_ports(self) -> None:
        ports = self.bridge.available_ports()
        self.port_combo["values"] = ports
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
        self.slot_title_var.set(f"{self.current_page.name} · Button {slot + 1}")
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
            button.configure(bg=self._theme()["card_active"] if index == slot else self._theme()["card_idle"])

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
            self.status_var.set(f"Updated {self.current_page.name} button {self.selected_slot + 1}")
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
            filetypes=[
                ("Apps and shortcuts", "*.exe *.lnk *.url *.bat *.cmd *.ps1"),
                ("All files", "*.*"),
            ],
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
        except Exception as exc:
            messagebox.showerror("Image failed", str(exc))
            return
        self.icon_path_var.set(processed)
        self.apply_current_edits()

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
            raise ValueError("Label Target is empty.")
        safe_label = mapping.label.replace("\\", "\\\\").replace('"', '\\"')
        return f'{mapping.label_target}.txt="{safe_label}"'

    def _on_theme_changed(self, _event: object) -> None:
        self.profile.theme_mode = self.theme_var.get().strip().lower()
        self._configure_theme()
        self.payload_entry.configure(
            bg=self._theme()["payload_bg"],
            fg=self._theme()["payload_fg"],
            insertbackground=self._theme()["payload_fg"],
        )
        self._render_grid()
        self._load_mapping_into_editor(self.selected_slot)

    def sync_selected_label(self) -> None:
        try:
            self.apply_current_edits()
            mapping = self.current_page.buttons[self.selected_slot]
            self.bridge.send_command(self._command_for_label(mapping))
            self.status_var.set(f"Synced label for button {self.selected_slot + 1}")
        except Exception as exc:
            messagebox.showerror("Sync failed", str(exc))

    def sync_all_labels(self) -> None:
        try:
            for page in self.profile.pages:
                for mapping in page.buttons:
                    if mapping.label_target:
                        self.bridge.send_command(self._command_for_label(mapping))
            self.status_var.set("Queued label sync for mapped buttons")
        except Exception as exc:
            messagebox.showerror("Sync failed", str(exc))

    def add_page(self) -> None:
        self.apply_current_edits()
        page_number = len(self.profile.pages) + 1
        page = DeckPage(
            name=f"Page {page_number}",
            nextion_page_id=len(self.profile.pages),
            buttons=[],
        )
        self.profile.pages.append(page)
        ensure_page_shape(self.profile)
        self.profile.active_page = len(self.profile.pages) - 1
        self._refresh_page_tabs()
        self._render_grid()
        self._load_mapping_into_editor(0)
        self.status_var.set(f"Added {page.name}")

    def rename_page(self) -> None:
        new_name = simpledialog.askstring("Rename page", "Page name:", initialvalue=self.current_page.name, parent=self.root)
        if not new_name:
            return
        self.current_page.name = new_name.strip()
        self.page_name_var.set(self.current_page.name)
        self._refresh_page_tabs()
        self.status_var.set(f"Renamed page to {self.current_page.name}")

    def duplicate_page(self) -> None:
        self.apply_current_edits()
        source = self.current_page
        clone_buttons = [ButtonMapping(**mapping.__dict__) for mapping in source.buttons]
        clone = DeckPage(
            name=f"{source.name} Copy",
            nextion_page_id=len(self.profile.pages),
            buttons=clone_buttons,
        )
        for mapping in clone.buttons:
            mapping.page_id = clone.nextion_page_id
            if mapping.label_target.startswith(f"page{source.nextion_page_id}."):
                mapping.label_target = mapping.label_target.replace(
                    f"page{source.nextion_page_id}.",
                    f"page{clone.nextion_page_id}.",
                    1,
                )
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
        name = self.current_page.name
        del self.profile.pages[self.profile.active_page]
        self.profile.active_page = max(0, self.profile.active_page - 1)
        ensure_page_shape(self.profile)
        self._refresh_page_tabs()
        self._render_grid()
        self._load_mapping_into_editor(0)
        self.status_var.set(f"Deleted {name}")

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

    def open_profile(self) -> None:
        chosen = filedialog.askopenfilename(
            title="Open profile",
            filetypes=[("JSON files", "*.json")],
            initialdir=str(Path("profiles").resolve()),
        )
        if not chosen:
            return
        self.profile_path = Path(chosen)
        self.profile = load_profile(self.profile_path)
        ensure_page_shape(self.profile)
        self.theme_var.set(self.profile.theme_mode or "dark")
        self.layout_var.set(self._layout_label())
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

    def on_close(self) -> None:
        self.bridge.disconnect()
        self.root.destroy()


def main() -> None:
    root = tk.Tk()
    App(root)
    root.mainloop()
