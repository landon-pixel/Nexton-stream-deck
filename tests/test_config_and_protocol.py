from pathlib import Path
import tempfile
import unittest

from nextion_stream_deck.config import create_default_profile, load_profile, save_profile
from nextion_stream_deck.metadata import _metadata_from_url
from nextion_stream_deck.protocol import NextionProtocol, encode_command


class ConfigTests(unittest.TestCase):
    def test_default_profile_has_one_page_with_expected_button_count(self) -> None:
        profile = create_default_profile(rows=2, cols=4)
        self.assertEqual(len(profile.pages), 1)
        self.assertEqual(len(profile.pages[0].buttons), 8)
        self.assertEqual(profile.pages[0].buttons[0].component_id, 1)
        self.assertEqual(profile.pages[0].buttons[-1].component_id, 8)

    def test_profile_round_trip(self) -> None:
        profile = create_default_profile()
        profile.pages.append(profile.pages[0])
        profile.theme_mode = "light"
        profile.pages[0].buttons[0].shortcut_keys = "ctrl+alt+1"
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "profile.json"
            save_profile(profile, path)
            loaded = load_profile(path)
        self.assertEqual(loaded.rows, profile.rows)
        self.assertEqual(loaded.pages[0].buttons[3].label, profile.pages[0].buttons[3].label)
        self.assertEqual(loaded.theme_mode, "light")
        self.assertEqual(loaded.pages[0].buttons[0].shortcut_keys, "ctrl+alt+1")

    def test_legacy_profile_migrates_to_pages(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "legacy.json"
            path.write_text(
                """
{
  "name": "Legacy",
  "baud_rate": 115200,
  "rows": 2,
  "cols": 2,
  "buttons": [
    { "slot": 0, "page_id": 0, "component_id": 1, "label": "A", "label_target": "", "action_type": "launch", "payload": "" }
  ]
}
""".strip(),
                encoding="utf-8",
            )
            loaded = load_profile(path)
        self.assertEqual(len(loaded.pages), 1)
        self.assertEqual(len(loaded.pages[0].buttons), 4)
        self.assertEqual(loaded.pages[0].buttons[0].label, "A")


class MetadataTests(unittest.TestCase):
    def test_url_metadata_import(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "Docs.url"
            path.write_text("[InternetShortcut]\nURL=https://example.com\n", encoding="utf-8")
            metadata = _metadata_from_url(path)
        self.assertEqual(metadata.action_type, "url")
        self.assertEqual(metadata.payload, "https://example.com")
        self.assertEqual(metadata.label, "Docs")


class ProtocolTests(unittest.TestCase):
    def test_touch_event_decodes(self) -> None:
        protocol = NextionProtocol()
        events = protocol.feed(bytes([0x65, 1, 7, 1, 0xFF, 0xFF, 0xFF]))
        self.assertEqual(len(events), 1)
        self.assertEqual(events[0].page_id, 1)
        self.assertEqual(events[0].component_id, 7)
        self.assertTrue(events[0].pressed)

    def test_command_encoding_appends_end_marker(self) -> None:
        encoded = encode_command('page0.b0.txt="OBS"')
        self.assertTrue(encoded.endswith(b"\xff\xff\xff"))


if __name__ == "__main__":
    unittest.main()
