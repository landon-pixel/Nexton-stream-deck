from __future__ import annotations

from dataclasses import dataclass


END_MARKER = b"\xff\xff\xff"
TOUCH_EVENT = 0x65


@dataclass
class NextionTouchEvent:
    page_id: int
    component_id: int
    event: int

    @property
    def pressed(self) -> bool:
        return self.event == 1


class NextionProtocol:
    def __init__(self) -> None:
        self._buffer = bytearray()

    def feed(self, chunk: bytes) -> list[NextionTouchEvent]:
        self._buffer.extend(chunk)
        events: list[NextionTouchEvent] = []
        while True:
            marker_index = self._buffer.find(END_MARKER)
            if marker_index == -1:
                break
            packet = bytes(self._buffer[:marker_index])
            del self._buffer[: marker_index + len(END_MARKER)]
            event = self._decode(packet)
            if event:
                events.append(event)
        return events

    @staticmethod
    def _decode(packet: bytes) -> NextionTouchEvent | None:
        if len(packet) != 4:
            return None
        if packet[0] != TOUCH_EVENT:
            return None
        return NextionTouchEvent(
            page_id=packet[1],
            component_id=packet[2],
            event=packet[3],
        )


def encode_command(command: str) -> bytes:
    return command.encode("ascii", errors="ignore") + END_MARKER
