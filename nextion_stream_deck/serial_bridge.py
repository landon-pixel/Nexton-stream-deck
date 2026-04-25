from __future__ import annotations

import queue
import threading
import time
from typing import Callable

import serial
from serial.tools import list_ports

from nextion_stream_deck.protocol import NextionProtocol, NextionTouchEvent, encode_command


EventCallback = Callable[[NextionTouchEvent], None]
StatusCallback = Callable[[str], None]


class NextionBridge:
    def __init__(self, on_event: EventCallback, on_status: StatusCallback) -> None:
        self._on_event = on_event
        self._on_status = on_status
        self._serial: serial.Serial | None = None
        self._protocol = NextionProtocol()
        self._thread: threading.Thread | None = None
        self._stop_event = threading.Event()
        self._write_queue: queue.Queue[str] = queue.Queue()

    @staticmethod
    def available_ports() -> list[str]:
        return [port.device for port in list_ports.comports()]

    @property
    def connected(self) -> bool:
        return bool(self._serial and self._serial.is_open)

    def connect(self, port: str, baud_rate: int) -> None:
        self.disconnect()
        self._serial = serial.Serial(port=port, baudrate=baud_rate, timeout=0.1)
        self._stop_event.clear()
        self._thread = threading.Thread(target=self._listen_loop, daemon=True)
        self._thread.start()
        self._on_status(f"Connected to {port} @ {baud_rate}")

    def disconnect(self) -> None:
        self._stop_event.set()
        if self._thread and self._thread.is_alive():
            self._thread.join(timeout=1.0)
        if self._serial:
            try:
                self._serial.close()
            except serial.SerialException:
                pass
        self._serial = None
        self._thread = None

    def send_command(self, command: str) -> None:
        if not self.connected:
            raise RuntimeError("Not connected to a Nextion display.")
        self._write_queue.put(command)

    def _listen_loop(self) -> None:
        while not self._stop_event.is_set():
            try:
                if self._serial is None:
                    break
                while not self._write_queue.empty():
                    command = self._write_queue.get_nowait()
                    self._serial.write(encode_command(command))
                chunk = self._serial.read(64)
                if chunk:
                    for event in self._protocol.feed(chunk):
                        self._on_event(event)
                else:
                    time.sleep(0.02)
            except (serial.SerialException, OSError) as exc:
                self._on_status(f"Serial error: {exc}")
                self.disconnect()
                break
