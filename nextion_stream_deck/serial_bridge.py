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
MAX_WRITE_QUEUE_SIZE = 128
MAX_WRITE_BATCH = 16


class NextionBridge:
    def __init__(self, on_event: EventCallback, on_status: StatusCallback) -> None:
        self._on_event = on_event
        self._on_status = on_status
        self._serial: serial.Serial | None = None
        self._protocol = NextionProtocol()
        self._thread: threading.Thread | None = None
        self._stop_event = threading.Event()
        self._write_queue: queue.Queue[str] = queue.Queue(maxsize=MAX_WRITE_QUEUE_SIZE)

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
        while not self._write_queue.empty():
            try:
                self._write_queue.get_nowait()
            except queue.Empty:
                break
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
        if self._write_queue.full():
            try:
                self._write_queue.get_nowait()
            except queue.Empty:
                pass
        self._write_queue.put_nowait(command)

    def _listen_loop(self) -> None:
        while not self._stop_event.is_set():
            try:
                if self._serial is None:
                    break
                wrote_command = False
                batch_count = 0
                while batch_count < MAX_WRITE_BATCH and not self._write_queue.empty():
                    command = self._write_queue.get_nowait()
                    self._serial.write(encode_command(command))
                    batch_count += 1
                    wrote_command = True
                waiting = getattr(self._serial, "in_waiting", 0)
                if waiting or wrote_command:
                    chunk = self._serial.read(max(64, min(waiting or 64, 256)))
                else:
                    chunk = self._serial.read(64)
                if chunk:
                    for event in self._protocol.feed(chunk):
                        self._on_event(event)
                else:
                    time.sleep(0.05 if self.connected else 0.1)
            except (serial.SerialException, OSError) as exc:
                self._on_status(f"Serial error: {exc}")
                self.disconnect()
                break
