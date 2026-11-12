from __future__ import annotations

import os
import threading
import uuid

from convertor.gui import instance


def address() -> str:
    return rf"\\.\pipe\DocumentConverter-test-{os.getpid()}-{uuid.uuid4().hex}"


def test_second_instance_hands_files_to_the_first() -> None:
    pipe = address()
    listener = instance.claim({"files": []}, address=pipe)
    assert listener is not None  # prima instanta asculta
    received = []
    got = threading.Event()
    server = instance.InstanceServer(listener, lambda m: (received.append(m), got.set()))
    server.start()
    try:
        assert instance.claim({"files": [r"C:\a.docx"], "auto": True}, address=pipe) is None
        assert got.wait(5)
        assert received == [{"files": [r"C:\a.docx"], "auto": True}]
    finally:
        server.stop()
    assert not server.is_alive()
    # dupa inchidere, o instanta noua devine ea cea principala
    again = instance.claim({"files": []}, address=pipe, attempts=3)
    assert again is not None
    again.close()


def test_send_without_running_instance() -> None:
    assert not instance.send_to_running({"files": []}, address())
