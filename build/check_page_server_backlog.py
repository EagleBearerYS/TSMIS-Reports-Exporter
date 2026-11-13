"""The window's page server holds a burst of connections it cannot accept yet.

pywebview serves ``ui/`` from a wsgiref server, and the page asks for its
stylesheet and twelve scripts six connections at a time. socketserver's stock
listen queue holds FIVE. Whenever this process was a few milliseconds late
accepting the first of them (the interpreter busy on another thread: a cold
start, a job, an on-access scan), the sixth connection was refused outright —
Chromium's net log named it, ``ERR_CONNECTION_REFUSED`` on the sixth file of the
page, 2 ms after it asked. That one script never ran, the page came up half
built, and nothing told the user why. `gui_api` widens the queue before any
window starts; this holds the behaviour rather than the number:

  * a server started the way pywebview starts it keeps a whole burst waiting
    while its accept loop is held, then serves every one of them;
  * the control — the same server on socketserver's stock queue — refuses part
    of the same burst, so the test is known to see the fault it guards against.

Run from the repo root:  build\\.venv\\Scripts\\python.exe build\\check_page_server_backlog.py
"""
from __future__ import annotations

import socket
import sys
import threading
import time
import wsgiref.simple_server as simple_server
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path[:0] = [str(ROOT / "scripts"), str(ROOT / "build"), str(ROOT)]

STOCK_QUEUE = simple_server.WSGIServer.request_queue_size   # before the app is imported

import gui_api  # noqa: E402  — importing the GUI core is what widens the queue
from webview import http as page_http  # noqa: E402

BURST = 24
_failures: list[str] = []


def check(name, cond, detail=""):
    print(f"  [{'OK ' if cond else 'FAIL'}] {name}")
    if not cond:
        if detail:
            print(f"       {detail}")
        _failures.append(name)


class _HeldAccept:
    """Stops every wsgiref server from ACCEPTING while closed — the state the
    window's process is in when another thread has the interpreter."""

    def __init__(self):
        self._open = threading.Event()
        self._open.set()
        real = simple_server.WSGIServer.get_request
        gate = self._open

        def get_request(server):
            gate.wait(30)
            return real(server)

        simple_server.WSGIServer.get_request = get_request

    def hold(self):
        self._open.clear()

    def release(self):
        self._open.set()

    def remove(self):
        self._open.set()
        del simple_server.WSGIServer.get_request


def _get(port, path="/index.html", sock=None):
    """One GET; returns the status line."""
    own = sock is None
    if own:
        sock = socket.create_connection(("127.0.0.1", port), timeout=10)
    try:
        sock.settimeout(10)
        sock.sendall(f"GET {path} HTTP/1.0\r\nHost: 127.0.0.1\r\n\r\n".encode("ascii"))
        data = b""
        while b"\r\n" not in data:
            chunk = sock.recv(4096)
            if not chunk:
                break
            data += chunk
        return data.split(b"\r\n", 1)[0].decode("ascii", "replace")
    finally:
        sock.close()


def _start_server(index):
    """A page server exactly as `webview.start` makes one; returns its port
    once it answers."""
    _address, _common, server = page_http.BottleServer.start_server([str(index)], None)
    deadline = time.monotonic() + 15
    while True:
        try:
            if "200" in _get(server.port):
                return server.port
        except OSError:  # silent-ok: not listening yet — retried until the deadline
            pass
        if time.monotonic() > deadline:
            raise RuntimeError("the page server never answered")
        time.sleep(0.05)


def _burst(port, count):
    """Open `count` connections at the same moment; returns (sockets, errors)."""
    connected, errors = [], []
    lock = threading.Lock()
    go = threading.Event()

    def one():
        go.wait()
        try:
            sock = socket.create_connection(("127.0.0.1", port), timeout=8)
        except OSError as e:
            with lock:
                errors.append(e)
        else:
            with lock:
                connected.append(sock)

    threads = [threading.Thread(target=one, daemon=True) for _ in range(count)]
    for t in threads:
        t.start()
    go.set()
    for t in threads:
        t.join(30)
    return connected, errors


def _held_burst(gate, port):
    """A burst against a server whose accept is held; then let it serve them.
    Returns (connected, refused, served_200)."""
    gate.hold()
    try:
        connected, errors = _burst(port, BURST)
    finally:
        gate.release()
    served = 0
    for sock in connected:
        try:
            served += "200" in _get(port, sock=sock)
        except OSError:  # silent-ok: counted as not served
            pass
    return len(connected), len(errors), served


def main():
    index = gui_api._ui_index_path()
    check("the UI entry page exists", index.exists(), str(index))
    check("pywebview builds its page server on the class the app widens",
          page_http.WSGIServer is simple_server.WSGIServer)
    widened = simple_server.WSGIServer.request_queue_size
    check("importing the GUI core widens the page server's listen queue",
          widened >= 64 and widened == getattr(gui_api, "PAGE_SERVER_BACKLOG", None),
          f"queue is {widened} (stock {STOCK_QUEUE})")

    gate = _HeldAccept()
    try:
        port = _start_server(index)
        connected, refused, served = _held_burst(gate, port)
        check(f"a burst of {BURST} connections all wait while the accept loop is held",
              connected == BURST and refused == 0,
              f"{connected} connected, {refused} refused")
        check("and every one of them is then served the page",
              served == BURST, f"{served} of {BURST} answered 200")

        # Control: the same server on the stock queue must show the fault, or
        # the two checks above prove nothing on this machine.
        simple_server.WSGIServer.request_queue_size = STOCK_QUEUE
        try:
            stock_port = _start_server(index)
        finally:
            simple_server.WSGIServer.request_queue_size = widened
        connected, refused, _served = _held_burst(gate, stock_port)
        check(f"control: on the stock queue of {STOCK_QUEUE} the same burst is partly refused",
              refused > 0 and connected < BURST,
              f"{connected} connected, {refused} refused")
    finally:
        gate.remove()


if __name__ == "__main__":
    print("the page server keeps a burst of connections waiting:")
    main()
    if _failures:
        print(f"\n{len(_failures)} check(s) FAILED")
        raise SystemExit(1)
    print("all good")
