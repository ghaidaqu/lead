from __future__ import annotations

import base64
import os
from http import HTTPStatus
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path


ROOT = Path("/Users/ghaida/lead6_host")
USERNAME = os.environ.get("LEAD6_AUTH_USER", "lead6")
PASSWORD = os.environ.get("LEAD6_AUTH_PASS", "")


class AuthHandler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(ROOT), **kwargs)

    def _authorized(self) -> bool:
        header = self.headers.get("Authorization", "")
        if not header.startswith("Basic "):
            return False
        try:
            decoded = base64.b64decode(header.split(" ", 1)[1]).decode("utf-8")
        except Exception:
            return False
        return decoded == f"{USERNAME}:{PASSWORD}"

    def _challenge(self):
        self.send_response(HTTPStatus.UNAUTHORIZED)
        self.send_header("WWW-Authenticate", 'Basic realm="lead6"')
        self.send_header("Content-Type", "text/plain; charset=utf-8")
        self.end_headers()
        self.wfile.write("Authentication required".encode("utf-8"))

    def do_GET(self):
        if not self._authorized():
            return self._challenge()
        return super().do_GET()

    def do_HEAD(self):
        if not self._authorized():
            return self._challenge()
        return super().do_HEAD()


def main():
    server = ThreadingHTTPServer(("0.0.0.0", 8001), AuthHandler)
    print(f"Serving {ROOT} on http://localhost:8001")
    print(f"Username: {USERNAME}")
    print("Password: set via LEAD6_AUTH_PASS")
    server.serve_forever()


if __name__ == "__main__":
    main()
