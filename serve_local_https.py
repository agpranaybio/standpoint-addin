#!/usr/bin/env python
"""Serve the task pane and a /health route over TLS, to price Option B.

    python addin/serve_local_https.py

Mac Word refused http://localhost:8000 in 7 milliseconds. That request was
three things at once -- plaintext, aimed at the local machine, and
cross-origin -- and only the failure was measured, not which of the three
caused it. The fixes differ:

    plaintext was the blocker   -> serve the API over TLS, keep the pane on
                                   GitHub Pages, and a pane fix stays a git push
    public->local was blocked   -> the app must serve the pane itself, and every
                                   pane fix needs a new installer

This script provides the TLS end of both experiments. Run it, then open the
pane from BOTH origins: the hosted copy tests the first case, the copy served
here tests the second.

IT IS A SPIKE, NOT THE SHIPPING CODE. The real version is a second uvicorn
listener over the actual app with a StaticFiles mount (uvicorn already takes
ssl_keyfile/ssl_certfile), and generates its certificate with the
`cryptography` package rather than the openssl binary, because Windows ships no
openssl. Doing that now would be building on the assumption this script exists
to test.

NOTHING IS LEFT BEHIND. The certificate goes to a temp directory, and the
command to UNTRUST it is printed next to the command to trust it -- a spike has
no business leaving a trusted root on a machine.
"""

from __future__ import annotations

import http.server
import json
import os
import ssl
import subprocess
import sys
import tempfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
HOST, PORT = "127.0.0.1", 8443
CERT_DIR = Path(tempfile.gettempdir()) / "standpoint-addin-spike"
CERT, KEY = CERT_DIR / "cert.pem", CERT_DIR / "key.pem"


def make_cert() -> None:
    """A self-signed certificate covering localhost AND 127.0.0.1.

    BOTH names, in subjectAltName, because a certificate is checked against the
    name the client actually asked for. One covering only `localhost` is
    rejected for `https://127.0.0.1` and the pane goes blank with no error --
    the failure mode this whole spike is trying not to be confused by.

    Also uses the SAN extension rather than the legacy common name, which
    modern clients ignore outright.
    """
    CERT_DIR.mkdir(parents=True, exist_ok=True)
    if CERT.exists() and KEY.exists():
        print(f"reusing the certificate in {CERT_DIR}")
        return
    print(f"generating a self-signed certificate in {CERT_DIR}")
    subprocess.run(
        ["openssl", "req", "-x509", "-newkey", "rsa:2048", "-nodes",
         "-keyout", str(KEY), "-out", str(CERT), "-days", "30",
         "-subj", "/CN=localhost/O=Standpoint.Bio spike",
         "-addext", "subjectAltName=DNS:localhost,IP:127.0.0.1",
         "-addext", "basicConstraints=critical,CA:TRUE",
         "-addext", "keyUsage=critical,digitalSignature,keyCertSign"],
        check=True, capture_output=True,
    )


class Handler(http.server.SimpleHTTPRequestHandler):
    """Static files from addin/, plus /health, plus a log of everything.

    THE LOG IS THE INSTRUMENT. A failed fetch tells JavaScript nothing, so the
    only way to place a failure is to look from this side:

        nothing logged      the browser refused to send it
        OPTIONS logged      a preflight, which the response must satisfy
        GET logged          we got through; any failure is on the response

    That distinction is what turned the Chrome result from a guess into a
    proof, so it is kept here.
    """

    def __init__(self, *a, **kw):
        super().__init__(*a, directory=str(HERE), **kw)

    def _note(self) -> None:
        origin = self.headers.get("Origin") or "(none)"
        pna = self.headers.get("Access-Control-Request-Private-Network")
        extra = "  PNA-preflight" if pna else ""
        print(f"  <- {self.command:7} {self.path:40} Origin: {origin}{extra}", flush=True)

    def _cors(self) -> None:
        # PERMISSIVE ON PURPOSE. The hosted pane calls this cross-origin, and a
        # missing header here would fail that probe for a reason that has
        # nothing to do with the question. We are measuring the browser's gate,
        # not this server's configuration. The shipping version is strict.
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "*")
        self.send_header("Access-Control-Allow-Private-Network", "true")

    def end_headers(self) -> None:
        self._cors()
        self.send_header("Cache-Control", "no-store")
        super().end_headers()

    def do_OPTIONS(self) -> None:
        self._note()
        self.send_response(204)
        self.end_headers()

    def do_GET(self) -> None:
        self._note()
        if self.path.split("?")[0] == "/health":
            # The same shape the real backend returns, so the pane needs no
            # special case and the probe exercises the real code path.
            body = json.dumps({"status": "ok", "warnings": [], "via": "spike"}).encode()
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
            return
        super().do_GET()

    def log_message(self, *a) -> None:      # the default stderr line, replaced by _note
        pass


class Server(http.server.HTTPServer):
    """Reports a REFUSED CERTIFICATE, which otherwise looks like silence.

    The log below is the instrument: "nothing arrived" is how a browser-level
    block is identified. But a client that rejects the certificate also leaves
    nothing, because the TLS handshake fails before there is any HTTP request
    to log -- and socketserver swallows that error entirely (SSLError subclasses
    OSError, and _handle_request_noblock catches OSError and simply returns).

    Two very different findings would therefore look identical:

        the engine refuses to reach a local machine  -> change the architecture
        the machine does not trust this certificate  -> trust it and re-run

    So the handshake failure gets a line of its own.
    """

    def get_request(self):
        try:
            return super().get_request()
        except ssl.SSLError as exc:
            print(f"  !! TLS HANDSHAKE REFUSED by the client "
                  f"({getattr(exc, 'reason', None) or exc}).\n"
                  f"     The certificate is not trusted on this machine yet — "
                  f"this is NOT a policy block. Run the trust command above.",
                  flush=True)
            raise


def trust_commands() -> tuple[str, str]:
    """(trust, untrust) for this platform. PRINTED, never run.

    Trusting a root certificate is the single most consequential thing in this
    spike, so it is the operator who does it, having read the command.
    """
    if sys.platform == "darwin":
        return (
            f'sudo security add-trusted-cert -d -r trustRoot '
            f'-k /Library/Keychains/System.keychain "{CERT}"',
            f'sudo security remove-trusted-cert -d "{CERT}"',
        )
    if sys.platform == "win32":
        return (
            f'certutil -addstore -user Root "{CERT}"',
            f'certutil -delstore -user Root "Standpoint.Bio spike"',
        )
    return (
        f'sudo cp "{CERT}" /usr/local/share/ca-certificates/standpoint-spike.crt '
        f'&& sudo update-ca-certificates',
        'sudo rm /usr/local/share/ca-certificates/standpoint-spike.crt '
        '&& sudo update-ca-certificates --fresh',
    )


def main() -> int:
    if not (HERE / "taskpane.html").is_file():
        print(f"no taskpane.html in {HERE}", file=sys.stderr)
        return 1
    try:
        make_cert()
    except FileNotFoundError:
        print("openssl is not on PATH. macOS and Linux ship it; on Windows use "
              "the one in Git Bash.", file=sys.stderr)
        return 1
    except subprocess.CalledProcessError as exc:
        print(f"openssl failed:\n{exc.stderr.decode(errors='replace')}", file=sys.stderr)
        return 1

    trust, untrust = trust_commands()
    print()
    print("=" * 72)
    print(f"  TRUST IT FIRST, or the pane will be BLANK with no error anywhere:")
    print(f"    {trust}")
    print()
    print(f"  AFTERWARDS, remove it again:")
    print(f"    {untrust}")
    print("=" * 72)
    print()
    print(f"  pane   https://{HOST}:{PORT}/taskpane.html")
    print(f"  health https://{HOST}:{PORT}/health")
    print()
    print("  Every request is logged below, and so is a refused certificate.")
    print("  If the pane shows a failure and NOTHING appears here, the browser")
    print("  never sent it — that is a policy block, not a certificate problem.")
    print()

    ctx = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
    ctx.load_cert_chain(str(CERT), str(KEY))
    srv = Server((HOST, PORT), Handler)
    srv.socket = ctx.wrap_socket(srv.socket, server_side=True)
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        print("\nstopped. Remember to untrust the certificate:")
        print(f"  {untrust}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
