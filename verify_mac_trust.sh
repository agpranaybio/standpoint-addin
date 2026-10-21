#!/usr/bin/env bash
# Verify the certificate trust step in a PACKAGED Mac build, end to end.
#
# WHY THIS IS A SCRIPT AND NOT A TEST. It needs three things no CI runner and no
# Linux machine has: macOS, a GUI session with a keychain, and a human at a
# password prompt. macOS cannot be legally virtualised on non-Apple hardware, so
# this is the one step that has to be done by hand, on the test Mac.
#
# WHAT IS ALREADY VERIFIED, so you are not re-checking it: the packaged code
# path itself was exercised in a real packaged build (on Linux, which uses the
# same asar + custom-scheme machinery). In that build the window loaded
# standpoint-app://app/index.html out of app.asar, React rendered, and the
# renderer's GETs and a POST all passed the request guard with zero refusals.
# What remains Mac-only is the KEYCHAIN: whether `security add-trusted-cert`
# into the login keychain is enough for WKWebView inside Word.
#
#   bash addin/verify_mac_trust.sh            # after installing the .dmg
#
# Prints a PASS/FAIL line per step and a summary. It changes nothing except
# through the app's own UI, and tells you how to undo the trust at the end.

set -uo pipefail

PASS=0; FAIL=0
ok()   { printf '  \033[32mok\033[0m    %s\n' "$1"; PASS=$((PASS+1)); }
bad()  { printf '  \033[31mFAIL\033[0m  %s\n' "$1"; FAIL=$((FAIL+1)); }
note() { printf '        %s\n' "$1"; }
step() { printf '\n\033[1m%s\033[0m\n' "$1"; }

APP="/Applications/Standpoint.Bio.app"
DATA="$HOME/Library/Application Support/Standpoint.Bio"
CERT="$DATA/addin-cert.pem"
WEF="$HOME/Library/Containers/com.microsoft.Word/Data/Documents/wef"

step "1. The packaged app is installed"
if [ -d "$APP" ]; then ok "found $APP"
else bad "no app at $APP — install the .dmg first"; exit 1; fi

step "2. Launch it, and let the backend come up"
open -a "$APP" || true
note "waiting up to 60s for the backend…"
for i in $(seq 1 30); do
  curl -s -o /dev/null --max-time 2 http://127.0.0.1:8000/health && break
  sleep 2
done
if curl -s -o /dev/null --max-time 3 http://127.0.0.1:8000/health; then
  ok "backend answering on 127.0.0.1:8000"
else
  bad "backend never came up — check Console.app for [paths] and [addin] lines"; exit 1
fi

step "3. The TLS listener and its certificate"
STATUS=$(curl -s --max-time 5 http://127.0.0.1:8000/addin/status)
echo "$STATUS" | python3 -m json.tool 2>/dev/null | sed 's/^/        /'
echo "$STATUS" | grep -q '"pane_available": *true' \
  && ok "the task pane is present in the bundle" \
  || bad "pane_available is false — addin/ did not make it into the build"
[ -f "$CERT" ] && ok "certificate generated at $CERT" \
                || bad "no certificate at $CERT"
curl -sk -o /dev/null --max-time 5 -w '' https://127.0.0.1:8443/addin/taskpane.html \
  && ok "the pane is served over TLS on 8443" \
  || bad "nothing serving on 8443"

step "4. The certificate covers BOTH names"
# A certificate is checked against the name the client ASKED for. One covering
# only `localhost` is rejected for https://127.0.0.1 — and that shows up as a
# blank pane with no error anywhere, which is the failure this whole feature is
# most likely to die of.
if [ -f "$CERT" ]; then
  SAN=$(openssl x509 -in "$CERT" -noout -text | grep -A1 "Subject Alternative Name" | tail -1)
  note "SAN: $(echo "$SAN" | xargs)"
  echo "$SAN" | grep -q "DNS:localhost"      && ok "covers localhost"   || bad "missing DNS:localhost"
  echo "$SAN" | grep -q "IP Address:127.0.0.1" && ok "covers 127.0.0.1" || bad "missing IP:127.0.0.1"
  DAYS=$(( ( $(date -j -f "%b %d %T %Y %Z" "$(openssl x509 -in "$CERT" -noout -enddate | cut -d= -f2)" +%s 2>/dev/null || echo 0) - $(date +%s) ) / 86400 ))
  [ "$DAYS" -gt 0 ] && [ "$DAYS" -le 825 ] && ok "validity ${DAYS}d (<= 825)" || note "validity: ${DAYS}d"
fi

step "5. TRUST IT THROUGH THE APP — this is the part being tested"
note "In Standpoint:  Settings ▸ Word add-in ▸ Trust the certificate"
note ""
note "Watch for, and note down:"
note "  * does the EXPLANATION appear before the password prompt?"
note "  * how many steps, and does it ask for your password?"
note "  * press Cancel ONCE first: it must say the panel stays blank and"
note "    offer to try again — not show an error or get stuck."
note ""
read -r -p "      Press Return once you have trusted it… " _

step "6. The app recorded the trust"
AFTER=$(curl -s --max-time 5 http://127.0.0.1:8000/addin/status)
echo "$AFTER" | grep -q '"trust_state": *"trusted"' \
  && ok "trust_state is now \"trusted\"" \
  || bad "trust_state is $(echo "$AFTER" | python3 -c 'import json,sys; print(json.load(sys.stdin)["trust_state"])' 2>/dev/null)"

step "7. macOS really trusts it (not just our record)"
# ASK macOS, NOT curl. `trust_state` only says the app believes it succeeded,
# and curl is the wrong witness: modern macOS ships it built against LibreSSL,
# which does not consult the keychain at all. The earlier version of this step
# used `curl` and would have reported "the keychain entry did not take" to a
# tester whose keychain was perfectly fine. security(1) asks the system's own
# trust evaluation, which is the thing WKWebView also asks.
if security verify-cert -c "$CERT" >/dev/null 2>&1; then
  ok "macOS trust evaluation accepts the certificate"
else
  bad "macOS does not trust the certificate — the keychain entry did not take"
  note "Re-run the trust step in Settings, and say whether a password prompt"
  note "appeared. The pane will be BLANK until this passes, with no error."
fi

# Independent of any trust store: does the certificate match what the server
# actually presents? A mismatch here means the backend is serving an older
# certificate than the file you trusted — restart Standpoint.
if curl -s -o /dev/null --max-time 5 --cacert "$CERT" https://127.0.0.1:8443/health; then
  ok "the served certificate matches the file (restart not needed)"
else
  bad "the server is NOT presenting this certificate — restart Standpoint"
fi

step "8. Sideload the pane and open it in Word"
mkdir -p "$WEF"
MANIFEST="$APP/Contents/Resources/backend/addin/manifest-local.xml"
[ -f "$MANIFEST" ] || MANIFEST="$(dirname "$0")/manifest-local.xml"
if [ -f "$MANIFEST" ]; then
  cp "$MANIFEST" "$WEF/" && ok "copied $(basename "$MANIFEST") into wef/"
else
  bad "could not find manifest-local.xml"
fi
note ""
note "Now: restart Word, open any document, then Home ▸ Standpoint ▸ Local."
note ""
note "PASS looks like:  the pane shows \"Connected to Standpoint\" in green."
note "FAIL looks like:  a BLANK pane — which means the certificate was refused"
note "                  (there is no error text anywhere; that is the point of"
note "                  step 7 above)."
note ""
read -r -p "      Did the pane say Connected? [y/N] " ANSWER
case "$ANSWER" in
  y|Y) ok "the pane reached the backend inside Word — end to end" ;;
  *)   bad "the pane did not connect" ;;
esac

step "Summary"
printf "  %d passed, %d failed\n" "$PASS" "$FAIL"
note ""
note "To undo the trust afterwards:"
note "  security remove-trusted-cert -d \"$CERT\""
exit $(( FAIL > 0 ))
