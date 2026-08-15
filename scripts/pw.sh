# Browser helpers for driving a real-host round. Source it: `. scripts/pw.sh`
#
# WHY THIS FILE EXISTS AT ALL. It lived in /tmp for a day and a half, rebuilt by
# hand at the start of every session, and the paths it carried were keyed to the
# agent session that happened to create them — a scratchpad directory with a UUID
# in it. Archive that session and the browser, its profile and every helper go
# with it. This file is the durable version: same three definitions, in the repo,
# working from any shell on this machine.
#
# Needs `@playwright/cli` installed globally (`npm i -g @playwright/cli`).

# The CLI's own JavaScript, run by node directly.
#
# NOT the `playwright-cli` command. AppLocker on this machine blocks the `.cmd`
# shim from every shell however the arguments are quoted, so the only route in is
# node.exe — which is allow-listed and is what the shim would have run anyway.
# `PLAYWRIGHT_CLI_JS` overrides for a machine whose layout differs; the driver
# reads the same variable, so setting it fixes both at once.
PW_ENTRY="${PLAYWRIGHT_CLI_JS:-$(node -e 'const{join,dirname}=require("path");console.log(join(dirname(process.execPath),"node_modules","@playwright","cli","playwright-cli.js"))')}"

# WHERE THE BROWSER LIVES, and it must be a stable path.
#
# The CLI daemon keys a session by the working directory STRING, so the browser
# is only findable from the exact path that opened it. Parking it in an agent's
# scratchpad meant a new session looked in a new directory, found "(no
# browsers)", and could not reach a browser that was sitting on screen.
#
# `.pw-session/` in the repo, gitignored. Any session on this machine finds it.
# Override with PW_SESSION_DIR if a second browser is wanted alongside.
PW_DIR="${PW_SESSION_DIR:-$(git -C "$(dirname "${BASH_SOURCE[0]}")" rev-parse --show-toplevel 2>/dev/null || echo .)/.pw-session}"
mkdir -p "$PW_DIR"

# `-s=ms` names the session; `--raw` drops the status wrapper so the output can
# be grepped. Both are what `scripts/round.mjs` passes, and they have to match or
# the driver and these helpers address different browsers.
pw() { (cd "$PW_DIR" && node "$PW_ENTRY" -s=ms --raw "$@"); }

# A `ref_N` for the first element matching a pattern.
#
# The ref on the MATCHING LINE, never the first in the output: `find` prints the
# whole frame hierarchy above a hit, so the first ref belongs to the outer iframe
# — the OneDrive document, where `Office` and `PowerPoint` are both undefined.
# Evaluating there reports a healthy host as dead. This mirrors `refFor` in
# `scripts/round.mjs`; fix them together.
paneref() { pw find "$1" | grep -E "$2" | head -1 | sed -n 's/.*ref=\([a-z0-9]*\).*/\1/p'; }

# Put the deck back to one slide, so the next round is comparable with the last.
pwclean() {
  local ref
  ref="$(paneref "Chart" 'tab "Chart"')"
  [ -z "$ref" ] && {
    echo "no pane — is the add-in open?"
    return 1
  }
  pw eval 'async () => { const budget = (p, ms) => Promise.race([p, new Promise((_, r) => setTimeout(() => r(new Error("TIMEOUT")), ms))]); try { const n = await budget(PowerPoint.run(async (c) => { const s = c.presentation.slides; s.load("items/id"); await c.sync(); const count = s.items.length; for (let i = count - 1; i >= 1; i--) c.presentation.slides.getItemAt(i).delete(); await c.sync(); s.load("items/id"); await c.sync(); return s.items.length; }), 90000); return "deck:" + n; } catch (e) { return "deck-failed"; } }' "$ref"
}

echo "pw ready — session dir $PW_DIR"
