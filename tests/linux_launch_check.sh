#!/usr/bin/env bash
# Checks the installed nous-atlas (.deb) on Linux. Run it under xvfb-run with a workbook:
#   xvfb-run -a bash tests/linux_launch_check.sh some.xlsx
# It opens the workbook, opens the file dialog (GTK, as on Linux Mint), then
# starts again as if in a Wayland session and checks it falls back to X11.
# Used by .github/workflows/build.yml.
set -uo pipefail
xlsx=$1
name=$(basename "$xlsx")
pid=
fail() { echo "FAIL: $*"; [ -n "$pid" ] && kill "$pid" 2>/dev/null; exit 1; }
wait_window() {  # wait_window <title regex>: prints the window id
    for _ in $(seq 30); do
        w=$(xdotool search --name "$1" 2>/dev/null | head -1)
        [ -n "$w" ] && { echo "$w"; return; }
        sleep 1
    done
}

# Opens the workbook, on a Cinnamon desktop so Qt uses the GTK theme and dialogs
XDG_CURRENT_DESKTOP=X-Cinnamon nous-atlas "$xlsx" & pid=$!
win=$(wait_window "$name")
kill -0 "$pid" 2>/dev/null || fail "nous-atlas exited early"
[ -n "$win" ] || fail "no window titled $name"
echo "window: $(xdotool getwindowname "$win")"
grep -q libqgtk3 "/proc/$pid/maps" || fail "Qt's GTK theme didn't load"
echo "GTK: $(grep -o '/[^ ]*libgtk-3\.so[^ ]*' "/proc/$pid/maps" | sort -u | xargs)"

# Ctrl+O opens the file dialog
xdotool windowfocus "$win"; sleep 1; xdotool key ctrl+o
dlg=$(wait_window '^Open workbook$')
[ -n "$dlg" ] || fail "Ctrl+O didn't open the file dialog"
echo "file dialog: '$(xdotool getwindowname "$dlg")'"
kill "$pid"; wait "$pid" 2>/dev/null

# A Wayland session falls back to X11 (XWayland); the Wayland plugin isn't shipped
XDG_SESSION_TYPE=wayland WAYLAND_DISPLAY=wayland-none nous-atlas "$xlsx" & pid=$!
win=$(wait_window "$name")
kill -0 "$pid" 2>/dev/null || fail "nous-atlas exited early in a Wayland session"
[ -n "$win" ] || fail "no window in a Wayland session"
echo "Wayland session -> X11 window: $(xdotool getwindowname "$win")"
kill "$pid"; wait "$pid" 2>/dev/null
echo "Launch checks OK"
