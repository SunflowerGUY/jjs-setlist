#!/usr/bin/env bash
# Starts JJ's Setlist (browser edition) on Linux and opens it in your web browser.
#
#   ./start-jjs-setlist.sh              start it (keep the terminal open while you use it)
#   ./start-jjs-setlist.sh --install    add "JJ's Setlist" to your applications menu
#   ./start-jjs-setlist.sh --uninstall  remove it from the menu again
#
# Settings and saved setlists are kept as files in this folder. Needs Python 3.8+
# (standard library only - nothing else to install).

set -u
HERE="$(dirname "$(readlink -f "$0")")"
cd "$HERE" || exit 1
APP="JJ's Setlist"
APPS="${XDG_DATA_HOME:-$HOME/.local/share}/applications"
ENTRY="$APPS/jjs-setlist-web.desktop"

# Show an error in the terminal, or in a pop-up when started from the menu.
fail() {
    printf '\n%b\n\n' "$1" >&2
    if [ ! -t 2 ]; then
        if command -v zenity >/dev/null 2>&1; then
            zenity --error --no-markup --title="$APP" --text="$(printf '%b' "$1")"
        elif command -v kdialog >/dev/null 2>&1; then
            kdialog --title "$APP" --error "$(printf '%b' "$1")"
        elif command -v notify-send >/dev/null 2>&1; then
            notify-send "$APP" "$(printf '%b' "$1")"
        fi
    elif [ -t 0 ]; then
        read -r -p "Press Enter to close..." _
    fi
    exit 1
}

# The command to install Python on this distribution.
python_install_cmd() {
    if   command -v apt-get >/dev/null 2>&1; then echo "sudo apt install python3"
    elif command -v dnf     >/dev/null 2>&1; then echo "sudo dnf install python3"
    elif command -v pacman  >/dev/null 2>&1; then echo "sudo pacman -S python"
    elif command -v zypper  >/dev/null 2>&1; then echo "sudo zypper install python3"
    else echo "install your distribution's python3 package"
    fi
}

case "${1:-}" in
    --uninstall)
        rm -f "$ENTRY"
        update-desktop-database "$APPS" >/dev/null 2>&1
        echo "Removed $APP from the applications menu."
        echo "Your setlists and settings in $HERE are still there."
        exit 0 ;;
    --install)
        chmod +x "$HERE/start-jjs-setlist.sh" 2>/dev/null
        mkdir -p "$APPS"
        # In a .desktop file, quotes and $ in paths must be escaped.
        esc() { printf '%s' "$1" | sed -e 's/[\\"`$]/\\&/g'; }
        cat > "$ENTRY" <<EOF
[Desktop Entry]
Type=Application
Name=JJ's Setlist
GenericName=Setlist Organiser
Comment=Build, save and print gig setlists in your web browser
Exec=bash "$(esc "$HERE/start-jjs-setlist.sh")"
Icon=$HERE/web/icon.png
Terminal=true
Categories=AudioVideo;Audio;Music;
Keywords=setlist;songs;gig;band;music;
EOF
        update-desktop-database "$APPS" >/dev/null 2>&1
        echo "Done! Find \"$APP\" in your applications menu"
        echo "(it may take a few seconds to appear, or a log out and back in)."
        echo "It opens a terminal window: keep that open while you use the app."
        exit 0 ;;
esac

command -v python3 >/dev/null 2>&1 ||
    fail "$APP needs Python 3, and it wasn't found. Install it with:\n    $(python_install_cmd)"
python3 -c 'import sys; sys.exit(sys.version_info < (3, 8))' ||
    fail "$APP needs Python 3.8 or newer (this is $(python3 -V 2>&1))."

python3 serve.py "$@" || fail "$APP stopped with an error (see above)."
