#!/usr/bin/env bash
# ============================================================
#  DotaArena 2026 - Zapuskator dlya Debian / Ubuntu Linux
#  Zapusk: bash start.sh
# ============================================================

set -e

DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" >/dev/null 2>&1 && pwd)"
cd "$DIR"

echo "============================================================"
echo "  Zapusk DotaArena 2026 i CCTV kamer (Linux)"
echo "============================================================"
echo ""

# 1. Proverka Python 3
if ! command -v python3 >/dev/null 2>&1; then
    echo "[!] Python 3 ne nayden v sisteme."
    echo "    Ustanovite ego komandoy: sudo apt update && sudo apt install -y python3 curl"
    exit 1
fi

# 2. Proverka i zagruzka go2rtc dlya Linux
if [ ! -f "./go2rtc" ]; then
    echo "[*] Poisk i zagruzka go2rtc dlya Linux..."
    UNAME_M=$(uname -m)
    case "$UNAME_M" in
        x86_64)
            GO2RTC_BIN="go2rtc_linux_amd64"
            ;;
        aarch64|arm64)
            GO2RTC_BIN="go2rtc_linux_arm64"
            ;;
        armv7l|armhf)
            GO2RTC_BIN="go2rtc_linux_arm"
            ;;
        *)
            GO2RTC_BIN="go2rtc_linux_amd64"
            ;;
    esac

    DOWNLOAD_URL="https://github.com/AlexxIT/go2rtc/releases/download/v1.9.14/${GO2RTC_BIN}"
    echo "[*] Skachivanie ${GO2RTC_BIN} s GitHub Releases..."

    if command -v curl >/dev/null 2>&1; then
        curl -L -s "$DOWNLOAD_URL" -o "./go2rtc"
    elif command -v wget >/dev/null 2>&1; then
        wget -q -O "./go2rtc" "$DOWNLOAD_URL"
    else
        echo "[!] Ne nayden curl ili wget. Ustanovite: sudo apt install -y curl"
        exit 1
    fi

    chmod +x "./go2rtc" 2>/dev/null || true
    echo "[OK] go2rtc uspeshno skachan i podgotovlen."
fi

# 3. Ochistka protsessov pri vyhode (Ctrl+C)
cleanup() {
    echo ""
    echo "[*] Ostanovka serverov..."
    if [ -n "$GO2RTC_PID" ] && kill -0 "$GO2RTC_PID" 2>/dev/null; then
        kill "$GO2RTC_PID" 2>/dev/null || true
    fi
    if [ -n "$SERVER_PID" ] && kill -0 "$SERVER_PID" 2>/dev/null; then
        kill "$SERVER_PID" 2>/dev/null || true
    fi
    echo "[OK] Vse servera ostanovleny."
    exit 0
}

trap cleanup INT TERM EXIT

# 4. Zapusk shlyuza potokovogo video go2rtc
echo "[1/2] Zapusk shlyuza kamer go2rtc (port 1984)..."
chmod +x ./go2rtc 2>/dev/null || true
./go2rtc -config ./go2rtc.yaml >/dev/null 2>&1 &
GO2RTC_PID=$!

sleep 1

# 5. Zapusk backend-servera server.py
echo "[2/2] Zapusk veb-servera server.py (port 8000)..."
python3 server.py &
SERVER_PID=$!

sleep 1

echo ""
echo "============================================================"
echo "  VSE SERVERA USPESHNO ZAPUSHCHENY:"
echo "  - Sayt i Baza Dannykh: http://localhost:8000"
echo "  - Pryamoy efir kamer:  http://127.0.0.1:1984"
echo "============================================================"
echo ""
echo "Dlya ostanovki nazhmite Ctrl+C"

# 6. Otkrytie brauzera (esli graficheskaya obolochka dostupna)
if command -v xdg-open >/dev/null 2>&1; then
    xdg-open "http://localhost:8000" >/dev/null 2>&1 || true
elif command -v sensible-browser >/dev/null 2>&1; then
    sensible-browser "http://localhost:8000" >/dev/null 2>&1 || true
fi

# Ozhidanie
wait "$SERVER_PID"
