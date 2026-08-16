#!/bin/bash
#
# Intune Custom Compliance - Discovery Script
# CrowdStrike Falcon - minimal health check for macOS
#
# Checks:
#   1. Sensor operational: true
#   2. Sensor status: loaded
#   3. "Established At" OR "Last Established At" within the last 30 days
#
# Upload to: Endpoint security > Device compliance > Scripts > Add > macOS
#   - Run script using logged-on credentials: NO  (falconctl requires root)
#   - Encoding: UTF-8 without BOM
#
# Outputs EXACTLY one line of JSON to stdout. Never add debug echo here.
#

export PATH="/usr/bin:/bin:/usr/sbin:/sbin"
export LC_ALL=C   # force English month names for date parsing

FALCONCTL="/Applications/Falcon.app/Contents/Resources/falconctl"
LOG="/var/log/intune_crowdstrike_discovery.log"

# Defaults chosen so a broken/missing sensor is reported as non-compliant
SensorOperational=false
SensorStatus="unknown"
ConnectionAgeDays=9999

log() { echo "$(date '+%Y-%m-%d %H:%M:%S') $1" >> "$LOG" 2>/dev/null; }

# Convert "  Established At: Aug 16, 2026 at 1:03:29 PM" -> epoch seconds
# Only the date part is parsed; day granularity is enough for a 30-day window.
to_epoch() {
    _line="$1"
    [ -z "$_line" ] && return 1
    _val=$(echo "$_line" | sed 's/^[^:]*://' | sed 's/ at .*//' | sed 's/^[[:space:]]*//;s/[[:space:]]*$//')
    [ -z "$_val" ] && return 1
    _epoch=$(date -j -f "%b %d, %Y" "$_val" "+%s" 2>/dev/null)
    if [ -n "$_epoch" ]; then
        echo "$_epoch"
        return 0
    fi
    return 1
}

log "--- discovery start ---"

if [ -x "$FALCONCTL" ]; then
    STATS=$("$FALCONCTL" stats 2>/dev/null)

    if [ -n "$STATS" ]; then

        # --- 1. Sensor operational ---
        if echo "$STATS" | grep -i "Sensor operational:" | grep -qi "true"; then
            SensorOperational=true
        fi

        # --- 2. Sensor status ---
        RAW_STATUS=$(echo "$STATS" | grep -i "Sensor status:" | head -n 1 \
                     | sed 's/^[^:]*://' | sed 's/^[[:space:]]*//;s/[[:space:]]*$//' \
                     | tr '[:upper:]' '[:lower:]')
        if [ -n "$RAW_STATUS" ]; then
            SensorStatus=$(echo "$RAW_STATUS" | tr -cd 'a-z0-9_-')
        fi

        # --- 3. Newest of "Established At" / "Last Established At" ---
        EST_LINE=$(echo "$STATS" | grep "Established At:" | grep -v "Last Established At:" | head -n 1)
        LAST_LINE=$(echo "$STATS" | grep "Last Established At:" | head -n 1)

        EST_EPOCH=$(to_epoch "$EST_LINE")
        LAST_EPOCH=$(to_epoch "$LAST_LINE")

        NEWEST=""
        if [ -n "$EST_EPOCH" ]; then NEWEST="$EST_EPOCH"; fi
        if [ -n "$LAST_EPOCH" ]; then
            if [ -z "$NEWEST" ] || [ "$LAST_EPOCH" -gt "$NEWEST" ]; then
                NEWEST="$LAST_EPOCH"
            fi
        fi

        if [ -n "$NEWEST" ]; then
            NOW=$(date "+%s")
            DIFF=$(( NOW - NEWEST ))
            [ "$DIFF" -lt 0 ] && DIFF=0
            ConnectionAgeDays=$(( DIFF / 86400 ))
        fi

        log "operational=$SensorOperational status=$SensorStatus est=$EST_EPOCH last=$LAST_EPOCH ageDays=$ConnectionAgeDays"
    else
        log "falconctl returned no output - sensor not loaded"
    fi
else
    log "falconctl not found - sensor not installed"
fi

# --- Emit single-line JSON (booleans/ints unquoted, strings quoted) ---
/bin/echo "{\"CrowdStrikeSensorOperational\":$SensorOperational,\"CrowdStrikeSensorStatus\":\"$SensorStatus\",\"CrowdStrikeConnectionAgeDays\":$ConnectionAgeDays}"

log "--- discovery end ---"
exit 0