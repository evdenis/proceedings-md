#!/usr/bin/env bash
set -euo pipefail

SAMPLE_DOCX="${1:?Usage: $0 <sample.docx> <reference.docx>}"
REFERENCE_DOCX="${2:?Usage: $0 <sample.docx> <reference.docx>}"
THRESHOLD="${THRESHOLD:-5}"
TEMP_DIR="$(mktemp -d)"
trap 'rm -rf "$TEMP_DIR"' EXIT

# Check dependencies
for cmd in libreoffice pdftoppm compare bc; do
    if ! command -v "$cmd" &>/dev/null; then
        echo "ERROR: '$cmd' is required but not found in PATH"
        exit 1
    fi
done

# Render DOCX to per-page PNGs in a subdirectory of TEMP_DIR
render_pages() {
    local docx="$1" name="$2"
    local outdir="$TEMP_DIR/$name"
    mkdir -p "$outdir"

    echo "Converting $name DOCX to PDF..." >&2
    libreoffice --headless --convert-to pdf --outdir "$outdir" "$docx" >/dev/null 2>&1
    local pdf="$outdir/$(basename "${docx%.docx}.pdf")"

    if [ ! -f "$pdf" ]; then
        echo "ERROR: PDF conversion failed for $name" >&2
        exit 1
    fi

    echo "Rendering $name pages at 300 DPI..." >&2
    pdftoppm -png -r 300 "$pdf" "$outdir/page"

    local count
    count=$(ls "$outdir"/page-*.png 2>/dev/null | wc -l)
    echo "Rendered $count pages for $name" >&2

    if [ "$count" -eq 0 ]; then
        echo "ERROR: No pages rendered for $name" >&2
        exit 1
    fi

    echo "$outdir"
}

SAMPLE_DIR=$(render_pages "$SAMPLE_DOCX" sample)
REFERENCE_DIR=$(render_pages "$REFERENCE_DOCX" reference)

echo ""
echo "Comparing pages..."
FAILURES=0

for CURRENT in "$SAMPLE_DIR"/page-*.png; do
    PAGE_NAME="$(basename "$CURRENT")"
    REFERENCE="$REFERENCE_DIR/$PAGE_NAME"

    if [ ! -f "$REFERENCE" ]; then
        echo "  SKIP  $PAGE_NAME (no reference page)"
        continue
    fi

    # Get total pixel count
    DIMENSIONS=$(identify -format '%w %h' "$CURRENT" 2>/dev/null)
    if [ -z "$DIMENSIONS" ]; then
        echo "  ERROR $PAGE_NAME (identify failed)"
        FAILURES=$((FAILURES + 1))
        continue
    fi
    TOTAL_PX=$(echo "$DIMENSIONS" | awk '{print $1 * $2}')

    # AE returns differing pixel count; Q16-HDRI format: "raw (normalized)"
    AE_OUTPUT=$(compare -metric AE "$REFERENCE" "$CURRENT" /dev/null 2>&1 || true)
    AE=$(echo "$AE_OUTPUT" | grep -oP '\([\d.]+\)' | tr -d '()')
    if [ -z "$AE" ]; then
        # Non-HDRI: plain integer
        AE=$(echo "$AE_OUTPUT" | grep -oP '^\d+' || true)
    fi
    if [ -z "$AE" ]; then
        echo "  ERROR $PAGE_NAME (comparison failed)"
        FAILURES=$((FAILURES + 1))
        continue
    fi

    PERCENT=$(echo "scale=1; $AE * 100 / $TOTAL_PX" | bc -l 2>/dev/null || echo "0")
    PERCENT_INT=$(echo "$PERCENT" | cut -d. -f1)
    if [ "${PERCENT_INT:-0}" -gt "$THRESHOLD" ]; then
        echo "  FAIL  $PAGE_NAME (AE: ${PERCENT}% > ${THRESHOLD}%)"
        if [ -n "${DIFF_DIR:-}" ]; then
            mkdir -p "$DIFF_DIR"
            compare "$REFERENCE" "$CURRENT" "$DIFF_DIR/$PAGE_NAME" 2>/dev/null || true
        fi
        FAILURES=$((FAILURES + 1))
    else
        echo "  PASS  $PAGE_NAME (AE: ${PERCENT}%)"
    fi
done

if [ "$FAILURES" -gt 0 ]; then
    echo ""
    echo "$FAILURES page(s) failed visual regression"
    exit 1
else
    echo ""
    echo "All pages within ${THRESHOLD}% AE threshold"
fi
