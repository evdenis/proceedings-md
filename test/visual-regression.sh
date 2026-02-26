#!/usr/bin/env bash
set -euo pipefail

DOCX="${1:?Usage: $0 <path-to-docx>}"
GOLDEN_DIR="$(dirname "$0")/golden"
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

# Convert DOCX → PDF
echo "Converting DOCX to PDF..."
libreoffice --headless --convert-to pdf --outdir "$TEMP_DIR" "$DOCX" >/dev/null 2>&1
PDF="$TEMP_DIR/$(basename "${DOCX%.docx}.pdf")"

if [ ! -f "$PDF" ]; then
    echo "ERROR: PDF conversion failed"
    exit 1
fi

# Convert PDF → per-page PNGs (300 DPI)
echo "Rendering pages at 300 DPI..."
pdftoppm -png -r 300 "$PDF" "$TEMP_DIR/page"

PAGE_COUNT=$(ls "$TEMP_DIR"/page-*.png 2>/dev/null | wc -l)
echo "Rendered $PAGE_COUNT pages"

if [ "$PAGE_COUNT" -eq 0 ]; then
    echo "ERROR: No pages rendered"
    exit 1
fi

# Compare or create golden reference
if [ -d "$GOLDEN_DIR" ] && ls "$GOLDEN_DIR"/page-*.png &>/dev/null; then
    echo "Comparing against golden reference..."
    FAILURES=0

    for CURRENT in "$TEMP_DIR"/page-*.png; do
        PAGE_NAME="$(basename "$CURRENT")"
        GOLDEN="$GOLDEN_DIR/$PAGE_NAME"

        if [ ! -f "$GOLDEN" ]; then
            echo "  NEW   $PAGE_NAME (no golden reference)"
            FAILURES=$((FAILURES + 1))
            continue
        fi

        RMSE=$(compare -metric RMSE "$GOLDEN" "$CURRENT" /dev/null 2>&1 | grep -oP '\([\d.]+\)' | tr -d '()' || true)
        if [ -z "$RMSE" ]; then
            echo "  ERROR $PAGE_NAME (comparison failed)"
            FAILURES=$((FAILURES + 1))
            continue
        fi

        PERCENT=$(echo "$RMSE * 100" | bc -l 2>/dev/null | cut -d. -f1 || echo "0")
        if [ "${PERCENT:-0}" -gt "$THRESHOLD" ]; then
            echo "  FAIL  $PAGE_NAME (RMSE: ${PERCENT}% > ${THRESHOLD}%)"
            FAILURES=$((FAILURES + 1))
        else
            echo "  PASS  $PAGE_NAME (RMSE: ${PERCENT}%)"
        fi
    done

    if [ "$FAILURES" -gt 0 ]; then
        echo ""
        echo "$FAILURES page(s) failed visual regression"
        exit 1
    else
        echo ""
        echo "All pages within ${THRESHOLD}% threshold"
    fi
else
    echo "No golden reference found. Saving current output as golden..."
    mkdir -p "$GOLDEN_DIR"
    cp "$TEMP_DIR"/page-*.png "$GOLDEN_DIR/"
    echo "Saved $PAGE_COUNT golden pages to $GOLDEN_DIR/"
fi
