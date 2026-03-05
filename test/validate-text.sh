#!/usr/bin/env bash
# Cross-document text content comparison
# Usage: bash test/validate-text.sh <sample.docx> <official.docx>

set -euo pipefail

SAMPLE_DOCX="${1:?Usage: bash test/validate-text.sh <sample.docx> <official.docx>}"
OFFICIAL_DOCX="${2:?Usage: bash test/validate-text.sh <sample.docx> <official.docx>}"

WORKDIR="$(mktemp -d)"
trap 'rm -rf "$WORKDIR"' EXIT

# Check dependencies
for cmd in libreoffice pdftotext pdfinfo wdiff; do
    if ! command -v "$cmd" &>/dev/null; then
        echo "ERROR: '$cmd' is required but not found in PATH"
        exit 1
    fi
done

PASSED=0
FAILED=0
WARNINGS=0

pass() { echo -e "  \033[32mPASS\033[0m  $1"; PASSED=$((PASSED + 1)); }
fail() { echo -e "  \033[31mFAIL\033[0m  $1"; echo "         $2"; FAILED=$((FAILED + 1)); }
warn() { echo -e "  \033[33mWARN\033[0m  $1"; echo "         $2"; WARNINGS=$((WARNINGS + 1)); }

# Convert DOCX to PDF to text
convert_to_text() {
    local docx="$1" name="$2"
    local abs_docx
    abs_docx="$(cd "$(dirname "$docx")" && pwd)/$(basename "$docx")"
    cp "$abs_docx" "$WORKDIR/${name}.docx"
    libreoffice --headless --convert-to pdf --outdir "$WORKDIR" "$WORKDIR/${name}.docx" >/dev/null 2>&1
    pdftotext -layout "$WORKDIR/${name}.pdf" "$WORKDIR/${name}.txt"
    echo "$WORKDIR/${name}"
}

echo ""
echo "  Text Content Cross-Comparison"
echo ""

SAMPLE="$(convert_to_text "$SAMPLE_DOCX" sample)"
OFFICIAL="$(convert_to_text "$OFFICIAL_DOCX" official)"

# ── a) Structural markers (ordering) ────────────────────────────────────────

MARKERS=(
    "DOI:"
    "Аннотация."
    "Ключевые слова:"
    "Для цитирования:"
    "Abstract."
    "Keywords:"
    "For citation:"
    "Введение"
    "Список литературы"
)

check_marker_order() {
    local file="$1"
    local prev_line=0 ok=true
    for marker in "${MARKERS[@]}"; do
        local line
        line=$(grep -n "$marker" "$file" | head -1 | cut -d: -f1 || true)
        if [ -z "$line" ]; then
            continue
        fi
        if [ "$line" -lt "$prev_line" ]; then
            ok=false
            break
        fi
        prev_line=$line
    done
    echo "$ok"
}

sample_order=$(check_marker_order "${SAMPLE}.txt")
official_order=$(check_marker_order "${OFFICIAL}.txt")

if [ "$sample_order" = "true" ] && [ "$official_order" = "true" ]; then
    pass "Structural marker ordering matches"
else
    details=""
    [ "$sample_order" != "true" ] && details="sample has wrong order; "
    [ "$official_order" != "true" ] && details="${details}official has wrong order"
    fail "Structural marker ordering" "$details"
fi

# Check that each marker in official also exists in sample
missing_markers=()
for marker in "${MARKERS[@]}"; do
    if grep -q "$marker" "${OFFICIAL}.txt" && ! grep -q "$marker" "${SAMPLE}.txt"; then
        missing_markers+=("$marker")
    fi
done
if [ ${#missing_markers[@]} -eq 0 ]; then
    pass "All structural markers present"
else
    fail "Structural markers missing from sample" "${missing_markers[*]}"
fi

# ── b) Shared content comparison ────────────────────────────────────────────

CONTENT_PATTERNS=(
    "И.И. Иванов"
    "П.П. Петров"
    "I.I. Ivanov"
    "P.P. Petrov"
    "Институт системного программирования"
    "Ivannikov Institute"
    "Московский государственный университет"
    "Lomonosov Moscow State University"
    "ORCID: 0000-0000-0000-0000"
    "Ермаков М. К."
)

content_missing=()
content_extra=()
for pattern in "${CONTENT_PATTERNS[@]}"; do
    in_official=false
    in_sample=false
    grep -q "$pattern" "${OFFICIAL}.txt" && in_official=true
    grep -q "$pattern" "${SAMPLE}.txt" && in_sample=true

    if $in_official && ! $in_sample; then
        content_missing+=("$pattern")
    elif $in_sample && ! $in_official; then
        content_extra+=("$pattern")
    fi
done

if [ ${#content_missing[@]} -eq 0 ]; then
    pass "Shared content present in sample"
else
    fail "Content missing from sample" "${content_missing[*]}"
fi

if [ ${#content_extra[@]} -gt 0 ]; then
    warn "Extra content in sample (not in official)" "${content_extra[*]}"
fi

# ── c) Page count comparison ────────────────────────────────────────────────

sample_pages=$(pdfinfo "${SAMPLE}.pdf" | grep "^Pages:" | awk '{print $2}')
official_pages=$(pdfinfo "${OFFICIAL}.pdf" | grep "^Pages:" | awk '{print $2}')

if [ "$sample_pages" -lt "$official_pages" ]; then
    fail "Page count" "sample has $sample_pages pages, official has $official_pages (content may be lost)"
elif [ "$sample_pages" -gt $((official_pages + 2)) ]; then
    warn "Page count" "sample has $sample_pages pages, official has $official_pages (possible formatting bloat)"
else
    pass "Page count: sample=$sample_pages, official=$official_pages"
fi

# ── d) Normalized text diff ─────────────────────────────────────────────────

normalize_text() {
    local file="$1"
    sed -E \
        -e '/^[[:space:]]*[0-9]+X?[[:space:]]*$/d' \
        -e '/Иванов И\.И\., Петров П\.П\. Заголовок статьи\. Труды ИСП РАН/d' \
        -e '/Ivanov I\.I\., Petrov P\.P\. Article title\. Trudy ISP RAN/d' \
        -e 's/[[:space:]]+/ /g' \
        -e 's/[[:space:]]+$//' \
        -e 's/[""„«»]/"/g' \
        -e 's/[–—]/-/g' \
        -e 's/([[:alpha:]]\.)[[:space:]]+([[:alpha:]]\.)/\1\2/g' \
        -e 's/([[:alpha:]]\.)[[:space:]]+([[:alpha:]]\.)/\1\2/g' \
        "$file" \
    | cat -s
}

normalize_text "${SAMPLE}.txt" > "${SAMPLE}.norm.txt"
normalize_text "${OFFICIAL}.txt" > "${OFFICIAL}.norm.txt"

# wdiff exit code: 0 = no differences, 1 = differences found, >1 = error
wdiff_stats=$(wdiff -s "${OFFICIAL}.norm.txt" "${SAMPLE}.norm.txt" 2>&1 || true)
# Stats line example: "file1: 1916 words  1807 92% common  12 1% deleted  97 5% changed"
stats_line=$(echo "$wdiff_stats" | grep "^${OFFICIAL}.norm.txt:" || true)
deleted=$(echo "$stats_line" | sed -E 's/.*[[:space:]]([0-9]+)[[:space:]]+[0-9]+%[[:space:]]+deleted.*/\1/')
changed=$(echo "$stats_line" | sed -E 's/.*[[:space:]]([0-9]+)[[:space:]]+[0-9]+%[[:space:]]+changed.*/\1/')
inserted=$(echo "$wdiff_stats" | grep "^${SAMPLE}.norm.txt:" | sed -E 's/.*[[:space:]]([0-9]+)[[:space:]]+[0-9]+%[[:space:]]+inserted.*/\1/' || true)
total_diff=$(( ${deleted:-0} + ${inserted:-0} + ${changed:-0} ))

MAX_DIFF_WORDS=105
if [ "$total_diff" -le "$MAX_DIFF_WORDS" ]; then
    pass "Word-level diff ($total_diff words differ: ${deleted:-0} deleted, ${inserted:-0} inserted, ${changed:-0} changed; max $MAX_DIFF_WORDS)"
else
    fail "Word-level diff" "$total_diff words differ (${deleted:-0} deleted, ${inserted:-0} inserted, ${changed:-0} changed) exceed threshold of $MAX_DIFF_WORDS"
fi

if [ "$total_diff" -gt 0 ]; then
    echo ""
    echo "  Word-level diff (changed regions only):"
    wdiff -3 "${OFFICIAL}.norm.txt" "${SAMPLE}.norm.txt" 2>/dev/null | head -80 || true
    echo ""
fi

# ── Summary ─────────────────────────────────────────────────────────────────

echo ""
echo "  $PASSED passed, $FAILED failed, $WARNINGS warnings"
echo ""
exit $( [ "$FAILED" -gt 0 ] && echo 1 || echo 0 )
