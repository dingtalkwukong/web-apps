#!/usr/bin/env bash
set -eu

ROOT=$(cd "$(dirname "$0")/../.." && pwd)
API="$ROOT/apps/api/documents/api.js"
BROWSER_BENCH="$ROOT/test/preview-optimizations/browser-preview-bench.js"
REAL_PREVIEW_TEST="$ROOT/test/preview-optimizations/real-preview-test.js"

fail() {
    echo "preview optimization check failed: $1" >&2
    exit 1
}

assert_contains() {
    file=$1
    pattern=$2
    grep -F -- "$pattern" "$file" >/dev/null || fail "$file does not contain: $pattern"
}

assert_matches() {
    file=$1
    pattern=$2
    grep -E -- "$pattern" "$file" >/dev/null || fail "$file does not match: $pattern"
}

lite_block() {
    awk '
        /if \(previewLite\)/ {flag=1}
        flag {print}
        flag && /Common\.Controllers\.Shortcuts/ {
            seen++
            if (seen == 2) exit
        }
    ' "$1"
}

assert_lite_modules() {
    file=$1
    block=$(lite_block "$file")

    for required in \
        "'Viewport'" \
        "'DocumentHolder'" \
        "'Toolbar'" \
        "'Statusbar'" \
        "'LeftMenu'" \
        "'Main'" \
        "'ViewTab'" \
        "'Search'" \
        "'Print'" \
        "'Common.Controllers.Fonts'" \
        "'Common.Controllers.Shortcuts'"
    do
        printf '%s\n' "$block" | grep -F "$required" >/dev/null || fail "$file previewLite is missing $required"
    done

    for heavy in \
        "RightMenu" \
        "Comments" \
        "Chat" \
        "History" \
        "Plugins" \
        "ReviewChanges" \
        "Protection" \
        "Draw" \
        "ExternalLinks" \
        "ExternalDiagramEditor" \
        "ExternalMergeEditor" \
        "ExternalOleEditor" \
        "PasteOptions" \
        "DocProtection" \
        "WBProtection" \
        "FormsTab" \
        "SlideMasterTab" \
        "Transitions" \
        "Animation" \
        "CellEditor" \
        "FormulaDialog" \
        "Spellcheck" \
        "PivotTable" \
        "DataTab" \
        "TableDesignTab" \
        "ChartTab" \
        "SparklineTab"
    do
        if printf '%s\n' "$block" | grep -F "$heavy" >/dev/null; then
            fail "$file previewLite loads heavy module $heavy"
        fi
    done
}

# PDF native iframe fast path.
assert_contains "$API" "function shouldUseNativePdfPreview(config)"
assert_contains "$API" "function isMobileNativePdfPreviewBrowser()"
assert_contains "$API" "function isNativePdfPreviewBrowserType(config)"
assert_contains "$API" "config.type !== 'mobile'"
assert_contains "$API" "nav.userAgentData && nav.userAgentData.mobile === true"
assert_contains "$API" "platform === 'MacIntel' && nav.maxTouchPoints > 1"
assert_contains "$API" "config.document && config.document[name] !== undefined"
assert_contains "$API" "getConfigFlag(config, 'openPdfInBrowser') !== false"
assert_contains "$API" "getConfigFlag(config, 'openPdfAsBinary') !== true"
assert_contains "$API" "config.document && config.document.isForm !== true"
assert_contains "$API" "isViewOnlyMode(config)"
assert_contains "$API" "function hasPdfEditOrReviewMode(config)"
assert_contains "$API" "editorConfig.mode !== 'fillforms'"
assert_contains "$API" "!hasPdfEditOrReviewMode(config)"
assert_contains "$API" "function sanitizePreviewTraceUrl(value)"
assert_contains "$API" "function getPreviewTraceElapsedMs(config)"
assert_contains "$API" "function previewCompleteTrace(config, data)"
assert_contains "$API" "payload[prop] = sanitizePreviewTraceValue(prop, data[prop]);"
assert_contains "$API" "previewElapsedMs: getPreviewTraceElapsedMs(config)"
assert_contains "$API" "totalPreviewMs = getPreviewTraceElapsedMs(config)"
assert_contains "$API" "isPreviewTraceNetworkEnabled(config)"
assert_contains "$API" "previewTrace(config, \"preview-complete\", data);"
assert_contains "$API" "function registerNativePdfCache(config, iframe, sourceUrl)"
assert_contains "$API" "getDocumentServerRootPath() + \"downloadfile-cache/register/\""
assert_contains "$API" "window.fetch(getNativePdfCacheRegisterUrl(config), {"
assert_contains "$API" "setNativePdfPreviewSrc(config, iframe, data.url, \"native-pdf-cache-hit-url\");"
assert_contains "$API" "setNativePdfPreviewSrc(config, iframe, sourceUrl, \"native-pdf-iframe-create\");"
assert_contains "$API" "previewTrace(config, \"native-pdf-cache-fallback\", {"
assert_contains "$API" "iframe.setAttribute(\"data-native-pdf-preview\", \"true\");"
assert_contains "$API" "iframe.setAttribute(\"data-onlyoffice-native-pdf-preview\", \"true\");"
assert_contains "$API" "_fireEvent('onAppReady');"
assert_contains "$API" "_fireEvent('onDocumentReady');"
assert_contains "$API" "if (!useNativePdfPreview && _config.document && (_config.document.isForm!==true && _config.document.isForm!==false))"
assert_contains "$API" "params += \"&mode=view\";"
assert_contains "$BROWSER_BENCH" "'pdf-unknown-form-native'"
assert_contains "$BROWSER_BENCH" "'pdf-large-native'"
assert_contains "$BROWSER_BENCH" "const LARGE_PDF_BYTES = 19 * 1024 * 1024;"
assert_contains "$BROWSER_BENCH" "function buildPreviewTimeSummary(results)"
assert_contains "$BROWSER_BENCH" "previewTimeSummary: buildPreviewTimeSummary(results)"
assert_contains "$BROWSER_BENCH" "editorOverheadByteReductionVsPdfFallbackPercent"
assert_contains "$BROWSER_BENCH" "delete doc.isForm;"
assert_contains "$BROWSER_BENCH" "large native PDF should load the large PDF URL directly"
assert_contains "$REAL_PREVIEW_TEST" "const TEST_NAME = '真实预览测试';"
assert_contains "$REAL_PREVIEW_TEST" "showSensitive: false"
assert_contains "$REAL_PREVIEW_TEST" "--show-sensitive"
assert_contains "$REAL_PREVIEW_TEST" "function sanitizeUrlForReport(value)"
assert_contains "$REAL_PREVIEW_TEST" "return '/document-url';"
assert_contains "$REAL_PREVIEW_TEST" "'PDF must not open full editor path'"
assert_contains "$REAL_PREVIEW_TEST" "'PDF must not use fillforms mode'"
assert_contains "$ROOT/test/preview-optimizations/api-preview-optimizations.test.js" "PDF Windows, macOS, and Linux desktop browsers use native PDF"
assert_contains "$ROOT/test/preview-optimizations/api-preview-optimizations.test.js" "PDF mobile browsers use the OnlyOffice read-only preview instead of native PDF"
assert_contains "$ROOT/test/preview-optimizations/api-preview-optimizations.test.js" "PDF mobile browsers forced to desktop still do not use native PDF"
assert_contains "$ROOT/test/preview-optimizations/api-preview-optimizations.test.js" "PDF mobile user agents do not use native PDF even with desktop config"

# PDF binary open path.
assert_contains "$API" "function shouldOpenPdfAsBinary(config)"
assert_contains "$API" "getConfigFlag(config, 'openPdfAsBinary') === true"
assert_contains "$API" "window.fetch(doc.url, {credentials: 'include'})"
assert_contains "$API" "command: 'openDocumentFromBinary'"
assert_contains "$API" "buffer: buffer"

# Office previewLite routing and opt-out.
assert_contains "$API" "function shouldUseOfficePreviewLite(config)"
assert_contains "$API" "getConfigFlag(config, 'openOfficePreviewLite') !== false"
assert_contains "$API" "correctedType !== 'mobile'"
assert_contains "$API" "correctedType !== 'embedded'"
assert_matches "$API" "xlsx.*pptx.*docx|docx.*xlsx.*pptx"
assert_contains "$API" "params += \"&previewLite=1\";"
assert_contains "$API" "iframe.setAttribute(\"data-office-preview-lite\", \"true\");"
assert_contains "$API" "iframe.setAttribute(\"data-onlyoffice-preview-lite\", \"true\");"

# Warmup keeps the old default and adds the explicit light path.
assert_contains "$API" "options && options.previewLite ? 'api/documents/preload-lite.html' : 'api/documents/preload.html'"
assert_contains "$API" "iframe.setAttribute(\"data-onlyoffice-preview-lite\", \"true\");"
assert_contains "$ROOT/apps/api/documents/preload-lite.html" "docserviceworker.js?__inline=true"
assert_contains "$ROOT/apps/api/documents/preload-lite.html" "apps/documenteditor/main/app.js"
assert_contains "$ROOT/apps/api/documents/preload-lite.html" "apps/spreadsheeteditor/main/app.js"
assert_contains "$ROOT/apps/api/documents/preload-lite.html" "apps/presentationeditor/main/app.js"

for editor in documenteditor spreadsheeteditor presentationeditor; do
    assert_contains "$ROOT/apps/$editor/main/app.js" "var previewLiteControllers = ["
    assert_contains "$ROOT/apps/$editor/main/app.js" "var previewLiteModules = ["
    assert_contains "$ROOT/apps/$editor/main/app_dev.js" "var previewLite = /(?:^|[?&])previewLite=1(?:&|$)/.test(window.location.search);"
    assert_contains "$ROOT/apps/$editor/main/app_dev.js" "controllers : previewLite ? ["
    assert_contains "$ROOT/apps/$editor/main/app_dev.js" "require(previewLite ? ["
    assert_contains "$ROOT/apps/$editor/main/app.js" "$editor/main/app_pack_lite"
    assert_contains "$ROOT/apps/$editor/main/app_dev.js" "$editor/main/app_pack_lite"
    assert_contains "$ROOT/apps/$editor/main/app_pack_lite.js" "Common.NotificationCenter.trigger('app-pack:loaded');"
done

assert_lite_modules "$ROOT/apps/documenteditor/main/app.js"
assert_lite_modules "$ROOT/apps/spreadsheeteditor/main/app.js"
assert_lite_modules "$ROOT/apps/presentationeditor/main/app.js"

# previewLite leaves heavy controllers unloaded, so startup and common event paths
# must tolerate those controllers being absent.
assert_contains "$ROOT/apps/spreadsheeteditor/main/app/controller/Viewport.js" "cellEditorController && cellEditorController.createView('CellEditor'"
for editor in documenteditor spreadsheeteditor presentationeditor; do
    assert_contains "$ROOT/apps/$editor/main/app/controller/Toolbar.js" "reviewController ? reviewController.createToolbarPanel() : null"
    assert_contains "$ROOT/apps/$editor/main/app/controller/Main.js" "rightMenuController ? rightMenuController.getView('RightMenu') : null"
    assert_contains "$ROOT/apps/$editor/main/app/controller/LeftMenu.js" "commentsController && commentsController.onAfterShow();"
done
assert_contains "$ROOT/apps/presentationeditor/main/app/controller/Toolbar.js" "drawController && drawController.setApi(me.api).setMode(config);"

if command -v node >/dev/null 2>&1; then
    node "$ROOT/test/preview-optimizations/api-preview-optimizations.test.js"
else
    fail "node is required for api-preview-optimizations.test.js"
fi

echo "preview optimization checks passed"
