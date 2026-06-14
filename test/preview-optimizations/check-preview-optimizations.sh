#!/usr/bin/env bash
set -eu

ROOT=$(cd "$(dirname "$0")/../.." && pwd)
API="$ROOT/apps/api/documents/api.js"

fail() {
    echo "preview optimization check failed: $1" >&2
    exit 1
}

assert_contains() {
    file=$1
    pattern=$2
    grep -F "$pattern" "$file" >/dev/null || fail "$file does not contain: $pattern"
}

assert_matches() {
    file=$1
    pattern=$2
    grep -E "$pattern" "$file" >/dev/null || fail "$file does not match: $pattern"
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
assert_contains "$API" "config.document && config.document[name] !== undefined"
assert_contains "$API" "getConfigFlag(config, 'openPdfInBrowser') !== false"
assert_contains "$API" "getConfigFlag(config, 'openPdfAsBinary') !== true"
assert_contains "$API" "config.document && config.document.isForm === false"
assert_contains "$API" "isViewOnlyMode(config)"
assert_contains "$API" "!hasEditOrReviewMode(config)"
assert_contains "$API" "iframe.src = config.document.url;"
assert_contains "$API" "iframe.setAttribute(\"data-native-pdf-preview\", \"true\");"
assert_contains "$API" "iframe.setAttribute(\"data-onlyoffice-native-pdf-preview\", \"true\");"
assert_contains "$API" "_fireEvent('onAppReady');"
assert_contains "$API" "_fireEvent('onDocumentReady');"
assert_contains "$API" "if (!useNativePdfPreview && _config.document && (_config.document.isForm!==true && _config.document.isForm!==false))"
assert_contains "$API" "params += \"&mode=view\";"

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
