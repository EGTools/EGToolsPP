// ToolTipWindow.h — the in-cell IntelliSense overlay.
//
// A topmost, non-activating popup that renders a function signature (with the
// current argument emphasised) plus help text, near the formula caret. Created
// and pumped on the IntelliSense thread.

#pragma once
#include "Catalog.h"

namespace egtools::intellisense
{
    // Show/refresh the tooltip for `f`, emphasising argument `argIndex`, with its
    // top-left at screen (x, y). Does not steal focus from the cell editor.
    void toolTipShow(int x, int y, const FuncInfo& f, int argIndex);

    // Hide the tooltip (no-op if already hidden).
    void toolTipHide();

    // Destroy the window and unregister its class (unload→reload hygiene).
    // Call from the IntelliSense thread after its message loop exits.
    void toolTipDestroy();
}
