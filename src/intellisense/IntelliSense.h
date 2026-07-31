// IntelliSense.h — self-implemented in-cell IntelliSense (Phase 0 / S4).
//
// xlOil does NOT provide self-contained in-cell IntelliSense; it only feeds the
// external Excel-DNA IntelliSense add-in. D2 requires a self-contained, no-.NET
// implementation, so we build our own via MSAA (oleacc) + WinEvent hooks.
//
// Stage 1 (this PoC): prove we can capture the formula text the user is typing.

#pragma once

namespace egtools::intellisense
{
    // Install/remove the WinEvent hook. Call from the add-in's AutoOpen/AutoClose
    // (Excel main thread).
    void install();
    void uninstall();
}
