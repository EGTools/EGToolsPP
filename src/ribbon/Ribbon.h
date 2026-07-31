// Ribbon.h — EGTools++ Ribbon UI (COM add-in) lifecycle.
//
// Adds the "EGTools++" tab with an "EGTools" group containing the two
// compatibility-conversion buttons. Backed by xlOil's COM add-in support
// (makeComAddin + connect), which works even for a static XLL. See
// plan/14_호환변환_리본.md.

#pragma once

namespace egtools::ribbon
{
    // Create + connect the ribbon COM add-in. Call once at AutoOpen (main thread,
    // after i18n is loaded so labels can be localised). Failures are swallowed —
    // a missing ribbon must never break function registration.
    void install();

    // Disconnect + release the ribbon add-in. Call at AutoClose.
    void uninstall();
}
