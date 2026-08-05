// Catalog.h — function metadata for IntelliSense (name, help, arguments).
//
// Backed by the i18n catalog (core/I18n), so descriptions/argument names follow
// the selected UI language and stay in sync with the embedded resources.

#pragma once
#include <string>
#include <vector>

#include "../core/I18n.h"   // egtools::i18n::RepeatSpec

namespace egtools::intellisense
{
    struct ArgInfo { std::wstring name; std::wstring help; bool optional = false; };
    struct FuncInfo
    {
        std::wstring name;        // canonical, e.g. "EG.HELLO"
        std::wstring help;
        std::vector<ArgInfo> args;        // one representative repetition if `repeat.has`
        egtools::i18n::RepeatSpec repeat; // variadic layout, drives R2 expansion
    };

    // Case-insensitive lookup by any alias (bare, EG.-, x-prefixed). Fills
    // `out` and returns true if known. Accepts native spellings too — use
    // lookupRegistered for anything user-facing.
    bool lookupFunction(const std::wstring& name, FuncInfo& out);

    // True only when `name` is EXACTLY the name this add-in registered on the
    // running host (core::registeredName). Native and add-in names coexist —
    // e.g. on a modern host XLOOKUP is native (Excel draws its own tooltip)
    // and ours is EG.XLOOKUP — so both the formula tooltip and the
    // autocomplete description overlay must engage only for our own names.
    bool lookupRegistered(const std::wstring& name, FuncInfo& out);
}
