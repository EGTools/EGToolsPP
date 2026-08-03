// Version.h — Excel edition / capability detection.
//
// Drives the function-naming rule (D5, revised): a "newer-Excel" function is
// registered WITHOUT the EG. prefix when the running Excel lacks it natively
// (drop-in replacement), and WITH the EG. prefix when Excel already has it
// (avoid the name clash). EGTools-only UDFs are always registered bare.

#pragma once
#include <string>
#include <vector>

namespace egtools::core
{
    // Normalised Excel edition year: 2010/2013/2016/2019/2021/2024, or 365 for
    // Microsoft 365 (treated as the newest). 0 if unknown.
    int excelYear();

    // True if the host supports native dynamic arrays / spill (Excel 2021/365).
    bool supportsDynamicArrays();

    // True if the running Excel natively provides function `bareName` (e.g.
    // "XLOOKUP"). Used to decide the EG. prefix. Unknown names → false
    // (treated as EGTools-only).
    bool hasNativeFunction(const std::wstring& bareName);

    // The internal name a modern Excel writes into .xlsx for `bareName`
    // (e.g. "XLOOKUP" -> "_xlfn.XLOOKUP", "FILTER" -> "_xlfn._xlws.FILTER").
    // Used to register a legacy-compat alias so modern-authored files resolve to
    // our UDF on older Excel. Returns "" for EGTools-only / non-shadowed names.
    // Table is empirically derived — see tests/extract_stored_names.ps1.
    std::wstring storedName(const std::wstring& bareName);

    // Bare names that collide with the host's own function-name table on legacy
    // hosts: a modern binary running under an older license refuses (SORT/FILTER,
    // "사용한 함수가 올바르지 않습니다") or hijacks the bare entry. These register
    // and convert with an "x" prefix there instead — xSORT/xFILTER/xLET, the same
    // convention EGTools VB used.
    bool needsXPrefix(const std::wstring& bareName);

    // True if `bareName` is table-eligible on modern hosts but its native
    // rollout is NOT yet general (e.g. IMPORTTEXT/IMPORTCSV — 365 preview
    // channels only, 2026-08). The restore-to-native conversion converts such
    // a function only after a positive runtime probe confirms the native name
    // actually resolves on this host (ribbon/Convert.cpp); everything else is
    // never probed. Flip the funcTable flag once the rollout completes.
    bool isStagedRollout(const std::wstring& bareName);

    // The name this add-in actually registers for `bareName` on the CURRENT host:
    // EG.<F> when the host has the native function, x<F> for keyword-conflict
    // names on non-native hosts, bare <F> otherwise (drop-in).
    std::wstring registeredName(const std::wstring& bareName);

    // The full set of "shadowed" function names (upper-case bare) — every
    // function that has a native counterpart Excel stores under an internal
    // token. This is the single source of truth for the compatibility-conversion
    // command (src/ribbon/Convert): the set of names it may rewrite. Adding a new
    // shadowed function to funcTable() automatically extends the conversion scope.
    const std::vector<std::wstring>& shadowedFunctionNames();
}
