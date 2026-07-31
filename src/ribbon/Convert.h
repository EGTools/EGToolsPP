// Convert.h — compatibility function-name conversion (Ribbon commands).
//
// Two user-triggered, batch "find & replace"-style passes over the ACTIVE
// workbook's cell formulas and defined names, rewriting the names of EGTools++
// shadowed functions so they resolve correctly across Excel versions. See
// plan/14_호환변환_리본.md and the 3-class model in the approved plan.
//
// Both commands: scan → MessageBox summary/confirm → backup (SaveCopyAs) →
// apply → completion notice (with backup path). Nothing is changed without the
// user's explicit Yes.

#pragma once

namespace egtools::ribbon
{
    // Direction A — "구버전 호환 적용": rewrite native/stored (`_xlfn.*`) and
    // `EG.*` tokens to the bare EGTools names, so a modern-authored (or
    // modern-EGTools-authored) document resolves to our compatibility UDFs on an
    // Excel that lacks the native functions.
    void convertToCompat();

    // Direction B — "내장 함수로 복원": rewrite EGTools-authored names to the
    // native built-ins where the running Excel provides them. IMAGE stays EGTools
    // (bare IMAGE → EG.IMAGE); GROUPBY/PIVOTBY toggle their text aggregator to a
    // function reference; unsupported cases are left untouched.
    void convertToNative();
}
