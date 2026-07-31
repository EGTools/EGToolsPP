// I18n.h — runtime multilingual catalog (ko·en·zh-CN·zh-TW·ja·es).
//
// JSON catalogs are embedded as .xll resources (resources/egtools.rc) and parsed
// at load. Language is detected from the Office UI language (registry) → Windows
// UI language → "en". See plan/03 §4 and plan/04 §6.

#pragma once
#include <string>
#include <vector>
#include <utility>

namespace egtools::i18n
{
    // One worksheet-function argument. `opt` marks an argument the user may omit;
    // it is language-independent, so it is sourced canonically from en.json (see
    // I18n.cpp func()). IntelliSense renders optional args as "[name]".
    struct ArgMeta { std::wstring name; std::wstring help; bool opt = false; };

    // Describes a function whose arguments repeat (e.g. LET's name/value pairs,
    // TEXTJOIN's text1,text2,…). The catalog stores ONE representative repetition
    // (head + period + tail = the listed args); this spec drives both the dynamic
    // registration (synthesise args up to `max`) and IntelliSense's R2 expansion.
    // Layout:  [head fixed args][repeating block of `period` args]…[tail fixed args]
    //   max == head + k*period + tail  for some integer k >= 1.
    // Like `opt`, it is language-independent (sourced from en.json). `has` is false
    // when the function has no repetition.
    struct RepeatSpec { int head = 0, period = 0, tail = 0, max = 0; bool has = false; };

    // Display/registration name for the `iter`-th (1-based) repetition of a
    // representative argument. iter==1 returns the name verbatim; iter>1 strips the
    // representative's trailing digits and appends `iter` ("text1"→"text3").
    std::wstring repeatArgName(const std::wstring& representative, int iter);

    // Expand one representative repetition (head+period+tail) into the full arg
    // list of length spec.max, synthesising names for repetitions 2..k (which are
    // marked optional). Used by the dynamic registration path.
    std::vector<ArgMeta> expandRepeatArgs(const std::vector<ArgMeta>& rep,
                                          const RepeatSpec& spec);

    // Parse all embedded language catalogs into memory. Call once at AutoOpen.
    void load();

    // Detect the UI language, set it current, and return the chosen code.
    std::wstring detectAndSet();

    void         setCurrent(const std::wstring& lang);
    std::wstring current();

    // Simple string lookup (current language → en → the key itself).
    std::wstring t(const std::wstring& key);

    // Function metadata: fills desc and argument metadata for `name`
    // (e.g. "EG.SEQ2D"). If `repeat` is non-null it receives the repetition spec
    // (with has=false when the function does not repeat). Returns false if the
    // function is unknown.
    bool func(const std::wstring& name,
              std::wstring& desc,
              std::vector<ArgMeta>& args,
              RepeatSpec* repeat = nullptr);
}
