// Spill.h — array output that adapts to the host Excel.
//
// The function-authoring principle (revised D5): new functions take arrays and
// return arrays. output() returns the array result so it spills:
//   * modern Excel (2021/365): just return — Excel spills natively.
//   * legacy Excel (<=2019): the forced array-formula (CSE) resize engine.
// See plan/04 §4 and plan/08 §4.2.

#pragma once
#include <string>

#include <xlOil/ExcelObj.h>

namespace egtools::core
{
    // RAII: marks `name` as the worksheet function currently executing on this
    // thread. Registry wraps every UDF with one of these so that output() can
    // attribute a legacy CSE-resize request to the UDF that produced the array —
    // only the OUTERMOST call of a cell's formula may resize; a UDF consumed as
    // an argument (e.g. xSORT in =SUM(TAKE(xSORT(...),3))) must return its array
    // in-memory only, else conflicting resizes recalc forever (see Spill.cpp).
    // `name` must outlive the scope (it's the wrapper lambda's captured copy).
    class UdfNameScope
    {
    public:
        explicit UdfNameScope(const std::wstring& name);
        ~UdfNameScope();
        UdfNameScope(const UdfNameScope&) = delete;
        UdfNameScope& operator=(const UdfNameScope&) = delete;
    private:
        const std::wstring* _prev;
    };

    // Return an array result as the XLL return value, spilling appropriately for
    // the host Excel version.
    xloil::ExcelObj* output(xloil::ExcelObj&& result);
}
