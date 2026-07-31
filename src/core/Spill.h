// Spill.h — array output that adapts to the host Excel.
//
// The function-authoring principle (revised D5): new functions take arrays and
// return arrays. output() returns the array result so it spills:
//   * modern Excel (2021/365): just return — Excel spills natively.
//   * legacy Excel (<=2019): the forced array-formula (CSE) resize engine.
// See plan/04 §4 and plan/08 §4.2.

#pragma once
#include <xlOil/ExcelObj.h>

namespace egtools::core
{
    // Return an array result as the XLL return value, spilling appropriately for
    // the host Excel version.
    xloil::ExcelObj* output(xloil::ExcelObj&& result);
}
