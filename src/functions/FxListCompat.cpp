// FxListCompat.cpp — M365 베타 "Lists in Cells" 호환 함수 (plan/27).
//   HAS(array, value) / HASANY(array, values) / HASALL(array, values) /
//   FLATTEN(array, [pad_value], [levels])
//
// 구버전 Excel은 목록 셀 자체를 읽지 못하므로(C API로 #VALUE!만 옴 — plan/26)
// 이 함수들은 변환 명령이 만든 텍스트 표기(core/ListText)와 일반 콤마 구분
// 텍스트, 범위, 배열상수를 받는다. 의미는 Insider 블로그 정의를 따른다:
//   HAS    — value가 array 어디든(중첩 포함) 있으면 TRUE.
//   HASANY — values 중 하나라도 있으면 TRUE.   HASALL — 모두 있으면 TRUE.
//   FLATTEN — 중첩을 한 단계씩 풀어 스필: 외곽 배열의 각 행에서 원소(스칼라=
//             1×1, 배열=그 모양)를 가로로 잇고 행들을 세로로 쌓되 폭은
//             pad_value(기본 #N/A)로 맞춘다. levels 기본=전부(0도 전부).
// 리프팅(plan/22): HAS의 value만 원소별, array/values는 통째로(집합).

#include "../core/Registry.h"
#include "../core/Apply.h"
#include "../core/Spill.h"
#include "../core/ArrayUtil.h"
#include "../core/ListText.h"

#include <xlOil/xlOil.h>
#include <xlOil/ExcelArray.h>
#include <vector>

using namespace xloil;
using egtools::core::ListNode;

namespace egtools::functions
{
    namespace
    {
        // FLATTEN의 외곽 배열: 범위/배열은 그 자체가 외곽(각 셀이 원소), 스칼라
        // 인수(목록 텍스트 한 개 포함)는 "셀 하나"이므로 1×1 외곽으로 감싼다 —
        // levels=1이 그 셀의 첫 중첩만 푸는 네이티브 모델과 맞추기 위함.
        ListNode outerArray(const ExcelObj& v)
        {
            ListNode n = egtools::core::listFromValue(v);
            if (v.isType(ExcelType::Multi)) return n;
            std::vector<ListNode> one;
            one.push_back(std::move(n));
            return ListNode::array(1, 1, std::move(one));
        }

        bool containsRec(const ListNode& n, const ExcelObj& v)
        {
            if (!n.isArray) return egtools::core::listItemEquals(n.value, v);
            for (const auto& it : n.items)
                if (containsRec(it, v)) return true;
            return false;
        }

        void leaves(const ListNode& n, std::vector<const ExcelObj*>& out)
        {
            if (!n.isArray)
            {
                if (!n.value.isType(ExcelType::Missing)) out.push_back(&n.value);
                return;
            }
            for (const auto& it : n.items) leaves(it, out);
        }

        // 한 단계 평탄화(파일 머리말 참조).
        ListNode flattenOnce(const ListNode& n, const ExcelObj& pad)
        {
            std::vector<std::vector<ListNode>> rows;   // 결과 행들(폭 가변)
            size_t width = 0;
            for (size_t r = 0; r < n.rows; ++r)
            {
                // 이 행의 블록들: (rows, cols, items)
                size_t h = 1, w = 0;
                struct Blk { size_t rows, cols; const ListNode* node; };
                std::vector<Blk> blocks;
                for (size_t c = 0; c < n.cols; ++c)
                {
                    const ListNode& child = n.at(r, c);
                    if (child.isArray)
                    {
                        blocks.push_back({ child.rows, child.cols, &child });
                        h = std::max(h, child.rows);
                        w += child.cols;
                    }
                    else { blocks.push_back({ 1, 1, &child }); w += 1; }
                }
                std::vector<std::vector<ListNode>> blockRows(h);
                for (auto& br : blockRows) br.reserve(w);
                for (const auto& b : blocks)
                {
                    for (size_t i = 0; i < h; ++i)
                        for (size_t j = 0; j < b.cols; ++j)
                        {
                            if (i < b.rows)
                                blockRows[i].push_back(b.node->isArray ? b.node->at(i, j) : *b.node);
                            else
                                blockRows[i].push_back(ListNode::scalar(ExcelObj(pad)));
                        }
                }
                for (auto& br : blockRows) { width = std::max(width, br.size()); rows.push_back(std::move(br)); }
            }
            std::vector<ListNode> items;
            items.reserve(rows.size() * width);
            for (auto& r : rows)
                for (size_t c = 0; c < width; ++c)
                    items.push_back(c < r.size() ? std::move(r[c]) : ListNode::scalar(ExcelObj(pad)));
            return ListNode::array(rows.size(), width, std::move(items));
        }

        // values 인수(범위/배열상수/목록 텍스트)의 리프 값들. 스칼라 오류면 nullopt
        // (호출자가 전파).
        std::optional<std::vector<ExcelObj>> valueSet(const ExcelObj& values, ListNode& keep)
        {
            if (values.isType(ExcelType::Err)) return std::nullopt;
            keep = egtools::core::listFromValue(values);
            std::vector<const ExcelObj*> ptrs;
            leaves(keep, ptrs);
            std::vector<ExcelObj> out;
            out.reserve(ptrs.size());
            for (auto p : ptrs) out.emplace_back(*p);
            return out;
        }
    }

    void registerListCompat()
    {
        // HAS(array, value) — value는 원소별 리프팅, array는 통째로 캡처.
        egtools::core::registerFn(L"HAS",
            [](const ExcelObj& array, const ExcelObj& value) -> ExcelObj*
            {
                if (array.isType(ExcelType::Err)) return returnValue(ExcelObj(array));
                if (value.isMissing()) return returnValue(CellError::Value);
                const ListNode tree = egtools::core::listFromValue(array);
                return egtools::core::mapLift([&](const ExcelObj& v) -> ExcelObj
                {
                    return ExcelObj(containsRec(tree, v));
                }, value);
            });

        // HASANY(array, values) / HASALL(array, values)
        auto hasSet = [](bool all)
        {
            return [all](const ExcelObj& array, const ExcelObj& values) -> ExcelObj*
            {
                if (array.isType(ExcelType::Err)) return returnValue(ExcelObj(array));
                if (values.isMissing()) return returnValue(CellError::Value);
                ListNode keep;
                auto vals = valueSet(values, keep);
                if (!vals) return returnValue(ExcelObj(values));
                const ListNode tree = egtools::core::listFromValue(array);
                if (all)
                {
                    for (const auto& v : *vals)
                        if (!containsRec(tree, v)) return returnValue(false);
                    return returnValue(true);
                }
                for (const auto& v : *vals)
                    if (containsRec(tree, v)) return returnValue(true);
                return returnValue(false);
            };
        };
        egtools::core::registerFn(L"HASANY", hasSet(false));
        egtools::core::registerFn(L"HASALL", hasSet(true));

        // FLATTEN(array, [pad_value], [levels])
        egtools::core::registerFn(L"FLATTEN",
            [](const ExcelObj& array, const ExcelObj& padA, const ExcelObj& levelsA) -> ExcelObj*
            {
                if (array.isType(ExcelType::Err)) return returnValue(ExcelObj(array));
                if (padA.isType(ExcelType::Multi)) return returnValue(CellError::Value);
                const ExcelObj pad = padA.isMissing() ? ExcelObj(CellError::NA) : ExcelObj(padA);
                const auto lv = egtools::core::optScalar<double>(levelsA, 0.0);
                if (!lv || *lv < 0) return returnValue(CellError::Value);
                const long levels = (long)*lv;   // 0 = 전부

                ListNode node = outerArray(array);
                for (long k = 0; (levels == 0 || k < levels) && node.hasNested(); ++k)
                    node = flattenOnce(node, pad);
                if (node.rows == 0 || node.cols == 0) return returnValue(CellError::Value);

                std::vector<ExcelObj> vals;
                vals.reserve(node.items.size());
                for (const auto& it : node.items)
                    vals.emplace_back(it.isArray
                        ? ExcelObj(std::wstring_view(egtools::core::formatListText(it)))
                        : ExcelObj(it.value));
                return egtools::core::output(egtools::core::makeArray(
                    (ExcelArrayBuilder::row_t)node.rows,
                    (ExcelArrayBuilder::col_t)node.cols, vals));
            });
    }
}
