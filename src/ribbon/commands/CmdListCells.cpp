// CmdListCells.cpp — "셀 내 목록 → 텍스트" 변환 (plan/27).
//
// M365 베타의 셀 내 목록/배열(리치 값 `_array`)은 구버전 Excel에서 #UNKNOWN!
// (2016류 #VALUE!)로만 보이고 COM·C API 어느 경로로도 항목을 읽을 수 없다
// (plan/26). 항목은 저장된 패키지의 xl/richData/rdarray.xml에만 있으므로,
// 통합문서 사본(.zip)을 풀어 vm → metadata → rdrichvalue → rdarray 체인을
// 직접 따라가 각 목록 셀을 core/ListText 표기로 바꿔 쓴다.
//
// 흐름: SaveCopyAs(임시 zip) → 해제 → 파싱 → 대상 집계 → 확인(예/아니요/취소:
// 수식 목록 셀 포함 여부) → 저장 게이트+백업(Convert::saveGateAndBackup) →
// 적용(PauseExcel) → 결과 안내. 셀의 현재 값이 오류일 때만 쓴다(저장 후
// 편집된 셀 보호). 수식 목록 셀은 udfFreeze처럼 원본 수식을 메모로 남긴다.

#include "Commands.h"
#include "CmdCommon.h"
#include "OoxmlPkg.h"
#include "../Convert.h"
#include "../../core/ListText.h"

#include <xlOil/AppObjects.h>
#include <xlOil/State.h>

#include <cstdlib>
#include <cwchar>
#include <filesystem>
#include <map>
#include <set>
#include <string>
#include <vector>

namespace egtools::commands
{
    namespace
    {
        using egtools::core::ListNode;
        namespace fs = std::filesystem;

        struct RichArray
        {
            size_t rows = 0, cols = 1;
            std::vector<std::pair<char, std::string>> vals;   // (t, 원문) — t: d s b e a r
        };

        struct Package
        {
            std::set<long> arrayStructs;                  // rdrichvaluestructure: t="_array" 인덱스
            std::vector<std::pair<long, long>> richValues; // rdrichvalue: (s, 첫 v)
            std::vector<RichArray> arrays;                // rdarray
            std::vector<long> rvbByFutureBk;              // metadata futureMetadata(XLRICHVALUE) bk → rvb i
            std::vector<std::pair<long, long>> valueMeta; // metadata valueMetadata bk → (rc t, rc v)
            long richTypeIndex = 0;                       // metadataTypes 중 XLRICHVALUE의 1-based 순번
        };

        // <tag …>…</tag> 요소들을 순회: fn(openStart, openEnd(> 위치), bodyEnd 또는 npos(자기 닫힘))
        template <class Fn>
        void forEachElement(const std::string& xml, const char* tagOpen, const char* tagClose, Fn fn)
        {
            size_t pos = 0;
            const std::string open(tagOpen);
            while (true)
            {
                const size_t a = xml.find(open, pos);
                if (a == std::string::npos) return;
                // "<a " 가 "<arrayData" 등에 걸리지 않도록 open은 공백/'>' 포함 형태로 준다
                const size_t b = xml.find('>', a);
                if (b == std::string::npos) return;
                const bool selfClosing = b > 0 && xml[b - 1] == '/';
                size_t end = std::string::npos;
                if (!selfClosing)
                {
                    end = xml.find(tagClose, b);
                    if (end == std::string::npos) return;
                }
                fn(a, b, end);
                pos = selfClosing ? b + 1 : end + 1;
            }
        }

        bool loadPackage(const fs::path& xl, Package& pk)
        {
            using namespace ooxml;
            // 구조: <s t="_array">…</s> 순번
            {
                const std::string s = readFileBytes(xl / L"richData" / L"rdrichvaluestructure.xml");
                if (s.empty()) return false;
                long idx = 0;
                forEachElement(s, "<s ", "</s>", [&](size_t a, size_t b, size_t)
                {
                    if (attrValue(s, a, b, "t") == "_array") pk.arrayStructs.insert(idx);
                    ++idx;
                });
                // 자기 닫힘 <s …/>도 순번에 넣어야 하므로 위 forEachElement는 둘 다 센다.
            }
            // 리치 값: <rv s="N"><v>idx</v>…</rv>
            {
                const std::string s = readFileBytes(xl / L"richData" / L"rdrichvalue.xml");
                if (s.empty()) return false;
                forEachElement(s, "<rv ", "</rv>", [&](size_t a, size_t b, size_t end)
                {
                    const long st = std::atol(attrValue(s, a, b, "s").c_str());
                    long first = -1;
                    if (end != std::string::npos)
                    {
                        const std::string v = tagBetween(s, b, end, "<v>", "</v>");
                        if (!v.empty()) first = std::atol(v.c_str());
                    }
                    pk.richValues.emplace_back(st, first);
                });
            }
            // 배열: <a r="R" c="C"><v t="s">…</v><v t="s"/>…</a>
            {
                const std::string s = readFileBytes(xl / L"richData" / L"rdarray.xml");
                if (s.empty()) return false;
                forEachElement(s, "<a ", "</a>", [&](size_t a, size_t b, size_t end)
                {
                    RichArray ra;
                    ra.rows = (size_t)std::atol(attrValue(s, a, b, "r").c_str());
                    const std::string c = attrValue(s, a, b, "c");
                    ra.cols = c.empty() ? 1 : (size_t)std::atol(c.c_str());
                    if (end != std::string::npos)
                    {
                        size_t p = b + 1;
                        while (true)
                        {
                            const size_t va = s.find("<v", p);
                            if (va == std::string::npos || va >= end) break;
                            const size_t vb = s.find('>', va);
                            if (vb == std::string::npos || vb >= end) break;
                            const std::string t = attrValue(s, va, vb, "t");
                            const char tc = t.empty() ? 'd' : t[0];
                            if (s[vb - 1] == '/')          // <v t="s"/> 빈 문자열
                            {
                                ra.vals.emplace_back(tc, "");
                                p = vb + 1;
                                continue;
                            }
                            const size_t ve = s.find("</v>", vb);
                            if (ve == std::string::npos || ve > end) break;
                            ra.vals.emplace_back(tc, unescapeXml(s.substr(vb + 1, ve - vb - 1)));
                            p = ve + 4;
                        }
                    }
                    pk.arrays.push_back(std::move(ra));
                });
            }
            // 메타데이터
            {
                const std::string s = readFileBytes(xl / L"metadata.xml");
                if (s.empty()) return false;
                long typeIdx = 0;
                forEachElement(s, "<metadataType ", "</metadataType>", [&](size_t a, size_t b, size_t)
                {
                    ++typeIdx;
                    if (attrValue(s, a, b, "name") == "XLRICHVALUE") pk.richTypeIndex = typeIdx;
                });
                if (pk.richTypeIndex == 0) return false;

                const size_t fm = s.find("<futureMetadata name=\"XLRICHVALUE\"");
                if (fm == std::string::npos) return false;
                const size_t fmEnd = s.find("</futureMetadata>", fm);
                size_t p = fm;
                while (true)
                {
                    const size_t a = s.find("<xlrd:rvb ", p);
                    if (a == std::string::npos || a >= fmEnd) break;
                    const size_t b = s.find('>', a);
                    pk.rvbByFutureBk.push_back(std::atol(attrValue(s, a, b, "i").c_str()));
                    p = b + 1;
                }

                const size_t vm = s.find("<valueMetadata");
                if (vm == std::string::npos) return false;
                const size_t vmEnd = s.find("</valueMetadata>", vm);
                p = vm;
                while (true)
                {
                    const size_t a = s.find("<rc ", p);
                    if (a == std::string::npos || a >= vmEnd) break;
                    const size_t b = s.find('>', a);
                    pk.valueMeta.emplace_back(std::atol(attrValue(s, a, b, "t").c_str()),
                                              std::atol(attrValue(s, a, b, "v").c_str()));
                    p = b + 1;
                }
            }
            return true;
        }

        // vm(1-based) → 배열 인덱스. 배열 리치 값이 아니면 -1.
        long arrayIndexForVm(const Package& pk, long vm)
        {
            if (vm < 1 || (size_t)vm > pk.valueMeta.size()) return -1;
            const auto [t, v] = pk.valueMeta[(size_t)vm - 1];
            if (t != pk.richTypeIndex) return -1;
            if (v < 0 || (size_t)v >= pk.rvbByFutureBk.size()) return -1;
            const long rv = pk.rvbByFutureBk[(size_t)v];
            if (rv < 0 || (size_t)rv >= pk.richValues.size()) return -1;
            const auto [st, first] = pk.richValues[(size_t)rv];
            if (!pk.arrayStructs.count(st)) return -1;
            if (first < 0 || (size_t)first >= pk.arrays.size()) return -1;
            return first;
        }

        ListNode buildNode(const Package& pk, long arrayIdx, int depth)
        {
            using namespace ooxml;
            if (depth > 32 || arrayIdx < 0 || (size_t)arrayIdx >= pk.arrays.size())
                return ListNode::scalar(xloil::ExcelObj(std::wstring_view(L"#UNKNOWN!")));
            const RichArray& ra = pk.arrays[(size_t)arrayIdx];
            std::vector<ListNode> items;
            items.reserve(ra.rows * ra.cols);
            for (size_t k = 0; k < ra.rows * ra.cols; ++k)
            {
                if (k >= ra.vals.size()) { items.push_back(ListNode::scalar(xloil::ExcelObj(xloil::CellError::NA))); continue; }
                const auto& [t, raw] = ra.vals[k];
                const std::wstring w = widen(raw);
                switch (t)
                {
                case 's': items.push_back(ListNode::scalar(xloil::ExcelObj(std::wstring_view(w)))); break;
                case 'b': items.push_back(ListNode::scalar(xloil::ExcelObj(w == L"1" || w == L"true"))); break;
                case 'a': items.push_back(buildNode(pk, std::atol(raw.c_str()), depth + 1)); break;
                case 'r':
                {
                    const long rv = std::atol(raw.c_str());
                    long arr = -1;
                    if (rv >= 0 && (size_t)rv < pk.richValues.size())
                    {
                        const auto [st, first] = pk.richValues[(size_t)rv];
                        if (pk.arrayStructs.count(st)) arr = first;
                    }
                    if (arr >= 0) items.push_back(buildNode(pk, arr, depth + 1));
                    else items.push_back(ListNode::scalar(xloil::ExcelObj(std::wstring_view(L"#UNKNOWN!"))));
                    break;
                }
                default:   // 'd' 숫자, 'e' 오류 — 토큰 분류(오류 이름→오류, 숫자→숫자)
                    items.push_back(ListNode::scalar(egtools::core::scalarFromToken(w)));
                    break;
                }
            }
            return ListNode::array(ra.rows, ra.cols, std::move(items));
        }

        struct Target
        {
            std::wstring sheet;
            std::wstring ref;
            bool formula = false;
            std::wstring text;
            xloil::ExcelObj scalar;   // 항목 하나짜리 숫자/논리 목록 → 값 자체로 쓴다(왕복 보존)
        };

        // 시트 XML에서 vm 셀을 모은다.
        void scanSheet(const Package& pk, const std::string& xml, const std::wstring& sheetName,
                       std::vector<Target>& out)
        {
            using namespace ooxml;
            forEachElement(xml, "<c ", "</c>", [&](size_t a, size_t b, size_t end)
            {
                const std::string vm = attrValue(xml, a, b, "vm");
                if (vm.empty()) return;
                const long arr = arrayIndexForVm(pk, std::atol(vm.c_str()));
                if (arr < 0) return;
                Target tg;
                tg.sheet = sheetName;
                tg.ref = widen(attrValue(xml, a, b, "r"));
                tg.formula = end != std::string::npos && xml.find("<f", b) < end;
                const ListNode node = buildNode(pk, arr, 0);
                tg.text = egtools::core::formatListText(node);
                if (node.isArray && node.rows == 1 && node.cols == 1 && !node.items[0].isArray)
                {
                    const xloil::ExcelObj& v = node.items[0].value;
                    if (v.isType(xloil::ExcelType::Num) || v.isType(xloil::ExcelType::Bool))
                        tg.scalar = v;
                }
                out.push_back(std::move(tg));
            });
        }

        // 값 쓰기 시 Excel의 자동 해석(=수식, TRUE/숫자 변환)을 막아야 하는 텍스트인가.
        bool needsTextFormat(const std::wstring& text)
        {
            if (text.empty()) return false;
            const wchar_t c = text[0];
            if (c == L'=' || c == L'+' || c == L'-' || c == L'@' || c == L'\'') return true;
            const xloil::ExcelObj v = egtools::core::scalarFromToken(text);
            return !v.isType(xloil::ExcelType::Str);
        }

        std::wstring fmtCount(const std::wstring& pattern, int a, int b = 0)
        {
            wchar_t buf[1024];
            swprintf_s(buf, pattern.c_str(), a, b);
            return buf;
        }
    }

    void listCellsToText()
    {
        using egtools::i18n::t;
        fs::path tempRoot;
        auto cleanup = [&]()
        {
            if (tempRoot.empty()) return;
            std::error_code ec;
            fs::remove_all(tempRoot, ec);
        };
        try
        {
            IDispatch* ad = appDisp();
            if (!ad) return;
            Releaser rApp{ ad };
            IDispatch* wb = getObject(ad, L"ActiveWorkbook");
            if (!wb) { msgInfo(t(L"cmd.listcells.none")); return; }
            Releaser rWb{ wb };

            // 1) 현재 상태의 사본(.zip) — SaveCopyAs는 미저장 편집도 담는다.
            wchar_t tempBuf[MAX_PATH]{};
            GetTempPathW(MAX_PATH, tempBuf);
            tempRoot = fs::path(tempBuf) / (L"egtools_list_" + std::to_wstring(GetTickCount64()));
            fs::create_directories(tempRoot / L"x");
            const fs::path zipPath = tempRoot / L"book.zip";
            {
                statusBar(ad, L"reading list cells...");
                VARIANT a = varBStr(zipPath.wstring());
                const bool ok = invokeRaw(wb, L"SaveCopyAs", DISPATCH_METHOD, nullptr, &a, 1);
                VariantClear(&a);
                if (!ok) throw std::runtime_error("SaveCopyAs");
            }
            if (!ooxml::shellUnzip(zipPath.wstring(), (tempRoot / L"x").wstring()))
                throw std::runtime_error("unzip");
            const fs::path xl = tempRoot / L"x" / L"xl";

            // 2) 리치 배열 파트가 없으면 목록 셀 없음
            if (ooxml::relTarget(xl / L"_rels" / L"workbook.xml.rels", "", "/rdArray").empty())
            {
                statusBarClear(ad);
                cleanup();
                msgInfo(t(L"cmd.listcells.none"));
                return;
            }
            Package pk;
            if (!loadPackage(xl, pk)) throw std::runtime_error("richData");

            // 3) 시트별 vm 셀 수집
            std::vector<Target> targets;
            {
                const std::string wbXml = ooxml::readFileBytes(xl / L"workbook.xml");
                size_t pos = 0;
                while (true)
                {
                    const size_t a = wbXml.find("<sheet ", pos);
                    if (a == std::string::npos) break;
                    const size_t b = wbXml.find('>', a);
                    if (b == std::string::npos) break;
                    pos = b;
                    const std::wstring name =
                        ooxml::widen(ooxml::unescapeXml(ooxml::attrValue(wbXml, a, b, "name")));
                    const std::string rid = ooxml::attrValue(wbXml, a, b, "r:id");
                    const std::string file =
                        ooxml::relTarget(xl / L"_rels" / L"workbook.xml.rels", rid, nullptr);
                    if (file.empty()) continue;
                    fs::path sheetPath = (file.size() > 1 && file[0] == '/')
                        ? tempRoot / L"x" / ooxml::widen(file.substr(1))
                        : xl / ooxml::widen(file);
                    const std::string sx = ooxml::readFileBytes(sheetPath);
                    if (sx.find("vm=\"") == std::string::npos) continue;
                    scanSheet(pk, sx, name, targets);
                }
            }
            statusBarClear(ad);
            cleanup();

            int nValue = 0, nFormula = 0;
            for (const auto& tg : targets) (tg.formula ? nFormula : nValue)++;
            if (targets.empty()) { msgInfo(t(L"cmd.listcells.none")); return; }

            // 4) 확인 — 수식 목록 셀 포함 여부
            bool includeFormula = false;
            if (nFormula > 0)
            {
                const int r = MessageBoxW(excelHwnd(),
                    fmtCount(t(L"cmd.listcells.confirmFormula"), nValue, nFormula).c_str(),
                    L"EGTools++", MB_YESNOCANCEL | MB_ICONQUESTION);
                if (r == IDCANCEL) return;
                includeFormula = (r == IDYES);
                if (!includeFormula && nValue == 0) return;
            }
            else if (MessageBoxW(excelHwnd(), fmtCount(t(L"cmd.listcells.confirm"), nValue).c_str(),
                                 L"EGTools++", MB_YESNO | MB_ICONQUESTION) != IDYES)
                return;

            // 5) 저장 게이트 + 백업(Convert와 동일 절차)
            std::wstring backupPath;
            if (!egtools::ribbon::detail::saveGateAndBackup(
                    t(L"ribbon.egpp.listToText.label").c_str(), backupPath))
                return;

            // 6) 적용
            long done = 0, skipped = 0;
            {
                auto& app = xloil::thisApp();
                xloil::PauseExcel pause(app);
                IDispatch* sheets = getObject(wb, L"Worksheets");
                if (!sheets) return;
                Releaser rSheets{ sheets };

                for (const auto& tg : targets)
                {
                    if (tg.formula && !includeFormula) continue;
                    VARIANT sn = varBStr(tg.sheet);
                    IDispatch* sh = getObject(sheets, L"Item", &sn, 1);
                    VariantClear(&sn);
                    if (!sh) { ++skipped; continue; }
                    Releaser rSh{ sh };
                    VARIANT rf = varBStr(tg.ref);
                    IDispatch* cell = getObject(sh, L"Range", &rf, 1);
                    VariantClear(&rf);
                    if (!cell) { ++skipped; continue; }
                    Releaser rc{ cell };

                    // 안전 검사: 현재 값이 오류(365 #UNKNOWN!/2016 #VALUE!)인 셀만
                    VARIANT cur; VariantInit(&cur);
                    getVar(cell, L"Value2", &cur);
                    const bool isErr = cur.vt == VT_ERROR;
                    VariantClear(&cur);
                    if (!isErr) { ++skipped; continue; }

                    std::wstring formula;
                    if (tg.formula)
                    {
                        formula = getBStr(cell, L"Formula");
                        // CSE 배열 영역은 udfFreeze처럼 영역 전체를 값으로 고정
                        if (getBool(cell, L"HasArray", false))
                        {
                            IDispatch* arr = getObject(cell, L"CurrentArray");
                            if (arr)
                            {
                                Releaser ra{ arr };
                                VARIANT vals; VariantInit(&vals);
                                if (getVar(arr, L"Value2", &vals)) putVar(arr, L"Value", vals);
                                VariantClear(&vals);
                            }
                        }
                    }

                    bool ok = false;
                    if (tg.scalar.isType(xloil::ExcelType::Num))
                        ok = putDouble(cell, L"Value2", tg.scalar.get<double>());
                    else if (tg.scalar.isType(xloil::ExcelType::Bool))
                        ok = putBool(cell, L"Value2", tg.scalar.get<bool>());
                    else
                    {
                        if (needsTextFormat(tg.text)) putBStr(cell, L"NumberFormat", L"@");
                        ok = putBStr(cell, L"Value2", tg.text);
                    }
                    if (!ok) { ++skipped; continue; }
                    ++done;

                    if (!formula.empty())
                    {
                        std::wstring note = L"EgListRemoved: " + formula;
                        IDispatch* cm = getObject(cell, L"Comment");
                        if (cm)
                        {
                            Releaser rcm{ cm };
                            VARIANT tr; VariantInit(&tr);
                            if (invokeRaw(cm, L"Text", DISPATCH_METHOD, &tr, nullptr, 0) &&
                                tr.vt == VT_BSTR && tr.bstrVal)
                                note += L"\n" + std::wstring(tr.bstrVal);
                            VariantClear(&tr);
                            callMethod(cm, L"Delete");
                        }
                        VARIANT na = varBStr(note);
                        callMethod(cell, L"AddComment", &na, 1);
                        VariantClear(&na);
                    }
                }
            }

            std::wstring msg = withCount(t(L"cmd.listcells.result"), done);
            if (skipped > 0) msg += L"\n" + withCount(t(L"cmd.listcells.skipped"), skipped);
            msg += L"\n\n" + backupPath;
            msgInfo(msg);
        }
        catch (...)
        {
            cleanup();
            msgWarn(t(L"cmd.listcells.fail"));
        }
    }
}
