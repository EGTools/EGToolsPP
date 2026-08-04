// CmdUtil.cpp — 도구 명령 (EGToolsVB C09_Util.vb 포팅, plan/23).
// M2: ApplyCheckBox / RecalcAll / OpenUrlInCell / RearrangeMemo /
//     DeleteNames / ChangeLineColors.
// M3: DeleteStyles / ClearError / RemoveZeroLengthString / EgUdfRemove(값 고정).
// 모양 뽑기(M4)는 이후.

#include "Commands.h"
#include "CmdCommon.h"
#include "Dialogs.h"
#include "../Convert.h"
#include "../../core/Version.h"

#include <shellapi.h>
#include <map>
#include <regex>
#include <set>
#include <vector>

namespace egtools::commands
{
    // CheckBox 셀서식 적용 (VB ApplyCheckBox, C09_Util.vb:293)
    // 0/1 유효성 검사 + [=1]✅/[=0]⬜ 표시 형식. 빈 셀은 0으로 채움.
    void applyCheckBox()
    {
        try
        {
            auto& app = xloil::thisApp();
            xloil::PauseExcel pause(app);

            IDispatch* ad = appDisp();
            if (!ad) return;
            Releaser rApp{ ad };
            IDispatch* sel = selectionRange(ad);
            if (!sel) return;
            Releaser rSel{ sel };

            IDispatch* valid = getObject(sel, L"Validation");
            if (valid)
            {
                Releaser rv{ valid };
                callMethod(valid, L"Delete");
                // Add(xlValidateWholeNumber=1, xlValidAlertStop=1, xlBetween=1, "0", "1")
                VARIANT a[5];
                a[0] = varLong(1); a[1] = varLong(1); a[2] = varLong(1);
                a[3] = varBStr(L"0"); a[4] = varBStr(L"1");
                callMethod(valid, L"Add", a, 5);
                VariantClear(&a[3]); VariantClear(&a[4]);
            }
            putBStr(sel, L"NumberFormat", L"[=1]\"✅\";[=0]\"⬜\";\"\"");

            // 빈 셀 → 0 + 가운데 정렬 (VB의 per-cell 루프와 동등:
            // 단일 셀은 직접, 그 외에는 SpecialCells(xlCellTypeBlanks) 일괄)
            const double cellCount = [&] {
                IDispatch* cells = getObject(sel, L"Cells");
                if (!cells) return 0.0;
                Releaser rc{ cells };
                return getDouble(cells, L"CountLarge", 0);
            }();
            if (cellCount == 1.0)
            {
                VARIANT v; VariantInit(&v);
                getVar(sel, L"Value2", &v);
                const bool empty = isEmptyVal(v);
                VariantClear(&v);
                if (empty)
                {
                    putLong(sel, L"Value", 0);
                    putLong(sel, L"HorizontalAlignment", -4108);  // xlHAlignCenter
                }
            }
            else if (cellCount > 1.0)
            {
                // SpecialCells는 대상이 없으면 실패 — 조용히 무시
                VARIANT a = varLong(4);                           // xlCellTypeBlanks
                VARIANT r; VariantInit(&r);
                if (invokeRaw(sel, L"SpecialCells",
                              DISPATCH_METHOD | DISPATCH_PROPERTYGET, &r, &a, 1) &&
                    r.vt == VT_DISPATCH && r.pdispVal)
                {
                    IDispatch* blanks = r.pdispVal;
                    putLong(blanks, L"Value", 0);
                    putLong(blanks, L"HorizontalAlignment", -4108);
                }
                VariantClear(&r);
            }
        }
        catch (...) {}
    }

    // 전체 재계산 (VB RecalcuateAllWB, C09_Util.vb:700)
    void recalcAll()
    {
        try
        {
            IDispatch* ad = appDisp();
            if (!ad) return;
            Releaser rApp{ ad };
            callMethod(ad, L"CalculateFull");
        }
        catch (...) {}
    }

    // 셀 내용의 URL 열기 (VB OpenUrlInCell, C09_Util.vb:319) — 컨텍스트 메뉴 포함
    void openUrlInCell()
    {
        using egtools::i18n::t;
        try
        {
            IDispatch* ad = appDisp();
            if (!ad) return;
            Releaser rApp{ ad };
            IDispatch* sel = selectionRange(ad);
            if (!sel)
            {
                msgWarn(t(L"cmd.openurl.single"));
                return;
            }
            Releaser rSel{ sel };
            IDispatch* cell = getObjectIdx2(sel, L"Cells", 1, 1);
            if (!cell) return;
            Releaser rc{ cell };

            VARIANT v; VariantInit(&v);
            getVar(cell, L"Value", &v);
            const std::wstring text = varToString(v);
            VariantClear(&v);
            if (text.empty()) return;

            static const std::wregex urlRe(LR"(https?://[^\s]+)");
            std::wsmatch m;
            if (std::regex_search(text, m, urlRe))
                ShellExecuteW(nullptr, L"open", m.str(0).c_str(), nullptr, nullptr,
                              SW_SHOWNORMAL);
        }
        catch (...) {}
    }

    // 메모 정돈 (VB RearrangeMemo, C09_Util.vb:243)
    // 모든 메모(Note)를 셀 오른쪽 바로 옆으로 이동 + 자동 크기.
    void rearrangeMemo()
    {
        using egtools::i18n::t;
        long count = 0;
        try
        {
            auto& app = xloil::thisApp();
            xloil::PauseExcel pause(app);

            IDispatch* ad = appDisp();
            if (!ad) return;
            Releaser rApp{ ad };
            IDispatch* sh = getObject(ad, L"ActiveSheet");
            if (!sh) return;
            Releaser rSh{ sh };
            IDispatch* comments = getObject(sh, L"Comments");
            if (!comments) return;
            Releaser rCm{ comments };

            const long n = getLong(comments, L"Count", 0);
            for (long i = 1; i <= n; ++i)
            {
                IDispatch* cm = getObjectIdx(comments, L"Item", i);
                if (!cm) continue;
                Releaser rItem{ cm };
                IDispatch* cell = getObject(cm, L"Parent");   // Comment.Parent = 셀
                if (!cell) continue;
                Releaser rCell{ cell };
                IDispatch* ma = getObject(cell, L"MergeArea");
                Releaser rMa{ ma };
                IDispatch* shape = getObject(cm, L"Shape");
                if (!shape) continue;
                Releaser rShp{ shape };

                putDouble(shape, L"Left",
                          getDouble(cell, L"Left") +
                          (ma ? getDouble(ma, L"Width") : getDouble(cell, L"Width")) + 2);
                putDouble(shape, L"Top", getDouble(cell, L"Top") + 2);
                putLong(shape, L"Placement", 2);              // xlMove
                if (IDispatch* tf = getObject(shape, L"TextFrame"))
                {
                    Releaser rtf{ tf };
                    putBool(tf, L"AutoSize", true);
                }
                ++count;
            }
            msgInfo(count > 0 ? withCount(t(L"cmd.memo.result"), count)
                              : t(L"cmd.memo.none"));
        }
        catch (...) {}
    }

    // 오류/숨김 이름 삭제 (VB DeleteNames, C09_Util.vb:75)
    void deleteNames()
    {
        using egtools::i18n::t;
        long count = 0;
        try
        {
            auto& app = xloil::thisApp();
            xloil::PauseExcel pause(app);

            IDispatch* ad = appDisp();
            if (!ad) return;
            Releaser rApp{ ad };
            IDispatch* wb = getObject(ad, L"ActiveWorkbook");
            if (!wb) return;
            Releaser rWb{ wb };
            IDispatch* names = getObject(wb, L"Names");
            if (!names) return;
            Releaser rNames{ names };

            const long total = getLong(names, L"Count", 0);
            for (long i = total; i >= 1; --i)
            {
                if (i % 100 == 0)
                    statusBar(ad, std::to_wstring(i) + L"/" + std::to_wstring(total) +
                                  L" processing...");
                IDispatch* item = getObjectIdx(names, L"Item", i);
                if (!item) continue;
                Releaser rItem{ item };

                // 내부 이름(_xl*.*)은 절대 건드리지 않음(VB 동일)
                const std::wstring name = getBStr(item, L"Name");
                if (name.compare(0, 3, L"_xl") == 0 && name.find(L'.') != std::wstring::npos)
                    continue;

                const std::wstring refersTo = getBStr(item, L"RefersTo");
                const bool broken =
                    refersTo.find(L"#NAME?") != std::wstring::npos ||
                    refersTo.find(L"#REF!") != std::wstring::npos ||
                    refersTo.find(L"#N/A") != std::wstring::npos ||
                    refersTo.find(L"BlankMacro1") != std::wstring::npos ||
                    !getBool(item, L"Visible", true);
                if (broken && callMethod(item, L"Delete"))
                    ++count;
            }
            statusBarClear(ad);
            msgInfo(count > 0 ? withCount(t(L"cmd.names.result"), count)
                              : t(L"cmd.names.none"));
        }
        catch (...) {}
    }

    // 테두리 색 일괄 변경 (VB ChangeLineColors, C09_Util.vb:274)
    // 선 종류는 유지하고 색만 바꾼다.
    void changeLineColors()
    {
        try
        {
            COLORREF color = RGB(0, 0, 0);
            if (!egtools::dialogs::pickColor(excelHwnd(), color)) return;

            auto& app = xloil::thisApp();
            xloil::PauseExcel pause(app);

            IDispatch* ad = appDisp();
            if (!ad) return;
            Releaser rApp{ ad };
            IDispatch* sel = selectionRange(ad);
            if (!sel) return;
            Releaser rSel{ sel };

            // xlEdgeTop=8, xlEdgeBottom=9, xlEdgeLeft=7, xlEdgeRight=10,
            // xlDiagonalDown=5, xlDiagonalUp=6 (VB와 동일 집합)
            static const long borders[] = { 8, 9, 7, 10, 5, 6 };
            constexpr long xlLineStyleNone = -4142;

            forEachCell(sel, [&](IDispatch* cell)
            {
                for (long idx : borders)
                {
                    IDispatch* b = getObjectIdx(cell, L"Borders", idx);
                    if (!b) continue;
                    Releaser rb{ b };
                    if (getLong(b, L"LineStyle", xlLineStyleNone) != xlLineStyleNone)
                        putLong(b, L"Color", (long)color);
                }
                return true;
            });
        }
        catch (...) {}
    }

    // 스타일 삭제 (VB DeleteStyles, C09_Util.vb:11)
    // 예=BuiltIn만 남김 / 아니오=BuiltIn+사용중 유지 / 취소.
    void deleteStyles()
    {
        using egtools::i18n::t;
        try
        {
            auto& app = xloil::thisApp();
            xloil::PauseExcel pause(app);

            IDispatch* ad = appDisp();
            if (!ad) return;
            Releaser rApp{ ad };
            IDispatch* wb = getObject(ad, L"ActiveWorkbook");
            if (!wb) return;
            Releaser rWb{ wb };

            bool removeAll = false;
            switch (MessageBoxW(excelHwnd(), t(L"cmd.styles.confirm").c_str(),
                                L"EGTools++",
                                MB_YESNOCANCEL | MB_ICONINFORMATION | MB_DEFBUTTON1))
            {
            case IDYES: removeAll = true; break;
            case IDNO:  removeAll = false; break;
            default:    return;
            }

            // "사용중 유지" 모드: 모든 시트 UsedRange를 셀 단위로 훑어 사용
            // 스타일 이름 수집(VB와 동일 — 대형 문서에서는 느림, StatusBar 표시).
            std::set<std::wstring> used;
            if (!removeAll)
            {
                IDispatch* sheets = getObject(wb, L"Sheets");
                if (sheets)
                {
                    Releaser rs{ sheets };
                    const long nSheets = getLong(sheets, L"Count", 0);
                    for (long s = 1; s <= nSheets; ++s)
                    {
                        IDispatch* sh = getObjectIdx(sheets, L"Item", s);
                        if (!sh) continue;
                        Releaser rSh{ sh };
                        statusBar(ad, getBStr(sh, L"Name") + L" scanning styles...");
                        IDispatch* ur = getObject(sh, L"UsedRange");
                        if (!ur) continue;
                        Releaser rUr{ ur };
                        IDispatch* cells = getObject(ur, L"Cells");
                        if (!cells) continue;
                        Releaser rc{ cells };
                        const long n = (long)getDouble(cells, L"CountLarge", 0);
                        for (long i = 1; i <= n; ++i)
                        {
                            IDispatch* cell = getObjectIdx(cells, L"Item", i);
                            if (!cell) continue;
                            Releaser rCell{ cell };
                            IDispatch* st = getObject(cell, L"Style");
                            if (!st) continue;
                            Releaser rSt{ st };
                            used.insert(getBStr(st, L"Name"));
                        }
                    }
                }
            }

            IDispatch* styles = getObject(wb, L"Styles");
            if (!styles) return;
            Releaser rStyles{ styles };
            const long total = getLong(styles, L"Count", 0);
            long deleted = 0, failed = 0;
            for (long i = total; i >= 1; --i)
            {
                if (i % 100 == 0)
                    statusBar(ad, std::to_wstring(i) + L"/" + std::to_wstring(total) +
                                  L" styles...");
                IDispatch* st = getObjectIdx(styles, L"Item", i);
                if (!st) continue;
                Releaser rSt{ st };
                if (getBool(st, L"BuiltIn", true)) continue;
                if (!removeAll && used.count(getBStr(st, L"Name"))) continue;
                putBool(st, L"Locked", false);
                if (callMethod(st, L"Delete")) ++deleted;
                else ++failed;
            }
            statusBarClear(ad);

            if (deleted > 0 || failed > 0)
            {
                std::wstring msg = withCount(t(L"cmd.styles.result"), deleted);
                if (failed > 0)
                    msg += L"\n" + withCount(t(L"cmd.styles.fail"), failed);
                msgInfo(msg);
            }
            else
                msgInfo(t(L"cmd.styles.none"));
        }
        catch (...) {}
    }

    // 오류 정리 (VB ClearError, C09_Util.vb:116)
    // 활성 시트 UsedRange에서 오류값을 내는 수식을 =IFERROR(수식,"")로 감싼다.
    // COM Value2 배열은 오류 셀을 VT_ERROR로 돌려주므로 그대로 판별한다
    // (VB의 ExcelReference 우회는 .NET 마샬링 문제 — C++에는 불필요).
    void clearErrors()
    {
        using egtools::i18n::t;
        long count = 0;
        try
        {
            auto& app = xloil::thisApp();
            xloil::PauseExcel pause(app);

            IDispatch* ad = appDisp();
            if (!ad) return;
            Releaser rApp{ ad };
            IDispatch* sh = getObject(ad, L"ActiveSheet");
            if (!sh) return;
            Releaser rSh{ sh };
            IDispatch* ur = getObject(sh, L"UsedRange");
            if (!ur) return;
            Releaser rUr{ ur };

            auto wrapFormula = [](const std::wstring& f) -> std::wstring
            {
                if (f.size() < 2 || f[0] != L'=') return L"";
                return L"=IFERROR(" + f.substr(1) + L",\"\")";
            };

            SafeArr2D values;
            if (values.load(ur, L"Value2"))
            {
                for (long r = values.r1; r <= values.r2; ++r)
                    for (long c = values.c1; c <= values.c2; ++c)
                    {
                        if (values.at(r, c).vt != VT_ERROR) continue;
                        IDispatch* cell = getObjectIdx2(ur, L"Cells", r, c);
                        if (!cell) continue;
                        Releaser rc{ cell };
                        if (!getBool(cell, L"HasFormula", false)) continue;

                        if (getBool(cell, L"HasArray", false))
                        {
                            IDispatch* arr = getObject(cell, L"CurrentArray");
                            if (!arr) continue;
                            Releaser ra{ arr };
                            IDispatch* first = getObjectIdx2(arr, L"Cells", 1, 1);
                            if (!first) continue;
                            Releaser rf{ first };
                            const std::wstring nf =
                                wrapFormula(getBStr(first, L"FormulaArray"));
                            if (!nf.empty() && putBStr(arr, L"FormulaArray", nf))
                                ++count;
                        }
                        else
                        {
                            const std::wstring nf =
                                wrapFormula(getBStr(cell, L"Formula"));
                            if (!nf.empty() && putBStr(cell, L"Formula", nf))
                                ++count;
                        }
                    }
            }
            else
            {
                // 단일 셀 UsedRange(스칼라) — 셀 하나만 검사
                VARIANT sv; VariantInit(&sv);
                getVar(ur, L"Value2", &sv);
                const bool isErr = sv.vt == VT_ERROR;
                VariantClear(&sv);
                if (isErr && getBool(ur, L"HasFormula", false))
                {
                    const std::wstring nf = wrapFormula(getBStr(ur, L"Formula"));
                    if (!nf.empty() && putBStr(ur, L"Formula", nf)) ++count;
                }
            }

            msgInfo(count > 0 ? withCount(t(L"cmd.clearerr.result"), count)
                              : t(L"cmd.clearerr.none"));
        }
        catch (...) {}
    }

    // 빈셀 정리 (VB RemoveZeroLengthString, C09_Util.vb:173)
    // 길이 0 문자열("")만 든 셀을 진짜 빈 셀로 만든다. 단일 셀 선택이면
    // UsedRange 전체가 대상. "수식은 제외할까요?" 예=수식 셀 보존.
    void removeEmptyStrings()
    {
        using egtools::i18n::t;
        long count = 0;
        try
        {
            auto& app = xloil::thisApp();
            xloil::PauseExcel pause(app);

            IDispatch* ad = appDisp();
            if (!ad) return;
            Releaser rApp{ ad };
            IDispatch* sel = selectionRange(ad);
            if (!sel) return;
            Releaser rSel{ sel };

            IDispatch* target = sel;
            IDispatch* urHold = nullptr;
            {
                IDispatch* cells = getObject(sel, L"Cells");
                Releaser rc{ cells };
                if (cells && getDouble(cells, L"CountLarge", 0) == 1.0)
                {
                    IDispatch* sh = getObject(ad, L"ActiveSheet");
                    if (sh)
                    {
                        Releaser rSh{ sh };
                        urHold = getObject(sh, L"UsedRange");
                        if (urHold) target = urHold;
                    }
                }
            }
            Releaser rUr{ urHold };

            const int r = MessageBoxW(excelHwnd(), t(L"cmd.zls.confirm").c_str(),
                                      L"EGTools++",
                                      MB_YESNOCANCEL | MB_ICONEXCLAMATION | MB_DEFBUTTON1);
            if (r == IDCANCEL) return;
            const bool excludeFormula = (r == IDYES);

            auto clearCell = [&](IDispatch* cell)
            {
                if (getBool(cell, L"HasArray", false))
                {
                    IDispatch* arr = getObject(cell, L"CurrentArray");
                    if (!arr) return false;
                    Releaser ra{ arr };
                    return callMethod(arr, L"ClearContents");
                }
                return callMethod(cell, L"ClearContents");
            };

            SafeArr2D values, formulas;
            if (values.load(target, L"Value2") && formulas.load(target, L"Formula"))
            {
                for (long rr = values.r1; rr <= values.r2; ++rr)
                    for (long cc = values.c1; cc <= values.c2; ++cc)
                    {
                        const VARIANT& v = values.at(rr, cc);
                        if (v.vt != VT_BSTR || (v.bstrVal && *v.bstrVal)) continue;
                        const VARIANT& f = formulas.at(rr, cc);
                        const bool hasFormula =
                            f.vt == VT_BSTR && f.bstrVal && *f.bstrVal == L'=';
                        if (hasFormula && excludeFormula) continue;
                        IDispatch* cell = getObjectIdx2(target, L"Cells", rr, cc);
                        if (!cell) continue;
                        Releaser rc{ cell };
                        if (clearCell(cell)) ++count;
                    }
            }
            else
            {
                VARIANT sv; VariantInit(&sv);
                getVar(target, L"Value2", &sv);
                const bool empty = sv.vt == VT_BSTR && (!sv.bstrVal || !*sv.bstrVal);
                VariantClear(&sv);
                if (empty)
                {
                    const std::wstring f = getBStr(target, L"Formula");
                    const bool hasFormula = !f.empty() && f[0] == L'=';
                    if (!(hasFormula && excludeFormula) && clearCell(target))
                        ++count;
                }
            }

            msgInfo(count > 0 ? withCount(t(L"cmd.zls.result"), count)
                              : t(L"cmd.zls.none"));
        }
        catch (...) {}
    }

    // UDF 값 고정 (VB EgUdfRemove, C09_Util.vb:712 — 배포용)
    // EGTools 함수가 든 수식을 값으로 바꾸고 원래 수식을 메모로 남긴다.
    // 대상 판정은 funcTable 단일 출처(core::shadowedFunctionNames — bare/EG./x
    // 접두 변형 포함). 하드코딩 함수 목록(VB) 대신 등록 목록과 자동 동기화.
    void udfFreeze()
    {
        using egtools::i18n::t;
        try
        {
            IDispatch* ad = appDisp();
            if (!ad) return;
            Releaser rApp{ ad };

            // 대상 토큰 집합(대문자 정규화)
            std::set<std::wstring> tokens;
            for (const auto& name : egtools::core::shadowedFunctionNames())
            {
                std::wstring u = name;
                for (auto& ch : u) ch = towupper(ch);
                tokens.insert(u);
                tokens.insert(L"EG." + u);
                if (egtools::core::needsXPrefix(name))
                    tokens.insert(L"X" + u);
            }

            static const std::wregex ident(LR"(([A-Za-z_][A-Za-z0-9_.]*)\s*\()");
            auto referencesUdf = [&](const std::wstring& formula)
            {
                for (auto it = std::wsregex_iterator(formula.begin(), formula.end(), ident);
                     it != std::wsregex_iterator(); ++it)
                {
                    std::wstring name = (*it)[1].str();
                    for (auto& ch : name) ch = towupper(ch);
                    if (tokens.count(name)) return true;
                }
                return false;
            };

            // PASS 1 — 스캔(읽기 전용): 시트별 대상 셀 주소 수집
            struct Target { long sheet; long row; long col; };
            std::vector<Target> targetsList;
            IDispatch* wb = getObject(ad, L"ActiveWorkbook");
            if (!wb) return;
            Releaser rWb{ wb };
            IDispatch* sheets = getObject(wb, L"Worksheets");
            if (!sheets) return;
            Releaser rSheets{ sheets };
            const long nSheets = getLong(sheets, L"Count", 0);
            for (long s = 1; s <= nSheets; ++s)
            {
                IDispatch* sh = getObjectIdx(sheets, L"Item", s);
                if (!sh) continue;
                Releaser rSh{ sh };
                statusBar(ad, getBStr(sh, L"Name") + L" scanning...");
                IDispatch* ur = getObject(sh, L"UsedRange");
                if (!ur) continue;
                Releaser rUr{ ur };
                const long baseRow = getLong(ur, L"Row", 1);
                const long baseCol = getLong(ur, L"Column", 1);

                SafeArr2D formulas;
                if (formulas.load(ur, L"Formula"))
                {
                    for (long rr = formulas.r1; rr <= formulas.r2; ++rr)
                        for (long cc = formulas.c1; cc <= formulas.c2; ++cc)
                        {
                            const VARIANT& f = formulas.at(rr, cc);
                            if (f.vt != VT_BSTR || !f.bstrVal || *f.bstrVal != L'=')
                                continue;
                            if (referencesUdf(f.bstrVal))
                                targetsList.push_back({ s, baseRow + rr - 1,
                                                        baseCol + cc - 1 });
                        }
                }
                else
                {
                    const std::wstring f = getBStr(ur, L"Formula");
                    if (!f.empty() && f[0] == L'=' && referencesUdf(f))
                        targetsList.push_back({ s, baseRow, baseCol });
                }
            }
            statusBarClear(ad);

            if (targetsList.empty())
            {
                msgInfo(t(L"cmd.udffreeze.none"));
                return;
            }

            wchar_t buf[512];
            swprintf_s(buf, t(L"cmd.udffreeze.confirm").c_str(),
                       (int)targetsList.size());
            if (MessageBoxW(excelHwnd(), buf, L"EGTools++",
                            MB_YESNO | MB_ICONQUESTION) != IDYES)
                return;

            // 저장 게이트 + 백업(Convert와 동일 절차)
            std::wstring backupPath;
            if (!egtools::ribbon::detail::saveGateAndBackup(
                    t(L"ribbon.egpp.udfFreeze.label").c_str(), backupPath))
                return;

            // PASS 2 — 적용
            long frozen = 0;
            {
                auto& app = xloil::thisApp();
                xloil::PauseExcel pause(app);

                for (const auto& tg : targetsList)
                {
                    IDispatch* sh = getObjectIdx(sheets, L"Item", tg.sheet);
                    if (!sh) continue;
                    Releaser rSh{ sh };
                    IDispatch* cell = getObjectIdx2(sh, L"Cells", tg.row, tg.col);
                    if (!cell) continue;
                    Releaser rc{ cell };

                    const std::wstring formula = getBStr(cell, L"Formula");
                    if (formula.empty() || formula[0] != L'=')
                        continue;               // 앞선 배열/스필 고정으로 이미 값이 됨

                    // 고정 범위 결정: CSE 배열 → CurrentArray, 스필 → 전체 스필
                    IDispatch* range = nullptr;
                    if (getBool(cell, L"HasArray", false))
                        range = getObject(cell, L"CurrentArray");
                    else if (getBool(cell, L"HasSpill", false))
                        range = getObject(cell, L"SpillingToRange");
                    if (!range) { range = cell; cell->AddRef(); }
                    Releaser rr{ range };

                    VARIANT vals; VariantInit(&vals);
                    if (!getVar(range, L"Value2", &vals)) continue;
                    const bool ok = putVar(range, L"Value", vals);
                    VariantClear(&vals);
                    if (!ok) continue;
                    ++frozen;

                    // 원래 수식을 메모로 보존(기존 메모는 뒤에 이어붙임)
                    std::wstring note = L"EgUdfRemoved: " + formula;
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

            std::wstring done = withCount(t(L"cmd.udffreeze.result"), frozen) +
                                L"\n\n" + backupPath;
            msgInfo(done);
        }
        catch (...) {}
    }

    // 모양 뽑기 (VB DrawFreeForm, C09_Util.vb:355)
    // 선택 영역에서 배경색이 칠해진 셀들의 외곽선을 경계 추적(contour tracing)
    // 상태머신으로 따라가며 자유형 도형을 만든다. VB의 순차 If(한 반복에서
    // 누적 조건이 연달아 발화)와 좌표 규약(vColor는 +1 패딩)을 그대로 보존.
    void drawFreeForm()
    {
        using egtools::i18n::t;
        constexpr long kNoColor = 16777215;     // RGB(255,255,255)
        try
        {
            IDispatch* ad = appDisp();
            if (!ad) return;
            Releaser rApp{ ad };
            IDispatch* sel = selectionRange(ad);
            if (!sel) { msgWarn(t(L"cmd.freeform.selectCells")); return; }
            Releaser rSel{ sel };

            long rowCount = 0, colCount = 0;
            {
                IDispatch* rows = getObject(sel, L"Rows");
                Releaser r1{ rows };
                if (rows) rowCount = getLong(rows, L"Count", 0);
                IDispatch* cols = getObject(sel, L"Columns");
                Releaser r2{ cols };
                if (cols) colCount = getLong(cols, L"Count", 0);
            }
            if (rowCount < 1 || colCount < 1) return;

            // 바탕색 비트맵 (vColor[r][c], VB와 동일한 +1 패딩 좌표계)
            std::vector<std::vector<int>> vColor((size_t)rowCount + 3,
                                                 std::vector<int>((size_t)colCount + 3, 0));
            long filled = 0;
            for (long r = 1; r <= rowCount; ++r)
                for (long c = 1; c <= colCount; ++c)
                {
                    IDispatch* cell = getObjectIdx2(sel, L"Cells", r, c);
                    if (!cell) continue;
                    Releaser rc{ cell };
                    IDispatch* interior = getObject(cell, L"Interior");
                    if (!interior) continue;
                    Releaser ri{ interior };
                    if (getLong(interior, L"Color", kNoColor) != kNoColor)
                    {
                        vColor[(size_t)r + 1][(size_t)c + 1] = 1;
                        ++filled;
                    }
                }
            if (filled == 0) { msgWarn(t(L"cmd.freeform.selectCells")); return; }

            // 셀 좌표/크기 캐시 — 기계 좌표(R,C)의 실제 셀은 Cells(R-1, C-1)
            std::map<long, double> colW, rowH;
            auto cellDisp = [&](long R, long C) -> IDispatch*
            { return getObjectIdx2(sel, L"Cells", R - 1, C - 1); };
            auto width = [&](long R, long C) -> double
            {
                auto it = colW.find(C);
                if (it != colW.end()) return it->second;
                IDispatch* cell = cellDisp(R, C);
                Releaser rc{ cell };
                const double w = cell ? getDouble(cell, L"Width") : 0;
                colW[C] = w;
                return w;
            };
            auto height = [&](long R, long C) -> double
            {
                auto it = rowH.find(R);
                if (it != rowH.end()) return it->second;
                IDispatch* cell = cellDisp(R, C);
                Releaser rc{ cell };
                const double h = cell ? getDouble(cell, L"Height") : 0;
                rowH[R] = h;
                return h;
            };
            auto empty = [&](long R, long C)
            {
                if (R < 0 || C < 0 || R >= (long)vColor.size() ||
                    C >= (long)vColor[0].size()) return true;
                return vColor[(size_t)R][(size_t)C] == 0;
            };
            auto isSet = [&](long R, long C) { return !empty(R, C); };

            IDispatch* sh = getObject(ad, L"ActiveSheet");
            if (!sh) return;
            Releaser rSh{ sh };
            IDispatch* shapes = getObject(sh, L"Shapes");
            if (!shapes) return;
            Releaser rShapes{ shapes };

            // 시작 셀 탐색 + FreeformBuilder 생성
            IDispatch* ffb = nullptr;
            double oX = 0, oY = 0, eX = 0, eY = 0;
            long R = 0, C = 0;
            wchar_t sD = L'\0';
            bool found = false, closedAtStart = false;
            for (R = 1; R <= rowCount + 2 && !found; ++R)
                for (C = 1; C <= colCount + 2; ++C)
                {
                    if (!isSet(R, C)) continue;
                    IDispatch* cell = cellDisp(R, C);
                    if (!cell) return;
                    Releaser rc{ cell };
                    oY = getDouble(cell, L"Top");
                    oX = getDouble(cell, L"Left");
                    eY = oY;
                    eX = oX + getDouble(cell, L"Width");

                    VARIANT a[3];
                    a[0] = varLong(1);              // msoEditingAuto
                    a[1] = varDouble(oX);
                    a[2] = varDouble(oY);
                    VARIANT r; VariantInit(&r);
                    if (!callMethod(shapes, L"BuildFreeform", a, 3, &r) ||
                        r.vt != VT_DISPATCH || !r.pdispVal)
                    {
                        VariantClear(&r);
                        return;
                    }
                    ffb = r.pdispVal;

                    auto addNode = [&](double x, double y)
                    {
                        VARIANT n[4];
                        n[0] = varLong(0);          // msoSegmentLine
                        n[1] = varLong(1);          // msoEditingAuto
                        n[2] = varDouble(x);
                        n[3] = varDouble(y);
                        callMethod(ffb, L"AddNodes", n, 4);
                    };

                    if (empty(R - 1, C))
                    {
                        eX = oX + getDouble(cell, L"Width");
                        sD = L'R';
                    }
                    if (empty(R - 1, C) && empty(R, C + 1))
                    {
                        addNode(eX, eY);
                        eY = eY + getDouble(cell, L"Height");
                        sD = L'D';
                    }
                    if (empty(R - 1, C) && empty(R, C + 1) && empty(R + 1, C))
                    {
                        // 고립된 단일 셀 — 사각형을 닫고 종료(VB GoTo EXIT_RUN)
                        addNode(eX, eY);
                        eX = eX - getDouble(cell, L"Width");
                        addNode(eX, eY);
                        eY = eY - getDouble(cell, L"Height");
                        addNode(eX, eY);
                        closedAtStart = true;
                    }
                    found = true;
                    break;
                }
            --R;                                    // for 루프 탈출 보정(VB Exit For 의미)
            if (!found || !ffb) return;
            Releaser rFfb{ ffb };

            auto addNode = [&](double x, double y)
            {
                VARIANT n[4];
                n[0] = varLong(0);
                n[1] = varLong(1);
                n[2] = varDouble(x);
                n[3] = varDouble(y);
                callMethod(ffb, L"AddNodes", n, 4);
            };

            if (!closedAtStart)
            {
                // 경계 추적 — VB의 순차 If 의미(같은 반복에서 누적 발화) 보존
                bool closed = false;
                for (long guard = 0; guard < 10000 && !closed; ++guard)
                {
                    switch (sD)
                    {
                    case L'R':
                        ++C;
                        if (empty(R - 1, C)) { eX += width(R, C); sD = L'R'; }
                        if (empty(R - 1, C) && empty(R, C + 1))
                        {
                            addNode(eX, eY);
                            if (oX == eX && oY == eY) { closed = true; break; }
                            eY += height(R, C);
                            sD = L'D';
                        }
                        if (empty(R - 1, C) && empty(R, C + 1) && empty(R + 1, C))
                        {
                            addNode(eX, eY);
                            if (oX == eX && oY == eY) { closed = true; break; }
                            eX -= width(R, C);
                            sD = L'L';
                        }
                        if (isSet(R, C - 1) && isSet(R - 1, C))
                        {
                            addNode(eX, eY);
                            if (oX == eX && oY == eY) { closed = true; break; }
                            sD = L'U';
                        }
                        break;
                    case L'L':
                        --C;
                        if (empty(R + 1, C)) { eX -= width(R, C); sD = L'L'; }
                        if (empty(R + 1, C) && empty(R, C - 1))
                        {
                            addNode(eX, eY);
                            if (oX == eX && oY == eY) { closed = true; break; }
                            eY -= height(R, C);
                            sD = L'U';
                        }
                        if (empty(R + 1, C) && empty(R, C - 1) && empty(R - 1, C))
                        {
                            addNode(eX, eY);
                            if (oX == eX && oY == eY) { closed = true; break; }
                            eX += width(R, C);
                            sD = L'R';
                        }
                        if (isSet(R, C + 1) && isSet(R + 1, C))
                        {
                            addNode(eX, eY);
                            if (oX == eX && oY == eY) { closed = true; break; }
                            sD = L'D';
                        }
                        break;
                    case L'D':
                        ++R;
                        if (empty(R, C + 1)) { eY += height(R, C); sD = L'D'; }
                        if (empty(R, C + 1) && empty(R + 1, C))
                        {
                            addNode(eX, eY);
                            if (oX == eX && oY == eY) { closed = true; break; }
                            eX -= width(R, C);
                            sD = L'L';
                        }
                        if (empty(R, C + 1) && empty(R + 1, C) && empty(R, C - 1))
                        {
                            addNode(eX, eY);
                            if (oX == eX && oY == eY) { closed = true; break; }
                            eY -= height(R, C);
                            sD = L'U';
                        }
                        if (isSet(R - 1, C) && isSet(R, C + 1))
                        {
                            addNode(eX, eY);
                            if (oX == eX && oY == eY) { closed = true; break; }
                            sD = L'R';
                        }
                        break;
                    case L'U':
                        --R;
                        if (empty(R, C - 1)) { eY -= height(R, C); sD = L'U'; }
                        if (empty(R, C - 1) && empty(R - 1, C))
                        {
                            addNode(eX, eY);
                            if (oX == eX && oY == eY) { closed = true; break; }
                            eX += width(R, C);
                            sD = L'R';
                        }
                        if (empty(R, C - 1) && empty(R - 1, C) && empty(R, C + 1))
                        {
                            addNode(eX, eY);
                            if (oX == eX && oY == eY) { closed = true; break; }
                            eY += height(R, C);
                            sD = L'D';
                        }
                        if (isSet(R + 1, C) && isSet(R, C - 1))
                        {
                            addNode(eX, eY);
                            if (oX == eX && oY == eY) { closed = true; break; }
                            sD = L'L';
                        }
                        break;
                    default:
                        closed = true;
                        break;
                    }
                }
            }

            // ConvertToShape + BackgroundStyle=10 (VB 동일)
            VARIANT r; VariantInit(&r);
            if (callMethod(ffb, L"ConvertToShape", nullptr, 0, &r) &&
                r.vt == VT_DISPATCH && r.pdispVal)
            {
                IDispatch* shp = r.pdispVal;
                putLong(shp, L"BackgroundStyle", 10);
            }
            else
                msgWarn(t(L"cmd.freeform.convertFail"));
            VariantClear(&r);
        }
        catch (...) {}
    }
}
