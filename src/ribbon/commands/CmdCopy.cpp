// CmdCopy.cpp — 보이는 셀 복사/붙여넣기 (EGToolsVB C01_Copy.vb 포팅, plan/23).
// M3: CopyVisibleCells / CopyAllCells(원본 기억).
// M4: PasteToVisibleCells 계열이 여기 저장된 원본을 사용한다.
//
// VB는 복사 원본 Range의 COM 참조를 모듈 전역에 보관했으나(원본 워크북이
// 닫히면 dangling), 여기서는 plan/23 §3-6에 따라 **주소 문자열**로 보관하고
// 사용 시점에 재해석한다.

#include "Commands.h"
#include "CmdCommon.h"

#include <vector>

namespace egtools::commands
{
    namespace
    {
        struct CopyState
        {
            bool valid = false;
            bool visibleOnly = false;            // 보이는 셀만 복사였는지
            std::wstring workbook, sheet;
            std::vector<std::wstring> areas;     // 절대 주소 목록
        };
        CopyState g_copy;                        // main-thread only

        // Range.Address(RowAbsolute, ColumnAbsolute) — VB와 동일하게 (False,False).
        std::wstring rangeAddress(IDispatch* range, bool absolute)
        {
            VARIANT a[2];
            VariantInit(&a[0]); VariantInit(&a[1]);
            a[0].vt = VT_BOOL; a[0].boolVal = absolute ? VARIANT_TRUE : VARIANT_FALSE;
            a[1].vt = VT_BOOL; a[1].boolVal = absolute ? VARIANT_TRUE : VARIANT_FALSE;
            VARIANT r; VariantInit(&r);
            std::wstring s;
            if (invokeRaw(range, L"Address", DISPATCH_PROPERTYGET, &r, a, 2) &&
                r.vt == VT_BSTR && r.bstrVal)
                s = r.bstrVal;
            VariantClear(&r);
            return s;
        }

        void rememberSource(IDispatch* ad, IDispatch* srcRange, bool visibleOnly)
        {
            g_copy = CopyState{};
            IDispatch* areas = getObject(srcRange, L"Areas");
            if (!areas) return;
            Releaser ra{ areas };
            const long n = getLong(areas, L"Count", 0);
            for (long i = 1; i <= n; ++i)
            {
                IDispatch* area = getObjectIdx(areas, L"Item", i);
                if (!area) continue;
                Releaser rArea{ area };
                g_copy.areas.push_back(rangeAddress(area, false));
            }
            if (IDispatch* sh = getObject(ad, L"ActiveSheet"))
            {
                Releaser r{ sh };
                g_copy.sheet = getBStr(sh, L"Name");
            }
            if (IDispatch* wb = getObject(ad, L"ActiveWorkbook"))
            {
                Releaser r{ wb };
                g_copy.workbook = getBStr(wb, L"Name");
            }
            g_copy.visibleOnly = visibleOnly;
            g_copy.valid = !g_copy.areas.empty();
        }
    }

    // 보이는 셀만 복사 (VB CopyVisibleCells, C01_Copy.vb:13)
    // 필터/숨김이 적용된 선택 영역에서 보이는 셀만 원본으로 기억 + 클립보드 복사.
    void copyVisibleCells()
    {
        try
        {
            IDispatch* ad = appDisp();
            if (!ad) return;
            Releaser rApp{ ad };
            IDispatch* sel = selectionRange(ad);
            if (!sel) return;
            Releaser rSel{ sel };

            // SpecialCells(xlCellTypeVisible=12)
            VARIANT a = varLong(12);
            VARIANT r; VariantInit(&r);
            if (!invokeRaw(sel, L"SpecialCells",
                           DISPATCH_METHOD | DISPATCH_PROPERTYGET, &r, &a, 1) ||
                r.vt != VT_DISPATCH || !r.pdispVal)
            {
                VariantClear(&r);
                return;
            }
            IDispatch* vis = r.pdispVal;
            Releaser rVis{ vis };

            rememberSource(ad, vis, true);
            callMethod(vis, L"Copy");
        }
        catch (...) {}
    }

    // 전체 복사 (VB CopyAllCells, C01_Copy.vb:38)
    // 선택 영역 전체(숨김 포함)를 원본으로 기억만 한다(VB와 동일 — 클립보드 미복사,
    // 붙여넣기 명령이 원본을 직접 읽는다).
    void copyAllCells()
    {
        try
        {
            IDispatch* ad = appDisp();
            if (!ad) return;
            Releaser rApp{ ad };
            IDispatch* sel = selectionRange(ad);
            if (!sel) return;
            Releaser rSel{ sel };
            rememberSource(ad, sel, false);
        }
        catch (...) {}
    }

    // ════════ M4: 보이는 셀에 붙여넣기 (VB PasteToVisibleCells, C01_Copy.vb:60) ════════

    namespace
    {
        // Excel 상수 (XlPasteType / XlCellType / Shift 방향 등)
        constexpr long xlPasteAll           = -4104;
        constexpr long xlPasteFormulas      = -4123;
        constexpr long xlPasteValues        = -4163;
        constexpr long xlPasteOperationNone = -4142;   // xlPasteSpecialOperationNone
        constexpr long xlCellTypeVisible    = 12;
        constexpr long xlCellTypeBlanks     = 4;
        constexpr long xlWhole              = 1;
        constexpr long xlShiftDown          = -4121;   // = xlDown
        constexpr long xlShiftUp            = -4162;   // = xlUp
        constexpr long xlShiftToRight       = -4161;   // = xlToRight
        constexpr long xlShiftToLeft        = -4159;   // = xlToLeft

        // 빈 셀 센티넬 (VB EmptyString, C01_Copy.vb:70)
        constexpr wchar_t kEmptyMark[] = L"!E!M!P!T!Y!";

        inline VARIANT varBool(bool b)
        {
            VARIANT a; VariantInit(&a);
            a.vt = VT_BOOL; a.boolVal = b ? VARIANT_TRUE : VARIANT_FALSE;
            return a;
        }

        // IDispatch를 돌려주는 메서드 호출 (Workbooks.Add / Sheets.Add 등).
        IDispatch* callObject(IDispatch* d, const wchar_t* name,
                              VARIANT* args = nullptr, UINT n = 0)
        {
            VARIANT r; VariantInit(&r);
            if (!invokeRaw(d, name, DISPATCH_METHOD, &r, args, n)) return nullptr;
            if (r.vt == VT_DISPATCH && r.pdispVal) return r.pdispVal;
            VariantClear(&r);
            return nullptr;
        }

        IDispatch* rangeFromAddr(IDispatch* parent, const std::wstring& addr)
        {
            VARIANT a = varBStr(addr);
            IDispatch* r = getObject(parent, L"Range", &a, 1);
            VariantClear(&a);
            return r;
        }

        IDispatch* rangeOffset(IDispatch* rng, long dr, long dc)
        {
            VARIANT a[2]; a[0] = varLong(dr); a[1] = varLong(dc);
            return getObject(rng, L"Offset", a, 2);
        }

        IDispatch* specialCells(IDispatch* rng, long type)
        {
            VARIANT a = varLong(type);
            return getObject(rng, L"SpecialCells", &a, 1);
        }

        // Range.Cells.CountLarge (다중 Area 전체 셀 수 — VB CopySource.Cells.Count)
        double cellCount(IDispatch* rng)
        {
            IDispatch* cells = getObject(rng, L"Cells");
            if (!cells) return 0;
            Releaser r{ cells };
            return getDouble(cells, L"CountLarge", 0);
        }

        // parent.Rows.Count / parent.Columns.Count
        long collCount(IDispatch* parent, const wchar_t* coll)
        {
            IDispatch* c = getObject(parent, coll);
            if (!c) return 0;
            Releaser r{ c };
            return getLong(c, L"Count", 0);
        }

        // PasteSpecial(Paste, Operation:=xlNone, SkipBlanks:=True, Transpose:=False)
        // — VB 87/199와 동일한 위치 인수.
        bool pasteSpecialSkipBlanks(IDispatch* rng, long pasteType)
        {
            VARIANT a[4];
            a[0] = varLong(pasteType);
            a[1] = varLong(xlPasteOperationNone);
            a[2] = varBool(true);
            a[3] = varBool(false);
            return callMethod(rng, L"PasteSpecial", a, 4);
        }

        // cell.EntireRow.Insert(shift) / cell.EntireColumn.Delete(shift) 등.
        bool entireOp(IDispatch* cell, const wchar_t* entire,
                      const wchar_t* op, long shift)
        {
            IDispatch* e = getObject(cell, entire);
            if (!e) return false;
            Releaser r{ e };
            VARIANT a = varLong(shift);
            return callMethod(e, op, &a, 1);
        }

        // ── 오토필터 백업/복원 (VB BackUpFilters/ReStoreFilters, C01_Copy.vb:226/262) ──
        // Criteria1은 VARIANT 그대로 보관: 문자열, xlFilterValues(7)의 1차원 배열 등.
        // xlFilterCellColor(8)만 VB처럼 Criteria1.Color(long)로 저장한다.
        struct FilterBackup
        {
            struct Crit
            {
                bool on = false;
                long op = 0;
                bool has1 = false, has2 = false;
                VARIANT c1, c2;
                Crit() { VariantInit(&c1); VariantInit(&c2); }
                Crit(Crit&& o) noexcept
                    : on(o.on), op(o.op), has1(o.has1), has2(o.has2), c1(o.c1), c2(o.c2)
                {
                    VariantInit(&o.c1); VariantInit(&o.c2);
                    o.has1 = o.has2 = false;
                }
                Crit(const Crit&) = delete;
                Crit& operator=(const Crit&) = delete;
                Crit& operator=(Crit&&) = delete;
                ~Crit() { VariantClear(&c1); VariantClear(&c2); }
            };

            std::wstring rangeAddr;      // VB cFilteredAddress
            std::vector<Crit> crits;
            bool active = false;

            void backup(IDispatch* sh)
            {
                if (!getBool(sh, L"AutoFilterMode", false)) return;     // VB 232
                IDispatch* af = getObject(sh, L"AutoFilter");
                if (!af) return;
                Releaser raf{ af };
                if (IDispatch* rng = getObject(af, L"Range"))
                {
                    Releaser r{ rng };
                    rangeAddr = getBStr(rng, L"Address");
                }
                if (rangeAddr.empty()) return;                          // VB 237
                IDispatch* fs = getObject(af, L"Filters");
                if (!fs) return;
                Releaser rfs{ fs };
                const long n = getLong(fs, L"Count", 0);
                for (long i = 1; i <= n; ++i)
                {
                    Crit cr;
                    if (IDispatch* f = getObjectIdx(fs, L"Item", i))
                    {
                        Releaser rf{ f };
                        if (getBool(f, L"On", false))                   // VB 241
                        {
                            cr.on = true;
                            cr.op = getLong(f, L"Operator", 0);
                            if (cr.op == 1 || cr.op == 2)               // xlAnd/xlOr: 조건 2개 (VB 244)
                            {
                                cr.has1 = getVar(f, L"Criteria1", &cr.c1);
                                cr.has2 = getVar(f, L"Criteria2", &cr.c2);
                            }
                            else if (cr.op == 8)                        // xlFilterCellColor (VB 247)
                            {
                                if (IDispatch* c = getObject(f, L"Criteria1"))
                                {
                                    Releaser rc{ c };
                                    cr.c1 = varLong(getLong(c, L"Color", 0));
                                    cr.has1 = true;
                                }
                            }
                            else    // 0/3~7/9~11: xlFilterValues 배열 포함 그대로 (VB 249)
                                cr.has1 = getVar(f, L"Criteria1", &cr.c1);
                        }
                    }
                    crits.push_back(std::move(cr));
                }
                active = true;
            }

            // 한 번만 복원(중복 호출 무해). VB와 동일하게 필드별 Range.AutoFilter 재적용.
            void restore(IDispatch* sh)
            {
                if (!active) return;
                active = false;
                if (!sh || rangeAddr.empty()) return;
                for (size_t i = 0; i < crits.size(); ++i)
                {
                    const Crit& cr = crits[i];
                    if (!cr.on) continue;                               // VB 270
                    IDispatch* rng = rangeFromAddr(sh, rangeAddr);
                    if (!rng) return;
                    Releaser rr{ rng };
                    VARIANT a[4];
                    a[0] = varLong((long)i + 1);                        // Field
                    if (cr.has1) a[1] = cr.c1;                          // 얕은 복사(소유권 유지)
                    else { VariantInit(&a[1]); a[1].vt = VT_ERROR; a[1].scode = DISP_E_PARAMNOTFOUND; }
                    a[2] = varLong(cr.op);                              // Operator
                    if ((cr.op == 1 || cr.op == 2) && cr.has2)
                    {
                        a[3] = cr.c2;
                        callMethod(rng, L"AutoFilter", a, 4);           // VB 274
                    }
                    else
                        callMethod(rng, L"AutoFilter", a, 3);           // VB 277
                }
            }
        };

        // VB Finally(210-221) 대응 RAII: 어떤 경로로 빠져나가도
        // 임시 워크북 폐기 → 필터 복원(미복원 시) → CutCopyMode 해제.
        struct PasteCleanup
        {
            IDispatch* app = nullptr;         // 소유 안 함
            IDispatch* srcSheet = nullptr;    // 소유 안 함
            FilterBackup* filters = nullptr;
            IDispatch* tempWb = nullptr;      // 소유(Close 후 Release)
            ~PasteCleanup()
            {
                if (tempWb)
                {
                    putBool(tempWb, L"Saved", true);          // VB 214
                    VARIANT a = varBool(false);
                    callMethod(tempWb, L"Close", &a, 1);      // VB 215 Close(False)
                    tempWb->Release();
                }
                if (filters && srcSheet) filters->restore(srcSheet);
                if (app) putBool(app, L"CutCopyMode", false);
            }
        };

        // 보이는 셀에 붙여넣기 본체 (VB PasteToVisibleCells, C01_Copy.vb:60).
        // 알고리즘: 원본(보이는 셀)을 임시 워크북에 붙인 뒤, 대상 영역의 숨김
        // 행/열 위치에 빈 행/열을 삽입해 모양을 맞추고, 빈 셀은 센티넬로 채워
        // SkipBlanks 붙여넣기로 대상의 보이는 셀에만 얹는다.
        void pasteToVisibleCells(long pasteType)
        {
            using egtools::i18n::t;
            try
            {
                // VB 72: 복사 원본이 없으면 안내 후 종료
                if (!g_copy.valid)
                {
                    msgWarn(t(L"cmd.paste.copyFirst"));
                    return;
                }

                auto& app = xloil::thisApp();
                xloil::PauseExcel pause(app);   // ScreenUpdating/DisplayAlerts 저장·복원 (VB 77,212,219,221)

                IDispatch* ad = appDisp();
                if (!ad) return;
                Releaser rApp{ ad };
                IDispatch* aSH = getObject(ad, L"ActiveSheet");
                if (!aSH) return;
                Releaser rAsh{ aSH };
                IDispatch* targets = selectionRange(ad);
                if (!targets) return;           // 셀 선택이 아니면 조용히 종료
                Releaser rTgt{ targets };

                // VB 82: 원본 워크북/시트 재해석. VB는 COM 참조를 그대로 써서
                // 원본이 닫히면 dangling — 여기서는 이름으로 다시 찾고 없으면 안내.
                IDispatch* sSH = nullptr;
                if (IDispatch* wbs = getObject(ad, L"Workbooks"))
                {
                    Releaser r{ wbs };
                    VARIANT nm = varBStr(g_copy.workbook);
                    IDispatch* wb = getObject(wbs, L"Item", &nm, 1);
                    VariantClear(&nm);
                    if (wb)
                    {
                        Releaser rw{ wb };
                        VARIANT sn = varBStr(g_copy.sheet);
                        sSH = getObject(wb, L"Worksheets", &sn, 1);
                        VariantClear(&sn);
                    }
                }
                if (!sSH) { msgWarn(t(L"cmd.paste.sourceGone")); return; }
                Releaser rSsh{ sSH };

                // 원본 Range 재해석 (다중 Area는 쉼표로 연결 — VB vSrcAddr(0)은
                // 첫 Area만 담는 버그가 있어 전체 주소를 쓴다)
                std::wstring srcAddr;
                for (const auto& s : g_copy.areas)
                {
                    if (!srcAddr.empty()) srcAddr += L",";
                    srcAddr += s;
                }
                IDispatch* src = rangeFromAddr(sSH, srcAddr);
                if (!src) { msgWarn(t(L"cmd.paste.sourceGone")); return; }
                Releaser rSrc{ src };

                // VB 86: 원본이 셀 하나면 대상 보이는 셀에 바로 붙여넣고 끝
                if (cellCount(src) == 1)
                {
                    if (getLong(ad, L"CutCopyMode", 0) == 0)
                        callMethod(src, L"Copy");   // 개선: VB는 클립보드가 죽었으면 오류
                    IDispatch* vis = specialCells(targets, xlCellTypeVisible);
                    if (!vis) return;
                    Releaser rv{ vis };
                    pasteSpecialSkipBlanks(vis, pasteType);
                    return;
                }

                FilterBackup filters;
                PasteCleanup cleanup;
                cleanup.app = ad;
                cleanup.srcSheet = sSH;
                cleanup.filters = &filters;

                // VB 91: 복사모드가 아니면 필터 백업 → 필터 해제 → 원본 재복사
                if (getLong(ad, L"CutCopyMode", 0) == 0)
                {
                    filters.backup(sSH);
                    putBool(sSH, L"AutoFilterMode", false);
                    callMethod(src, L"Copy");
                }

                // VB 100: 대상 좌표. 대상이 셀 하나면 시트 전체로 확장
                const long lbR = getLong(targets, L"Row", 1);
                const long lbC = getLong(targets, L"Column", 1);
                long tRows = 0, tCols = 0;
                if (cellCount(targets) == 1)
                {
                    tRows = collCount(aSH, L"Rows");
                    tCols = collCount(aSH, L"Columns");
                }
                else
                {
                    tRows = lbR + collCount(targets, L"Rows") - 1;
                    tCols = lbC + collCount(targets, L"Columns") - 1;
                }

                // VB 111: 여전히 복사모드가 아니면 진행 불가 (VB c0102)
                if (getLong(ad, L"CutCopyMode", 0) == 0)
                {
                    msgWarn(t(L"cmd.paste.copyFirst"));
                    return;
                }

                // VB 114: 임시 워크북 + 임시 시트에 원본 붙여넣기
                if (IDispatch* wbs = getObject(ad, L"Workbooks"))
                {
                    Releaser r{ wbs };
                    cleanup.tempWb = callObject(wbs, L"Add");
                }
                if (!cleanup.tempWb) return;

                IDispatch* tSH = nullptr;
                if (IDispatch* sheets = getObject(cleanup.tempWb, L"Sheets"))
                {
                    Releaser r{ sheets };
                    tSH = callObject(sheets, L"Add");
                }
                if (!tSH) return;
                Releaser rTsh{ tSH };

                {
                    IDispatch* a1 = rangeFromAddr(tSH, L"A1");
                    if (!a1) return;
                    Releaser r{ a1 };
                    VARIANT p = varLong(pasteType);
                    if (!callMethod(a1, L"PasteSpecial", &p, 1))    // VB 116
                        return;
                }

                long cRows = 0, cCols = 0;
                {
                    IDispatch* ur = getObject(tSH, L"UsedRange");
                    if (!ur) return;
                    Releaser r{ ur };
                    cRows = collCount(ur, L"Rows");
                    cCols = collCount(ur, L"Columns");

                    // VB 122: 빈 셀 센티넬 — SkipBlanks가 빈 셀을 건너뛰어
                    // 대상 값이 남는 문제 예방 (빈 셀 없으면 실패 무시)
                    if (IDispatch* blanks = specialCells(ur, xlCellTypeBlanks))
                    {
                        Releaser rb{ blanks };
                        putBStr(blanks, L"Formula", kEmptyMark);
                    }
                }
                if (cRows < 1 || cCols < 1) return;

                // VB 129: 숨김 스캔 전에 원본 필터 복구
                // (원본=대상 시트인 경우 행 숨김 상태가 필터를 반영해야 함)
                filters.restore(sSH);

                // VB 137-154: 대상 영역의 행/열 숨김 상태 수집
                // (보이는 칸 수가 원본 크기에 닿으면 중단 — 인덱스 0 = lbR/lbC)
                std::vector<char> hidRows, hidCols;
                {
                    long iCnt = 0;
                    for (long r = lbR; r <= tRows; ++r)
                    {
                        bool hid = false;
                        if (IDispatch* row = getObjectIdx(aSH, L"Rows", r))
                        {
                            Releaser rr{ row };
                            hid = getBool(row, L"Hidden", false);
                        }
                        hidRows.push_back(hid ? 1 : 0);
                        if (!hid && ++iCnt == cRows) break;
                    }
                    iCnt = 0;
                    for (long c = lbC; c <= tCols; ++c)
                    {
                        bool hid = false;
                        if (IDispatch* col = getObjectIdx(aSH, L"Columns", c))
                        {
                            Releaser rc{ col };
                            hid = getBool(col, L"Hidden", false);
                        }
                        hidCols.push_back(hid ? 1 : 0);
                        if (!hid && ++iCnt == cCols) break;
                    }
                }

                // VB 158-175: 대상 숨김 행 위치에 빈 행 삽입, 대상 범위 초과분은 삭제
                // (VB의 `i > ubR - lbR`는 경계에서 배열 범위를 1 넘는 잠재 버그 —
                //  여기서는 수집된 크기 기준으로 초과분을 삭제 분기로 보낸다)
                {
                    IDispatch* aCell = rangeFromAddr(tSH, L"A1");
                    if (!aCell) return;
                    long i = 0;
                    while (i < cRows)
                    {
                        if (i >= (long)hidRows.size())          // 초과 행 삭제 (VB 161)
                        {
                            IDispatch* next = rangeOffset(aCell, 1, 0);
                            if (!next) break;
                            if (IDispatch* prev = rangeOffset(next, -1, 0))
                            {
                                Releaser rp{ prev };
                                entireOp(prev, L"EntireRow", L"Delete", xlShiftUp);
                            }
                            aCell->Release(); aCell = next;
                        }
                        else if (hidRows[(size_t)i])            // 숨김 행 → 빈 행 삽입 (VB 165)
                        {
                            entireOp(aCell, L"EntireRow", L"Insert", xlShiftDown);
                            ++cRows;
                        }
                        else
                        {
                            IDispatch* next = rangeOffset(aCell, 1, 0);
                            if (!next) break;
                            aCell->Release(); aCell = next;
                        }
                        ++i;
                    }
                    cRows = getLong(aCell, L"Row", 1) - 1;      // VB 175
                    aCell->Release();
                }

                // VB 177-194: 열도 동일하게 처리
                {
                    IDispatch* aCell = rangeFromAddr(tSH, L"A1");
                    if (!aCell) return;
                    long i = 0;
                    while (i < cCols)
                    {
                        if (i >= (long)hidCols.size())          // 초과 열 삭제 (VB 180)
                        {
                            IDispatch* next = rangeOffset(aCell, 0, 1);
                            if (!next) break;
                            if (IDispatch* prev = rangeOffset(next, 0, -1))
                            {
                                Releaser rp{ prev };
                                entireOp(prev, L"EntireColumn", L"Delete", xlShiftToLeft);
                            }
                            aCell->Release(); aCell = next;
                        }
                        else if (hidCols[(size_t)i])            // 숨김 열 → 빈 열 삽입 (VB 184)
                        {
                            entireOp(aCell, L"EntireColumn", L"Insert", xlShiftToRight);
                            ++cCols;
                        }
                        else
                        {
                            IDispatch* next = rangeOffset(aCell, 0, 1);
                            if (!next) break;
                            aCell->Release(); aCell = next;
                        }
                        ++i;
                    }
                    cCols = getLong(aCell, L"Column", 1) - 1;   // VB 194
                    aCell->Release();
                }
                if (cRows < 1 || cCols < 1) return;

                // VB 197-199: 조정된 임시 영역 복사 → 대상 첫 셀에 SkipBlanks 붙여넣기
                {
                    IDispatch* c1 = rangeFromAddr(tSH, L"A1");
                    IDispatch* c2 = getObjectIdx2(tSH, L"Cells", cRows, cCols);
                    Releaser r1{ c1 }, r2{ c2 };
                    if (!c1 || !c2) return;
                    VARIANT a[2];
                    VariantInit(&a[0]); a[0].vt = VT_DISPATCH; a[0].pdispVal = c1;
                    VariantInit(&a[1]); a[1].vt = VT_DISPATCH; a[1].pdispVal = c2;
                    IDispatch* block = getObject(tSH, L"Range", a, 2);
                    if (!block) return;
                    Releaser rb{ block };
                    if (!callMethod(block, L"Copy")) return;
                }
                callMethod(aSH, L"Activate");                   // VB 198
                {
                    IDispatch* first = getObjectIdx2(targets, L"Cells", 1, 1);
                    if (!first) return;
                    Releaser r{ first };
                    pasteSpecialSkipBlanks(first, pasteType);   // VB 199
                }

                // VB 202: 대상 보이는 셀에서 센티넬 제거 (실패 무시)
                if (IDispatch* sel = selectionRange(ad))
                {
                    Releaser rs{ sel };
                    if (IDispatch* vis = specialCells(sel, xlCellTypeVisible))
                    {
                        Releaser rv{ vis };
                        VARIANT a[3];
                        a[0] = varBStr(kEmptyMark);
                        a[1] = varBStr(L"");
                        a[2] = varLong(xlWhole);                // LookAt:=xlWhole
                        callMethod(vis, L"Replace", a, 3);
                        VariantClear(&a[0]); VariantClear(&a[1]);
                    }
                }
            }
            catch (...) {}   // 정리는 PasteCleanup/PauseExcel RAII가 보장
        }
    }

    // 보이는 셀에 값/수식/모두 붙여넣기 (VB PasteVisible*, C01_Copy.vb:48-58)
    void pasteVisibleValues()   { pasteToVisibleCells(xlPasteValues); }
    void pasteVisibleFormulas() { pasteToVisibleCells(xlPasteFormulas); }
    void pasteVisibleAll()      { pasteToVisibleCells(xlPasteAll); }
}
