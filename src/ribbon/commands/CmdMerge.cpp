// CmdMerge.cpp — 병합/나누기 명령 (EGToolsVB C02_Merge.vb 포팅, plan/23).
// M2: SplitFill(나누고 채우기). M3: MergeClone(연속값 병합).
// M4: MergeJoin/MergeRow/MergeColumn(내용병합), FakeMerge(Clone)(내용감춘 병합),
//     SplitRows/SplitColumns(행/열 나누기).

#include "Commands.h"
#include "CmdCommon.h"
#include "Dialogs.h"

#include <optional>
#include <vector>

namespace egtools::commands
{
    // 나누고 채우기 (VB SplitFill, C02_Merge.vb:501)
    // 병합을 해제하고 병합돼 있던 값을 (첫 열에) 다시 채운다.
    // 수식은 값으로 바뀐다(VB와 동일).
    void splitFill()
    {
        try
        {
            auto& app = xloil::thisApp();
            xloil::PauseExcel pause(app);

            IDispatch* ad = appDisp();
            if (!ad) return;
            Releaser rApp{ ad };
            IDispatch* sel = selectionRange(ad);
            if (!sel) return;                       // 셀 선택이 아니면 조용히 종료(VB 동일)
            Releaser rSel{ sel };

            forEachCell(sel, [](IDispatch* cell)
            {
                VARIANT v; VariantInit(&v);
                getVar(cell, L"Value2", &v);

                IDispatch* ma = getObject(cell, L"MergeArea");
                if (!ma) { VariantClear(&v); return true; }
                Releaser rMa{ ma };
                // VB: MergeArea.Columns(1) — 여러 열 병합이면 왼쪽 첫 열에만 채움
                IDispatch* col1 = getObjectIdx(ma, L"Columns", 1);
                if (!col1) { VariantClear(&v); return true; }
                Releaser rCol{ col1 };

                callMethod(col1, L"UnMerge");
                if (!isEmptyVal(v))
                    putVar(col1, L"Value", v);
                VariantClear(&v);
                return true;
            });
        }
        catch (...) {}   // VB와 동일: 조용히 종료(상태는 PauseExcel이 복원)
    }

    namespace
    {
        std::wstring valToStr(const VARIANT& v)
        {
            if (isEmptyVal(v)) return L"";
            return varToString(v);
        }

        // 선택 범위에서 세로로 연속된 같은 값 구간(2칸 이상, 빈칸 제외)의 절대
        // 주소 목록 (VB ContinuousRunAddresses, C02_Merge.vb:140 — 계층형:
        // 오른쪽 열의 구간은 왼쪽 열의 그룹 경계에서 끊는다).
        std::vector<std::wstring> continuousRunAddresses(IDispatch* sel)
        {
            std::vector<std::wstring> out;
            IDispatch* areas = getObject(sel, L"Areas");
            if (!areas) return out;
            Releaser ra{ areas };
            const long nAreas = getLong(areas, L"Count", 0);
            for (long a = 1; a <= nAreas; ++a)
            {
                IDispatch* area = getObjectIdx(areas, L"Item", a);
                if (!area) continue;
                Releaser rArea{ area };

                long nR = 0, nK = 0;
                {
                    IDispatch* rows = getObject(area, L"Rows");
                    if (rows) { Releaser r{ rows }; nR = getLong(rows, L"Count", 0); }
                    IDispatch* cols = getObject(area, L"Columns");
                    if (cols) { Releaser r{ cols }; nK = getLong(cols, L"Count", 0); }
                }
                if (nR < 2) continue;

                SafeArr2D vals;
                if (!vals.load(area, L"Value2")) continue;

                std::vector<bool> cumBreak((size_t)nR + 2, false);
                for (long c = 1; c <= nK; ++c)
                {
                    for (long r = 2; r <= nR; ++r)
                        if (valToStr(vals.at(r, c)) != valToStr(vals.at(r - 1, c)))
                            cumBreak[r] = true;

                    long r = 1;
                    while (r <= nR)
                    {
                        long r2 = r;
                        while (r2 < nR && !cumBreak[(size_t)r2 + 1]) ++r2;
                        if (r2 > r && !valToStr(vals.at(r, c)).empty())
                        {
                            // area.Range()는 상대좌표 재해석 — 셀 절대주소 조합(VB 동일)
                            IDispatch* c1 = getObjectIdx2(area, L"Cells", r, c);
                            IDispatch* c2 = getObjectIdx2(area, L"Cells", r2, c);
                            Releaser rc1{ c1 }, rc2{ c2 };
                            if (c1 && c2)
                                out.push_back(getBStr(c1, L"Address") + L":" +
                                              getBStr(c2, L"Address"));
                        }
                        r = r2 + 1;
                    }
                }
            }
            return out;
        }

        // 선택이 표(ListObject) 영역과 겹치면 true (VB 가드와 동일).
        bool intersectsTable(IDispatch* ad, IDispatch* sh, IDispatch* sel)
        {
            IDispatch* tables = getObject(sh, L"ListObjects");
            if (!tables) return false;
            Releaser rt{ tables };
            const long n = getLong(tables, L"Count", 0);
            for (long i = 1; i <= n; ++i)
            {
                IDispatch* tbl = getObjectIdx(tables, L"Item", i);
                if (!tbl) continue;
                Releaser rTbl{ tbl };
                IDispatch* tr = getObject(tbl, L"Range");
                if (!tr) continue;
                Releaser rTr{ tr };

                VARIANT args[2];
                VariantInit(&args[0]); VariantInit(&args[1]);
                args[0].vt = VT_DISPATCH; args[0].pdispVal = sel;
                args[1].vt = VT_DISPATCH; args[1].pdispVal = tr;
                VARIANT res; VariantInit(&res);
                const bool got = invokeRaw(ad, L"Intersect", DISPATCH_METHOD,
                                           &res, args, 2);
                bool hit = false;
                if (got && res.vt == VT_DISPATCH && res.pdispVal)
                {
                    IDispatch* isect = res.pdispVal;
                    IDispatch* cells = getObject(isect, L"Cells");
                    if (cells)
                    {
                        Releaser rc{ cells };
                        hit = getDouble(cells, L"CountLarge", 0) > 0;
                    }
                }
                VariantClear(&res);
                if (hit) return true;
            }
            return false;
        }
    }

    // 연속값 병합 (VB MergeClone, C02_Merge.vb:9)
    // 세로로 같은 값이 연속되는 구간을 실제 병합. 여러 열이면 계층형(왼쪽 열
    // 그룹 경계에서 오른쪽 열 구간을 끊음). 빈 셀 구간은 병합하지 않음.
    void mergeClone()
    {
        using egtools::i18n::t;
        try
        {
            auto& app = xloil::thisApp();
            xloil::PauseExcel pause(app);

            IDispatch* ad = appDisp();
            if (!ad) return;
            Releaser rApp{ ad };
            IDispatch* sel = selectionRange(ad);
            if (!sel) { msgWarn(t(L"cmd.merge.needRange")); return; }
            Releaser rSel{ sel };

            {
                IDispatch* rows = getObject(sel, L"Rows");
                Releaser rr{ rows };
                if (rows && getLong(rows, L"Count", 0) == 1)
                {
                    msgWarn(t(L"cmd.merge.oneRow"));
                    return;
                }
            }

            IDispatch* sh = getObject(ad, L"ActiveSheet");
            if (!sh) return;
            Releaser rSh{ sh };
            if (intersectsTable(ad, sh, sel))
            {
                msgWarn(t(L"cmd.merge.table"));
                return;
            }

            for (const auto& addr : continuousRunAddresses(sel))
            {
                VARIANT a = varBStr(addr);
                IDispatch* range = getObject(sh, L"Range", &a, 1);
                VariantClear(&a);
                if (!range) continue;
                Releaser rRange{ range };
                callMethod(range, L"Merge");
            }
        }
        catch (...) {}
    }

    // ═══════════════════ M4: 내용병합/내용감춘 병합/나누기 ═══════════════════
    namespace
    {
        bool endsWith(const std::wstring& s, const std::wstring& suf)
        {
            return !suf.empty() && s.size() >= suf.size() &&
                   s.compare(s.size() - suf.size(), suf.size(), suf) == 0;
        }

        std::vector<std::wstring> wsplit(const std::wstring& s,
                                         const std::wstring& delim)
        {
            std::vector<std::wstring> out;
            if (delim.empty()) { out.push_back(s); return out; }
            size_t pos = 0;
            for (;;)
            {
                const size_t p = s.find(delim, pos);
                if (p == std::wstring::npos) { out.push_back(s.substr(pos)); break; }
                out.push_back(s.substr(pos, p - pos));
                pos = p + delim.size();
            }
            return out;
        }

        // rng.Range("A1") — 병합 셀의 대표(첫) 셀.
        IDispatch* subRangeA1(IDispatch* rng)
        {
            VARIANT a = varBStr(L"A1");
            IDispatch* r = getObject(rng, L"Range", &a, 1);
            VariantClear(&a);
            return r;
        }

        // Range.Offset(dr, dc).
        IDispatch* rangeOffset(IDispatch* range, long dr, long dc)
        {
            VARIANT a[2];
            a[0] = varLong(dr);
            a[1] = varLong(dc);
            return getObject(range, L"Offset", a, 2);
        }

        // 구분 문자 입력 (VB InputBox EGT c0210/c0211, 기본 ","). 취소=false.
        // VB InputBox는 취소 시 ""를 돌려줘 구분자 없이 계속 진행됐지만, 여기서는
        // 취소=중단으로 처리(빈 구분자를 원하면 값을 지우고 확인하면 된다).
        bool askDelimiter(std::wstring& delim)
        {
            using egtools::i18n::t;
            delim = L",";
            return egtools::dialogs::inputText(excelHwnd(),
                       t(L"cmd.merge.sepTitle"), t(L"cmd.merge.sepPrompt"), delim);
        }

        // 선택 주소 문자열 목록 (VB sel.Areas(i).Address 수집).
        std::vector<std::wstring> areaAddresses(IDispatch* sel)
        {
            std::vector<std::wstring> out;
            IDispatch* areas = getObject(sel, L"Areas");
            if (!areas) return out;
            Releaser ra{ areas };
            const long n = getLong(areas, L"Count", 0);
            for (long i = 1; i <= n; ++i)
            {
                IDispatch* area = getObjectIdx(areas, L"Item", i);
                if (!area) continue;
                Releaser r{ area };
                out.push_back(getBStr(area, L"Address"));
            }
            return out;
        }

        // 선택 안에 병합된 셀(MergeArea.Cells.Count > 1)이 있는지 (VB c0213 검사).
        bool hasMergedCells(IDispatch* sel)
        {
            bool merged = false;
            forEachCell(sel, [&](IDispatch* cell)
            {
                IDispatch* ma = getObject(cell, L"MergeArea");
                if (!ma) return true;
                Releaser rMa{ ma };
                IDispatch* cs = getObject(ma, L"Cells");
                if (!cs) return true;
                Releaser rCs{ cs };
                if (getDouble(cs, L"CountLarge", 0) > 1) { merged = true; return false; }
                return true;
            });
            return merged;
        }

        // ── 글자 단위 서식 (VB FormatList/FormatTextAs) ─────────────────────
        // Characters(i,1).Font의 11속성 (VB C02_Merge.vb:323-335).
        struct CharFmt
        {
            std::wstring name, style;
            double size = 0, color = 0, tint = 0;
            long underline = 0;
            bool bold = false, italic = false, strike = false;
            bool sub = false, super = false;
        };

        // 셀 텍스트가 Characters로 접근 가능한지 (숫자 셀 등은 실패 — VB의
        // Try Characters.Count Catch 대응).
        bool cellHasCharacters(IDispatch* cell)
        {
            IDispatch* chars = getObject(cell, L"Characters");
            if (!chars) return false;
            Releaser r{ chars };
            return getLong(chars, L"Count", 0) > 0;
        }

        // cell.Characters(idx,1).Font 읽기. hasChars=false면 셀 Font로 폴백(VB 동일).
        CharFmt charFontFmt(IDispatch* cell, bool hasChars, long idx)
        {
            CharFmt f;
            IDispatch* font = nullptr;
            if (hasChars)
            {
                if (IDispatch* ch = getObjectIdx2(cell, L"Characters", idx, 1))
                {
                    Releaser rc{ ch };
                    font = getObject(ch, L"Font");
                }
            }
            if (!font) font = getObject(cell, L"Font");
            if (!font) return f;
            Releaser rf{ font };
            f.name      = getBStr(font, L"Name");
            f.size      = getDouble(font, L"Size");
            f.color     = getDouble(font, L"Color");
            f.style     = getBStr(font, L"FontStyle");
            f.bold      = getBool(font, L"Bold");
            f.italic    = getBool(font, L"Italic");
            f.underline = getLong(font, L"Underline");
            f.strike    = getBool(font, L"Strikethrough");
            f.sub       = getBool(font, L"Subscript");
            f.super     = getBool(font, L"Superscript");
            f.tint      = getDouble(font, L"TintAndShade");
            return f;
        }

        // rng.Characters(idx,1).Font에 저장된 서식 적용 (VB FormatTextAs 본문).
        void applyCharFmt(IDispatch* rng, long idx, const CharFmt& f)
        {
            IDispatch* ch = getObjectIdx2(rng, L"Characters", idx, 1);
            if (!ch) return;
            Releaser rc{ ch };
            IDispatch* font = getObject(ch, L"Font");
            if (!font) return;
            Releaser rf{ font };
            putBStr(font, L"Name", f.name);
            putDouble(font, L"Size", f.size);
            putDouble(font, L"Color", f.color);
            putBStr(font, L"FontStyle", f.style);
            putBool(font, L"Bold", f.bold);
            putBool(font, L"Italic", f.italic);
            putLong(font, L"Underline", f.underline);
            putBool(font, L"Strikethrough", f.strike);
            // Subscript/Superscript는 배타적 — 하나만 설정해야 반영됨(VB 402-410)
            if (f.sub)
                putBool(font, L"Subscript", true);
            else if (f.super)
                putBool(font, L"Superscript", true);
            else
            {
                putBool(font, L"Subscript", false);
                putBool(font, L"Superscript", false);
            }
            putDouble(font, L"TintAndShade", f.tint);
        }

        // 저장된 글자 서식을 rng 텍스트에 재적용 (VB FormatTextAs, C02_Merge.vb:386).
        // nullopt 슬롯(구분자 위치)은 건너뛴다. Characters 루프는 느리므로 상태
        // 표시줄에 진행을 표시한다.
        void formatTextAs(IDispatch* ad, IDispatch* rng,
                          const std::vector<std::optional<CharFmt>>& fmts)
        {
            long count = 0;
            if (IDispatch* chars = getObject(rng, L"Characters"))
            {
                Releaser r{ chars };
                count = getLong(chars, L"Count", 0);
            }
            if (count > (long)fmts.size()) count = (long)fmts.size();
            for (long i = 1; i <= count; ++i)
            {
                if (i % 20 == 0)
                    statusBar(ad, std::to_wstring(i) + L"/" + std::to_wstring(count) +
                                  L" formatting...");
                if (!fmts[(size_t)i - 1]) continue;
                applyCharFmt(rng, i, *fmts[(size_t)i - 1]);
            }
        }

        // ── MergeJoin (VB MergeJoin, C02_Merge.vb:276) ──────────────────────
        // 사전검사 (VB 288-291). 통과=true, 실패는 안내 후 false.
        bool mergeJoinChecks(IDispatch* rng)
        {
            using egtools::i18n::t;
            {
                IDispatch* areas = getObject(rng, L"Areas");
                Releaser ra{ areas };
                // VB c0205/c0206("하나의 구역만 통합 가능") — 전용 키가 없어 needRange 사용
                if (!areas || getLong(areas, L"Count", 0) != 1)
                {
                    msgWarn(t(L"cmd.merge.needRange"));
                    return false;
                }
            }
            {
                IDispatch* rows = getObject(rng, L"Rows");
                Releaser rr{ rows };
                if (rows && getLong(rows, L"Count", 0) > 100)
                {
                    msgWarn(t(L"cmd.merge.tooManyRows"));   // VB c0207
                    return false;
                }
            }
            {
                // VB c0208("병합할 셀이 하나 뿐입니다") — 선택 전체가 이미 한 병합 셀
                IDispatch* c11 = getObjectIdx2(rng, L"Cells", 1, 1);
                if (c11)
                {
                    Releaser rc{ c11 };
                    IDispatch* ma = getObject(c11, L"MergeArea");
                    if (ma)
                    {
                        Releaser rm{ ma };
                        if (getBStr(ma, L"Address") == getBStr(rng, L"Address"))
                        {
                            msgWarn(t(L"cmd.merge.alreadyMerged"));
                            return false;
                        }
                    }
                }
            }
            return true;
        }

        // MergeJoin 본체 (VB 304-377). 행 안은 delimC, 행 사이는 \n으로 이어
        // 붙여 하나의 병합 셀에 넣는다. keepFmt=true면 글자 단위 Font 11속성을
        // 저장했다가 병합 후 재적용.
        void mergeJoinCore(IDispatch* ad, IDispatch* rng, bool keepFmt,
                           const std::wstring& delimC)
        {
            const std::wstring delimR = L"\n";     // 행 구분은 \n 고정(VB delimR=vbLf)
            std::wstring text;
            std::vector<std::optional<CharFmt>> fmts;   // text와 1:1 정렬

            IDispatch* rows = getObject(rng, L"Rows");
            if (!rows) return;
            Releaser rRows{ rows };
            const long nRows = getLong(rows, L"Count", 0);
            for (long r = 1; r <= nRows; ++r)
            {
                if (keepFmt)
                    statusBar(ad, std::to_wstring(r) + L"/" + std::to_wstring(nRows) +
                                  L" scanning formats...");
                IDispatch* row = getObjectIdx(rows, L"Item", r);
                if (!row) continue;
                Releaser rRow{ row };
                IDispatch* cells = getObject(row, L"Cells");
                if (!cells) continue;
                Releaser rCells{ cells };
                const long nCells = (long)getDouble(cells, L"CountLarge", 0);
                for (long i = 1; i <= nCells; ++i)
                {
                    IDispatch* cell = getObjectIdx(cells, L"Item", i);
                    if (!cell) continue;
                    Releaser rc{ cell };
                    VARIANT v; VariantInit(&v);
                    getVar(cell, L"Value2", &v);
                    const std::wstring s = valToStr(v);
                    VariantClear(&v);
                    // 빈 셀은 텍스트/서식 둘 다 건너뜀 (VB는 빈 셀에도 구분자
                    // 서식 슬롯을 추가해 이후 서식이 밀리는 버그 — 정렬 유지 수정)
                    if (s.empty()) continue;

                    text += s;
                    if (keepFmt)
                    {
                        const bool hasChars = cellHasCharacters(cell);
                        for (size_t c = 1; c <= s.size(); ++c)
                            fmts.emplace_back(charFontFmt(cell, hasChars, (long)c));
                    }
                    text += delimC;
                    if (keepFmt)
                        fmts.insert(fmts.end(), delimC.size(), std::nullopt);
                }
                // 행 끝의 남은 구분자 제거 (VB 346-351)
                if (endsWith(text, delimC))
                {
                    text.resize(text.size() - delimC.size());
                    if (keepFmt) fmts.resize(fmts.size() - delimC.size());
                }
                text += delimR;
                if (keepFmt)
                    fmts.insert(fmts.end(), delimR.size(), std::nullopt);
            }
            // 마지막 행 구분자 제거 (VB 359-364)
            if (endsWith(text, delimR))
            {
                text.resize(text.size() - delimR.size());
                if (keepFmt) fmts.resize(fmts.size() - delimR.size());
            }

            callMethod(rng, L"ClearContents");
            callMethod(rng, L"Merge");
            if (text.empty()) return;

            putBStr(rng, L"NumberFormat", L"@");   // 텍스트 서식(VB 369)
            IDispatch* a1 = subRangeA1(rng);
            if (!a1) return;
            Releaser rA1{ a1 };
            VARIANT val = varBStr(text);
            putVar(a1, L"Value", val);
            VariantClear(&val);
            if (keepFmt)
                formatTextAs(ad, a1, fmts);
            else
            {
                putBStr(rng, L"Style", L"Normal");
                putBool(rng, L"WrapText", true);
            }
        }

        // MergeJoinText/All 진입부 (VB MergeJoinText/All + MergeJoin 가드).
        void mergeJoinEntry(bool keepFmt)
        {
            using egtools::i18n::t;
            try
            {
                IDispatch* ad = appDisp();
                if (!ad) return;
                Releaser rApp{ ad };
                IDispatch* sel = selectionRange(ad);
                if (!sel) { msgWarn(t(L"cmd.merge.needRange")); return; }   // VB c0201
                Releaser rSel{ sel };

                IDispatch* sh = getObject(ad, L"ActiveSheet");
                if (!sh) return;
                Releaser rSh{ sh };
                if (intersectsTable(ad, sh, sel))
                {
                    msgWarn(t(L"cmd.merge.table"));                          // VB c0203
                    return;
                }
                if (!mergeJoinChecks(sel)) return;

                // 열이 2개 이상이면 세로칸 구분자 질의 (VB 304)
                std::wstring delimC = L",";
                {
                    IDispatch* cols = getObject(sel, L"Columns");
                    Releaser rc{ cols };
                    if (cols && getLong(cols, L"Count", 0) > 1 &&
                        !askDelimiter(delimC))
                        return;
                }

                xloil::PauseExcel pause(xloil::thisApp());
                mergeJoinCore(ad, sel, keepFmt, delimC);
                statusBarClear(ad);
            }
            catch (...) {}
        }

        // MergeRow/MergeColumn 진입부 (VB MergeRow/MergeColumn, C02_Merge.vb:427,469)
        // 각 행(열)마다 MergeJoin을 반복한다. VB는 병합 셀 검사를 단위마다 직전에
        // 했지만 여기서는 전체를 사전검사해 하나라도 있으면 시작 전에 중단한다.
        // 구분자 질의도 VB는 행마다 InputBox를 반복하지만 한 번만 물어 재사용(UX 개선).
        void mergeUnitsEntry(bool keepFmt, bool byRows)
        {
            using egtools::i18n::t;
            try
            {
                IDispatch* ad = appDisp();
                if (!ad) return;
                Releaser rApp{ ad };
                IDispatch* sel = selectionRange(ad);
                if (!sel) { msgWarn(t(L"cmd.merge.needRange")); return; }   // VB c0201
                Releaser rSel{ sel };

                IDispatch* sh = getObject(ad, L"ActiveSheet");
                if (!sh) return;
                Releaser rSh{ sh };
                if (intersectsTable(ad, sh, sel))
                {
                    msgWarn(t(L"cmd.merge.table"));                          // VB c0203
                    return;
                }
                if (hasMergedCells(sel))
                {
                    msgWarn(t(L"cmd.merge.alreadyMerged"));                  // VB c0213
                    return;
                }

                // 행끼리 병합은 단위(행)가 여러 열일 때 구분자 필요. 열끼리 병합의
                // 단위(열)는 항상 1열이라 VB에서도 질의가 발생하지 않는다.
                std::wstring delimC = L",";
                if (byRows)
                {
                    IDispatch* cols = getObject(sel, L"Columns");
                    Releaser rc{ cols };
                    if (cols && getLong(cols, L"Count", 0) > 1 &&
                        !askDelimiter(delimC))
                        return;
                }

                xloil::PauseExcel pause(xloil::thisApp());
                IDispatch* units = getObject(sel, byRows ? L"Rows" : L"Columns");
                if (!units) return;
                Releaser rUnits{ units };
                const long n = getLong(units, L"Count", 0);
                for (long i = 1; i <= n; ++i)
                {
                    statusBar(ad, std::to_wstring(i) + L"/" + std::to_wstring(n) +
                                  L" merging...");
                    IDispatch* unit = getObjectIdx(units, L"Item", i);
                    if (!unit) continue;
                    Releaser rU{ unit };
                    // VB는 실패해도 다음 단위로 계속(단위마다 MsgBox 반복) —
                    // 여기서는 첫 실패에서 중단한다.
                    if (!mergeJoinChecks(unit)) break;
                    mergeJoinCore(ad, unit, keepFmt, delimC);
                }
                statusBarClear(ad);
            }
            catch (...) {}
        }

        // ── FakeMerge (VB UnmergeAndFill/DoFakeMerge, C02_Merge.vb:194-264) ──
        // rng 안의 병합 셀만 해제하고 각 병합 영역의 첫 열에 값을 채운다
        // (splitFill과 달리 일반 셀의 수식/값은 건드리지 않음).
        void unmergeAndFill(IDispatch* rng)
        {
            forEachCell(rng, [](IDispatch* cell)
            {
                if (!getBool(cell, L"MergeCells")) return true;
                IDispatch* ma = getObject(cell, L"MergeArea");
                if (!ma) return true;
                Releaser rMa{ ma };
                VARIANT v; VariantInit(&v);
                if (IDispatch* c11 = getObjectIdx2(ma, L"Cells", 1, 1))
                {
                    Releaser rc{ c11 };
                    getVar(c11, L"Value2", &v);
                }
                if (IDispatch* col1 = getObjectIdx(ma, L"Columns", 1))
                {
                    Releaser rCol{ col1 };
                    callMethod(col1, L"UnMerge");
                    if (!isEmptyVal(v)) putVar(col1, L"Value", v);
                }
                VariantClear(&v);
                return true;
            });
        }

        // 각 주소 영역을 임시 워크북(시트 복사본)에서 실제 병합한 뒤 그 "서식만"
        // 원본에 붙여넣어, 값을 유지한 채 병합된 것처럼 보이게 한다.
        //   cloneRuns=false : 각 영역 전체를 하나로 병합 (내용감춘 병합)
        //   cloneRuns=true  : 각 영역을 나누고 채우기 후 세로 연속값 구간별로
        //                     계층형 병합 (연속값 감춘 병합)
        // 새 워크북으로 처리하므로 원본 워크북의 시트 구성을 건드리지 않는다.
        void doFakeMerge(IDispatch* ad, IDispatch* srcWs,
                         const std::vector<std::wstring>& addrs, bool cloneRuns)
        {
            if (!srcWs || addrs.empty()) return;

            // 현재 시트를 인수 없이 복사 → 시트 1장짜리 새 워크북 생성·활성화(VB 212)
            if (!callMethod(srcWs, L"Copy")) return;
            IDispatch* tmpWs = getObject(ad, L"ActiveSheet");
            if (!tmpWs) return;
            Releaser rTmpWs{ tmpWs };

            // RAII: 어느 경로로 나가도 임시 워크북을 저장 없이 닫고 원본으로 복귀
            struct TmpWbCloser
            {
                IDispatch* ad; IDispatch* wb; IDispatch* srcWs;
                ~TmpWbCloser()
                {
                    putBool(ad, L"CutCopyMode", false);
                    if (wb)
                    {
                        VARIANT a; VariantInit(&a);
                        a.vt = VT_BOOL; a.boolVal = VARIANT_FALSE;   // SaveChanges:=False
                        invokeRaw(wb, L"Close", DISPATCH_METHOD, nullptr, &a, 1);
                        wb->Release();
                    }
                    callMethod(srcWs, L"Activate");
                }
            } closer{ ad, getObject(tmpWs, L"Parent"), srcWs };

            for (const auto& addr : addrs)
            {
                statusBar(ad, addr + L" processing...");
                VARIANT a = varBStr(addr);
                IDispatch* tmpArea = getObject(tmpWs, L"Range", &a, 1);
                VariantClear(&a);
                if (!tmpArea) continue;
                Releaser rTa{ tmpArea };

                if (cloneRuns)
                {
                    // 복사본에서 나누고 채우기 후 연속값 구간별 실제 병합(VB 221-224)
                    unmergeAndFill(tmpArea);
                    for (const auto& run : continuousRunAddresses(tmpArea))
                    {
                        VARIANT ra = varBStr(run);
                        IDispatch* r = getObject(tmpWs, L"Range", &ra, 1);
                        VariantClear(&ra);
                        if (!r) continue;
                        Releaser rr{ r };
                        callMethod(r, L"Merge");
                    }
                }
                else
                {
                    callMethod(tmpArea, L"UnMerge");
                    callMethod(tmpArea, L"Merge");
                }

                // 병합된 서식만 원본의 동일 영역에 붙여넣기 — 값은 유지(VB 232-233)
                callMethod(tmpArea, L"Copy");
                VARIANT sa = varBStr(addr);
                IDispatch* srcArea = getObject(srcWs, L"Range", &sa, 1);
                VariantClear(&sa);
                if (srcArea)
                {
                    Releaser rSa{ srcArea };
                    VARIANT p = varLong(-4122);        // xlPasteFormats
                    callMethod(srcArea, L"PasteSpecial", &p, 1);
                }
            }
            putBool(ad, L"CutCopyMode", false);
            statusBarClear(ad);
        }

        // FakeMerge/FakeMergeClone 진입부 (VB C02_Merge.vb:49,86).
        void fakeMergeEntry(bool cloneRuns)
        {
            using egtools::i18n::t;
            try
            {
                IDispatch* ad = appDisp();
                if (!ad) return;
                Releaser rApp{ ad };
                IDispatch* sel = selectionRange(ad);
                if (!sel) { msgWarn(t(L"cmd.merge.needRange")); return; }   // VB c0201
                Releaser rSel{ sel };

                IDispatch* sh = getObject(ad, L"ActiveSheet");
                if (!sh) return;
                Releaser rSh{ sh };

                if (cloneRuns)    // 연속값 감춘 병합만 추가 가드(VB FakeMergeClone)
                {
                    IDispatch* rows = getObject(sel, L"Rows");
                    Releaser rr{ rows };
                    if (rows && getLong(rows, L"Count", 0) == 1)
                    {
                        msgWarn(t(L"cmd.merge.oneRow"));                     // VB c0202
                        return;
                    }
                    if (intersectsTable(ad, sh, sel))
                    {
                        msgWarn(t(L"cmd.merge.table"));                      // VB c0203
                        return;
                    }
                }

                const auto addrs = areaAddresses(sel);
                if (addrs.empty()) return;

                xloil::PauseExcel pause(xloil::thisApp());
                doFakeMerge(ad, sh, addrs, cloneRuns);
            }
            catch (...) {}
        }

        // ── SplitRows/SplitColumns (VB C02_Merge.vb:538,642) ────────────────
        // 병합 셀 하나의 텍스트를 구분자로 나눠 행(열)으로 분배한다. 분배할
        // 조각이 병합 행(열) 수보다 많으면 행(열)을 삽입한다.
        //   byRows=true  : \n 기준 행 분배 + EntireRow.Insert/AutoFit
        //   byRows=false : 구분자 질의(기본 ",") 열 분배 + EntireColumn.Insert/AutoFit
        void splitImpl(bool keepFmt, bool byRows)
        {
            using egtools::i18n::t;
            try
            {
                IDispatch* ad = appDisp();
                if (!ad) return;
                Releaser rApp{ ad };
                IDispatch* rng = selectionRange(ad);
                if (!rng) { msgWarn(t(L"cmd.merge.needRange")); return; }
                Releaser rRng{ rng };

                // "오직 하나의 (병합)셀만 처리 가능" (VB c0214 — 전용 키가 없어
                // needRange 사용). SplitRows는 셀 2개 이상이면, SplitColumns는
                // 선택이 첫 셀의 병합 영역과 다르면 전체가 한 병합 셀이어야 한다.
                bool needMergedCheck = false;
                if (byRows)
                {
                    IDispatch* cells = getObject(rng, L"Cells");
                    Releaser rc{ cells };
                    needMergedCheck =
                        cells && getDouble(cells, L"CountLarge", 0) > 1;
                }
                else
                {
                    std::wstring maAddr;
                    if (IDispatch* c11 = getObjectIdx2(rng, L"Cells", 1, 1))
                    {
                        Releaser rc{ c11 };
                        if (IDispatch* ma = getObject(c11, L"MergeArea"))
                        {
                            Releaser rm{ ma };
                            maAddr = getBStr(ma, L"Address");
                        }
                    }
                    needMergedCheck = getBStr(rng, L"Address") != maAddr;
                }
                if (needMergedCheck)
                {
                    VARIANT mc; VariantInit(&mc);
                    getVar(rng, L"MergeCells", &mc);   // Null(혼합)/False 모두 거부
                    const bool merged =
                        mc.vt == VT_BOOL && mc.boolVal == VARIANT_TRUE;
                    VariantClear(&mc);
                    if (!merged) { msgWarn(t(L"cmd.merge.needRange")); return; }
                }

                std::wstring delim = L"\n";   // 행 나누기는 \n 고정(VB vbLf)
                if (!byRows)
                {
                    if (!askDelimiter(delim)) return;   // VB c0210/c0211
                    if (delim.empty()) return;          // VB: Len(delimC)=0 → Exit Sub
                }

                long mergedUnits = 0;
                {
                    IDispatch* u = getObject(rng, byRows ? L"Rows" : L"Columns");
                    if (u) { Releaser r{ u }; mergedUnits = getLong(u, L"Count", 0); }
                }

                std::wstring text;
                {
                    IDispatch* a1 = subRangeA1(rng);
                    if (!a1) return;
                    Releaser rA1{ a1 };
                    VARIANT v; VariantInit(&v);
                    getVar(a1, L"Value2", &v);
                    text = valToStr(v);
                    VariantClear(&v);
                }
                const auto parts = wsplit(text, delim);
                // 나눌 조각이 하나뿐 (VB c0204/c0209 — 전용 키가 없어 oneRow 사용)
                if (parts.size() <= 1) { msgWarn(t(L"cmd.merge.oneRow")); return; }

                xloil::PauseExcel pause(xloil::thisApp());

                // 조각별 글자 서식 수집 — iCnt는 원본 텍스트에서의 위치.
                // (VB는 빈 조각에서 구분자 위치를 건너뛰지 않고, 열 나누기는 다문자
                // 구분자에서 +1만 해 서식이 밀리는 버그 — 항상 구분자 길이만큼 전진)
                std::vector<std::vector<std::optional<CharFmt>>> partFmts;
                if (keepFmt)
                {
                    IDispatch* a1 = subRangeA1(rng);
                    Releaser rA1{ a1 };
                    const bool hasChars = a1 && cellHasCharacters(a1);
                    long iCnt = 0;
                    for (size_t p = 0; p < parts.size(); ++p)
                    {
                        statusBar(ad, std::to_wstring(p + 1) + L"/" +
                                      std::to_wstring(parts.size()) +
                                      L" scanning formats...");
                        std::vector<std::optional<CharFmt>> fa;
                        fa.reserve(parts[p].size());
                        for (size_t c = 0; c < parts[p].size(); ++c)
                        {
                            ++iCnt;
                            if (a1) fa.emplace_back(charFontFmt(a1, hasChars, iCnt));
                        }
                        partFmts.push_back(std::move(fa));
                        iCnt += (long)delim.size();    // 구분자 건너뜀
                    }
                }

                callMethod(rng, L"ClearContents");
                if (getBool(rng, L"MergeCells")) callMethod(rng, L"UnMerge");

                const wchar_t* entire = byRows ? L"EntireRow" : L"EntireColumn";
                for (size_t i = 0; i < parts.size(); ++i)
                {
                    statusBar(ad, std::to_wstring(i + 1) + L"/" +
                                  std::to_wstring(parts.size()) + L" splitting...");
                    // 삽입으로 좌표가 밀리므로 VB처럼 매 단계 Offset을 다시 구한다
                    auto offsetTgt = [&]() -> IDispatch*
                    {
                        IDispatch* a1 = subRangeA1(rng);
                        if (!a1) return nullptr;
                        Releaser rA1{ a1 };
                        return rangeOffset(a1, byRows ? (long)i : 0,
                                               byRows ? 0 : (long)i);
                    };
                    if ((long)i >= mergedUnits)    // 조각이 더 많음 → 행(열) 삽입
                    {
                        if (IDispatch* t0 = offsetTgt())
                        {
                            Releaser rT0{ t0 };
                            if (IDispatch* e = getObject(t0, entire))
                            {
                                Releaser re{ e };
                                callMethod(e, L"Insert");
                            }
                        }
                    }
                    IDispatch* tgt = offsetTgt();
                    if (!tgt) continue;
                    Releaser rT{ tgt };
                    putBStr(tgt, L"NumberFormat", L"@");
                    VARIANT val = varBStr(parts[i]);
                    putVar(tgt, L"Value", val);
                    VariantClear(&val);
                    if (IDispatch* e = getObject(tgt, entire))
                    {
                        Releaser re{ e };
                        callMethod(e, L"AutoFit");
                    }
                    if (keepFmt && i < partFmts.size())
                        formatTextAs(ad, tgt, partFmts[i]);
                }
                statusBarClear(ad);
            }
            catch (...) {}
        }
    }

    // 서식없이/서식대로 내용병합 (VB MergeJoinText/All, C02_Merge.vb:266,271)
    void mergeJoinText() { mergeJoinEntry(false); }
    void mergeJoinAll()  { mergeJoinEntry(true); }

    // 서식없이/서식대로 행끼리 병합 (VB MergeRowText/All, C02_Merge.vb:459,464)
    void mergeRowText() { mergeUnitsEntry(false, true); }
    void mergeRowAll()  { mergeUnitsEntry(true, true); }

    // 서식없이/서식대로 열끼리 병합 (VB MergeColumnText/All, C02_Merge.vb:417,422)
    void mergeColumnText() { mergeUnitsEntry(false, false); }
    void mergeColumnAll()  { mergeUnitsEntry(true, false); }

    // 내용감춘 병합 / 연속값 감춘 병합 (VB FakeMerge/FakeMergeClone, C02_Merge.vb:49,86)
    void fakeMerge()      { fakeMergeEntry(false); }
    void fakeMergeClone() { fakeMergeEntry(true); }

    // 서식없이/서식대로 행 나누기 (VB SplitRowText/All, C02_Merge.vb:530,534)
    void splitRowText() { splitImpl(false, true); }
    void splitRowAll()  { splitImpl(true, true); }

    // 서식없이/서식대로 열 나누기 (VB SplitColumnText/All, C02_Merge.vb:633,638)
    void splitColumnText() { splitImpl(false, false); }
    void splitColumnAll()  { splitImpl(true, false); }
}
