// CmdForm.cpp — 양식 명령 (EGToolsVB C05_MailMerge.vb MakeLabels 포팅, plan/23).
// 라벨지 만들기: 양식영역의 {{필드명}}을 목록영역 데이터로 치환하며 새 시트에
// 양식을 반복 복사한다(서식 포함, 열너비/행높이 복제).

#include "Commands.h"
#include "CmdCommon.h"
#include "Dialogs.h"

#include <filesystem>
#include <map>
#include <optional>
#include <regex>
#include <set>
#include <string>
#include <vector>

namespace egtools::commands
{
    namespace
    {
        std::wstring upperCase(std::wstring s)
        {
            for (auto& ch : s) ch = towupper(ch);
            return s;
        }

        // VB GetSheetName(NewName:=True) 대응 — 중복 시 base_1, base_2 …
        std::wstring newSheetName(IDispatch* wb, const std::wstring& base)
        {
            std::set<std::wstring> names;
            if (IDispatch* sheets = getObject(wb, L"Sheets"))
            {
                Releaser rs{ sheets };
                const long n = getLong(sheets, L"Count", 0);
                for (long i = 1; i <= n; ++i)
                {
                    IDispatch* sh = getObjectIdx(sheets, L"Item", i);
                    if (!sh) continue;
                    Releaser r{ sh };
                    names.insert(upperCase(getBStr(sh, L"Name")));
                }
            }
            if (!names.count(upperCase(base))) return base;
            for (int i = 1; ; ++i)
            {
                std::wstring cand = base + L"_" + std::to_wstring(i);
                if (!names.count(upperCase(cand))) return cand;
            }
        }

        void replaceAll(std::wstring& s, const std::wstring& from,
                        const std::wstring& to)
        {
            if (from.empty()) return;
            size_t pos = 0;
            while ((pos = s.find(from, pos)) != std::wstring::npos)
            {
                s.replace(pos, from.size(), to);
                pos += to.size();
            }
        }

        // 현재 Selection.CurrentRegion 주소(양식/목록 InputBox 기본값, VB 동일).
        std::wstring currentRegionAddr(IDispatch* ad)
        {
            IDispatch* sel = selectionRange(ad);
            if (!sel) return L"";
            Releaser rs{ sel };
            IDispatch* cr = getObject(sel, L"CurrentRegion");
            if (!cr) return L"";
            Releaser rc{ cr };
            return getBStr(cr, L"Address");
        }
    }

    // 라벨지 만들기 (VB MakeLabels, C05_MailMerge.vb:274)
    // 양식영역 → {{필드명}} 셀만 기억, 목록영역 1행 = 필드명(빈 헤더는 EmptyN).
    // 새 시트에 양식을 가로 CopyCols개씩 반복 복사하며 필드를 치환한다.
    void makeLabels()
    {
        using egtools::i18n::t;
        try
        {
            const bool ko = egtools::i18n::current() == L"ko";

            IDispatch* ad = appDisp();
            if (!ad) return;
            Releaser rApp{ ad };
            IDispatch* sh = getObject(ad, L"ActiveSheet");
            if (!sh) return;
            Releaser rSh{ sh };
            IDispatch* wb = getObject(sh, L"Parent");
            if (!wb) return;
            Releaser rWb{ wb };

            // 1) 양식영역
            IDispatch* formRng = inputBoxRange(ad, t(L"cmd.labels.formPrompt"),
                                               t(L"cmd.labels.formTitle"),
                                               currentRegionAddr(ad));
            if (!formRng) return;                   // 취소 → 조용히 종료
            Releaser rForm{ formRng };

            // 양식 값: {{...}}가 든 셀만 템플릿으로 보존(그 외는 복사 서식만 사용)
            static const std::wregex fieldRe(LR"(\{\{[^}]*\}\})");
            long formRows = 0, formCols = 0;
            std::vector<std::vector<std::optional<std::wstring>>> form;
            {
                SafeArr2D vals;
                if (vals.load(formRng, L"Value2"))
                {
                    formRows = vals.r2 - vals.r1 + 1;
                    formCols = vals.c2 - vals.c1 + 1;
                    form.assign((size_t)formRows + 1,
                                std::vector<std::optional<std::wstring>>((size_t)formCols + 1));
                    for (long r = 1; r <= formRows; ++r)
                        for (long c = 1; c <= formCols; ++c)
                        {
                            const VARIANT& v = vals.at(vals.r1 + r - 1, vals.c1 + c - 1);
                            if (v.vt != VT_BSTR || !v.bstrVal) continue;
                            const std::wstring s = v.bstrVal;
                            if (std::regex_search(s, fieldRe)) form[r][c] = s;
                        }
                }
                else
                {
                    // 단일 셀 양식(스칼라). 빈 셀 하나면 종료(VB c0518 관례).
                    VARIANT sv; VariantInit(&sv);
                    getVar(formRng, L"Value2", &sv);
                    const bool empty = isEmptyVal(sv);
                    std::wstring s = (sv.vt == VT_BSTR && sv.bstrVal) ? sv.bstrVal : L"";
                    VariantClear(&sv);
                    if (empty) return;
                    formRows = formCols = 1;
                    form.assign(2, std::vector<std::optional<std::wstring>>(2));
                    if (std::regex_search(s, fieldRe)) form[1][1] = s;
                }
            }

            // 2) 목록영역(1행 = 필드명 헤더 + 데이터 행)
            IDispatch* listRng = inputBoxRange(ad, t(L"cmd.labels.listPrompt"),
                                               t(L"cmd.labels.listTitle"),
                                               currentRegionAddr(ad));
            if (!listRng) return;
            Releaser rList{ listRng };

            long listRows = 0, listCols = 0;
            std::vector<std::vector<std::wstring>> list;    // [1..rows][1..cols]
            {
                SafeArr2D vals;
                if (!vals.load(listRng, L"Value2")) return; // 단일 셀 목록 → 종료
                listRows = vals.r2 - vals.r1 + 1;
                listCols = vals.c2 - vals.c1 + 1;
                list.assign((size_t)listRows + 1,
                            std::vector<std::wstring>((size_t)listCols + 1));
                for (long r = 1; r <= listRows; ++r)
                    for (long c = 1; c <= listCols; ++c)
                    {
                        const VARIANT& v = vals.at(vals.r1 + r - 1, vals.c1 + c - 1);
                        if (v.vt == VT_ERROR || isEmptyVal(v)) continue;
                        list[r][c] = varToString(v);
                    }
            }

            // 3) 가로로 배치할 라벨 수
            double acrossD = 1;
            if (!egtools::dialogs::inputNumber(excelHwnd(),
                    t(L"cmd.labels.acrossTitle"), t(L"cmd.labels.acrossPrompt"),
                    acrossD, 1, 10, 1))
                return;
            const long copyCols = (long)acrossD;

            xloil::PauseExcel pause(xloil::thisApp());

            // 필드명 → 열 번호(빈 헤더는 EmptyN — VB 동일)
            std::map<std::wstring, long> fields;
            long emptyCol = 0;
            for (long c = 1; c <= listCols; ++c)
            {
                std::wstring name = list[1][c];
                if (name.empty()) name = L"Empty" + std::to_wstring(++emptyCol);
                fields.emplace(name, c);            // 중복 헤더는 첫 열 우선
            }

            // 4) 새 시트(양식 시트 이름 + _n)
            std::wstring baseName;
            if (IDispatch* parent = getObject(formRng, L"Parent"))
            { Releaser rp{ parent }; baseName = getBStr(parent, L"Name"); }
            const std::wstring newName = newSheetName(wb, baseName);

            IDispatch* sheets = getObject(wb, L"Sheets");
            if (!sheets) return;
            Releaser rSheets{ sheets };
            VARIANT addR; VariantInit(&addR);
            if (!callMethod(sheets, L"Add", nullptr, 0, &addR)) return;
            if (addR.vt != VT_DISPATCH || !addR.pdispVal) { VariantClear(&addR); return; }
            IDispatch* newSh = addR.pdispVal;       // 소유권 이전
            Releaser rNew{ newSh };
            putBStr(newSh, L"Name", newName);

            // 열 너비를 양식과 동일하게 복제
            for (long i = 0; i < copyCols; ++i)
                for (long j = 1; j <= formCols; ++j)
                {
                    IDispatch* src = getObjectIdx(formRng, L"Columns", j);
                    if (!src) continue;
                    Releaser rs{ src };
                    IDispatch* dst = getObjectIdx(newSh, L"Columns", i * formCols + j);
                    if (!dst) continue;
                    Releaser rd{ dst };
                    putDouble(dst, L"ColumnWidth", getDouble(src, L"ColumnWidth"));
                }

            // 5) 목록 2행부터 라벨 복사+치환
            long iRow = 1, iCol = 1;
            for (long iList = 2; iList <= listRows; ++iList)
            {
                if ((iList - 2) % copyCols == 0)
                {
                    iCol = 1;
                    if (iList > 2) iRow += formRows;
                    // 행 높이를 양식과 동일하게 복제
                    for (long iR = 1; iR <= formRows; ++iR)
                    {
                        IDispatch* src = getObjectIdx(formRng, L"Rows", iR);
                        if (!src) continue;
                        Releaser rs{ src };
                        IDispatch* dst = getObjectIdx(newSh, L"Rows", iRow - 1 + iR);
                        if (!dst) continue;
                        Releaser rd{ dst };
                        putDouble(dst, L"RowHeight", getDouble(src, L"RowHeight"));
                    }
                }
                else
                    iCol += formCols;

                // 양식 복사(서식 포함)
                if (IDispatch* dest = getObjectIdx2(newSh, L"Cells", iRow, iCol))
                {
                    Releaser rd{ dest };
                    VARIANT da; VariantInit(&da);
                    da.vt = VT_DISPATCH; da.pdispVal = dest;
                    callMethod(formRng, L"Copy", &da, 1);
                }

                // {{필드명}} 치환
                for (long iR = 1; iR <= formRows; ++iR)
                    for (long iC = 1; iC <= formCols; ++iC)
                    {
                        if (!form[iR][iC]) continue;
                        const std::wstring& tmpl = *form[iR][iC];
                        std::wstring value = tmpl;
                        for (auto it = std::wsregex_iterator(tmpl.begin(), tmpl.end(),
                                                             fieldRe);
                             it != std::wsregex_iterator(); ++it)
                        {
                            const std::wstring tokenStr = it->str();
                            const std::wstring fieldName =
                                tokenStr.substr(2, tokenStr.size() - 4);
                            auto f = fields.find(fieldName);
                            if (f == fields.end())
                            {
                                msgWarn(fieldName +
                                        (ko ? L" 필드명이 목록에 존재하지 않습니다."
                                            : L" field name does not exist in the list."));
                                return;             // VB: Throw → 중단
                            }
                            replaceAll(value, L"{{" + fieldName + L"}}",
                                       list[iList][f->second]);
                        }
                        IDispatch* cell = getObjectIdx2(newSh, L"Cells",
                                                        iRow + iR - 1, iCol + iC - 1);
                        if (!cell) continue;
                        Releaser rc{ cell };
                        VARIANT v = varBStr(value);
                        putVar(cell, L"Value", v);
                        VariantClear(&v);
                    }
            }

            msgInfo(t(L"cmd.labels.done"));
        }
        catch (...) {}
    }

    namespace
    {
        std::wstring sanitizeSheetNameMM(std::wstring s)
        {
            for (auto& ch : s)
                if (wcschr(L":\\/?*[]", ch)) ch = L'_';
            if (s.size() > 31) s.resize(31);
            return s;
        }
        std::wstring sanitizeFileNameMM(std::wstring s)
        {
            for (auto& ch : s)
                if (wcschr(L"\\/:*?\"<>|\r\n\t", ch)) ch = L'_';
            return s;
        }
        std::wstring trimWs(std::wstring s)
        {
            const auto b = s.find_first_not_of(L" \t\r\n");
            const auto e = s.find_last_not_of(L" \t\r\n");
            return b == std::wstring::npos ? L"" : s.substr(b, e - b + 1);
        }
    }

    // 메일머지 v1 (VB MailMerge, C05_MailMerge.vb:16)
    // 자료 목록의 행마다 양식 시트를 복제해 {{필드}}를 치환하고, 옵션에 따라
    // 개별 파일 저장(xlsx — 암호 열 지원 / 첫 헤더가 "PDF"면 PDF)·개별 인쇄를
    // 수행한다. 옵션 0이면 복제 시트를 남긴다(그 외에는 삭제).
    // **메일 발송(v2)은 제외**(plan/23 결정 1 — SMTP 설정은 v2에서 암호 제외
    // 레지스트리 저장 방침, §3-3). VB의 삭제 대상 시트명 버그(sName vs
    // newShName)는 교정.
    void mailMerge()
    {
        using egtools::i18n::t;
        namespace fs = std::filesystem;
        try
        {
            IDispatch* ad = appDisp();
            if (!ad) return;
            Releaser rApp{ ad };

            // 1) 옵션 (VB 0~7 중 메일 비트 제외: 0=시트 복제만, 1=파일저장,
            //    2=인쇄, 3=저장+인쇄)
            double optD = 1;
            if (!egtools::dialogs::inputNumber(excelHwnd(), t(L"cmd.mm.optTitle"),
                                               t(L"cmd.mm.optPrompt"), optD, 0, 3, 1))
                return;
            const int opt = (int)optD;
            const bool doSave = (opt & 1) != 0;
            const bool doPrint = (opt & 2) != 0;
            const bool deleteSheet = opt != 0;

            // 2) 인쇄 설정 (xlDialogPrinterSetup = 9, VB 동일)
            if (doPrint)
            {
                bool ok = false;
                if (IDispatch* dialogs = getObject(ad, L"Dialogs"))
                {
                    Releaser rd{ dialogs };
                    if (IDispatch* dlg = getObjectIdx(dialogs, L"Item", 9))
                    {
                        Releaser rDlg{ dlg };
                        VARIANT r; VariantInit(&r);
                        if (invokeRaw(dlg, L"Show", DISPATCH_METHOD, &r, nullptr, 0))
                            ok = r.vt == VT_BOOL && r.boolVal != VARIANT_FALSE;
                        VariantClear(&r);
                    }
                }
                if (!ok) { msgWarn(t(L"cmd.mm.printCancel")); return; }
            }

            // 3) 양식 시트 지정(선택한 영역의 부모 시트 전체를 양식으로)
            IDispatch* formPick = inputBoxRange(ad, t(L"cmd.mm.selectForm"),
                                                t(L"cmd.mm.formTitle"), L"");
            if (!formPick) return;
            Releaser rFormPick{ formPick };
            IDispatch* formSh = getObject(formPick, L"Parent");
            if (!formSh) return;
            Releaser rFormSh{ formSh };
            IDispatch* wb = getObject(formSh, L"Parent");
            if (!wb) return;
            Releaser rWb{ wb };
            IDispatch* formUr = getObject(formSh, L"UsedRange");
            if (!formUr) return;
            Releaser rFormUr{ formUr };

            // 4) 자료 목록(CurrentRegion, 첫 행 = 필드명)
            IDispatch* listPick = inputBoxRange(ad, t(L"cmd.mm.selectList"),
                                                t(L"cmd.mm.listTitle"), L"");
            if (!listPick) return;
            Releaser rListPick{ listPick };
            IDispatch* listCr = getObject(listPick, L"CurrentRegion");
            if (!listCr) return;
            Releaser rListCr{ listCr };

            auto& app = xloil::thisApp();
            xloil::PauseExcel pause(app);

            SafeArr2D listVals;
            if (!listVals.load(listCr, L"Value2")) return;   // 목록은 2차원이어야 함
            const long listRows = listVals.r2;
            const long listCols = listVals.c2;
            if (listRows < 2) return;

            // 헤더 → 열 (대소문자 무시, VB Match 동일)
            std::map<std::wstring, long> headers;
            for (long c = 1; c <= listCols; ++c)
                headers[upperCase(trimWs(varToString(listVals.at(1, c))))] = c;

            // 5) 양식의 {{필드}} 스캔 → 필드별 (열, 셀주소 목록)
            struct FieldRef { long col; std::vector<std::wstring> addrs; };
            std::map<std::wstring, FieldRef> fields;   // 키: 원문 필드명
            {
                SafeArr2D formVals;
                const bool arr = formVals.load(formUr, L"Value2");
                const long fr = arr ? formVals.r2 : 1;
                const long fc = arr ? formVals.c2 : 1;
                for (long r = 1; r <= fr; ++r)
                    for (long c = 1; c <= fc; ++c)
                    {
                        std::wstring v;
                        if (arr) v = varToString(formVals.at(r, c));
                        else
                        {
                            VARIANT sv; VariantInit(&sv);
                            getVar(formUr, L"Value2", &sv);
                            v = varToString(sv);
                            VariantClear(&sv);
                        }
                        size_t ls = v.find(L"{{");
                        while (ls != std::wstring::npos)
                        {
                            const size_t rs = v.find(L"}}", ls);
                            if (rs == std::wstring::npos) break;
                            const std::wstring field = v.substr(ls + 2, rs - ls - 2);
                            const auto hit = headers.find(upperCase(trimWs(field)));
                            if (hit == headers.end())
                            {
                                wchar_t buf[512];
                                swprintf_s(buf, t(L"cmd.mm.fieldMissing").c_str(),
                                           field.c_str());
                                msgWarn(buf);
                                return;
                            }
                            auto& frefEntry = fields[field];
                            frefEntry.col = hit->second;
                            IDispatch* cell = getObjectIdx2(formUr, L"Cells", r, c);
                            if (cell)
                            {
                                Releaser rc{ cell };
                                frefEntry.addrs.push_back(getBStr(cell, L"Address"));
                            }
                            ls = v.find(L"{{", ls + 1);
                        }
                    }
            }

            // 6) 암호 열(저장 시)
            long passCol = 0;
            if (doSave)
            {
                auto it = headers.find(L"PASSWORD");
                if (it == headers.end()) it = headers.find(L"암호");
                if (it != headers.end()) passCol = it->second;
            }

            const bool asPdf =
                upperCase(trimWs(varToString(listVals.at(1, 1)))) == L"PDF";
            const std::wstring wbPath = getBStr(wb, L"Path");
            if (doSave && wbPath.empty())
            {
                msgWarn(t(L"cmd.mm.outputDirFail"));
                return;
            }
            const std::wstring outDir = wbPath + L"\\Output";
            if (doSave)
            {
                std::error_code ec;
                fs::create_directories(outDir, ec);
                if (!fs::exists(outDir))
                {
                    msgWarn(t(L"cmd.mm.outputDirFail"));
                    return;
                }
            }

            // 7) 데이터 행 루프
            long made = 0;
            for (long r = 2; r <= listRows; ++r)
            {
                const std::wstring rawName = trimWs(varToString(listVals.at(r, 1)));
                if (rawName.empty()) continue;
                statusBar(ad, rawName + L" processing...");

                const std::wstring newName =
                    newSheetName(wb, sanitizeSheetNameMM(rawName));

                // 양식 시트 복제(맨 뒤로)
                {
                    IDispatch* sheets = getObject(wb, L"Sheets");
                    if (!sheets) break;
                    Releaser rs{ sheets };
                    IDispatch* last = getObjectIdx(sheets, L"Item",
                                                   getLong(sheets, L"Count", 1));
                    if (!last) break;
                    Releaser rl{ last };
                    VARIANT a[2];
                    VariantInit(&a[0]); VariantInit(&a[1]);
                    a[0].vt = VT_ERROR; a[0].scode = DISP_E_PARAMNOTFOUND;
                    a[1].vt = VT_DISPATCH; a[1].pdispVal = last;
                    if (!callMethod(formSh, L"Copy", a, 2)) break;
                }
                IDispatch* newSh = getObject(ad, L"ActiveSheet");
                if (!newSh) break;
                Releaser rNew{ newSh };
                putBStr(newSh, L"Name", newName);

                // 필드 치환
                for (const auto& [field, ref] : fields)
                    for (const auto& addr : ref.addrs)
                    {
                        VARIANT aa = varBStr(addr);
                        IDispatch* cell = getObject(newSh, L"Range", &aa, 1);
                        VariantClear(&aa);
                        if (!cell) continue;
                        Releaser rc{ cell };
                        VARIANT cv; VariantInit(&cv);
                        getVar(cell, L"Value2", &cv);
                        std::wstring text = varToString(cv);
                        VariantClear(&cv);
                        replaceAll(text, L"{{" + field + L"}}",
                                   varToString(listVals.at(r, ref.col)));
                        putBStr(cell, L"Value2", text);
                    }

                if (doPrint) callMethod(newSh, L"PrintOut");

                if (doSave)
                {
                    // 단일 시트 새 워크북으로 복사 후 저장
                    if (callMethod(newSh, L"Copy"))
                    {
                        IDispatch* outWb = getObject(ad, L"ActiveWorkbook");
                        Releaser rOut{ outWb };
                        const std::wstring fileBase = sanitizeFileNameMM(rawName);
                        if (outWb)
                        {
                            if (IDispatch* outSh = getObject(ad, L"ActiveSheet"))
                            {
                                Releaser ros{ outSh };
                                putBStr(outSh, L"Name", sanitizeSheetNameMM(rawName));
                            }
                            if (asPdf)
                            {
                                VARIANT a[2];
                                a[0] = varLong(0);           // xlTypePDF
                                a[1] = varBStr(outDir + L"\\" + fileBase + L".pdf");
                                callMethod(outWb, L"ExportAsFixedFormat", a, 2);
                                VariantClear(&a[1]);
                            }
                            else
                            {
                                // SaveAs(Filename, FileFormat=51, [Password])
                                VARIANT a[3];
                                a[0] = varBStr(outDir + L"\\" + fileBase + L".xlsx");
                                a[1] = varLong(51);
                                VariantInit(&a[2]);
                                a[2].vt = VT_ERROR; a[2].scode = DISP_E_PARAMNOTFOUND;
                                std::wstring pw;
                                if (passCol > 0)
                                    pw = varToString(listVals.at(r, passCol));
                                if (!pw.empty()) a[2] = varBStr(pw);
                                callMethod(outWb, L"SaveAs", a, 3);
                                VariantClear(&a[0]);
                                if (a[2].vt == VT_BSTR) VariantClear(&a[2]);
                            }
                            VARIANT f; VariantInit(&f);
                            f.vt = VT_BOOL; f.boolVal = VARIANT_FALSE;
                            callMethod(outWb, L"Close", &f, 1);
                        }
                    }
                }

                callMethod(wb, L"Activate");
                if (deleteSheet)
                {
                    // VB는 sName으로 삭제(중복 접미사 시 원본 오삭제 위험) —
                    // 복제된 시트(newName)를 삭제하도록 교정.
                    IDispatch* sheets = getObject(wb, L"Sheets");
                    if (sheets)
                    {
                        Releaser rs{ sheets };
                        VARIANT nm = varBStr(newName);
                        IDispatch* target = getObject(sheets, L"Item", &nm, 1);
                        VariantClear(&nm);
                        if (target)
                        {
                            Releaser rt{ target };
                            callMethod(target, L"Delete");
                        }
                    }
                }
                ++made;
            }

            statusBarClear(ad);
            callMethod(wb, L"Activate");
            callMethod(formSh, L"Activate");
            msgInfo(made > 0 ? withCount(t(L"cmd.mm.done"), made)
                             : t(L"cmd.mm.none"));
        }
        catch (...) {}
    }
}
