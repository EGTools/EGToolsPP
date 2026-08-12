// CmdPicture.cpp — 그림/사진 명령 (EGToolsVB C03_Picture.vb 포팅, plan/23).
// M2: InsertPicture / PictureFitToCell / PicturesFitToCellAll.
// M3: InsertPicturesFromFolder / InsertPicturesIntoForm. 추출 저장(M4)은 이후.

#include "Commands.h"
#include "CmdCommon.h"
#include "Dialogs.h"

#include <algorithm>
#include <cstdio>
#include <cstdlib>
#include <filesystem>
#include <map>
#include <set>
#include <vector>

namespace egtools::commands
{
    namespace
    {
        constexpr long msoPicture = 13;
        constexpr long msoLinkedPicture = 11;

        // 여백 입력(VB c0303~c0307 프롬프트). 취소/비숫자 → false.
        bool askMargin(double& margin, double maxV)
        {
            using egtools::i18n::t;
            return egtools::dialogs::inputNumber(
                excelHwnd(), t(L"cmd.pic.marginTitle"), t(L"cmd.pic.marginPrompt"),
                margin, 0.0, maxV, 0.3);
        }

        // 셰이프 컬렉션(Shapes/ShapeRange)에 그림이 하나라도 있는지.
        bool hasPicture(IDispatch* shapes)
        {
            const long n = getLong(shapes, L"Count", 0);
            for (long i = 1; i <= n; ++i)
            {
                IDispatch* shape = getObjectIdx(shapes, L"Item", i);
                if (!shape) continue;
                Releaser rs{ shape };
                const long type = getLong(shape, L"Type", 0);
                if (type == msoPicture || type == msoLinkedPicture) return true;
            }
            return false;
        }

        // 셰이프 하나를 좌상단 셀(병합 영역)에 맞춘다. 회전 90/270 보정 포함
        // (VB C03_Picture.vb:304-321과 동일 수식).
        void fitShape(IDispatch* shape, double off)
        {
            const long type = getLong(shape, L"Type", 0);
            if (type != msoPicture && type != msoLinkedPicture) return;

            IDispatch* tlc = getObject(shape, L"TopLeftCell");
            if (!tlc) return;
            Releaser rTlc{ tlc };
            IDispatch* ma = getObject(tlc, L"MergeArea");
            if (!ma) return;
            Releaser rMa{ ma };

            const double maW = getDouble(ma, L"Width");
            const double maH = getDouble(ma, L"Height");
            const double maT = getDouble(ma, L"Top");
            const double maL = getDouble(ma, L"Left");
            const double rot = getDouble(shape, L"Rotation");

            putLong(shape, L"LockAspectRatio", 0);
            if (rot == 0.0 || rot == 180.0)
            {
                putDouble(shape, L"Width", maW - off * 2);
                putDouble(shape, L"Height", maH - off * 2);
                putDouble(shape, L"Top", maT + off);
                putDouble(shape, L"Left", maL + off);
            }
            else    // 90/270도: 가로세로 교환 + 중심 보정
            {
                putDouble(shape, L"Width", maH - off * 2);
                putDouble(shape, L"Height", maW - off * 2);
                putDouble(shape, L"Top", maT + off + (maH - maW) / 2);
                putDouble(shape, L"Left", maL + off - (maH - maW) / 2);
            }
        }
    }

    // 이미지 삽입 (VB InsertPicture, C03_Picture.vb:10)
    // 파일 하나를 골라 현재 선택 영역에 0.2pt 여백으로 맞춰 삽입.
    void insertPicture()
    {
        try
        {
            std::wstring file;
            if (!egtools::dialogs::pickImageFile(excelHwnd(), file)) return;

            auto& app = xloil::thisApp();
            xloil::PauseExcel pause(app);

            IDispatch* ad = appDisp();
            if (!ad) return;
            Releaser rApp{ ad };
            IDispatch* sel = getObject(ad, L"Selection");
            if (!sel) return;
            Releaser rSel{ sel };

            const double l = getDouble(sel, L"Left");
            const double t = getDouble(sel, L"Top");
            const double w = getDouble(sel, L"Width");
            const double h = getDouble(sel, L"Height");

            IDispatch* sh = getObject(ad, L"ActiveSheet");
            if (!sh) return;
            Releaser rSh{ sh };
            IDispatch* shapes = getObject(sh, L"Shapes");
            if (!shapes) return;
            Releaser rShapes{ shapes };

            // AddPicture(Filename, LinkToFile=msoFalse, SaveWithDocument=msoCTrue,
            //            Left+0.2, Top+0.2, Width-0.4, Height-0.4)
            VARIANT a[7];
            a[0] = varBStr(file);
            a[1] = varLong(0);
            a[2] = varLong(1);
            a[3] = varDouble(l + 0.2);
            a[4] = varDouble(t + 0.2);
            a[5] = varDouble(w - 0.4);
            a[6] = varDouble(h - 0.4);
            VARIANT r; VariantInit(&r);
            const bool ok = callMethod(shapes, L"AddPicture", a, 7, &r);
            VariantClear(&a[0]);

            if (ok && r.vt == VT_DISPATCH && r.pdispVal)
            {
                IDispatch* pic = r.pdispVal;
                putLong(pic, L"LockAspectRatio", 0);
                const size_t slash = file.find_last_of(L"\\/");
                putBStr(pic, L"Name",
                        slash == std::wstring::npos ? file : file.substr(slash + 1));
            }
            VariantClear(&r);
        }
        catch (...) {}
    }

    // 선택 이미지 셀 맞춤 (VB PictureFitToCell, C03_Picture.vb:277)
    void pictureFitToCell()
    {
        using egtools::i18n::t;
        try
        {
            IDispatch* ad = appDisp();
            if (!ad) return;
            Releaser rApp{ ad };
            IDispatch* sel = getObject(ad, L"Selection");
            if (!sel) { msgInfo(t(L"cmd.pic.noneSelected")); return; }
            Releaser rSel{ sel };
            IDispatch* shapes = getObject(sel, L"ShapeRange");
            if (!shapes) { msgInfo(t(L"cmd.pic.noneSelected")); return; }
            Releaser rShapes{ shapes };
            if (!hasPicture(shapes)) { msgInfo(t(L"cmd.pic.noneSelected")); return; }

            double off = 0.3;
            if (!askMargin(off, 10.0)) return;

            auto& app = xloil::thisApp();
            xloil::PauseExcel pause(app);

            const long n = getLong(shapes, L"Count", 0);
            for (long i = 1; i <= n; ++i)
            {
                IDispatch* shape = getObjectIdx(shapes, L"Item", i);
                if (!shape) continue;
                Releaser rs{ shape };
                fitShape(shape, off);
            }
        }
        catch (...) {}
    }

    // 시트 전체 이미지 셀 맞춤 (VB PicturesFitToCellAll, C03_Picture.vb:327)
    void pictureFitAll()
    {
        using egtools::i18n::t;
        try
        {
            IDispatch* ad = appDisp();
            if (!ad) return;
            Releaser rApp{ ad };
            IDispatch* sh = getObject(ad, L"ActiveSheet");
            if (!sh) return;
            Releaser rSh{ sh };
            IDispatch* shapes = getObject(sh, L"Shapes");
            if (!shapes) return;
            Releaser rShapes{ shapes };
            if (!hasPicture(shapes)) { msgInfo(t(L"cmd.pic.fitAllNone")); return; }

            if (!msgOkCancel(t(L"cmd.pic.fitAllConfirm"))) return;
            double off = 0.3;
            if (!askMargin(off, 5.0)) return;

            auto& app = xloil::thisApp();
            xloil::PauseExcel pause(app);

            const long n = getLong(shapes, L"Count", 0);
            for (long i = 1; i <= n; ++i)
            {
                IDispatch* shape = getObjectIdx(shapes, L"Item", i);
                if (!shape) continue;
                Releaser rs{ shape };
                const long type = getLong(shape, L"Type", 0);
                if (type != msoPicture && type != msoLinkedPicture) continue;
                statusBar(ad, getBStr(shape, L"Name") + L" is processing...");
                fitShape(shape, off);
            }
            statusBarClear(ad);
            msgInfo(t(L"cmd.pic.fitAllDone"));
        }
        catch (...) {}
    }

    namespace
    {
        // 이미지 확장자 화이트리스트(VB와 동일 집합, 소문자·점 없음).
        bool isImageExt(std::wstring ext)
        {
            if (!ext.empty() && ext[0] == L'.') ext.erase(0, 1);
            for (auto& ch : ext) ch = towlower(ch);
            static const std::set<std::wstring> k = {
                L"jpg", L"jpeg", L"gif", L"png", L"bmp", L"tif", L"tiff", L"emf",
                L"wmf", L"jfif", L"jpe", L"dib", L"rle", L"emz", L"wmz", L"pcz",
                L"svg", L"pct", L"pict", L"wpg" };
            return k.count(ext) > 0;
        }

        std::wstring upper(std::wstring s)
        {
            for (auto& ch : s) ch = towupper(ch);
            return s;
        }

        // Shapes.AddPicture(file, link=0, save, L, T, W, H) → 그림 IDispatch.
        IDispatch* addPicture(IDispatch* shapes, const std::wstring& file,
                              long saveWithDoc, double l, double t, double w, double h)
        {
            VARIANT a[7];
            a[0] = varBStr(file);
            a[1] = varLong(0);
            a[2] = varLong(saveWithDoc);
            a[3] = varDouble(l);
            a[4] = varDouble(t);
            a[5] = varDouble(w);
            a[6] = varDouble(h);
            VARIANT r; VariantInit(&r);
            const bool ok = callMethod(shapes, L"AddPicture", a, 7, &r);
            VariantClear(&a[0]);
            if (ok && r.vt == VT_DISPATCH && r.pdispVal) return r.pdispVal;
            VariantClear(&r);
            return nullptr;
        }

        // 통합문서 안에서 안 겹치는 시트 이름 (VB GetSheetName 대응).
        std::wstring uniqueSheetName(IDispatch* wb, const std::wstring& base)
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
                    names.insert(upper(getBStr(sh, L"Name")));
                }
            }
            for (int i = 1; ; ++i)
            {
                std::wstring cand = base + L"_" + std::to_wstring(i);
                if (!names.count(upper(cand))) return cand;
            }
        }
    }

    // 폴더 이미지 삽입 (VB InsertPicturesFromFolder, C03_Picture.vb:52)
    // 시트 셀 값과 같은 이름의 이미지 파일을 찾아 그 셀(병합 크기)에 삽입.
    void insertPicturesFromFolder()
    {
        using egtools::i18n::t;
        long inserted = 0;
        try
        {
            std::wstring dir;
            if (!egtools::dialogs::pickFolder(excelHwnd(), dir)) return;   // 취소=조용히

            auto& app = xloil::thisApp();
            xloil::PauseExcel pause(app);

            IDispatch* ad = appDisp();
            if (!ad) return;
            Releaser rApp{ ad };
            IDispatch* sh = getObject(ad, L"ActiveSheet");
            if (!sh) return;
            Releaser rSh{ sh };
            IDispatch* shapes = getObject(sh, L"Shapes");
            if (!shapes) return;
            Releaser rShapes{ shapes };

            // UsedRange 값을 한 번에 읽는다(VB와 동일 — A1부터).
            IDispatch* ur = getObject(sh, L"UsedRange");
            if (!ur) return;
            Releaser rUr{ ur };
            const long lastRow = getLong(ur, L"Row", 1) +
                [&]{ IDispatch* rows = getObject(ur, L"Rows");
                     Releaser r{ rows };
                     return rows ? getLong(rows, L"Count", 1) : 1; }() - 1;
            const long lastCol = getLong(ur, L"Column", 1) +
                [&]{ IDispatch* cols = getObject(ur, L"Columns");
                     Releaser r{ cols };
                     return cols ? getLong(cols, L"Count", 1) : 1; }() - 1;

            IDispatch* corner = getObjectIdx2(sh, L"Cells", lastRow, lastCol);
            if (!corner) return;
            Releaser rCorner{ corner };
            const std::wstring fullAddr = L"$A$1:" + getBStr(corner, L"Address");
            VARIANT aa = varBStr(fullAddr);
            IDispatch* all = getObject(sh, L"Range", &aa, 1);
            VariantClear(&aa);
            if (!all) return;
            Releaser rAll{ all };

            const double off = 0.3;
            auto tryInsert = [&](long r, long c, const std::wstring& cellText)
            {
                // Dir(path\값.*) 첫 일치 파일(VB 동일)
                WIN32_FIND_DATAW fd{};
                HANDLE h = FindFirstFileW((dir + L"\\" + cellText + L".*").c_str(), &fd);
                if (h == INVALID_HANDLE_VALUE) return;
                const std::wstring fileName = fd.cFileName;
                FindClose(h);

                const size_t dot = fileName.find_last_of(L'.');
                if (dot == std::wstring::npos) return;
                if (!isImageExt(fileName.substr(dot))) return;
                const std::wstring stem = fileName.substr(0, dot);
                const std::wstring cellU = upper(cellText);
                if (cellU != upper(stem) && cellU != upper(fileName)) return;

                IDispatch* cell = getObjectIdx2(sh, L"Cells", r, c);
                if (!cell) return;
                Releaser rc{ cell };
                IDispatch* ma = getObject(cell, L"MergeArea");
                Releaser rMa{ ma };
                const double w = ma ? getDouble(ma, L"Width") : getDouble(cell, L"Width");
                const double hgt = ma ? getDouble(ma, L"Height") : getDouble(cell, L"Height");

                IDispatch* pic = addPicture(shapes, dir + L"\\" + fileName, 1,
                                            getDouble(cell, L"Left") + off,
                                            getDouble(cell, L"Top") + off,
                                            w - off * 2, hgt - off * 2);
                if (!pic) return;
                Releaser rPic{ pic };
                putLong(pic, L"LockAspectRatio", 0);
                putBStr(pic, L"Name", stem);
                ++inserted;
            };

            SafeArr2D vals;
            if (vals.load(all, L"Value2"))
            {
                for (long r = vals.r1; r <= vals.r2; ++r)
                    for (long c = vals.c1; c <= vals.c2; ++c)
                    {
                        const VARIANT& v = vals.at(r, c);
                        if (v.vt == VT_ERROR || isEmptyVal(v)) continue;
                        const std::wstring s = varToString(v);
                        if (!s.empty()) tryInsert(r, c, s);
                    }
            }

            msgInfo(inserted > 0 ? dir + t(L"cmd.picfolder.done")
                                 : t(L"cmd.picfolder.none"));
        }
        catch (...) {}
    }

    // 양식에 삽입 (VB InsertPicturesIntoForm, C03_Picture.vb:138)
    // {{}} 자리에 폴더의 이미지들을 양식을 복제하며 순서대로 삽입.
    void insertPicturesIntoForm()
    {
        using egtools::i18n::t;
        try
        {
            IDispatch* ad = appDisp();
            if (!ad) return;
            Releaser rApp{ ad };
            IDispatch* sh = getObject(ad, L"ActiveSheet");
            if (!sh) return;
            Releaser rSh{ sh };

            // 1) 양식 영역 지정 (Range 선택 InputBox)
            std::wstring defAddr;
            if (IDispatch* sel = selectionRange(ad))
            {
                Releaser r{ sel };
                defAddr = getBStr(sel, L"Address");
            }
            IDispatch* formRng = inputBoxRange(ad, t(L"cmd.picform.selectForm"),
                                               t(L"cmd.picform.selectFormTitle"), defAddr);
            if (!formRng) return;                       // 취소
            Releaser rForm{ formRng };

            // 2) {{}} 위치 파악
            long rows = 0, cols = 0;
            {
                IDispatch* rr = getObject(formRng, L"Rows");
                Releaser r1{ rr };
                if (rr) rows = getLong(rr, L"Count", 0);
                IDispatch* cc = getObject(formRng, L"Columns");
                Releaser r2{ cc };
                if (cc) cols = getLong(cc, L"Count", 0);
            }
            if (rows == 0 || cols == 0) return;

            std::vector<std::vector<bool>> token((size_t)rows + 1,
                                                 std::vector<bool>((size_t)cols + 1, false));
            long tokenCount = 0;
            auto isToken = [](std::wstring s)
            {
                const auto b = s.find_first_not_of(L" \t");
                const auto e = s.find_last_not_of(L" \t");
                if (b == std::wstring::npos) return false;
                return s.substr(b, e - b + 1) == L"{{}}";
            };
            {
                SafeArr2D vals;
                if (vals.load(formRng, L"Value2"))
                {
                    for (long r = 1; r <= rows; ++r)
                        for (long c = 1; c <= cols; ++c)
                        {
                            const VARIANT& v = vals.at(r, c);
                            if (v.vt == VT_BSTR && v.bstrVal && isToken(v.bstrVal))
                            { token[r][c] = true; ++tokenCount; }
                        }
                }
                else
                {
                    VARIANT sv; VariantInit(&sv);
                    getVar(formRng, L"Value2", &sv);
                    if (sv.vt == VT_BSTR && sv.bstrVal && isToken(sv.bstrVal))
                    { token[1][1] = true; tokenCount = 1; }
                    VariantClear(&sv);
                }
            }
            if (tokenCount == 0) { msgWarn(t(L"cmd.picform.noToken")); return; }

            // 3) 복제 방법(1=시트, 2=아래로, 3=오른쪽으로)
            double methodD = 1;
            if (!egtools::dialogs::inputNumber(excelHwnd(),
                    t(L"cmd.picform.methodTitle"), t(L"cmd.picform.methodPrompt"),
                    methodD, 1, 3, 1))
                return;
            const int method = (int)methodD;
            if (method < 1 || method > 3) { msgWarn(t(L"cmd.picform.badMethod")); return; }

            // 4) 이미지 폴더
            std::wstring dir;
            if (!egtools::dialogs::pickFolder(excelHwnd(), dir)) return;
            std::vector<std::wstring> images;
            try
            {
                for (const auto& e : std::filesystem::directory_iterator(dir))
                {
                    if (!e.is_regular_file()) continue;
                    if (isImageExt(e.path().extension().wstring()))
                        images.push_back(e.path().wstring());
                }
            }
            catch (...) {}
            std::sort(images.begin(), images.end());
            if (images.empty()) { msgWarn(t(L"cmd.picform.noImages")); return; }

            auto& app = xloil::thisApp();
            xloil::PauseExcel pause(app);

            IDispatch* wb = getObject(sh, L"Parent");
            if (!wb) return;
            Releaser rWb{ wb };
            const std::wstring formAddr = getBStr(formRng, L"Address");
            const long formRow = getLong(formRng, L"Row", 1);
            const long formCol = getLong(formRng, L"Column", 1);
            const std::wstring baseSheetName = getBStr(sh, L"Name");

            const long totalPages =
                (long)((images.size() + tokenCount - 1) / tokenCount);
            size_t next = 0;

            for (long page = 1; page <= totalPages; ++page)
            {
                // 복제 방법에 따라 이번 페이지의 양식 영역(copyRng) 확보
                IDispatch* copyRng = nullptr;
                if (method == 1)
                {
                    // 시트 복제: sh.Copy(After:=wb.Sheets(count))
                    IDispatch* sheets = getObject(wb, L"Sheets");
                    if (!sheets) break;
                    Releaser rs{ sheets };
                    const long nSheets = getLong(sheets, L"Count", 0);
                    IDispatch* last = getObjectIdx(sheets, L"Item", nSheets);
                    if (!last) break;
                    Releaser rl{ last };
                    VARIANT a[2];
                    VariantInit(&a[0]); VariantInit(&a[1]);
                    a[0].vt = VT_ERROR; a[0].scode = DISP_E_PARAMNOTFOUND;  // Before
                    a[1].vt = VT_DISPATCH; a[1].pdispVal = last;            // After
                    if (!callMethod(sh, L"Copy", a, 2)) break;

                    IDispatch* newSh = getObject(ad, L"ActiveSheet");
                    if (!newSh) break;
                    Releaser rNew{ newSh };
                    putBStr(newSh, L"Name", uniqueSheetName(wb, baseSheetName));
                    VARIANT fa = varBStr(formAddr);
                    copyRng = getObject(newSh, L"Range", &fa, 1);
                    VariantClear(&fa);
                }
                else
                {
                    const long anchorRow = method == 2 ? formRow + rows * page : formRow;
                    const long anchorCol = method == 2 ? formCol : formCol + cols * page;
                    IDispatch* active = getObject(ad, L"ActiveSheet");
                    if (!active) break;
                    Releaser rAct{ active };
                    IDispatch* anchor = getObjectIdx2(active, L"Cells", anchorRow, anchorCol);
                    if (!anchor) break;
                    Releaser rAnchor{ anchor };
                    if (IDispatch* clearRng = rangeResize(anchor, rows, cols))
                    {
                        Releaser rc{ clearRng };
                        callMethod(clearRng, L"Clear");
                    }
                    // 행/열 전체 복사(높이·너비 유지 — VB 동일)
                    IDispatch* whole = getObject(formRng,
                        method == 2 ? L"EntireRow" : L"EntireColumn");
                    if (!whole) break;
                    Releaser rWhole{ whole };
                    IDispatch* dest = method == 2
                        ? getObjectIdx2(active, L"Cells", anchorRow, 1)
                        : getObjectIdx2(active, L"Cells", 1, anchorCol);
                    if (!dest) break;
                    Releaser rDest{ dest };
                    VARIANT da; VariantInit(&da);
                    da.vt = VT_DISPATCH; da.pdispVal = dest;
                    callMethod(whole, L"Copy", &da, 1);
                    copyRng = rangeResize(anchor, rows, cols);
                }
                if (!copyRng) break;
                Releaser rCopy{ copyRng };

                // {{}} 위치에 이미지 삽입(소진 후 남은 토큰은 지움)
                IDispatch* parentSh = getObject(copyRng, L"Parent");
                if (!parentSh) break;
                Releaser rPar{ parentSh };
                IDispatch* shapes = getObject(parentSh, L"Shapes");
                if (!shapes) break;
                Releaser rShp{ shapes };

                for (long r = 1; r <= rows; ++r)
                    for (long c = 1; c <= cols; ++c)
                    {
                        if (!token[r][c]) continue;
                        IDispatch* cell = getObjectIdx2(copyRng, L"Cells", r, c);
                        if (!cell) continue;
                        Releaser rc{ cell };
                        if (next >= images.size())
                        {
                            callMethod(cell, L"ClearContents");
                            continue;
                        }
                        const std::wstring& img = images[next];
                        statusBar(ad, img);
                        IDispatch* ma = getObject(cell, L"MergeArea");
                        Releaser rMa{ ma };
                        IDispatch* area = ma ? ma : cell;
                        const double off = 0.3;
                        IDispatch* pic = addPicture(shapes, img, -1,
                                                    getDouble(area, L"Left") + off,
                                                    getDouble(area, L"Top") + off,
                                                    getDouble(area, L"Width") - off * 2,
                                                    getDouble(area, L"Height") - off * 2);
                        if (pic)
                        {
                            Releaser rPic{ pic };
                            putLong(pic, L"LockAspectRatio", 0);
                            const size_t slash = img.find_last_of(L"\\/");
                            const std::wstring name =
                                slash == std::wstring::npos ? img : img.substr(slash + 1);
                            putBStr(pic, L"Name", name);
                            putBStr(area, L"Value", name);
                            ++next;
                        }
                    }
            }
            statusBarClear(ad);
        }
        catch (...) {}
    }

    namespace
    {
        // ── OOXML 문자열 파싱 헬퍼(추출 저장용 — 외부 XML 라이브러리 미사용) ──
        std::string tagBetween(const std::string& s, size_t from, size_t to,
                               const std::string& open, const std::string& close)
        {
            const size_t a = s.find(open, from);
            if (a == std::string::npos || a >= to) return "";
            const size_t b = s.find(close, a + open.size());
            if (b == std::string::npos || b > to) return "";
            return s.substr(a + open.size(), b - a - open.size());
        }

        // 요소 구간 [from,to)에서 attr="값" 추출(요소 태그 내 첫 일치).
        std::string attrValue(const std::string& s, size_t from, size_t to,
                              const std::string& attr)
        {
            const std::string key = attr + "=\"";
            const size_t a = s.find(key, from);
            if (a == std::string::npos || a >= to) return "";
            const size_t b = s.find('"', a + key.size());
            if (b == std::string::npos || b > to) return "";
            return s.substr(a + key.size(), b - a - key.size());
        }

        std::string readFileBytes(const std::filesystem::path& p)
        {
            std::string out;
            FILE* f = _wfopen(p.c_str(), L"rb");
            if (!f) return out;
            char buf[8192];
            size_t n;
            while ((n = fread(buf, 1, sizeof(buf), f)) > 0) out.append(buf, n);
            fclose(f);
            return out;
        }

        std::wstring widen(const std::string& s)
        {
            if (s.empty()) return L"";
            const int n = MultiByteToWideChar(CP_UTF8, 0, s.c_str(), (int)s.size(),
                                              nullptr, 0);
            std::wstring w((size_t)n, L'\0');
            MultiByteToWideChar(CP_UTF8, 0, s.c_str(), (int)s.size(), w.data(), n);
            return w;
        }

        std::wstring sanitizeFileName(std::wstring s)
        {
            for (auto& ch : s)
                if (wcschr(L"\\/:*?\"<>|\r\n\t", ch)) ch = L'_';
            return s;
        }

        // Shell COM으로 ZIP 해제(외부 라이브러리 의존 없음). 완료 대기 폴링.
        bool shellUnzip(const std::wstring& zipPath, const std::wstring& destDir)
        {
            CLSID clsid{};
            if (FAILED(CLSIDFromProgID(L"Shell.Application", &clsid))) return false;
            IDispatch* shell = nullptr;
            if (FAILED(CoCreateInstance(clsid, nullptr, CLSCTX_INPROC_SERVER,
                                        IID_IDispatch, (void**)&shell)))
                return false;
            Releaser rShell{ shell };

            auto folderOf = [&](const std::wstring& p) -> IDispatch*
            {
                VARIANT a = varBStr(p);
                VARIANT r; VariantInit(&r);
                const bool ok = invokeRaw(shell, L"NameSpace", DISPATCH_METHOD, &r, &a, 1);
                VariantClear(&a);
                if (ok && r.vt == VT_DISPATCH && r.pdispVal) return r.pdispVal;
                VariantClear(&r);
                return nullptr;
            };
            IDispatch* zipFolder = folderOf(zipPath);
            if (!zipFolder) return false;
            Releaser rz{ zipFolder };
            IDispatch* dest = folderOf(destDir);
            if (!dest) return false;
            Releaser rd{ dest };

            VARIANT items; VariantInit(&items);
            if (!invokeRaw(zipFolder, L"Items", DISPATCH_METHOD, &items, nullptr, 0) ||
                items.vt != VT_DISPATCH || !items.pdispVal)
            {
                VariantClear(&items);
                return false;
            }
            const long total = getLong(items.pdispVal, L"Count", 0);

            // 4=진행창 없음, 16=모두 예, 512=디렉터리 확인 없음, 1024=UI 오류 없음
            VARIANT args[2];
            args[0] = items;                        // 소유권은 아래 VariantClear가 정리
            args[1] = varLong(4 | 16 | 512 | 1024);
            const bool started = invokeRaw(dest, L"CopyHere", DISPATCH_METHOD,
                                           nullptr, args, 2);
            VariantClear(&items);
            if (!started) return false;

            // CopyHere는 비동기 — 최상위 항목 수가 채워질 때까지 대기(최대 30초)
            for (int i = 0; i < 300; ++i)
            {
                long have = 0;
                try
                {
                    for (auto it = std::filesystem::directory_iterator(destDir);
                         it != std::filesystem::directory_iterator(); ++it)
                        ++have;
                }
                catch (...) {}
                if (have >= total && total > 0) return true;
                Sleep(100);
            }
            return false;
        }
    }

    // 그림 추출 저장 (VB ExportAllPictures, C03_Picture.vb:387 — 비파괴 개선판)
    // VB는 다른 시트를 지우고 저장 후 원본을 다시 여는 파괴적 흐름이었으나,
    // 여기서는 SaveCopyAs 사본(.zip)을 그대로 파싱해 원본을 건드리지 않는다.
    // drawing1.xml 하드코딩도 workbook→sheet→drawing 관계 해석으로 대체
    // (그림이 어느 drawing*.xml에 있어도 동작). 원본 화질 그대로 보존.
    void exportAllPictures()
    {
        using egtools::i18n::t;
        namespace fs = std::filesystem;
        long exported = 0;
        fs::path tempRoot;
        try
        {
            IDispatch* ad = appDisp();
            if (!ad) return;
            Releaser rApp{ ad };
            IDispatch* sh = getObject(ad, L"ActiveSheet");
            if (!sh) return;
            Releaser rSh{ sh };
            const std::wstring sheetName = getBStr(sh, L"Name");
            IDispatch* wb = getObject(ad, L"ActiveWorkbook");
            if (!wb) return;
            Releaser rWb{ wb };

            // 그림 유무 선확인
            {
                IDispatch* shapes = getObject(sh, L"Shapes");
                Releaser rs{ shapes };
                if (!shapes || !hasPicture(shapes))
                { msgInfo(t(L"cmd.export.none")); return; }
            }

            std::wstring destDir;
            if (!egtools::dialogs::pickFolder(excelHwnd(), destDir)) return;

            auto& app = xloil::thisApp();
            xloil::PauseExcel pause(app);
            statusBar(ad, L"exporting pictures...");

            // 1) SaveCopyAs 사본(.zip) + 해제
            wchar_t tempBuf[MAX_PATH]{};
            GetTempPathW(MAX_PATH, tempBuf);
            tempRoot = fs::path(tempBuf) /
                       (L"egtools_export_" + std::to_wstring(GetTickCount64()));
            fs::create_directories(tempRoot / L"x");
            const fs::path zipPath = tempRoot / L"book.zip";
            {
                VARIANT a = varBStr(zipPath.wstring());
                const bool ok = invokeRaw(wb, L"SaveCopyAs", DISPATCH_METHOD,
                                          nullptr, &a, 1);
                VariantClear(&a);
                if (!ok) throw std::runtime_error("SaveCopyAs");
            }
            if (!shellUnzip(zipPath.wstring(), (tempRoot / L"x").wstring()))
                throw std::runtime_error("unzip");
            const fs::path xl = tempRoot / L"x" / L"xl";

            // 2) workbook.xml에서 활성 시트의 r:id → sheet*.xml → drawing*.xml
            const std::string wbXml = readFileBytes(xl / L"workbook.xml");
            std::string sheetRid;
            {
                const std::string nameU8 = [&] {
                    const std::wstring& w = sheetName;
                    const int n = WideCharToMultiByte(CP_UTF8, 0, w.c_str(), (int)w.size(),
                                                      nullptr, 0, nullptr, nullptr);
                    std::string s((size_t)n, '\0');
                    WideCharToMultiByte(CP_UTF8, 0, w.c_str(), (int)w.size(),
                                        s.data(), n, nullptr, nullptr);
                    return s;
                }();
                size_t pos = 0;
                while (true)
                {
                    const size_t a = wbXml.find("<sheet ", pos);
                    if (a == std::string::npos) break;
                    const size_t b = wbXml.find('>', a);
                    if (b == std::string::npos) break;
                    pos = b;
                    if (attrValue(wbXml, a, b, "name") == nameU8)
                    {
                        sheetRid = attrValue(wbXml, a, b, "r:id");
                        break;
                    }
                }
            }
            if (sheetRid.empty()) throw std::runtime_error("sheet rid");

            auto relTarget = [&](const fs::path& relsFile, const std::string& rid,
                                 const char* typeContains) -> std::string
            {
                const std::string xml = readFileBytes(relsFile);
                size_t pos = 0;
                while (true)
                {
                    const size_t a = xml.find("<Relationship ", pos);
                    if (a == std::string::npos) return "";
                    const size_t b = xml.find('>', a);
                    if (b == std::string::npos) return "";
                    pos = b;
                    if (!rid.empty() && attrValue(xml, a, b, "Id") != rid) continue;
                    if (typeContains &&
                        attrValue(xml, a, b, "Type").find(typeContains) == std::string::npos)
                        continue;
                    return attrValue(xml, a, b, "Target");
                }
            };

            std::string sheetFile = relTarget(xl / L"_rels" / L"workbook.xml.rels",
                                              sheetRid, nullptr);
            if (sheetFile.empty()) throw std::runtime_error("sheet target");
            const std::string sheetBase =
                sheetFile.substr(sheetFile.find_last_of('/') + 1);   // sheetN.xml

            const std::string drawingTarget = relTarget(
                xl / L"worksheets" / L"_rels" / widen(sheetBase + ".rels"),
                "", "/drawing");
            if (drawingTarget.empty()) { msgInfo(t(L"cmd.export.none")); goto cleanup; }
            {
                const std::string drawingBase =
                    drawingTarget.substr(drawingTarget.find_last_of('/') + 1);

                // 3) drawing*.xml 앵커 파싱 + rels 매핑
                const std::string dXml =
                    readFileBytes(xl / L"drawings" / widen(drawingBase));
                std::map<std::string, std::string> ridToMedia;   // rId → media 파일명
                {
                    const std::string relXml = readFileBytes(
                        xl / L"drawings" / L"_rels" / widen(drawingBase + ".rels"));
                    size_t pos = 0;
                    while (true)
                    {
                        const size_t a = relXml.find("<Relationship ", pos);
                        if (a == std::string::npos) break;
                        const size_t b = relXml.find('>', a);
                        if (b == std::string::npos) break;
                        pos = b;
                        std::string target = attrValue(relXml, a, b, "Target");
                        const size_t slash = target.find_last_of('/');
                        if (slash != std::string::npos) target = target.substr(slash + 1);
                        ridToMedia[attrValue(relXml, a, b, "Id")] = target;
                    }
                }

                std::set<std::wstring> usedNames;
                size_t pos = 0;
                while (true)
                {
                    size_t a = dXml.find("<xdr:twoCellAnchor", pos);
                    size_t close = std::string::npos;
                    if (a != std::string::npos)
                        close = dXml.find("</xdr:twoCellAnchor>", a);
                    else
                    {
                        a = dXml.find("<xdr:oneCellAnchor", pos);
                        if (a == std::string::npos) break;
                        close = dXml.find("</xdr:oneCellAnchor>", a);
                    }
                    if (close == std::string::npos) break;
                    pos = close + 8;

                    const std::string rid = [&] {
                        const size_t blip = dXml.find("<a:blip ", a);
                        if (blip == std::string::npos || blip > close) return std::string();
                        const size_t be = dXml.find('>', blip);
                        return attrValue(dXml, blip, be, "r:embed");
                    }();
                    if (rid.empty() || !ridToMedia.count(rid)) continue;
                    const std::string media = ridToMedia[rid];

                    const long fromCol =
                        atol(tagBetween(dXml, a, close, "<xdr:col>", "</xdr:col>").c_str());
                    const long fromRow =
                        atol(tagBetween(dXml, a, close, "<xdr:row>", "</xdr:row>").c_str());

                    // 대상 파일명: 좌상단 셀 값(라이브 시트에서) → 없으면 그림이름_RxCy
                    std::wstring base;
                    {
                        IDispatch* cell = getObjectIdx2(sh, L"Cells",
                                                        fromRow + 1, fromCol + 1);
                        if (cell)
                        {
                            Releaser rc{ cell };
                            VARIANT v; VariantInit(&v);
                            getVar(cell, L"Value", &v);
                            if (!isEmptyVal(v)) base = varToString(v);
                            VariantClear(&v);
                        }
                    }
                    if (base.empty())
                    {
                        std::string nm;
                        const size_t pr = dXml.find("<xdr:cNvPr ", a);
                        if (pr != std::string::npos && pr < close)
                            nm = attrValue(dXml, pr, dXml.find('>', pr), "name");
                        base = widen(nm) + L"_R" + std::to_wstring(fromRow + 1) +
                               L"C" + std::to_wstring(fromCol + 1);
                    }
                    base = sanitizeFileName(base);

                    const std::string extA =
                        media.substr(media.find_last_of('.') == std::string::npos
                                         ? media.size()
                                         : media.find_last_of('.'));
                    std::wstring name = base + widen(extA);
                    for (int k = 1; usedNames.count(name); ++k)
                        name = base + L"_" + std::to_wstring(k) + widen(extA);
                    usedNames.insert(name);

                    const fs::path src = xl / L"media" / widen(media);
                    std::error_code ec;
                    if (fs::copy_file(src, fs::path(destDir) / name,
                                      fs::copy_options::overwrite_existing, ec))
                        ++exported;
                }
            }

        cleanup:
            statusBarClear(ad);
            std::error_code ec;
            fs::remove_all(tempRoot, ec);
            msgInfo(exported > 0
                        ? withCount(t(L"cmd.export.done"), exported) + L"\n" + destDir
                        : t(L"cmd.export.none"));
            return;
        }
        catch (...)
        {
            if (!tempRoot.empty())
            {
                std::error_code ec;
                fs::remove_all(tempRoot, ec);
            }
            msgWarn(t(L"cmd.export.fail"));
        }
    }
}
