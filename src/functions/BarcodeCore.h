// BarcodeCore.h — shared barcode helpers used by FxBarcode.cpp and the
// eg_barcode_selftest console tool: type-name mapping, zint encoding to a
// module bitmap, GS1 AI tables (READGS1 stream parsing) and a minimal 24-bpp
// BMP writer.
//
// 생성은 zint(external/zint, BSD-3-Clause), 판독은 zxing-cpp(READBARCODE).
// zxing-cpp의 writer는 DataMatrix에 FNC1 코드워드(232)가 없고 QR의 FNC1 모드도
// MultiFormatWriter로는 켤 수 없어 GS1 2D를 규격대로 만들 수 없었다(2026-09-12
// 전환). zint는 GS1-128·GS1 DataMatrix·GS1 QR 모두 FNC1 첫 자리 + 구분자를
// 규격대로 넣고(스캐너 심볼로지 ]C1/]d2/]Q3) AI 형식·길이 린터를 내장한다.
#pragma once

#include <zint.h>

#include <algorithm>
#include <cmath>
#include <cstdint>
#include <cstdio>
#include <cwchar>
#include <cwctype>
#include <map>
#include <optional>
#include <string>
#include <vector>

namespace egtools::barcode
{
    // ── type-name → zint symbology (VB BARCODE table) ─────────────────────
    // is2D: 규격 비율(모듈 정사각) 유지 대상. gs1: 괄호 AI 표기 "(01)…(10)…"를
    // GS1 모드로 인코딩(GS1PARENS_MODE — zint가 AI 유효성 검사).
    struct TypeInfo { int symbology; bool is2D; bool gs1; };

    inline std::optional<TypeInfo> parseType(std::wstring t)
    {
        for (auto& c : t) c = (wchar_t)towupper(c);
        // strip spaces and hyphens for tolerant matching ("QR CODE", "CODE-128")
        std::wstring k;
        for (wchar_t c : t) if (c != L' ' && c != L'-') k.push_back(c);
        static const std::map<std::wstring, TypeInfo> m = {
            { L"1", { BARCODE_CODE128, false, false } },   { L"CODE128", { BARCODE_CODE128, false, false } },
            { L"2", { BARCODE_CODE39, false, false } },    { L"CODE39",  { BARCODE_CODE39, false, false } },
            { L"3", { BARCODE_EAN13, false, false } },     { L"EAN13",   { BARCODE_EAN13, false, false } },
            { L"4", { BARCODE_EAN8, false, false } },      { L"EAN8",    { BARCODE_EAN8, false, false } },
            { L"5", { BARCODE_C25INTER, false, false } },  { L"ITF",     { BARCODE_C25INTER, false, false } },
            { L"INTERLEAVED2OF5", { BARCODE_C25INTER, false, false } },
            { L"2OF5", { BARCODE_C25INTER, false, false } },
            { L"CODE93", { BARCODE_CODE93, false, false } },
            { L"CODABAR", { BARCODE_CODABAR, false, false } },
            { L"UPCA", { BARCODE_UPCA, false, false } },   { L"UPC_A", { BARCODE_UPCA, false, false } },
            { L"UPCE", { BARCODE_UPCE, false, false } },   { L"UPC_E", { BARCODE_UPCE, false, false } },
            { L"10", { BARCODE_PDF417, true, false } },    { L"PDF417", { BARCODE_PDF417, true, false } },
            { L"11", { BARCODE_QRCODE, true, false } },    { L"QRCODE", { BARCODE_QRCODE, true, false } },
            { L"QR", { BARCODE_QRCODE, true, false } },
            { L"12", { BARCODE_DATAMATRIX, true, false } },{ L"DATAMATRIX", { BARCODE_DATAMATRIX, true, false } },
            { L"DM", { BARCODE_DATAMATRIX, true, false } },
            { L"13", { BARCODE_AZTEC, true, false } },     { L"AZTEC", { BARCODE_AZTEC, true, false } },
            { L"21", { BARCODE_GS1_128, false, true } },   { L"GS1128", { BARCODE_GS1_128, false, true } },
            { L"GS1CODE128", { BARCODE_GS1_128, false, true } },
            { L"22", { BARCODE_DATAMATRIX, true, true } }, { L"GS1DATAMATRIX", { BARCODE_DATAMATRIX, true, true } },
            { L"GS1DM", { BARCODE_DATAMATRIX, true, true } },
            { L"23", { BARCODE_QRCODE, true, true } },     { L"GS1QRCODE", { BARCODE_QRCODE, true, true } },
            { L"GS1QR", { BARCODE_QRCODE, true, true } },
        };
        auto it = m.find(k);
        if (it == m.end()) return std::nullopt;
        return it->second;
    }

    // ── GS1 application identifier tables (READGS1 decoded-stream parsing) ──
    // 인코딩 쪽 검증은 zint 린터가 하므로 여기서는 판독 문자열을 AI 단위로
    // 자를 때만 쓴다(고정 길이 AI는 구분자 없이 이어지므로 길이를 알아야 함).
    inline const std::map<std::wstring, int>& fixedLenAIs()
    {
        static const std::map<std::wstring, int> m = {
            {L"00",18},{L"01",14},{L"02",14},{L"03",14},{L"04",16},
            {L"11",6},{L"12",6},{L"13",6},{L"14",6},{L"15",6},{L"16",6},{L"17",6},{L"18",6},{L"19",6},
            {L"20",2},
            {L"31",8},{L"32",8},{L"33",8},{L"34",8},{L"35",8},{L"36",8},
            {L"402",17},{L"410",13},{L"411",13},{L"412",13},{L"413",13},{L"414",13},{L"415",13},
            {L"422",3},{L"423",15},{L"424",3},{L"425",3},{L"426",3},
            {L"7001",13},{L"7003",10},
            {L"8001",14},{L"8003",16},{L"8005",6},{L"8006",18},{L"8008",12},{L"8017",18},
            {L"8018",18},{L"8100",6},{L"8101",10},{L"8102",2} };
        return m;
    }

    inline const std::map<std::wstring, int>& varLenAIs()
    {
        static const std::map<std::wstring, int> m = {
            {L"10",20},{L"21",20},{L"22",29},{L"240",30},{L"241",30},{L"242",6},{L"243",20},
            {L"250",30},{L"251",30},{L"253",30},{L"254",20},{L"255",25},
            {L"30",8},{L"37",8},
            {L"400",30},{L"401",30},{L"403",30},{L"420",20},{L"421",12},{L"427",3},
            {L"7002",30},{L"7004",4},{L"7030",30},{L"7031",30},{L"7032",30},{L"7033",30},
            {L"7034",30},{L"7035",30},{L"7036",30},{L"7037",30},{L"7038",30},{L"7039",30},
            {L"710",20},{L"711",20},{L"712",20},
            {L"8002",20},{L"8004",30},{L"8007",30},{L"8019",10},{L"8020",25},{L"8110",70},{L"8200",70},
            {L"90",30},{L"91",30},{L"92",30},{L"93",30},{L"94",30},{L"95",30},{L"96",30},
            {L"97",30},{L"98",30},{L"99",30} };
        return m;
    }

    // 판독 원문 스트림 "[FNC1]AI값AI값…"(가변 길이 AI 뒤 GS 등 제어문자 구분자)을
    // (AI, 값) 목록으로 자른다. 앞의 제어문자·심볼로지 식별자(]C1/]d2/]Q3)는
    // 건너뛴다. 고정 길이 AI는 길이만큼, 가변 길이 AI는 구분자 또는 최대
    // 길이까지. 표에 없는 AI(4자리까지 불일치) → false. READGS1·GS1 인코딩 공용.
    // READBARCODE CODETEXT 표기(<GS>, <FNC1>, <FCN>, <RS>, <EOT> …)를 글자 그대로
    // 붙여 넣은 입력도 받도록 해당 토큰을 제어문자(구분자)로 바꾼다.
    inline std::wstring expandControlTokens(const std::wstring& s)
    {
        static const wchar_t* tokens[] = { L"<GS>", L"<FNC1>", L"<FCN>", L"<RS>", L"<EOT>",
                                           L"<FS>", L"<US>", L"<TAB>", L"<CR>", L"<LF>" };
        std::wstring out;
        for (size_t i = 0; i < s.size(); ++i)
        {
            if (s[i] == L'<')
            {
                bool hit = false;
                for (const wchar_t* t : tokens)
                {
                    const size_t n = wcslen(t);
                    if (s.compare(i, n, t) == 0 || (i + n <= s.size() && _wcsnicmp(s.c_str() + i, t, n) == 0))
                    { out.push_back((wchar_t)0x1D); i += n - 1; hit = true; break; }
                }
                if (hit) continue;
            }
            out.push_back(s[i]);
        }
        return out;
    }

    inline bool parseGs1Stream(const std::wstring& input,
                               std::vector<std::pair<std::wstring, std::wstring>>& parts,
                               bool* unknownAi = nullptr)
    {
        const std::wstring src = expandControlTokens(input);
        size_t i = 0;
        if (src.size() >= 3 && src[0] == L']' && iswalpha(src[1]) && iswdigit(src[2])) i = 3;
        while (i < src.size() && src[i] < 32) ++i;
        std::wstring ai, val;
        bool inValue = false;
        int expected = 0;
        auto flush = [&]() {
            if (!ai.empty()) parts.emplace_back(ai, val);
            ai.clear(); val.clear(); inValue = false; expected = 0;
        };
        for (; i < src.size(); ++i)
        {
            const wchar_t ch = src[i];
            if (ch < 32) { flush(); continue; }
            if (inValue)
            {
                val.push_back(ch);
                if ((int)val.size() == expected) flush();
            }
            else
            {
                ai.push_back(ch);
                auto fit = fixedLenAIs().find(ai);
                if (fit != fixedLenAIs().end()) { expected = fit->second; inValue = true; }
                else
                {
                    auto vit = varLenAIs().find(ai);
                    if (vit != varLenAIs().end()) { expected = vit->second; inValue = true; }
                    else if (ai.size() >= 4) { if (unknownAi) *unknownAi = true; return false; }
                }
            }
        }
        flush();
        return !parts.empty();
    }

    // GS1 입력을 zint GS1PARENS_MODE가 받는 "(AI)값…" 표기로 정규화한다.
    //   "(01)…(10)…"  괄호 표기 → 그대로
    //   "[01]…[10]…"  대괄호 표기 → 괄호로 치환(GS1 문자 집합에 대괄호는 없음)
    //   그 밖         스캐너 원문 스트림(GS/FNC1 구분자, 선행 심볼로지 식별자 허용)
    //                 → AI 길이 표로 잘라 괄호 표기 구성
    // AI 형식·길이·검사숫자 검증은 이후 zint 린터가 한다.
    inline bool normalizeGs1(const std::wstring& input, std::wstring& out)
    {
        std::wstring s = input;
        while (!s.empty() && (s.front() == L' ' || s.front() == L'\t')) s.erase(0, 1);
        while (!s.empty() && (s.back() == L' ' || s.back() == L'\t')) s.pop_back();
        if (s.empty()) return false;
        if (s.front() == L'(') { out = s; return true; }
        if (s.front() == L'[')
        {
            out.clear();
            for (size_t i = 0; i < s.size(); ++i)
            {
                if (s[i] == L'[')
                {
                    const size_t close = s.find(L']', i + 1);
                    if (close == std::wstring::npos || close - i - 1 < 2 || close - i - 1 > 4) return false;
                    for (size_t k = i + 1; k < close; ++k) if (!iswdigit(s[k])) return false;
                    out += L"(" + s.substr(i + 1, close - i - 1) + L")";
                    i = close;
                }
                else if (s[i] == L']') return false;
                else out.push_back(s[i]);
            }
            return true;
        }
        std::vector<std::pair<std::wstring, std::wstring>> parts;
        if (!parseGs1Stream(s, parts)) return false;
        out.clear();
        for (auto& [ai, val] : parts) out += L"(" + ai + L")" + val;
        return true;
    }

    // ── encoding (zint) ─────────────────────────────────────────────────────
    // 모듈 비트맵: 0 = black, 1 = white, row-major w×h.
    struct ModuleBitmap { int w = 0, h = 0; std::vector<uint8_t> px; };

    inline std::string toUtf8(const std::wstring& s)
    {
        std::string out;
        for (size_t i = 0; i < s.size(); ++i)
        {
            uint32_t c = s[i];
            if (c >= 0xD800 && c <= 0xDBFF && i + 1 < s.size()
                && s[i + 1] >= 0xDC00 && s[i + 1] <= 0xDFFF)
            {
                c = 0x10000 + ((c - 0xD800) << 10) + (s[i + 1] - 0xDC00);
                ++i;
            }
            if (c < 0x80) out.push_back((char)c);
            else if (c < 0x800) { out.push_back((char)(0xC0 | (c >> 6))); out.push_back((char)(0x80 | (c & 0x3F))); }
            else if (c < 0x10000) { out.push_back((char)(0xE0 | (c >> 12))); out.push_back((char)(0x80 | ((c >> 6) & 0x3F))); out.push_back((char)(0x80 | (c & 0x3F))); }
            else { out.push_back((char)(0xF0 | (c >> 18))); out.push_back((char)(0x80 | ((c >> 12) & 0x3F))); out.push_back((char)(0x80 | ((c >> 6) & 0x3F))); out.push_back((char)(0x80 | (c & 0x3F))); }
        }
        return out;
    }

    // zint로 인코딩해 모듈 비트맵을 만든다(모듈 1px). 2D는 zint row_height를
    // 정수 배율로 반영(PDF417 행높이 3X, 매트릭스 심볼은 1), 1D는 한 줄(높이는
    // 호출자가 늘림). margin: 모듈 단위 여백 — 2D는 사방, 1D는 좌우만.
    // GS1은 괄호 AI 표기를 그대로 넘기고(GS1PARENS_MODE) zint가 AI 형식·길이를
    // 검사한다; 일반 텍스트는 UTF-8(UNICODE_MODE, 필요 시 zint가 ECI 삽입).
    // 실패 시 false(err에 zint 메시지).
    inline bool encodeModules(const std::wstring& text, const TypeInfo& ti, int margin,
                              ModuleBitmap& out, std::string* err = nullptr)
    {
        struct Sym { zint_symbol* p; ~Sym() { if (p) ZBarcode_Delete(p); } } s{ ZBarcode_Create() };
        if (!s.p) { if (err) *err = "ZBarcode_Create failed"; return false; }
        s.p->symbology = ti.symbology;
        s.p->input_mode = ti.gs1 ? (GS1_MODE | GS1PARENS_MODE) : UNICODE_MODE;

        std::string bytes;
        if (ti.gs1)
        {
            // 괄호·대괄호·원문 스트림 → "(AI)값" 표기(normalizeGs1). GS1 문자 집합은
            // 인쇄 가능 ASCII — 그 밖의 문자는 zint 오류가 되므로 먼저 거른다.
            std::wstring norm;
            if (!normalizeGs1(text, norm)) { if (err) *err = "unrecognized GS1 input"; return false; }
            for (wchar_t c : norm)
            {
                if (c < 0x20 || c > 0x7E) { if (err) *err = "GS1 data must be printable ASCII"; return false; }
                bytes.push_back((char)c);
            }
        }
        else bytes = toUtf8(text);
        if (bytes.empty()) { if (err) *err = "empty input"; return false; }

        const int rc = ZBarcode_Encode(s.p, (const unsigned char*)bytes.data(), (int)bytes.size());
        // GS1 린터 위반(예: (01) GTIN 검사숫자 오류)은 zint가 인코딩은 하되
        // ZINT_WARN_NONCOMPLIANT로 알린다 — GS1 모드에서는 실패로 취급한다.
        // (ZINT_WARN_USES_ECI 등 다른 경고는 정상: 비 Latin-1 텍스트의 ECI 삽입.)
        if (rc >= ZINT_ERROR || (ti.gs1 && rc == ZINT_WARN_NONCOMPLIANT))
            { if (err) *err = s.p->errtxt; return false; }
        const int rows = s.p->rows, width = s.p->width;
        if (rows <= 0 || width <= 0 || rows > 200 || width > 1152)
            { if (err) *err = "unexpected symbol size"; return false; }

        std::vector<int> rowPx((size_t)rows, 1);
        int total = 0;
        for (int r = 0; r < rows; ++r)
        {
            int hpx = 1;
            if (ti.is2D)
            {
                const float f = s.p->row_height[r];
                hpx = f > 0.0f ? (int)std::lround(f) : (ti.symbology == BARCODE_PDF417 ? 3 : 1);
                if (hpx < 1) hpx = 1;
            }
            rowPx[(size_t)r] = hpx;
            total += hpx;
        }

        if (margin < 0) margin = 0;
        const int mx = margin, my = ti.is2D ? margin : 0;
        out.w = width + 2 * mx;
        out.h = total + 2 * my;
        out.px.assign((size_t)out.w * out.h, 1);
        int y = my;
        for (int r = 0; r < rows; ++r)
        {
            for (int x = 0; x < width; ++x)
                if ((s.p->encoded_data[r][x >> 3] >> (x & 7)) & 1)
                    for (int k = 0; k < rowPx[(size_t)r]; ++k)
                        out.px[(size_t)(y + k) * out.w + mx + x] = 0;
            y += rowPx[(size_t)r];
        }
        return true;
    }

    // 정수 배율 확대(sx 가로, sy 세로). 2D는 sx==sy로 모듈 정사각 유지.
    inline ModuleBitmap scaleBitmap(const ModuleBitmap& in, int sx, int sy)
    {
        if (sx < 1) sx = 1;
        if (sy < 1) sy = 1;
        ModuleBitmap out;
        out.w = in.w * sx; out.h = in.h * sy;
        out.px.resize((size_t)out.w * out.h);
        for (int y = 0; y < out.h; ++y)
        {
            const uint8_t* src = &in.px[(size_t)(y / sy) * in.w];
            uint8_t* dst = &out.px[(size_t)y * out.w];
            for (int x = 0; x < out.w; ++x) dst[x] = src[x / sx];
        }
        return out;
    }

    // 셀 그림용 픽셀 크기로 확대: 2D는 긴 변이 target2D 이상이 되는 정수배(비율 유지),
    // 1D는 폭 target1DW 이상 정수배 + 높이 target1DH(바 높이는 규격상 자유).
    inline ModuleBitmap scaleForPicture(const ModuleBitmap& mb, bool is2D,
                                        int target2D = 512, int target1DW = 800, int target1DH = 200)
    {
        if (is2D)
        {
            const int longest = (std::max)(mb.w, mb.h);
            const int k = (std::max)(1, (target2D + longest - 1) / longest);
            return scaleBitmap(mb, k, k);
        }
        const int sx = (std::max)(1, (target1DW + mb.w - 1) / mb.w);
        const int sy = (std::max)(1, target1DH / mb.h);
        return scaleBitmap(mb, sx, sy);
    }

    // ── minimal 24-bpp bottom-up BMP writer ─────────────────────────────────
    // pixels: row-major top-down, 0 = black, nonzero = white.
    inline bool writeBmp24(const std::wstring& path, int w, int h,
                           const std::vector<uint8_t>& pixels)
    {
        if (w <= 0 || h <= 0 || (int)pixels.size() < w * h) return false;
        const int stride = (w * 3 + 3) & ~3;
        const uint32_t dataSize = (uint32_t)stride * h;
        const uint32_t fileSize = 54 + dataSize;

        FILE* f = nullptr;
        if (_wfopen_s(&f, path.c_str(), L"wb") != 0 || !f) return false;

        uint8_t hdr[54] = {};
        hdr[0] = 'B'; hdr[1] = 'M';
        *(uint32_t*)(hdr + 2) = fileSize;
        *(uint32_t*)(hdr + 10) = 54;
        *(uint32_t*)(hdr + 14) = 40;
        *(int32_t*)(hdr + 18) = w;
        *(int32_t*)(hdr + 22) = h;               // bottom-up
        *(uint16_t*)(hdr + 26) = 1;
        *(uint16_t*)(hdr + 28) = 24;
        *(uint32_t*)(hdr + 34) = dataSize;
        *(int32_t*)(hdr + 38) = 2835;            // ~72 dpi
        *(int32_t*)(hdr + 42) = 2835;
        fwrite(hdr, 1, 54, f);

        std::vector<uint8_t> row((size_t)stride, 0);
        for (int y = h - 1; y >= 0; --y)
        {
            for (int x = 0; x < w; ++x)
            {
                const uint8_t v = pixels[(size_t)y * w + x] ? 255 : 0;
                row[(size_t)x * 3 + 0] = v;
                row[(size_t)x * 3 + 1] = v;
                row[(size_t)x * 3 + 2] = v;
            }
            fwrite(row.data(), 1, (size_t)stride, f);
        }
        fclose(f);
        return true;
    }

    // Encode straight to a BMP file at picture scale. Returns false on any failure.
    inline bool encodeToBmp(const std::wstring& text, const TypeInfo& ti, int margin,
                            const std::wstring& path, std::string* err = nullptr)
    {
        ModuleBitmap mb;
        if (!encodeModules(text, ti, margin, mb, err)) return false;
        const ModuleBitmap img = scaleForPicture(mb, ti.is2D);
        return writeBmp24(path, img.w, img.h, img.px);
    }
}
