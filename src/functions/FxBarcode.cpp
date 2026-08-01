// FxBarcode.cpp — barcode functions (plan/22 B2), ported from VB FX5_Barcode
// with zxing-cpp replacing ZXing.NET.
//
//   BARCODE(text, [type], [option])          — multi-format encode → picture in cell
//   QRCODE(text, [margin], [gs1])            — QR encode → picture
//   DATAMATRIX(text, [margin], [gs1])        — DataMatrix encode → picture
//   CODE128(text, [show_text], [gs1], [text_size]) — Code128 encode → picture
//   READBARCODE(source, [result_type])       — decode image file/URL
//   READGS1(gs1_text, [concat])              — split GS1 AI stream into (AI),value rows
//
// Encoding writes a temporary 24-bpp BMP and reuses the IMAGE insertion path
// (ImageInsert.h): deferred main-thread Shapes.AddPicture sized to the calling
// cell, temp file deleted after embedding. GS1: Code128 uses the writer's FNC1
// escape (U+00F1, full compliance); QR/DataMatrix fall back to ASCII 29 group
// separators (readers commonly accept; documented in the manual).
// Decoding loads pixels via WIC (PNG/JPG/BMP/GIF) — no image library needed.

// windows.h FIRST with full GDI: xlOil's WindowsSlim.h pre-includes windows.h
// with NOGDI, so any xlOil header before this would lock GDI types out.
#include <windows.h>
#include <wincodec.h>
#include <urlmon.h>

#include "../core/Registry.h"
#include "../core/Spill.h"
#include "../core/ArrayUtil.h"
#include "BarcodeCore.h"
#include "ImageInsert.h"

#include <ZXing/ReadBarcode.h>

#include <xlOil/xlOil.h>
#include <xlOil/Caller.h>
#include <xlOil/ExcelArray.h>
#include <string>
#include <vector>

#pragma comment(lib, "windowscodecs.lib")
#pragma comment(lib, "urlmon.lib")
#pragma comment(lib, "gdi32.lib")

using namespace xloil;
namespace bc = egtools::barcode;

namespace egtools::functions
{
    namespace
    {
        // ── small shared helpers ─────────────────────────────────────────────
        std::wstring tempBmpPath()
        {
            wchar_t tmp[MAX_PATH]{};
            GetTempPathW(MAX_PATH, tmp);
            static LONG counter = 0;
            return std::wstring(tmp) + L"egbar_" + std::to_wstring(GetTickCount64())
                 + L"_" + std::to_wstring(InterlockedIncrement(&counter)) + L".bmp";
        }

        void splitCaller(const std::wstring& addr, std::wstring& cell)
        {
            size_t bang = addr.rfind(L'!');
            cell = (bang == std::wstring::npos) ? addr : addr.substr(bang + 1);
        }

        std::wstring fromUtf8(const std::string& s)
        {
            if (s.empty()) return L"";
            int n = MultiByteToWideChar(CP_UTF8, 0, s.data(), (int)s.size(), nullptr, 0);
            std::wstring w((size_t)n, L'\0');
            MultiByteToWideChar(CP_UTF8, 0, s.data(), (int)s.size(), w.data(), n);
            return w;
        }

        // Append a white band with centered black text below a pixel buffer
        // (0 = black / 1 = white, row-major w×h). Uses a GDI DIB for the glyphs.
        bool appendTextBand(std::vector<uint8_t>& px, int w, int& h,
                            const std::wstring& text, int bandH)
        {
            if (bandH <= 0 || text.empty()) return true;

            BITMAPINFO bi{};
            bi.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
            bi.bmiHeader.biWidth = w;
            bi.bmiHeader.biHeight = -bandH;          // top-down
            bi.bmiHeader.biPlanes = 1;
            bi.bmiHeader.biBitCount = 32;
            bi.bmiHeader.biCompression = BI_RGB;

            void* bits = nullptr;
            HDC dc = CreateCompatibleDC(nullptr);
            if (!dc) return false;
            HBITMAP bmp = CreateDIBSection(dc, &bi, DIB_RGB_COLORS, &bits, nullptr, 0);
            if (!bmp) { DeleteDC(dc); return false; }
            HGDIOBJ oldBmp = SelectObject(dc, bmp);

            RECT rc{ 0, 0, w, bandH };
            HBRUSH white = (HBRUSH)GetStockObject(WHITE_BRUSH);
            FillRect(dc, &rc, white);

            LOGFONTW lf{};
            lf.lfHeight = -(bandH * 7 / 10);
            lf.lfWeight = FW_NORMAL;
            wcscpy_s(lf.lfFaceName, L"Arial");
            HFONT font = CreateFontIndirectW(&lf);
            HGDIOBJ oldFont = SelectObject(dc, font);
            SetBkMode(dc, TRANSPARENT);
            SetTextColor(dc, RGB(0, 0, 0));
            DrawTextW(dc, text.c_str(), (int)text.size(), &rc,
                      DT_CENTER | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX);
            GdiFlush();

            px.resize((size_t)w * (h + bandH), 1);
            const uint32_t* src = (const uint32_t*)bits;
            for (int y = 0; y < bandH; ++y)
                for (int x = 0; x < w; ++x)
                {
                    const uint32_t v = src[(size_t)y * w + x];
                    // luminance threshold — anti-aliased glyph edges go black
                    const int lum = ((v >> 16 & 0xFF) + (v >> 8 & 0xFF) + (v & 0xFF)) / 3;
                    px[(size_t)(h + y) * w + x] = lum < 160 ? 0 : 1;
                }
            h += bandH;

            SelectObject(dc, oldFont); DeleteObject(font);
            SelectObject(dc, oldBmp); DeleteObject(bmp);
            DeleteDC(dc);
            return true;
        }

        // Encode → temp BMP → queue deferred insertion into the calling cell.
        // Returns "" on success (VB parity) or an error ExcelObj*.
        ExcelObj* encodeAndInsert(const std::wstring& text, ZXing::BarcodeFormat fmt,
                                  bool is2D, int margin, bool showText,
                                  const std::wstring& showTextStr, int bandH)
        {
            const int W = is2D ? 512 : 800;
            const int H = is2D ? 512 : 200;

            auto m = bc::encodeMatrix(text, fmt, W, H, margin);
            if (!m) return returnValue(CellError::Value);
            int mw = m->width(), mh = m->height();
            auto px = bc::matrixToPixels(*m);
            if (showText && !appendTextBand(px, mw, mh, showTextStr, bandH))
                return returnValue(CellError::Value);

            const std::wstring path = tempBmpPath();
            if (!bc::writeBmp24(path, mw, mh, px)) return returnValue(CellError::Value);

            const std::wstring fullAddr = xloil::CallerInfo().address();
            std::wstring cell; splitCaller(fullAddr, cell);
            if (cell.empty()) { DeleteFileW(path.c_str()); return returnValue(CellError::Ref); }

            // mode 0 = keep aspect ratio within the cell (2D); 1 = fill cell (1D).
            queueInsertPicture(fullAddr, cell, path, is2D ? 0 : 1, /*deleteLocalAfter*/ true);
            return returnValue(ExcelObj(std::wstring_view(L"")));
        }

        // GS1 text for the target format. Code128 gets full FNC1 compliance via
        // the writer's U+00F1 escape (leading FNC1 included); QR/DM approximate
        // with ASCII 29 separators.
        bool gs1Text(const std::wstring& input, ZXing::BarcodeFormat fmt, std::wstring& out)
        {
            std::wstring err;
            if (fmt == ZXing::BarcodeFormat::Code128)
            {
                std::wstring s = bc::gs1ToStream(input, (wchar_t)0x00F1, err);
                if (s.empty()) return false;
                out = std::wstring(1, (wchar_t)0x00F1) + s;
            }
            else
            {
                out = bc::gs1ToStream(input, (wchar_t)0x001D, err);
                if (out.empty()) return false;
            }
            return true;
        }

        // ── READBARCODE: WIC pixel load ──────────────────────────────────────
        struct ComScope
        {
            bool uninit = false;
            ComScope()
            {
                HRESULT hr = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
                uninit = SUCCEEDED(hr);            // balance only when we took a ref
            }
            ~ComScope() { if (uninit) CoUninitialize(); }
        };

        template <class T> struct ComPtr
        {
            T* p = nullptr;
            ~ComPtr() { if (p) p->Release(); }
            T** operator&() { return &p; }
            T* operator->() { return p; }
            explicit operator bool() const { return p != nullptr; }
        };

        bool loadPixelsBgr(const std::wstring& file, std::vector<uint8_t>& buf, UINT& w, UINT& h)
        {
            ComScope com;
            ComPtr<IWICImagingFactory> factory;
            if (FAILED(CoCreateInstance(CLSID_WICImagingFactory, nullptr, CLSCTX_INPROC_SERVER,
                                        IID_IWICImagingFactory, (void**)&factory)))
                return false;
            ComPtr<IWICBitmapDecoder> decoder;
            if (FAILED(factory->CreateDecoderFromFilename(file.c_str(), nullptr, GENERIC_READ,
                                                          WICDecodeMetadataCacheOnDemand, &decoder)))
                return false;
            ComPtr<IWICBitmapFrameDecode> frame;
            if (FAILED(decoder->GetFrame(0, &frame))) return false;
            ComPtr<IWICBitmapSource> converted;
            if (FAILED(WICConvertBitmapSource(GUID_WICPixelFormat24bppBGR,
                                              frame.p, &converted)))
                return false;
            if (FAILED(converted->GetSize(&w, &h)) || w == 0 || h == 0) return false;
            const UINT stride = w * 3;
            buf.resize((size_t)stride * h);
            return SUCCEEDED(converted->CopyPixels(nullptr, stride, (UINT)buf.size(), buf.data()));
        }

        // Escape ASCII control chars like the VB CODETEXT view (<GS>, <TAB>, …).
        std::wstring escapeControls(const std::wstring& s)
        {
            static const wchar_t* names[33] = {
                L"<NUL>",L"<SOH>",L"<STX>",L"<ETX>",L"<EOT>",L"<ENQ>",L"<ACK>",L"<BEL>",
                L"<BS>",L"<TAB>",L"<LF>",L"<VT>",L"<FF>",L"<CR>",L"<SO>",L"<SI>",
                L"<DLE>",L"<DC1>",L"<DC2>",L"<DC3>",L"<DC4>",L"<NAK>",L"<SYN>",L"<ETB>",
                L"<CAN>",L"<EM>",L"<SUB>",L"<ESC>",L"<FS>",L"<GS>",L"<RS>",L"<US>",L"<DEL>" };
            std::wstring out;
            for (wchar_t c : s)
            {
                if (c < 32) out += names[c];
                else if (c == 127) out += names[32];
                else out.push_back(c);
            }
            return out;
        }
    }

    void registerBarcode()
    {
        namespace core = egtools::core;

        // ── encode family (macro-type: xlfCaller + deferred COM insert) ──────
        auto encodeFamily = [](const std::wstring& text, bc::TypeInfo ti,
                               int margin, bool showText, int bandH,
                               const std::wstring& original) -> ExcelObj*
        {
            std::wstring real = text;
            if (ti.gs1 && !gs1Text(text, ti.fmt, real)) return returnValue(CellError::Value);
            return encodeAndInsert(real, ti.fmt, ti.is2D, margin,
                                   showText && !ti.is2D, original, bandH);
        };

        // BARCODE(text, [type], [option]) — option: number → margin, TRUE/FALSE → show text.
        core::registerFn(L"BARCODE",
            [encodeFamily](const ExcelObj& textA, const ExcelObj& typeA, const ExcelObj& optA) -> ExcelObj*
            {
                const std::wstring text = textA.toString();
                if (text.empty()) return returnValue(CellError::Value);
                const std::wstring tname = typeA.isMissing() ? L"QRCODE" : typeA.toString();
                auto ti = bc::parseType(tname);
                if (!ti) return returnValue(CellError::Value);

                int margin = 0; bool showText = false;
                if (!optA.isMissing())
                {
                    if (optA.isType(ExcelType::Bool)) showText = optA.get<bool>(false);
                    else { margin = optA.get<int>(0); if (margin < 0) margin = 0; }
                }
                return encodeFamily(text, *ti, margin, showText, 60, text);
            },
            /*macro*/ true, /*threadsafe*/ false);

        // QRCODE(text, [margin], [gs1])
        core::registerFn(L"QRCODE",
            [encodeFamily](const ExcelObj& textA, const ExcelObj& marginA, const ExcelObj& gs1A) -> ExcelObj*
            {
                const std::wstring text = textA.toString();
                if (text.empty()) return returnValue(CellError::Value);
                bc::TypeInfo ti{ ZXing::BarcodeFormat::QRCode, true, gs1A.get<bool>(false) };
                int margin = marginA.get<int>(0); if (margin < 0) margin = 0;
                return encodeFamily(text, ti, margin, false, 0, text);
            },
            /*macro*/ true, /*threadsafe*/ false);

        // DATAMATRIX(text, [margin], [gs1])
        core::registerFn(L"DATAMATRIX",
            [encodeFamily](const ExcelObj& textA, const ExcelObj& marginA, const ExcelObj& gs1A) -> ExcelObj*
            {
                const std::wstring text = textA.toString();
                if (text.empty()) return returnValue(CellError::Value);
                bc::TypeInfo ti{ ZXing::BarcodeFormat::DataMatrix, true, gs1A.get<bool>(false) };
                int margin = marginA.get<int>(0); if (margin < 0) margin = 0;
                return encodeFamily(text, ti, margin, false, 0, text);
            },
            /*macro*/ true, /*threadsafe*/ false);

        // CODE128(text, [show_text], [gs1], [text_size])
        core::registerFn(L"CODE128",
            [encodeFamily](const ExcelObj& textA, const ExcelObj& showA,
                           const ExcelObj& gs1A, const ExcelObj& sizeA) -> ExcelObj*
            {
                const std::wstring text = textA.toString();
                if (text.empty()) return returnValue(CellError::Value);
                bc::TypeInfo ti{ ZXing::BarcodeFormat::Code128, false, gs1A.get<bool>(false) };
                const bool showText = showA.get<bool>(false);
                int bandH = sizeA.get<int>(0);
                bandH = bandH > 0 ? bandH * 5 : 60;      // VB font-size knob, scaled to 800px art
                return encodeFamily(text, ti, 0, showText, bandH, text);
            },
            /*macro*/ true, /*threadsafe*/ false);

        // ── READBARCODE(source, [result_type]) ───────────────────────────────
        core::registerFn(L"READBARCODE",
            [](const ExcelObj& srcA, const ExcelObj& typeA) -> ExcelObj*
            {
                const std::wstring source = srcA.toString();
                if (source.empty()) return returnValue(CellError::Value);

                // result_type grid (default 1 = TEXT)
                std::vector<std::vector<ExcelObj>> shape;
                if (typeA.isType(ExcelType::Multi))
                {
                    ExcelArray a(typeA);
                    for (ExcelArray::row_t r = 0; r < a.nRows(); ++r)
                    {
                        std::vector<ExcelObj> row;
                        for (ExcelArray::col_t c = 0; c < a.nCols(); ++c) row.emplace_back(a.at(r, c));
                        shape.push_back(std::move(row));
                    }
                }
                else if (typeA.isMissing()) shape.push_back({ ExcelObj(1) });
                else shape.push_back({ ExcelObj(typeA) });

                // load image (URL → temp download)
                std::wstring path = source;
                bool isTemp = false;
                if (source.rfind(L"http://", 0) == 0 || source.rfind(L"https://", 0) == 0)
                {
                    wchar_t tmp[MAX_PATH]{};
                    GetTempPathW(MAX_PATH, tmp);
                    path = std::wstring(tmp) + L"egbrd_" + std::to_wstring(GetTickCount64()) + L".img";
                    if (URLDownloadToFileW(nullptr, source.c_str(), path.c_str(), 0, nullptr) != S_OK)
                        return returnValue(CellError::Null);
                    isTemp = true;
                }
                else if (GetFileAttributesW(source.c_str()) == INVALID_FILE_ATTRIBUTES)
                    return returnValue(CellError::NA);

                std::vector<uint8_t> buf; UINT w = 0, h = 0;
                const bool loaded = loadPixelsBgr(path, buf, w, h);
                if (isTemp) DeleteFileW(path.c_str());
                if (!loaded) return returnValue(CellError::NA);

                ZXing::ReaderOptions opts;
                opts.setTryHarder(true);
                opts.setTryRotate(true);
                auto res = ZXing::ReadBarcode(
                    ZXing::ImageView(buf.data(), (int)w, (int)h, ZXing::ImageFormat::BGR), opts);
                if (!res.isValid()) return returnValue(CellError::NA);

                const std::wstring text = fromUtf8(res.text());
                const std::wstring fmt = fromUtf8(ZXing::ToString(res.format()));
                const std::wstring symId = fromUtf8(res.symbologyIdentifier());

                auto cellFor = [&](const ExcelObj& t) -> ExcelObj
                {
                    std::wstring key = t.toString();
                    for (auto& c : key) c = (wchar_t)towupper(c);
                    if (key == L"1" || key == L"TEXT") return ExcelObj(std::wstring_view(text));
                    if (key == L"2" || key == L"CODETEXT")
                    {
                        std::wstring body = text;
                        if (!body.empty() && body[0] < 32)
                        {
                            body = body.substr(1);
                            if (symId != L"]d2") body = L"<FCN>" + body;
                        }
                        return ExcelObj(std::wstring_view(symId + escapeControls(body)));
                    }
                    if (key == L"3" || key == L"FORMAT") return ExcelObj(std::wstring_view(fmt));
                    if (key == L"4" || key == L"SYMBOLOGYID") return ExcelObj(std::wstring_view(symId));
                    if (key == L"5" || key == L"RAWBYTES")
                    {
                        const auto& bytes = res.bytes();
                        if (bytes.empty()) return ExcelObj(CellError::NA);
                        std::wstring hex;
                        static const wchar_t* kHex = L"0123456789ABCDEF";
                        for (size_t i = 0; i < bytes.size(); ++i)
                        {
                            if (i) hex.push_back(L' ');
                            hex.push_back(kHex[bytes[i] >> 4]);
                            hex.push_back(kHex[bytes[i] & 0x0F]);
                        }
                        return ExcelObj(std::wstring_view(hex));
                    }
                    return ExcelObj(CellError::Value);
                };

                const auto rows = (ExcelArrayBuilder::row_t)shape.size();
                const auto cols = (ExcelArrayBuilder::col_t)shape[0].size();
                std::vector<ExcelObj> flat;
                for (auto& row : shape)
                    for (auto& t : row) flat.emplace_back(cellFor(t));
                return core::output(core::makeArray(rows, cols, flat));
            },
            /*macro*/ false, /*threadsafe*/ false);   // WIC COM + network

        // ── READGS1(gs1_text, [concat]) ──────────────────────────────────────
        core::registerFn(L"READGS1",
            [](const ExcelObj& srcA, const ExcelObj& joinA) -> ExcelObj*
            {
                std::wstring src = srcA.toString();
                while (!src.empty() && (src.front() == L' ' || src.front() == L'\t')) src.erase(0, 1);
                while (!src.empty() && (src.back() == L' ' || src.back() == L'\t')) src.pop_back();
                if (src.empty()) return returnValue(CellError::Value);
                const bool join = joinA.get<bool>(false);

                std::vector<std::pair<std::wstring, std::wstring>> parts;
                if (src.find(L'(') != std::wstring::npos)
                {
                    // bracketed input: (AI)value(AI)value…
                    size_t pos = 0;
                    while ((pos = src.find(L'(', pos)) != std::wstring::npos)
                    {
                        size_t close = src.find(L')', pos + 1);
                        if (close == std::wstring::npos) break;
                        size_t next = src.find(L'(', close + 1);
                        const std::wstring ai = src.substr(pos + 1, close - pos - 1);
                        const std::wstring val = src.substr(close + 1,
                            (next == std::wstring::npos ? src.size() : next) - close - 1);
                        parts.emplace_back(ai, val);
                        pos = (next == std::wstring::npos) ? src.size() : next;
                    }
                }
                else
                {
                    // decoded stream: [FNC1] AI value … with GS separators
                    size_t i = 0;
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
                            auto fit = bc::fixedLenAIs().find(ai);
                            if (fit != bc::fixedLenAIs().end()) { expected = fit->second; inValue = true; }
                            else
                            {
                                auto vit = bc::varLenAIs().find(ai);
                                if (vit != bc::varLenAIs().end()) { expected = vit->second; inValue = true; }
                                else if (ai.size() >= 4) return returnValue(CellError::Num);
                            }
                        }
                    }
                    flush();
                }
                if (parts.empty()) return returnValue(CellError::NA);

                if (join)
                {
                    std::wstring one;
                    for (auto& [ai, val] : parts) one += L"(" + ai + L")" + val;
                    return returnValue(ExcelObj(std::wstring_view(one)));
                }
                std::vector<ExcelObj> flat;
                for (auto& [ai, val] : parts)
                {
                    flat.emplace_back(std::wstring_view(L"(" + ai + L")"));
                    flat.emplace_back(std::wstring_view(val));
                }
                return core::output(core::makeArray(
                    (ExcelArrayBuilder::row_t)parts.size(), 2, flat));
            });
    }
}
