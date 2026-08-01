// BarcodeCore.h — shared barcode helpers used by FxBarcode.cpp and the
// eg_barcode_selftest console tool: type-name mapping, GS1 AI stream
// conversion, zxing-cpp encoding and a minimal 24-bpp BMP writer.
#pragma once

#include <ZXing/BarcodeFormat.h>
#include <ZXing/BitMatrix.h>
#include <ZXing/MultiFormatWriter.h>
#include <ZXing/CharacterSet.h>

#include <cstdio>
#include <cstdint>
#include <map>
#include <optional>
#include <regex>
#include <string>
#include <vector>

namespace egtools::barcode
{
    // ── type-name → format (VB BARCODE table, writers zxing-cpp supports) ────
    // Returns format plus flags: is2D, gs1 (pre-set for the GS1-* aliases).
    struct TypeInfo { ZXing::BarcodeFormat fmt; bool is2D; bool gs1; };

    inline std::optional<TypeInfo> parseType(std::wstring t)
    {
        for (auto& c : t) c = (wchar_t)towupper(c);
        // strip spaces and hyphens for tolerant matching ("QR CODE", "CODE-128")
        std::wstring k;
        for (wchar_t c : t) if (c != L' ' && c != L'-') k.push_back(c);
        using F = ZXing::BarcodeFormat;
        static const std::map<std::wstring, TypeInfo> m = {
            { L"1", { F::Code128, false, false } },        { L"CODE128", { F::Code128, false, false } },
            { L"2", { F::Code39, false, false } },         { L"CODE39",  { F::Code39, false, false } },
            { L"3", { F::EAN13, false, false } },          { L"EAN13",   { F::EAN13, false, false } },
            { L"4", { F::EAN8, false, false } },           { L"EAN8",    { F::EAN8, false, false } },
            { L"5", { F::ITF, false, false } },            { L"ITF", { F::ITF, false, false } },
            { L"INTERLEAVED2OF5", { F::ITF, false, false } }, { L"2OF5", { F::ITF, false, false } },
            { L"CODE93", { F::Code93, false, false } },
            { L"CODABAR", { F::Codabar, false, false } },
            { L"UPCA", { F::UPCA, false, false } },        { L"UPC_A", { F::UPCA, false, false } },
            { L"UPCE", { F::UPCE, false, false } },        { L"UPC_E", { F::UPCE, false, false } },
            { L"10", { F::PDF417, true, false } },         { L"PDF417", { F::PDF417, true, false } },
            { L"11", { F::QRCode, true, false } },         { L"QRCODE", { F::QRCode, true, false } },
            { L"QR", { F::QRCode, true, false } },
            { L"12", { F::DataMatrix, true, false } },     { L"DATAMATRIX", { F::DataMatrix, true, false } },
            { L"DM", { F::DataMatrix, true, false } },
            { L"13", { F::Aztec, true, false } },          { L"AZTEC", { F::Aztec, true, false } },
            { L"21", { F::Code128, false, true } },        { L"GS1128", { F::Code128, false, true } },
            { L"GS1CODE128", { F::Code128, false, true } },
            { L"22", { F::DataMatrix, true, true } },      { L"GS1DATAMATRIX", { F::DataMatrix, true, true } },
            { L"GS1DM", { F::DataMatrix, true, true } },
            { L"23", { F::QRCode, true, true } },          { L"GS1QRCODE", { F::QRCode, true, true } },
            { L"GS1QR", { F::QRCode, true, true } },
        };
        auto it = m.find(k);
        if (it == m.end()) return std::nullopt;
        return it->second;
    }

    // ── GS1 application identifier tables (ported from VB FX5_Barcode) ──────
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

    // "(01)00852670007076(10)LOT1" → AI stream with `sep` after variable-length
    // AIs (except the last). Returns empty on a length violation.
    // sep: FNC1 escape the target encoder understands (Code128 writer: U+00F1;
    // QR/DataMatrix fallback: ASCII 29 group separator).
    inline std::wstring gs1ToStream(const std::wstring& input, wchar_t sep, std::wstring& err)
    {
        static const std::wregex pat(LR"(\((\d{2,4})\)([^\(]+))");
        std::wstring out;
        auto begin = std::wsregex_iterator(input.begin(), input.end(), pat);
        auto end = std::wsregex_iterator();
        std::vector<std::pair<std::wstring, std::wstring>> parts;
        for (auto it = begin; it != end; ++it)
            parts.emplace_back((*it)[1].str(), (*it)[2].str());
        if (parts.empty()) { err = L"no (AI)value pairs"; return L""; }

        for (size_t i = 0; i < parts.size(); ++i)
        {
            const auto& [ai, data] = parts[i];
            auto fit = fixedLenAIs().find(ai);
            if (fit != fixedLenAIs().end())
            {
                if ((int)data.size() != fit->second) { err = L"AI(" + ai + L") fixed length mismatch"; return L""; }
                out += ai; out += data;
            }
            else
            {
                auto vit = varLenAIs().find(ai);
                if (vit != varLenAIs().end() && (int)data.size() > vit->second)
                    { err = L"AI(" + ai + L") exceeds max length"; return L""; }
                out += ai; out += data;
                if (i + 1 < parts.size()) out.push_back(sep);
            }
        }
        return out;
    }

    // ── encoding ────────────────────────────────────────────────────────────
    inline std::optional<ZXing::BitMatrix> encodeMatrix(
        const std::wstring& text, ZXing::BarcodeFormat fmt,
        int width, int height, int margin)
    {
        try
        {
            ZXing::MultiFormatWriter writer(fmt);
            writer.setEncoding(ZXing::CharacterSet::UTF8);
            writer.setMargin(margin);
            return writer.encode(text, width, height);
        }
        catch (...) { return std::nullopt; }
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

    // BitMatrix (get(x,y): true = black module) → pixel buffer (0 black / 1 white).
    inline std::vector<uint8_t> matrixToPixels(const ZXing::BitMatrix& m)
    {
        const int w = m.width(), h = m.height();
        std::vector<uint8_t> px((size_t)w * h, 1);
        for (int y = 0; y < h; ++y)
            for (int x = 0; x < w; ++x)
                if (m.get(x, y)) px[(size_t)y * w + x] = 0;
        return px;
    }

    // Encode straight to a BMP file. Returns false on any failure.
    inline bool encodeToBmp(const std::wstring& text, ZXing::BarcodeFormat fmt,
                            int width, int height, int margin, const std::wstring& path)
    {
        auto m = encodeMatrix(text, fmt, width, height, margin);
        if (!m) return false;
        return writeBmp24(path, m->width(), m->height(), matrixToPixels(*m));
    }
}
