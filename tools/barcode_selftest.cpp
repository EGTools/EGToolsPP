// barcode_selftest.cpp — console round-trip helper for the barcode core.
// Encodes sample barcodes to BMP files in %TEMP% using the same BarcodeCore.h
// the add-in uses; tests/ then decode them through READBARCODE (headless Excel)
// to verify both directions. Never shipped (EXCLUDE_FROM_ALL).

#include "functions/BarcodeCore.h"

#include <windows.h>
#include <cstdio>
#include <string>

namespace bc = egtools::barcode;

int wmain()
{
    wchar_t tmp[MAX_PATH]{};
    GetTempPathW(MAX_PATH, tmp);
    const std::wstring dir = tmp;

    struct Case { const wchar_t* name; const wchar_t* text; ZXing::BarcodeFormat fmt; int w, h, margin; };
    const Case cases[] = {
        { L"qr",     L"EGTools++ QR 한글", ZXing::BarcodeFormat::QRCode,     512, 512, 2 },
        { L"c128",   L"ABC-12345",         ZXing::BarcodeFormat::Code128,    800, 200, 0 },
        { L"dm",     L"DM-TEST-99",        ZXing::BarcodeFormat::DataMatrix, 512, 512, 2 },
        { L"ean13",  L"4006381333931",     ZXing::BarcodeFormat::EAN13,      800, 200, 0 },
    };

    int fail = 0;
    for (const auto& c : cases)
    {
        const std::wstring path = dir + L"egselftest_" + c.name + L".bmp";
        if (bc::encodeToBmp(c.text, c.fmt, c.w, c.h, c.margin, path))
            wprintf(L"OK %s %s\n", c.name, path.c_str());
        else { wprintf(L"FAIL %s\n", c.name); ++fail; }
    }

    // GS1-128 stream (FNC1 escapes) — encode only; decoding checks symbology id.
    std::wstring err;
    std::wstring gs1 = bc::gs1ToStream(L"(01)04012345678901(10)LOT42", (wchar_t)0x00F1, err);
    if (!gs1.empty())
    {
        const std::wstring path = dir + L"egselftest_gs1128.bmp";
        std::wstring full = std::wstring(1, (wchar_t)0x00F1) + gs1;
        if (bc::encodeToBmp(full, ZXing::BarcodeFormat::Code128, 800, 200, 0, path))
            wprintf(L"OK gs1128 %s\n", path.c_str());
        else { wprintf(L"FAIL gs1128 encode\n"); ++fail; }
    }
    else { wprintf(L"FAIL gs1 stream: %s\n", err.c_str()); ++fail; }

    return fail;
}
