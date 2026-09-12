// barcode_selftest.cpp — console round-trip helper for the barcode core.
// Encodes sample barcodes with zint (the same BarcodeCore.h the add-in uses)
// to BMP files in %TEMP% and prints one machine-readable line per case:
//   CASE<TAB>name<TAB>bmp path<TAB>expected TEXT<TAB>expected symbology id
// tools/dev/barcode_roundtrip.ps1 decodes them through READBARCODE (headless
// Excel, zxing-cpp) and compares — proves generation(zint) ↔ reading(zxing)
// agree, including GS1 FNC1 placement (]C1/]d2/]Q3). Cases that must be
// REJECTED by the GS1 linter print REJECT lines instead. Never shipped.

#include "functions/BarcodeCore.h"

#include <windows.h>
#include <cstdio>
#include <fcntl.h>
#include <io.h>
#include <string>

namespace bc = egtools::barcode;

int wmain()
{
    _setmode(_fileno(stdout), _O_U8TEXT);   // UTF-8 stdout (Korean sample text)
    wchar_t tmp[MAX_PATH]{};
    GetTempPathW(MAX_PATH, tmp);
    const std::wstring dir = tmp;

    // expect: READBARCODE result_type 1(TEXT). GS1 symbols come back from zxing in
    // bracketed (AI)value form, i.e. equal to the input. "" = print only.
    struct Case { const wchar_t* name; const wchar_t* type; const wchar_t* text;
                  int margin; const wchar_t* expect; const wchar_t* symId; };
    const Case cases[] = {
        { L"code128",  L"CODE128", L"ABC-12345",     0, L"ABC-12345",     L"]C0" },
        { L"code39",   L"CODE39",  L"CODE39",        0, L"CODE39",        L"]A0" },
        { L"code93",   L"CODE93",  L"CODE93",        0, L"CODE93",        L"]G0" },
        { L"ean13",    L"EAN13",   L"4006381333931", 0, L"4006381333931", L"]E0" },
        { L"ean13_12", L"EAN13",   L"400638133393",  0, L"4006381333931", L"]E0" },
        { L"ean8",     L"EAN8",    L"96385074",      0, L"96385074",      L"]E4" },
        { L"upca",     L"UPCA",    L"036000291452",  0, L"036000291452",  L"]E0" },
        { L"upce",     L"UPCE",    L"0123456",       0, L"01234565",      L"]E0" },
        { L"itf",      L"ITF",     L"1234567890",    0, L"1234567890",    L"]I0" },
        { L"codabar",  L"CODABAR", L"A12345B",       0, L"A12345B",       L"]F0" },
        { L"qr",       L"QRCODE",  L"EGTools++ QR 한글", 0, L"EGTools++ QR 한글", L"]Q1" },
        { L"qr_m4",    L"QRCODE",  L"margin",        4, L"margin",        L"]Q1" },
        { L"dm",       L"DATAMATRIX", L"DM-TEST-99", 0, L"DM-TEST-99",    L"]d1" },
        { L"aztec",    L"AZTEC",   L"Aztec sample",  0, L"Aztec sample",  L"]z0" },
        { L"pdf417",   L"PDF417",  L"PDF417 sample text", 0, L"PDF417 sample text", L"]L2" },
        { L"gs1128",   L"GS1128",  L"(01)04012345678901(10)LOT42", 0,
                                   L"(01)04012345678901(10)LOT42", L"]C1" },
        { L"gs1128s",  L"GS1128",  L"(10)LOT42(21)SER1", 0, L"(10)LOT42(21)SER1", L"]C1" },
        { L"gs1dm",    L"GS1DATAMATRIX", L"(01)04012345678901(17)261231(10)LOT42", 0,
                                   L"(01)04012345678901(17)261231(10)LOT42", L"]d2" },
        { L"gs1dms",   L"GS1DM",   L"(10)LOT42(21)SER1", 0, L"(10)LOT42(21)SER1", L"]d2" },
        { L"gs1qr",    L"GS1QRCODE", L"(01)04012345678901(10)LOT42", 0,
                                   L"(01)04012345678901(10)LOT42", L"]Q3" },
        { L"gs1qrs",   L"GS1QR",   L"(10)LOT42(21)SER1", 0, L"(10)LOT42(21)SER1", L"]Q3" },
        // alternative GS1 inputs → same symbol as the bracketed form
        { L"gs1_sq",   L"GS1DM",   L"[01]04012345678901[10]LOT42", 0,
                                   L"(01)04012345678901(10)LOT42", L"]d2" },
        { L"gs1_raw",  L"GS1DM",   L"0104012345678901\x1D""10LOT42", 0,
                                   L"(01)04012345678901(10)LOT42", L"]d2" },
        { L"gs1_raws", L"GS1QR",   L"10LOT42\x1D""21SER1", 0, L"(10)LOT42(21)SER1", L"]Q3" },
        { L"gs1_rawid",L"GS1128",  L"]C10104012345678901\x1D""10LOT42", 0,
                                   L"(01)04012345678901(10)LOT42", L"]C1" },
        { L"gs1_fix",  L"GS1DM",   L"010401234567890117261231", 0,
                                   L"(01)04012345678901(17)261231", L"]d2" },
        // user sample: bracketed vs CODETEXT-style raw (literal "<GS>" token and real 0x1D)
        { L"udi_br",   L"GS1DM",   L"(01)28809083947822(10)CCN6I003(17)290911", 0,
                                   L"(01)28809083947822(10)CCN6I003(17)290911", L"]d2" },
        { L"udi_tok",  L"GS1DM",   L"]d2012880908394782210CCN6I003<GS>17290911", 0,
                                   L"(01)28809083947822(10)CCN6I003(17)290911", L"]d2" },
        { L"udi_gs",   L"GS1DM",   L"]d2012880908394782210CCN6I003\x1D""17290911", 0,
                                   L"(01)28809083947822(10)CCN6I003(17)290911", L"]d2" },
    };

    int fail = 0;
    for (const auto& c : cases)
    {
        auto ti = bc::parseType(c.type);
        if (!ti) { wprintf(L"FAIL %s: unknown type %s\n", c.name, c.type); ++fail; continue; }
        const std::wstring path = dir + L"egselftest_" + c.name + L".bmp";
        std::string err;
        if (bc::encodeToBmp(c.text, *ti, c.margin, path, &err))
            wprintf(L"CASE\t%s\t%s\t%s\t%s\n", c.name, path.c_str(), c.expect, c.symId);
        else { wprintf(L"FAIL %s: %hs\n", c.name, err.c_str()); ++fail; }
    }

    // GS1 linter must reject these (wrong length / check digit / unknown AI /
    // non-ASCII) — the add-in maps rejection to #VALUE!.
    struct Bad { const wchar_t* name; const wchar_t* type; const wchar_t* text; };
    const Bad bad[] = {
        { L"gs1_len",    L"GS1DM",  L"(01)0401234567890(10)LOT" },
        { L"gs1_chk",    L"GS1128", L"(01)04012345678902" },
        { L"gs1_ai",     L"GS1QR",  L"(999)ABC" },
        { L"gs1_ascii",  L"GS1DM",  L"(10)한글" },
        { L"ean13_chk",  L"EAN13",  L"4006381333930" },
        { L"gs1_rawai",  L"GS1DM",  L"4999ABC" },
    };
    for (const auto& b : bad)
    {
        auto ti = bc::parseType(b.type);
        bc::ModuleBitmap mb; std::string err;
        if (ti && !bc::encodeModules(b.text, *ti, 0, mb, &err))
            wprintf(L"REJECT\t%s\t%hs\n", b.name, err.c_str());
        else { wprintf(L"FAIL %s: should have been rejected\n", b.name); ++fail; }
    }

    wprintf(L"SELFTEST %s (%d failures)\n", fail ? L"FAIL" : L"OK", fail);
    return fail;
}
