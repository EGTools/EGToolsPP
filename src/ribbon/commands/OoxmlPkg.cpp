// OoxmlPkg.cpp — 패키지 읽기 헬퍼 (OoxmlPkg.h). CmdPicture.cpp의 추출 저장
// 헬퍼를 공용으로 옮긴 것(동작 동일).

#include "OoxmlPkg.h"
#include "CmdCommon.h"

#include <cstdio>
#include <cstdlib>

namespace egtools::commands::ooxml
{
    std::string tagBetween(const std::string& s, size_t from, size_t to,
                           const std::string& open, const std::string& close)
    {
        const size_t a = s.find(open, from);
        if (a == std::string::npos || a >= to) return "";
        const size_t b = s.find(close, a + open.size());
        if (b == std::string::npos || b > to) return "";
        return s.substr(a + open.size(), b - a - open.size());
    }

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

    std::string unescapeXml(const std::string& s)
    {
        if (s.find('&') == std::string::npos) return s;
        std::string out;
        out.reserve(s.size());
        for (size_t i = 0; i < s.size(); ++i)
        {
            if (s[i] != '&') { out.push_back(s[i]); continue; }
            const size_t semi = s.find(';', i);
            if (semi == std::string::npos || semi - i > 10) { out.push_back('&'); continue; }
            const std::string ent = s.substr(i + 1, semi - i - 1);
            if      (ent == "amp")  out.push_back('&');
            else if (ent == "lt")   out.push_back('<');
            else if (ent == "gt")   out.push_back('>');
            else if (ent == "quot") out.push_back('"');
            else if (ent == "apos") out.push_back('\'');
            else if (!ent.empty() && ent[0] == '#')
            {
                const unsigned long cp = (ent.size() > 1 && (ent[1] == 'x' || ent[1] == 'X'))
                    ? std::strtoul(ent.c_str() + 2, nullptr, 16)
                    : std::strtoul(ent.c_str() + 1, nullptr, 10);
                // 코드포인트 → UTF-8
                if (cp < 0x80) out.push_back((char)cp);
                else if (cp < 0x800)
                { out.push_back((char)(0xC0 | (cp >> 6))); out.push_back((char)(0x80 | (cp & 0x3F))); }
                else if (cp < 0x10000)
                {
                    out.push_back((char)(0xE0 | (cp >> 12)));
                    out.push_back((char)(0x80 | ((cp >> 6) & 0x3F)));
                    out.push_back((char)(0x80 | (cp & 0x3F)));
                }
                else
                {
                    out.push_back((char)(0xF0 | (cp >> 18)));
                    out.push_back((char)(0x80 | ((cp >> 12) & 0x3F)));
                    out.push_back((char)(0x80 | ((cp >> 6) & 0x3F)));
                    out.push_back((char)(0x80 | (cp & 0x3F)));
                }
            }
            else { out.push_back('&'); continue; }   // 모르는 참조는 그대로
            i = semi;
        }
        return out;
    }

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

    std::string relTarget(const std::filesystem::path& relsFile,
                          const std::string& rid, const char* typeContains)
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
    }
}
