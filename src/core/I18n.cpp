#include "I18n.h"
#include "resource.h"   // resource ids (resources/ is on the include path)

#include <windows.h>
#include <xlOil/State.h>
#include <map>
#include <mutex>
#include <string>
#include <cwctype>
#include <nlohmann/json.hpp>

using json = nlohmann::json;

namespace egtools::i18n
{
    namespace
    {
        std::map<std::wstring, json> g_cat;     // lang code -> parsed catalog
        std::wstring                 g_current = L"en";
        std::mutex                   g_mtx;

        std::wstring toW(const std::string& s)
        {
            if (s.empty()) return L"";
            int n = MultiByteToWideChar(CP_UTF8, 0, s.data(), (int)s.size(), nullptr, 0);
            std::wstring w(n, 0);
            MultiByteToWideChar(CP_UTF8, 0, s.data(), (int)s.size(), w.data(), n);
            return w;
        }
        std::string toU8(const std::wstring& w)
        {
            if (w.empty()) return "";
            int n = WideCharToMultiByte(CP_UTF8, 0, w.data(), (int)w.size(), nullptr, 0, nullptr, nullptr);
            std::string s(n, 0);
            WideCharToMultiByte(CP_UTF8, 0, w.data(), (int)w.size(), s.data(), n, nullptr, nullptr);
            return s;
        }

        // Handle of the module that contains this code (our .xll).
        HMODULE selfModule()
        {
            HMODULE m = nullptr;
            GetModuleHandleExW(
                GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS |
                GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
                reinterpret_cast<LPCWSTR>(&selfModule), &m);
            return m;
        }

        std::string loadResource(int id)
        {
            HMODULE mod = selfModule();
            HRSRC h = FindResourceW(mod, MAKEINTRESOURCEW(id), RT_RCDATA);
            if (!h) return "";
            HGLOBAL g = LoadResource(mod, h);
            DWORD sz = SizeofResource(mod, h);
            const char* p = (const char*)LockResource(g);
            if (!p || sz == 0) return "";
            return std::string(p, sz);
        }

        LCID officeUILcid()
        {
            // 실행 중인 Excel의 메이저 버전(2010=14 / 2013=15 / 2016+=16) 키만
            // 읽는다 — 버전 하드코딩 시 2010/2013에서 표시 언어를 못 읽고, 여러
            // Office가 공존하면 다른 버전의 설정을 읽는 문제. 값이 없으면(표시
            // 언어를 "Windows와 일치"로 둔 경우 포함) 0을 돌려 호출부가
            // GetUserDefaultUILanguage로 폴백한다.
            const int major = xloil::Environment::excelProcess().version;
            if (major < 12) return 0;   // 2007 미만은 LanguageResources 레이아웃 상이
            wchar_t path[96];
            swprintf_s(path, L"Software\\Microsoft\\Office\\%d.0\\Common\\LanguageResources",
                       major);
            DWORD val = 0, sz = sizeof(val);
            if (RegGetValueW(HKEY_CURRENT_USER, path,
                    L"UILanguage", RRF_RT_REG_DWORD, nullptr, &val, &sz) == ERROR_SUCCESS)
                return (LCID)val;
            return 0;
        }

        std::wstring lcidToCode(LCID lcid)
        {
            WORD prim = PRIMARYLANGID(LANGIDFROMLCID(lcid));
            WORD sub  = SUBLANGID(LANGIDFROMLCID(lcid));
            switch (prim)
            {
            case LANG_KOREAN:   return L"ko";
            case LANG_JAPANESE: return L"ja";
            case LANG_SPANISH:  return L"es";
            case LANG_ENGLISH:  return L"en";
            case LANG_CHINESE:
                if (sub == SUBLANG_CHINESE_TRADITIONAL ||
                    sub == SUBLANG_CHINESE_HONGKONG ||
                    sub == SUBLANG_CHINESE_MACAU)
                    return L"zh-TW";
                return L"zh-CN";
            default:            return L"en";
            }
        }
    }

    void load()
    {
        struct Entry { const wchar_t* code; int id; };
        static const Entry table[] = {
            { L"en",    IDR_I18N_EN },
            { L"ko",    IDR_I18N_KO },
            { L"zh-CN", IDR_I18N_ZH_CN },
            { L"zh-TW", IDR_I18N_ZH_TW },
            { L"ja",    IDR_I18N_JA },
            { L"es",    IDR_I18N_ES },
        };
        std::lock_guard<std::mutex> lk(g_mtx);
        for (const auto& e : table)
        {
            std::string s = loadResource(e.id);
            if (s.empty()) continue;
            try { g_cat[e.code] = json::parse(s); } catch (...) {}
        }
    }

    std::wstring detectAndSet()
    {
        LCID lcid = officeUILcid();
        if (!lcid) lcid = GetUserDefaultUILanguage();
        std::wstring code = lcidToCode(lcid);
        {
            std::lock_guard<std::mutex> lk(g_mtx);
            if (g_cat.find(code) == g_cat.end()) code = L"en";
            g_current = code;
        }
        return code;
    }

    void setCurrent(const std::wstring& lang)
    {
        std::lock_guard<std::mutex> lk(g_mtx);
        if (g_cat.find(lang) != g_cat.end()) g_current = lang;
    }

    std::wstring current()
    {
        std::lock_guard<std::mutex> lk(g_mtx);
        return g_current;
    }

    std::wstring t(const std::wstring& key)
    {
        const std::string k = toU8(key);
        std::lock_guard<std::mutex> lk(g_mtx);
        auto get = [&](const std::wstring& lang) -> std::wstring {
            auto it = g_cat.find(lang);
            if (it == g_cat.end()) return L"";
            if (it->second.contains(k) && it->second[k].is_string())
                return toW(it->second[k].get<std::string>());
            return L"";
        };
        std::wstring v = get(g_current); if (!v.empty()) return v;
        v = get(L"en");                  if (!v.empty()) return v;
        return key;
    }

    bool func(const std::wstring& name,
              std::wstring& desc,
              std::vector<ArgMeta>& args,
              RepeatSpec* repeat)
    {
        if (repeat) *repeat = RepeatSpec{};
        const std::string n = toU8(name);
        std::lock_guard<std::mutex> lk(g_mtx);
        auto tryLang = [&](const std::wstring& lang) -> bool {
            auto it = g_cat.find(lang);
            if (it == g_cat.end()) return false;
            const json& j = it->second;
            if (!j.contains("func") || !j["func"].contains(n)) return false;
            const json& fi = j["func"][n];
            desc = (fi.contains("desc") && fi["desc"].is_string())
                       ? toW(fi["desc"].get<std::string>()) : L"";
            args.clear();
            if (fi.contains("args") && fi["args"].is_array())
                for (const auto& a : fi["args"])
                    args.push_back(ArgMeta{
                        a.contains("name") ? toW(a["name"].get<std::string>()) : L"",
                        a.contains("help") ? toW(a["help"].get<std::string>()) : L"",
                        a.value("opt", false) });
            return true;
        };

        std::wstring matched;
        if (tryLang(g_current)) matched = g_current;
        else if (tryLang(L"en")) matched = L"en";
        else return false;

        // Optionality and the repeat spec are language-independent and live only
        // in en.json. Overlay both from the en catalog (catalogs are parallel; the
        // arg `opt` overlay is positional and bounds-guarded).
        {
            auto it = g_cat.find(L"en");
            if (it != g_cat.end() && it->second.contains("func") &&
                it->second["func"].contains(n))
            {
                const json& en = it->second["func"][n];
                if (matched != L"en" && en.contains("args") && en["args"].is_array())
                {
                    const json& ea = en["args"];
                    for (size_t i = 0; i < args.size() && i < ea.size(); ++i)
                        args[i].opt = ea[i].value("opt", false);
                }
                if (repeat && en.contains("repeat") && en["repeat"].is_object())
                {
                    const json& r = en["repeat"];
                    repeat->head   = r.value("head", 0);
                    repeat->period = r.value("period", 0);
                    repeat->tail   = r.value("tail", 0);
                    repeat->max    = r.value("max", 0);
                    repeat->has    = repeat->period > 0;
                }
            }
        }
        return true;
    }

    std::wstring repeatArgName(const std::wstring& representative, int iter)
    {
        if (iter <= 1) return representative;
        // Strip a trailing run of digits, then append the iteration number.
        size_t end = representative.size();
        while (end > 0 && iswdigit(representative[end - 1])) --end;
        return representative.substr(0, end) + std::to_wstring(iter);
    }

    std::vector<ArgMeta> expandRepeatArgs(const std::vector<ArgMeta>& rep,
                                          const RepeatSpec& spec)
    {
        std::vector<ArgMeta> out;
        if (!spec.has || (int)rep.size() != spec.head + spec.period + spec.tail
            || spec.max < spec.head + spec.period + spec.tail)
            return rep;   // malformed — fall back to the representative list

        const int iters = (spec.max - spec.head - spec.tail) / spec.period;
        out.reserve(spec.max);
        for (int i = 0; i < spec.head; ++i) out.push_back(rep[i]);
        for (int k = 1; k <= iters; ++k)
            for (int p = 0; p < spec.period; ++p)
            {
                const ArgMeta& u = rep[spec.head + p];
                out.push_back(ArgMeta{
                    repeatArgName(u.name, k),
                    k == 1 ? u.help : std::wstring(),
                    k == 1 ? u.opt : true });
            }
        for (int i = 0; i < spec.tail; ++i)
            out.push_back(rep[spec.head + spec.period + i]);
        return out;
    }
}
