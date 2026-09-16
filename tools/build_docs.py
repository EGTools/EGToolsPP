#!/usr/bin/env python3
"""Build the multilingual manual site for GitHub Pages.

docs/<lang>/ 마크다운(SoT, gen_manual.ps1 산출물)을 언어별 MkDocs Material
사이트로 빌드해 _site/<lang>/ 에 놓고, 루트에 언어 감지 index.html 을 만든다.
nav 그룹 순서·표기는 docs/_content/categories.json 이 단일 출처.
함수별 평면 URL(/<lang>/<FN>/)은 분류 폴더 URL로 리다이렉트하는 스텁을 만들어
IntelliSense 링크가 분류를 몰라도 되고, 재분류에도 URL이 안정적이다.
언어 추가 시 LANGS 한 줄 + categories.json labels 항목을 추가하면 된다.

사용:  pip install mkdocs-material
       python tools/build_docs.py
"""
import json
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DOCS = ROOT / "docs"
SITE = ROOT / "_site"
# 커스텀 도메인(manual.egtools.kr) 루트에 서빙 — 구 egtools.github.io/EGToolsPP/ 는 GitHub가 301 연결
BASE_PATH = ""
SITE_URL = "https://manual.egtools.kr" + BASE_PATH

# (docs 폴더명, Material theme.language, 전환기 표기, site_name)
LANGS = [
    ("kr", "ko", "한국어", "EGTools++ 매뉴얼"),
    ("en", "en", "English", "EGTools++ Manual"),
    ("ja", "ja", "日本語", "EGTools++ マニュアル"),
    ("es", "es", "Español", "Manual de EGTools++"),
    ("zh-CN", "zh", "简体中文", "EGTools++ 手册"),
    ("zh-TW", "zh-TW", "繁體中文", "EGTools++ 手冊"),
]

# 검색 결과 스니펫용 meta description (언어별, 함수 개수는 변동하므로 명시하지 않음)
DESCRIPTIONS = {
    "kr": "구형 Excel(2010+)에서 최신 Excel 함수를 쓰게 해 주는 EGTools++ 추가기능의 함수·리본 도움말 — XLOOKUP, FILTER, TEXTSPLIT, GROUPBY, 정규식 함수 등.",
    "en": "Function and ribbon help for EGTools++, the add-in that brings modern Excel functions — XLOOKUP, FILTER, TEXTSPLIT, GROUPBY, REGEX and more — to legacy Excel 2010+.",
    "ja": "旧バージョンの Excel(2010+)で最新の Excel 関数を使える EGTools++ アドインの関数・リボン ヘルプ — XLOOKUP、FILTER、TEXTSPLIT、GROUPBY、正規表現関数など。",
    "es": "Ayuda de funciones y cinta de EGTools++, el complemento que lleva las funciones modernas de Excel — XLOOKUP, FILTER, TEXTSPLIT, GROUPBY, REGEX y más — a Excel 2010+.",
    "zh-CN": "EGTools++ 加载项的函数与功能区帮助 — 让旧版 Excel(2010+)使用 XLOOKUP、FILTER、TEXTSPLIT、GROUPBY、正则表达式等新版函数。",
    "zh-TW": "EGTools++ 增益集的函數與功能區說明 — 讓舊版 Excel(2010+)使用 XLOOKUP、FILTER、TEXTSPLIT、GROUPBY、正規表示式等新版函數。",
}

# 페이지별 SEO 메타(tools/docs_seo.py 훅이 소비) — 언어별 문구의 단일 출처.
#   hreflang: 페이지 단위 rel=alternate 코드, locale: og:locale,
#   fnTitle : 함수 페이지 <title> 패턴(h1 은 함수명 그대로 유지),
#   homeTitle: 언어별 홈 <title>(site_name 만으로는 검색 질의를 못 담는다),
#   tail    : description 끝에 붙는 한 문장(길이가 남을 때만).
SEO = {
    "kr": {
        "hreflang": "ko", "locale": "ko_KR", "fnTitle": "{FN} 함수",
        "homeTitle": "EGTools++ 매뉴얼 — 구형 Excel에서 쓰는 최신 함수",
        "tail": "EGTools++ 추가기능으로 Excel 2010 이상에서 사용할 수 있습니다.",
    },
    "en": {
        "hreflang": "en", "locale": "en_US", "fnTitle": "{FN} function",
        "homeTitle": "EGTools++ Manual — modern Excel functions for legacy Excel",
        "tail": "Available in Excel 2010 and later with the EGTools++ add-in.",
    },
    "ja": {
        "hreflang": "ja", "locale": "ja_JP", "fnTitle": "{FN} 関数",
        "homeTitle": "EGTools++ マニュアル — 旧バージョンの Excel で使える最新関数",
        "tail": "EGTools++ アドインで Excel 2010 以降でも使えます。",
    },
    "es": {
        "hreflang": "es", "locale": "es_ES", "fnTitle": "función {FN}",
        "homeTitle": "Manual de EGTools++ — funciones modernas para Excel antiguo",
        "tail": "Disponible en Excel 2010 y posteriores con el complemento EGTools++.",
    },
    "zh-CN": {
        "hreflang": "zh-Hans", "locale": "zh_CN", "fnTitle": "{FN} 函数",
        "homeTitle": "EGTools++ 手册 — 让旧版 Excel 使用新版函数",
        "tail": "通过 EGTools++ 加载项，可在 Excel 2010 及更高版本中使用。",
    },
    "zh-TW": {
        "hreflang": "zh-Hant", "locale": "zh_TW", "fnTitle": "{FN} 函數",
        "homeTitle": "EGTools++ 手冊 — 讓舊版 Excel 使用新版函數",
        "tail": "透過 EGTools++ 增益集，可在 Excel 2010 及更新版本中使用。",
    },
}
# 모든 페이지에 공통으로 붙는 keywords — 브랜드어(EGTools/EGTools++)를 맨 앞에 둬서
# "EGTools" 질의로도 어느 페이지든 잡히게 한다. 함수 페이지는 앞에 함수명/EG.함수명이 붙는다.
KEYWORDS = {
    "kr": "EGTools, EGTools++, 엑셀 추가기능, Excel 추가 기능, xll, 엑셀 함수, Excel 2010, 동적 배열",
    "en": "EGTools, EGTools++, Excel add-in, xll, Excel functions, Excel 2010, dynamic arrays",
    "ja": "EGTools, EGTools++, Excel アドイン, xll, Excel 関数, Excel 2010, 動的配列",
    "es": "EGTools, EGTools++, complemento de Excel, xll, funciones de Excel, Excel 2010, matrices dinámicas",
    "zh-CN": "EGTools, EGTools++, Excel 加载项, xll, Excel 函数, Excel 2010, 动态数组",
    "zh-TW": "EGTools, EGTools++, Excel 增益集, xll, Excel 函數, Excel 2010, 動態陣列",
}

# 언어를 못 맞춘 검색 사용자에게 보여줄 hreflang="x-default" 대상
DEFAULT_LANG = "en"
HOOK = ROOT / "tools" / "docs_seo.py"


CATS = json.loads((DOCS / "_content" / "categories.json").read_text(encoding="utf-8"))

# 평면 리다이렉트 폴더(/<lang>/<FN>/)와 최상위 섹션 폴더가 Windows 대소문자
# 비구분에서 충돌하면 안 된다 (functions/ 도입 전 IMAGE vs image 사례, 2026-09-03).
_clash = [fn for fn in CATS["functions"] if fn.lower() in {"functions", "ribbon"}]
assert not _clash, f"top-level folder clashes with function name(s): {_clash}"


def seo_block(lang_dir: str) -> str:
    """mkdocs extra.seo — 훅이 쓰는 값만 담는다. 문자열은 JSON 으로 인용(YAML 호환)."""
    e = SEO[lang_dir]
    lines = [
        f"    root: {json.dumps(SITE_URL)}",
        f"    dir: {json.dumps(lang_dir)}",
        f"    lang: {json.dumps(e['hreflang'])}",
        f"    locale: {json.dumps(e['locale'])}",
        f"    default: {json.dumps(DEFAULT_LANG)}",
        f"    fnTitle: {json.dumps(e['fnTitle'], ensure_ascii=False)}",
        f"    homeTitle: {json.dumps(e['homeTitle'], ensure_ascii=False)}",
        f"    tail: {json.dumps(e['tail'], ensure_ascii=False)}",
        f"    keywords: {json.dumps(KEYWORDS[lang_dir], ensure_ascii=False)}",
        "    langs:",
    ]
    lines += [
        f"      - {{dir: {json.dumps(d)}, hreflang: {json.dumps(SEO[d]['hreflang'])}}}"
        for d, _, _, _ in LANGS
    ]
    return "\n".join(lines)


def nav_yaml(lang_dir: str) -> str:
    """2단 nav: 대분류(Excel 호환/Google/EGTools 전용) → 소그룹(버전/기능) → 페이지.
    대분류 제목은 <lang>-strings.json(secExcel/secGoogle/secEg), 소그룹은 categories.json labels."""
    strings = json.loads(
        (DOCS / "_content" / f"{lang_dir}-strings.json").read_text(encoding="utf-8")
    )
    sec_title = {"excel": strings["secExcel"], "google": strings["secGoogle"], "eg": strings["secEg"]}
    labels = CATS["labels"][lang_dir]
    by_cat = {}
    for fn, cat in CATS["functions"].items():
        by_cat.setdefault(cat, []).append(fn)
    lines = ["nav:", "  - README.md"]
    for sec in CATS["sections"]:
        lines.append(f'  - "{sec_title[sec["key"]]}":')
        single = len(sec["slugs"]) == 1
        for slug in sec["slugs"]:
            fns = sorted(by_cat.get(slug, []))
            if not fns:
                continue
            if single:  # 단일 소그룹(Google)은 페이지를 대분류 바로 아래에
                lines.extend(f'    - "functions/{slug}/{fn}.md"' for fn in fns)
            else:
                lines.append(f'    - "{labels[slug]}":')
                lines.extend(f'      - "functions/{slug}/{fn}.md"' for fn in fns)
    ribbon_pages = sorted(
        p.name for p in (DOCS / lang_dir / "ribbon").glob("*.md") if p.name != "README.md"
    )
    lines.append(f'  - "{labels["ribbon"]}":')
    lines.append('    - "ribbon/README.md"')
    lines.extend(f'    - "ribbon/{p}"' for p in ribbon_pages)
    return "\n".join(lines)


def yaml_config(lang_dir: str, mat_lang: str, site_name: str) -> str:
    # 전환기(alternate)는 모든 언어 공통 목록
    alternates = "\n".join(
        f"    - name: {label}\n      link: {BASE_PATH}/{d}/\n      lang: {ml}"
        for d, ml, label, _ in LANGS
    )
    docs_dir = (DOCS / lang_dir).as_posix()
    site_dir = (SITE / lang_dir).as_posix()
    # toc slugify를 pymdownx.slugs로 바꿔 GitHub식 유니코드 앵커(#보이는-셀만-복사)를 유지한다.
    return f"""\
site_name: {site_name}
site_description: {DESCRIPTIONS[lang_dir]}
site_url: {SITE_URL}/{lang_dir}/
repo_url: https://github.com/EGTools/EGToolsPP
repo_name: EGTools/EGToolsPP
edit_uri: edit/main/docs/{lang_dir}/
docs_dir: {docs_dir}
site_dir: {site_dir}
theme:
  name: material
  language: {mat_lang}
  features:
    - navigation.top
    - search.highlight
    - content.code.copy
    - content.action.edit
  palette:
    - media: "(prefers-color-scheme: light)"
      scheme: default
      toggle:
        icon: material/brightness-7
        name: Dark mode
    - media: "(prefers-color-scheme: dark)"
      scheme: slate
      toggle:
        icon: material/brightness-4
        name: Light mode
markdown_extensions:
  - toc:
      permalink: true
      slugify: !!python/object/apply:pymdownx.slugs.slugify
        kwds:
          case: lower
hooks:
  - {HOOK.as_posix()}
extra:
  homepage: https://egtools.kr/
  seo:
{seo_block(lang_dir)}
  alternate:
{alternates}
{nav_yaml(lang_dir)}
"""


def redirect_stub(target: str) -> str:
    return f"""\
<!doctype html>
<html>
<head>
<meta charset="utf-8">
<link rel="canonical" href="{target}">
<meta http-equiv="refresh" content="0; url={target}">
<script>location.replace("{target}");</script>
</head>
<body><a href="{target}">Moved</a></body>
</html>
"""


def write_redirects(lang_dir: str) -> None:
    """평면 함수 URL /<lang>/<FN>/ -> /<lang>/<cat>/<FN>/ 스텁 생성."""
    for fn, cat in CATS["functions"].items():
        stub_dir = SITE / lang_dir / fn
        stub_dir.mkdir(parents=True, exist_ok=False)
        (stub_dir / "index.html").write_text(
            redirect_stub(f"{BASE_PATH}/{lang_dir}/functions/{cat}/{fn}/"), encoding="utf-8"
        )


def landing_page() -> str:
    links = "\n".join(
        f'      <li><a href="{BASE_PATH}/{d}/">{label}</a></li>'
        for d, _, label, _ in LANGS
    )
    # 언어 선택 루트는 자체가 x-default — 각 언어판을 hreflang 으로 짝지어 준다.
    alts = "\n".join(
        f'<link rel="alternate" href="{SITE_URL}/{d}/" hreflang="{SEO[d]["hreflang"]}">'
        for d, _, _, _ in LANGS
    )
    # navigator.language → docs 폴더 매핑 (기본 en)
    return f"""\
<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{SEO["en"]["homeTitle"]}</title>
<meta name="description" content="{DESCRIPTIONS["en"]}">
<meta name="keywords" content="{KEYWORDS["en"]}">
<link rel="canonical" href="{SITE_URL}/">
{alts}
<link rel="alternate" href="{SITE_URL}/" hreflang="x-default">
<style>
  body {{ font-family: system-ui, sans-serif; max-width: 32rem; margin: 15vh auto; padding: 0 1rem; }}
  h1 {{ font-size: 1.4rem; }}
  ul {{ line-height: 2; }}
</style>
<script>
  (function () {{
    var map = {{ ko: "kr", ja: "ja", es: "es" }};
    var l = (navigator.language || "en").toLowerCase();
    var target;
    if (l.indexOf("zh") === 0) {{
      target = (l.indexOf("tw") >= 0 || l.indexOf("hant") >= 0 || l.indexOf("hk") >= 0) ? "zh-TW" : "zh-CN";
    }} else {{
      target = map[l.split("-")[0]] || "en";
    }}
    location.replace("{BASE_PATH}/" + target + "/");
  }})();
</script>
</head>
<body>
  <h1>EGTools++ Manual</h1>
  <p>Select your language:</p>
  <ul>
{links}
  </ul>
</body>
</html>
"""


def sitemap_index() -> str:
    """언어별 MkDocs sitemap을 묶는 루트 sitemap index — Search Console에 한 번만 제출."""
    entries = "\n".join(
        f"  <sitemap><loc>{SITE_URL}/{d}/sitemap.xml</loc></sitemap>" for d, _, _, _ in LANGS
    )
    return f"""\
<?xml version="1.0" encoding="UTF-8"?>
<sitemapindex xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">
{entries}
</sitemapindex>
"""


def robots_txt() -> str:
    return f"User-agent: *\nAllow: /\n\nSitemap: {SITE_URL}/sitemap.xml\n"


def main() -> int:
    if SITE.exists():
        shutil.rmtree(SITE)
    SITE.mkdir()

    with tempfile.TemporaryDirectory() as tmp:
        for lang_dir, mat_lang, _, site_name in LANGS:
            cfg = Path(tmp) / f"mkdocs-{lang_dir}.yml"
            cfg.write_text(yaml_config(lang_dir, mat_lang, site_name), encoding="utf-8")
            print(f"== building {lang_dir} ==", flush=True)
            r = subprocess.run(
                [sys.executable, "-m", "mkdocs", "build", "-f", str(cfg)],
                cwd=ROOT,
            )
            if r.returncode != 0:
                print(f"build failed: {lang_dir}", file=sys.stderr)
                return r.returncode
            write_redirects(lang_dir)

    (SITE / "index.html").write_text(landing_page(), encoding="utf-8")
    (SITE / "sitemap.xml").write_text(sitemap_index(), encoding="utf-8")
    (SITE / "robots.txt").write_text(robots_txt(), encoding="utf-8")
    # docs/_root/* 는 사이트 루트에 그대로 복사 (Search Console 확인 파일 등)
    root_extra = DOCS / "_root"
    if root_extra.is_dir():
        for f in root_extra.iterdir():
            if f.is_file():
                shutil.copy(f, SITE / f.name)
    print(f"done -> {SITE}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
