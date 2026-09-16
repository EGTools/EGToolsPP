#!/usr/bin/env python3
"""Build the egtools.kr product home (landing page) for GitHub Pages.

egtools.kr 은 제품 홈, manual.egtools.kr 은 함수 레퍼런스로 역할을 나눈다. 리다이렉트
표지판이던 egtools.kr 에 실제 색인 대상 콘텐츠를 두어 브랜드·일반 질의를 받게 한다.

문구는 docs/_content/landing.json, 언어/SEO 표와 함수 분류는 tools/build_docs.py ·
docs/_content/categories.json 이 단일 출처 — 함수 개수도 여기서 계산하므로 안 늙는다.

출력은 _landing/ 이고, 그 내용을 EGTools/egtools.github.io 저장소 main 에 올리면
Pages 가 egtools.kr 로 서빙한다(CNAME 포함).

사용:  python tools/build_landing.py
"""
import collections
import json
import shutil
import sys
from pathlib import Path  # noqa: F401  (build_docs 의 ROOT 와 함께 쓰임)

from build_docs import CATS, DEFAULT_LANG, KEYWORDS, LANGS, ROOT, SEO
from build_docs import SITE_URL as MANUAL_URL

OUT = ROOT / "_landing"
HOME_URL = "https://egtools.kr"
REPO_URL = "https://github.com/EGTools/EGToolsPP"
RELEASES_URL = f"{REPO_URL}/releases/latest"
CAFE_URL = "https://cafe.naver.com/egtools"

COPY = json.loads((ROOT / "docs" / "_content" / "landing.json").read_text(encoding="utf-8"))

# 홈에 내세우는 함수 — 매뉴얼의 실제 분류 경로로 링크한다(평면 리다이렉트를 안 거치게).
FEATURED = [
    "XLOOKUP", "XMATCH", "FILTER", "SORT", "SORTBY", "UNIQUE", "SEQUENCE",
    "TEXTSPLIT", "TEXTBEFORE", "TEXTAFTER", "TEXTJOIN", "GROUPBY", "PIVOTBY", "MAXIFS",
    "REGEXTEST", "REGEXEXTRACT", "REGEXREPLACE", "VSTACK", "HSTACK", "TAKE", "DROP",
    "WRAPROWS", "CHOOSECOLS", "LET", "IFS", "SWITCH", "IMAGE", "WEBSERVICE", "FILTERXML",
    "QUERY", "GOOGLETRANSLATE", "IMPORTRANGE", "QRCODE", "BARCODE", "ISVISIBLE",
    "SUMIFCOLOR", "UNPIVOT", "COMPARELIST", "TOLUNAR", "KOREANHOLIDAYS", "CPK",
]
_missing = [fn for fn in FEATURED if fn not in CATS["functions"]]
assert not _missing, f"FEATURED 에 없는 함수: {_missing}"


def counts() -> dict:
    """categories.json 에서 함수 개수 — 함수를 추가해도 홈이 자동으로 맞는다."""
    sec_of = {slug: s["key"] for s in CATS["sections"] for slug in s["slugs"]}
    per = collections.Counter(sec_of[cat] for cat in CATS["functions"].values())
    return {
        "{TOTAL}": str(len(CATS["functions"])),
        "{EXCEL}": str(per["excel"]),
        "{GOOGLE}": str(per["google"]),
        "{EG}": str(per["eg"]),
    }


N = counts()


def sub(text: str) -> str:
    for k, v in N.items():
        text = text.replace(k, v)
    return text


CSS = """
:root {
  --bg: #ffffff; --bg-alt: #f6f7f9; --fg: #1b1f24; --fg-dim: #58616b;
  --line: #e2e5ea; --accent: #1f6feb; --accent-fg: #ffffff; --code-bg: #eef1f5;
}
@media (prefers-color-scheme: dark) {
  :root {
    --bg: #14181d; --bg-alt: #1b2026; --fg: #e6e9ee; --fg-dim: #9aa4b0;
    --line: #2b323a; --accent: #5b9cff; --accent-fg: #0d1117; --code-bg: #232a32;
  }
}
* { box-sizing: border-box; }
body {
  margin: 0; background: var(--bg); color: var(--fg);
  font: 16px/1.65 system-ui, -apple-system, "Segoe UI", "Malgun Gothic", "Hiragino Sans",
        "Microsoft YaHei", sans-serif;
  -webkit-text-size-adjust: 100%;
}
a { color: var(--accent); }
code {
  background: var(--code-bg); border-radius: 4px; padding: 0.1em 0.35em;
  font: 0.88em/1.4 ui-monospace, "Cascadia Mono", Consolas, monospace;
}
.wrap { max-width: 56rem; margin: 0 auto; padding: 0 1.25rem; }
header { border-bottom: 1px solid var(--line); }
header .wrap {
  display: flex; flex-wrap: wrap; gap: 0.5rem 1rem;
  align-items: center; justify-content: space-between; padding-block: 0.9rem;
}
.brand { font-weight: 700; font-size: 1.15rem; letter-spacing: -0.01em; }
.brand span { color: var(--accent); }
header nav { display: flex; flex-wrap: wrap; gap: 0.9rem; font-size: 0.9rem; }
.hero { padding-block: 3.5rem 2.5rem; }
h1 { font-size: 2.6rem; line-height: 1.15; margin: 0 0 0.4rem; letter-spacing: -0.02em; }
.tag { font-size: 1.3rem; font-weight: 600; color: var(--fg); margin: 0 0 1rem; }
.lead { font-size: 1.05rem; color: var(--fg-dim); max-width: 44rem; margin: 0 0 1.75rem; }
.cta { display: flex; flex-wrap: wrap; gap: 0.75rem; }
.btn {
  display: inline-block; padding: 0.6rem 1.15rem; border-radius: 7px;
  border: 1px solid var(--line); text-decoration: none; font-weight: 600;
}
.btn.primary { background: var(--accent); color: var(--accent-fg); border-color: var(--accent); }
.note { font-size: 0.85rem; color: var(--fg-dim); margin: 0.85rem 0 0; }
.stats {
  list-style: none; display: flex; flex-wrap: wrap; gap: 1.5rem 2.5rem;
  margin: 2.5rem 0 0; padding: 1.5rem 0 0; border-top: 1px solid var(--line);
}
.stats b { display: block; font-size: 1.9rem; line-height: 1.1; letter-spacing: -0.02em; }
.stats span { font-size: 0.85rem; color: var(--fg-dim); }
section { padding-block: 2.5rem; border-top: 1px solid var(--line); }
section.hero { border-top: 0; }
h2 { font-size: 1.5rem; margin: 0 0 1.25rem; letter-spacing: -0.01em; }
.grid { display: grid; gap: 1rem; grid-template-columns: repeat(auto-fit, minmax(16rem, 1fr)); }
.card { background: var(--bg-alt); border: 1px solid var(--line); border-radius: 9px; padding: 1.1rem 1.15rem; }
.card h3 { margin: 0 0 0.4rem; font-size: 1rem; }
.card p { margin: 0; font-size: 0.92rem; color: var(--fg-dim); }
.chips { display: flex; flex-wrap: wrap; gap: 0.45rem; margin: 0 0 1.25rem; }
.chips a {
  border: 1px solid var(--line); background: var(--bg-alt); border-radius: 6px;
  padding: 0.25rem 0.55rem; text-decoration: none; color: var(--fg);
  font: 0.86rem/1.5 ui-monospace, "Cascadia Mono", Consolas, monospace;
}
.chips a:hover { border-color: var(--accent); color: var(--accent); }
ol.steps { margin: 0; padding-left: 1.4rem; }
ol.steps li { margin-bottom: 0.5rem; }
ul.links { list-style: none; margin: 0; padding: 0; }
ul.links li { padding: 0.45rem 0; border-bottom: 1px solid var(--line); }
footer { border-top: 1px solid var(--line); padding-block: 2rem 3rem; color: var(--fg-dim); font-size: 0.85rem; }
footer .langs { display: flex; flex-wrap: wrap; gap: 0.9rem; margin-bottom: 1rem; }
@media (max-width: 34rem) {
  h1 { font-size: 2.05rem; }
  .tag { font-size: 1.12rem; }
  .hero { padding-block: 2.25rem 1.75rem; }
}
"""


def head(lang_dir: str, url: str, title: str, desc: str) -> str:
    """언어별 페이지 공통 head — description·keywords·hreflang·OG·JSON-LD."""
    e = SEO[lang_dir]
    alts = "".join(
        f'<link rel="alternate" href="{HOME_URL}/{d}/" hreflang="{SEO[d]["hreflang"]}">\n'
        for d, _, _, _ in LANGS
    )
    ld = {
        "@context": "https://schema.org",
        "@type": "SoftwareApplication",
        "name": "EGTools++",
        "description": desc,
        "applicationCategory": "BusinessApplication",
        "operatingSystem": "Windows",
        "softwareRequirements": "Microsoft Excel 2010 or later",
        "url": url,
        "downloadUrl": RELEASES_URL,
        "license": "https://www.apache.org/licenses/LICENSE-2.0",
        "inLanguage": e["hreflang"],
        "offers": {"@type": "Offer", "price": "0", "priceCurrency": "USD"},
        "sameAs": REPO_URL,
    }
    return f"""<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{title}</title>
<meta name="description" content="{desc}">
<meta name="keywords" content="{KEYWORDS[lang_dir]}">
<link rel="canonical" href="{url}">
{alts}<link rel="alternate" href="{HOME_URL}/{DEFAULT_LANG}/" hreflang="x-default">
<meta property="og:type" content="website">
<meta property="og:title" content="{title}">
<meta property="og:description" content="{desc}">
<meta property="og:url" content="{url}">
<meta property="og:site_name" content="EGTools++">
<meta property="og:locale" content="{e["locale"]}">
<meta name="twitter:card" content="summary">
<meta name="twitter:title" content="{title}">
<meta name="twitter:description" content="{desc}">
<script type="application/ld+json">{json.dumps(ld, ensure_ascii=False, separators=(",", ":"))}</script>
<style>{CSS}</style>"""


def lang_page(lang_dir: str) -> str:
    c = COPY[lang_dir]
    html_lang = SEO[lang_dir]["hreflang"]
    url = f"{HOME_URL}/{lang_dir}/"
    manual = f"{MANUAL_URL}/{lang_dir}/"
    title = sub(c["title"])
    desc = sub(c["description"])

    stats = "".join(
        f"<li><b>{n}</b><span>{lbl}</span></li>"
        for n, lbl in zip([N["{TOTAL}"], N["{EXCEL}"], N["{GOOGLE}"], N["{EG}"]], c["stats"])
    )
    cards = "".join(
        f'<div class="card"><h3>{sub(t)}</h3><p>{sub(b)}</p></div>' for t, b in c["features"]
    )
    chips = "".join(
        f'<a href="{manual}functions/{CATS["functions"][fn]}/{fn}/">{fn}</a>' for fn in FEATURED
    )
    steps = "".join(f"<li>{s}</li>" for s in c["installSteps"])
    links = "".join(
        f'<li><a href="{href}">{text}</a></li>'
        for href, text in [
            (manual, c["linkManual"]),
            (CAFE_URL, c["linkCafe"]),
            (REPO_URL, c["linkGithub"]),
            (f"{REPO_URL}/discussions", c["linkDiscuss"]),
            (f"{REPO_URL}/issues", c["linkIssues"]),
        ]
    )
    others = "".join(
        f'<a href="{HOME_URL}/{d}/" hreflang="{SEO[d]["hreflang"]}">{lb}</a>'
        for d, _, lb, _ in LANGS
        if d != lang_dir
    )
    return f"""<!doctype html>
<html lang="{html_lang}">
<head>
{head(lang_dir, url, title, desc)}
</head>
<body>
<header><div class="wrap">
  <div class="brand">EGTools<span>++</span></div>
  <nav>
    <a href="#features">{sub(c["featuresTitle"])}</a>
    <a href="#install">{c["installTitle"]}</a>
    <a href="{manual}">{c["ctaManual"]}</a>
    <a href="{REPO_URL}">GitHub</a>
  </nav>
</div></header>
<main class="wrap">
  <section class="hero">
    <h1>EGTools++</h1>
    <p class="tag">{c["hero"]}</p>
    <p class="lead">{c["lead"]}</p>
    <div class="cta">
      <a class="btn primary" href="{RELEASES_URL}">{c["ctaDownload"]}</a>
      <a class="btn" href="{manual}">{c["ctaManual"]}</a>
    </div>
    <p class="note">{c["ctaNote"]}</p>
    <ul class="stats">{stats}</ul>
  </section>
  <section id="features">
    <h2>{sub(c["featuresTitle"])}</h2>
    <div class="grid">{cards}</div>
  </section>
  <section id="functions">
    <h2>{c["fnsTitle"]}</h2>
    <p class="note" style="margin:0 0 1rem">{c["fnsNote"]}</p>
    <div class="chips">{chips}</div>
    <p><a href="{manual}">{sub(c["fnsMore"])} &rarr;</a></p>
  </section>
  <section id="install">
    <h2>{c["installTitle"]}</h2>
    <p class="note" style="margin:0 0 1rem">{c["installNote"]}</p>
    <ol class="steps">{steps}</ol>
  </section>
  <section id="links">
    <h2>{c["linksTitle"]}</h2>
    <ul class="links">{links}</ul>
  </section>
</main>
<footer><div class="wrap">
  <div class="langs">{others}</div>
  <p style="margin:0">{c["footer"]}</p>
</div></footer>
</body>
</html>
"""


def detector_page() -> str:
    """루트(egtools.kr/) — 브라우저 언어로 보내고, 스스로 x-default 가 된다."""
    links = "".join(
        f'      <li><a href="/{d}/" hreflang="{SEO[d]["hreflang"]}">{label}</a></li>\n'
        for d, _, label, _ in LANGS
    )
    alts = "".join(
        f'<link rel="alternate" href="{HOME_URL}/{d}/" hreflang="{SEO[d]["hreflang"]}">\n'
        for d, _, _, _ in LANGS
    )
    desc = sub(COPY["en"]["description"])
    return f"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{sub(COPY["en"]["title"])}</title>
<meta name="description" content="{desc}">
<meta name="keywords" content="{KEYWORDS["en"]}">
<link rel="canonical" href="{HOME_URL}/">
{alts}<link rel="alternate" href="{HOME_URL}/" hreflang="x-default">
<style>{CSS}</style>
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
    location.replace("/" + target + "/");
  }})();
</script>
</head>
<body>
<main class="wrap">
  <section class="hero">
    <h1>EGTools++</h1>
    <p class="tag">{COPY["en"]["hero"]}</p>
    <ul class="links" style="max-width:20rem">
{links}    </ul>
  </section>
</main>
</body>
</html>
"""


def sitemap() -> str:
    """언어판마다 xhtml:link 로 서로를 가리키는 sitemap — hreflang 을 이중으로 알려 준다."""
    alts = "".join(
        f'    <xhtml:link rel="alternate" hreflang="{SEO[d]["hreflang"]}" href="{HOME_URL}/{d}/"/>\n'
        for d, _, _, _ in LANGS
    ) + f'    <xhtml:link rel="alternate" hreflang="x-default" href="{HOME_URL}/"/>\n'
    urls = f"  <url>\n    <loc>{HOME_URL}/</loc>\n{alts}  </url>\n"
    for d, _, _, _ in LANGS:
        urls += f"  <url>\n    <loc>{HOME_URL}/{d}/</loc>\n{alts}  </url>\n"
    return (
        '<?xml version="1.0" encoding="UTF-8"?>\n'
        '<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9"\n'
        '        xmlns:xhtml="http://www.w3.org/1999/xhtml">\n'
        f"{urls}</urlset>\n"
    )


def main() -> int:
    if OUT.exists():
        shutil.rmtree(OUT)
    OUT.mkdir()
    (OUT / "CNAME").write_text("egtools.kr\n", encoding="utf-8")
    (OUT / "index.html").write_text(detector_page(), encoding="utf-8")
    for lang_dir, _, _, _ in LANGS:
        d = OUT / lang_dir
        d.mkdir()
        (d / "index.html").write_text(lang_page(lang_dir), encoding="utf-8")
    (OUT / "sitemap.xml").write_text(sitemap(), encoding="utf-8")
    (OUT / "robots.txt").write_text(
        f"User-agent: *\nAllow: /\n\nSitemap: {HOME_URL}/sitemap.xml\n", encoding="utf-8"
    )
    print(f"done -> {OUT}  ({len(LANGS)} langs, {N['{TOTAL}']} functions)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
