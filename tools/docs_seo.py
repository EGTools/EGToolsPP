#!/usr/bin/env python3
"""MkDocs 훅: 다국어 매뉴얼의 페이지별 검색엔진 메타데이터.

tools/build_docs.py 가 생성하는 언어별 mkdocs 설정에서 hooks 로 로드되고,
값은 모두 그 설정의 extra.seo 에서 받는다(언어 목록·문구의 단일 출처는 build_docs.py).

Material 기본값으로 부족한 것만 보완한다:
  1) description — 기본은 site_description 하나가 모든 페이지에 붙어 검색 스니펫이
     전부 같아진다. 페이지 첫 단락(함수 설명)으로 페이지별 description 을 만든다.
  2) hreflang — Material 의 extra.alternate 는 언어 '루트'(/kr/)만 가리켜서 번역 짝을
     알려주지 못한다(6개 언어가 서로 중복 경쟁). 같은 경로의 6개 언어 절대 URL +
     x-default 로 교체한다. 세 언어 모두 페이지 집합이 동일해야 유효 — gen_manual.ps1
     이 6개 언어를 함께 생성하므로 보장된다.
  3) Open Graph / 트위터 카드 — 링크 공유 시 제목·설명 스니펫.
  4) JSON-LD — 함수/리본 페이지 TechArticle, 언어별 홈 SoftwareApplication.
  5) <title> 에 '함수'(언어별 표기)를 넣어 "XLOOKUP 함수" 류 질의에 맞춘다. h1 은 그대로.

meta name="keywords" 는 Google/Bing 은 무시하지만(공식 입장) 브랜드어 "EGTools" 를
모든 페이지에 공통으로 남겨 두는 값이라 유지한다(사용자 요청, 2026-09-16).
"""
import html
import json
import re

DESC_MAX = 160

# Material 이 extra.alternate 로 찍는 언어 루트 hreflang 링크(rel=prev/next 는 건드리지 않음).
_ALT_RE = re.compile(r'[ \t]*<link\s+rel="alternate"[^>]*hreflang="[^"]*"[^>]*>\n?')
_TITLE_RE = re.compile(r"<title>.*?</title>", re.S)
_H1_RE = re.compile(r"^#\s+(.+)$", re.M)
# 함수 페이지 머리의 분류 줄: "**종류**: … · **네이티브 도입**: …"
_CATLINE_RE = re.compile(r"\*\*[^*\n]+\*\*\s*[:：]")


def _plain(md: str) -> str:
    """검색 스니펫용 평문 — 주석/링크/강조 표기 제거 후 공백 정규화."""
    md = re.sub(r"<!--.*?-->", "", md, flags=re.S)
    md = re.sub(r"\[([^\]]+)\]\([^)]*\)", r"\1", md)
    md = md.replace("**", "").replace("`", "").replace("*", "")
    return " ".join(md.split())


def _first_paragraph(markdown: str) -> str:
    """제목/분류 줄('**종류**: ...')/표/코드/목록을 건너뛴 첫 설명 단락.

    홈(README)의 소개 문단은 '**EGTools++**는 …' 처럼 굵게 시작하므로 ** 만으로 걸러내면
    안 된다 — '**라벨**: 값' 꼴(함수 페이지 분류 줄)만 건너뛴다.
    """
    for block in re.split(r"\n[ \t]*\n", markdown):
        b = block.strip()
        # '- '/'* ' 로 목록만 걸러낸다 — '**EGTools++**는 …' 을 목록으로 오인하면 안 된다.
        if not b or b.startswith(("#", "|", "```", ">", "- ", "* ", "<!--", "!!!", "=== ")):
            continue
        if _CATLINE_RE.match(b):
            continue
        return _plain(b)
    return ""


def _clip(text: str, limit: int = DESC_MAX) -> str:
    if len(text) <= limit:
        return text
    cut = text[:limit]
    sp = cut.rfind(" ")
    if sp > limit * 0.6:  # CJK 는 공백이 거의 없어 단어 경계를 못 찾는다 — 그때는 그냥 자른다
        cut = cut[:sp]
    return cut.rstrip(" ,.;·、，") + "…"


def _is_function(page) -> bool:
    return page.file.src_uri.startswith("functions/")


def on_page_markdown(markdown, page, config, files):
    """페이지별 description 을 page.meta 에 심는다(Material 이 렌더 시 사용).

    page.title 은 render() 에서야 채워지므로 여기서는 h1 을 직접 읽는다.
    """
    seo = config["extra"]["seo"]
    meta = page.meta if page.meta is not None else {}
    if not meta.get("description"):
        m = _H1_RE.search(markdown)
        name = _plain(m.group(1)) if m else page.file.name
        body = _first_paragraph(markdown)
        # 홈은 h1 이 사이트명과 겹치므로 소개 문단만 쓴다.
        home = page.file.src_uri in ("README.md", "index.md")
        desc = body if (home and body) else (f"{name} — {body}" if body else name)
        tail = seo["tail"]
        if body and len(desc) + 1 + len(tail) <= DESC_MAX:
            desc = f"{desc} {tail}"
        meta["description"] = _clip(desc)
        page.meta = meta
    return markdown


def _head_title(page, config, seo) -> str:
    if page.is_homepage:
        return seo["homeTitle"]
    title = str(page.title)
    if _is_function(page):
        title = seo["fnTitle"].replace("{FN}", title)
    return f"{title} - {config['site_name']}"


def _hreflangs(page, seo):
    """같은 문서의 언어별 절대 URL — 경로 집합이 6개 언어 동일하므로 언어 조각만 바꾼다."""
    url = page.url  # 홈은 "", 나머지는 "functions/excel-2021/XLOOKUP/" 꼴
    pairs = [(f"{seo['root']}/{e['dir']}/{url}", e["hreflang"]) for e in seo["langs"]]
    pairs.append((f"{seo['root']}/{seo['default']}/{url}", "x-default"))
    return pairs


def _keywords(page, seo) -> str:
    """공통 브랜드 키워드 + 함수 페이지는 함수명/EG.함수명(최신 Excel 에서의 실제 등록명)."""
    common = seo["keywords"]
    if not _is_function(page):
        return common
    fn = str(page.title)
    return f"{fn}, {seo['fnTitle'].replace('{FN}', fn)}, EG.{fn}, {common}"


def _jsonld(page, config, seo) -> dict:
    app = {
        "@type": "SoftwareApplication",
        "name": "EGTools++",
        "applicationCategory": "BusinessApplication",
        "operatingSystem": "Windows",
        "softwareRequirements": "Microsoft Excel 2010 or later",
        "url": "https://github.com/EGTools/EGToolsPP",
        "downloadUrl": "https://github.com/EGTools/EGToolsPP/releases/latest",
        "offers": {"@type": "Offer", "price": "0", "priceCurrency": "USD"},
    }
    desc = (page.meta or {}).get("description", config["site_description"])
    if page.is_homepage:
        return {
            "@context": "https://schema.org",
            **app,
            "description": desc,
            "inLanguage": seo["lang"],
            "url": page.canonical_url,
            "sameAs": "https://github.com/EGTools/EGToolsPP",
        }
    return {
        "@context": "https://schema.org",
        "@type": "TechArticle",
        "headline": str(page.title),
        "description": desc,
        "url": page.canonical_url,
        "inLanguage": seo["lang"],
        "isPartOf": {
            "@type": "WebSite",
            "name": config["site_name"],
            "url": f"{seo['root']}/{seo['dir']}/",
        },
        "keywords": _keywords(page, seo),
        "about": app,
    }


def _head_extra(page, config, seo) -> str:
    esc = lambda s: html.escape(str(s), quote=True)
    desc = (page.meta or {}).get("description", config["site_description"])
    title = _head_title(page, config, seo)
    lines = [f'<link rel="alternate" href="{esc(u)}" hreflang="{esc(h)}">' for u, h in _hreflangs(page, seo)]
    lines += [
        f'<meta name="keywords" content="{esc(_keywords(page, seo))}">',
        '<meta property="og:type" content="article">',
        f'<meta property="og:title" content="{esc(title)}">',
        f'<meta property="og:description" content="{esc(desc)}">',
        f'<meta property="og:url" content="{esc(page.canonical_url)}">',
        f'<meta property="og:site_name" content="{esc(config["site_name"])}">',
        f'<meta property="og:locale" content="{esc(seo["locale"])}">',
        '<meta name="twitter:card" content="summary">',
        f'<meta name="twitter:title" content="{esc(title)}">',
        f'<meta name="twitter:description" content="{esc(desc)}">',
        '<script type="application/ld+json">'
        + json.dumps(_jsonld(page, config, seo), ensure_ascii=False, separators=(",", ":"))
        + "</script>",
    ]
    return "".join(f"    {ln}\n" for ln in lines)


def on_post_page(output, page, config):
    seo = config["extra"]["seo"]
    out, n = _ALT_RE.subn("", output)
    if not n:
        raise RuntimeError(
            "Material 의 언어 루트 hreflang 링크를 못 찾았다 — 테마 마크업이 바뀐 듯하다. "
            "그대로 두면 페이지별 hreflang 과 충돌하므로 tools/docs_seo.py _ALT_RE 를 고칠 것."
        )
    if "</head>" not in out:
        raise RuntimeError(f"</head> 없음: {page.file.src_uri}")
    out = out.replace("</head>", _head_extra(page, config, seo) + "  </head>", 1)
    title = _head_title(page, config, seo)
    out, n = _TITLE_RE.subn(lambda _m: f"<title>{html.escape(title)}</title>", out, count=1)
    if not n:
        raise RuntimeError(f"<title> 없음: {page.file.src_uri}")
    return out
