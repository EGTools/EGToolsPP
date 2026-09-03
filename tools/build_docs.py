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
BASE_PATH = "/EGToolsPP"  # GitHub project-pages base path
SITE_URL = "https://egtools.github.io" + BASE_PATH

# (docs 폴더명, Material theme.language, 전환기 표기, site_name)
LANGS = [
    ("kr", "ko", "한국어", "EGTools++ 매뉴얼"),
    ("en", "en", "English", "EGTools++ Manual"),
    ("ja", "ja", "日本語", "EGTools++ マニュアル"),
    ("es", "es", "Español", "Manual de EGTools++"),
    ("zh-CN", "zh", "简体中文", "EGTools++ 手册"),
    ("zh-TW", "zh-TW", "繁體中文", "EGTools++ 手冊"),
]

CATS = json.loads((DOCS / "_content" / "categories.json").read_text(encoding="utf-8"))

# 평면 리다이렉트 폴더(/<lang>/<FN>/)와 최상위 섹션 폴더가 Windows 대소문자
# 비구분에서 충돌하면 안 된다 (functions/ 도입 전 IMAGE vs image 사례, 2026-09-03).
_clash = [fn for fn in CATS["functions"] if fn.lower() in {"functions", "ribbon"}]
assert not _clash, f"top-level folder clashes with function name(s): {_clash}"


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
extra:
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
    # navigator.language → docs 폴더 매핑 (기본 en)
    return f"""\
<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>EGTools++ Manual</title>
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
