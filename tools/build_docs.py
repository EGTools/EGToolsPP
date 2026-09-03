#!/usr/bin/env python3
"""Build the multilingual manual site for GitHub Pages.

manual/<lang>/ 마크다운(SoT, gen_manual.ps1 산출물)을 언어별 MkDocs Material
사이트로 빌드해 _site/<lang>/ 에 놓고, 루트에 언어 감지 index.html 을 만든다.
언어 추가 시 LANGS 한 줄만 추가하면 된다(폴더 규약은 gen_manual.ps1 참조).

사용:  pip install mkdocs-material
       python tools/build_docs.py
"""
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SITE = ROOT / "_site"
BASE_PATH = "/EGToolsPP"  # GitHub project-pages base path
SITE_URL = "https://egtools.github.io" + BASE_PATH

# (manual 폴더명, Material theme.language, 전환기 표기, site_name)
LANGS = [
    ("kr", "ko", "한국어", "EGTools++ 매뉴얼"),
    ("en", "en", "English", "EGTools++ Manual"),
    ("ja", "ja", "日本語", "EGTools++ マニュアル"),
    ("es", "es", "Español", "Manual de EGTools++"),
    ("zh-CN", "zh", "简体中文", "EGTools++ 手册"),
    ("zh-TW", "zh-TW", "繁體中文", "EGTools++ 手冊"),
]


def yaml_config(lang_dir: str, mat_lang: str, site_name: str) -> str:
    # 전환기(alternate)는 모든 언어 공통 목록
    alternates = "\n".join(
        f"    - name: {label}\n      link: {BASE_PATH}/{d}/\n      lang: {ml}"
        for d, ml, label, _ in LANGS
    )
    docs_dir = (ROOT / "manual" / lang_dir).as_posix()
    site_dir = (SITE / lang_dir).as_posix()
    # toc slugify를 pymdownx.slugs로 바꿔 GitHub식 유니코드 앵커(#보이는-셀만-복사)를 유지한다.
    return f"""\
site_name: {site_name}
site_url: {SITE_URL}/{lang_dir}/
repo_url: https://github.com/EGTools/EGToolsPP
repo_name: EGTools/EGToolsPP
edit_uri: edit/main/manual/{lang_dir}/
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
"""


def landing_page() -> str:
    links = "\n".join(
        f'      <li><a href="{BASE_PATH}/{d}/">{label}</a></li>'
        for d, _, label, _ in LANGS
    )
    # navigator.language → manual 폴더 매핑 (기본 en)
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

    (SITE / "index.html").write_text(landing_page(), encoding="utf-8")
    print(f"done -> {SITE}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
