# EGTools++

**구형 Excel(2010+)에서 최신 Excel 함수를 쓰게 해 주는 단일 `.xll` 추가기능** (Windows 전용)

Microsoft 365에만 있는 `XLOOKUP`, `FILTER`, `TEXTSPLIT`, `GROUPBY`, `REGEX*` 같은 함수들을
Excel 2010~2021 어디서든 동일하게 사용할 수 있습니다. 설치 파일·런타임·관리자 권한이
필요 없는 **자체 완결형 단일 파일**입니다.

| 산출물 | 대상 |
|---|---|
| `EGTools++64.xll` | 64비트 Excel |
| `EGTools++32.xll` | 32비트 Excel |

최신 버전은 [Releases](https://github.com/EGTools/EGToolsPP/releases)에서 받을 수 있습니다.

## 주요 기능

- **함수 128개** — 최신 Excel 호환 69개: 조회(`XLOOKUP`/`XMATCH`), 동적배열(`FILTER`/`SORT`/`UNIQUE`/`SEQUENCE`),
  배열 조작(`VSTACK`/`TAKE`/`WRAPROWS` 등), 텍스트(`TEXTSPLIT`/`TEXTBEFORE`/`TEXTJOIN`),
  집계(`GROUPBY`/`PIVOTBY`/`MAXIFS`), 정규식(`REGEXTEST`/`REGEXEXTRACT`/`REGEXREPLACE`),
  웹(`WEBSERVICE`/`FILTERXML`), 논리·정보·수학·비트 함수와 `LET`, `IMAGE`까지 +
  Google Sheets 호환 7개(`QUERY`/`GOOGLETRANSLATE`/`IMPORT*`) + EGTools 고유 함수 52개
- **버전 적응 이름 등록** — 호스트에 내장 함수가 있으면 `EG.` 접두사로, 없으면 원래 이름
  그대로 등록되어 구버전에서 드롭인(drop-in)으로 동작
- **동적배열 Spill 호환** — 최신 Excel에서는 네이티브 spill, 구버전에서는 자동 CSE 배열
  확장으로 동일한 결과
- **셀 내 IntelliSense** — 수식 입력 중 함수·인수 도움말 툴팁 표시
- **리본 `EGTools` 탭** — 보이는 셀만 복사/붙여넣기, 셀 병합/나누기, 그림 삽입·셀 맞춤·일괄
  추출, 달력/일정표 생성(공휴일 표시), 라벨·메일머지(SMTP 메일 발송), 스타일/이름/오류
  정리 등 실무 유틸리티와 셀 컨텍스트 메뉴·날짜 선택기
- **호환 변환** — 통합 문서의 최신 함수 수식을 호환 함수명으로 일괄 변환/복원
  (변환 전 자동 백업)
- **다국어 지원** — 한국어·English·日本語·简体中文·繁體中文·Español UI 자동 감지

## 매뉴얼

함수별 사용법·예제와 리본 기능 설명을 언어별로 제공합니다:
[한국어](manual/kr/README.md) · [English](manual/en/README.md) · [日本語](manual/ja/README.md) ·
[Español](manual/es/README.md) · [简体中文](manual/zh-CN/README.md) · [繁體中文](manual/zh-TW/README.md)

## 커뮤니티

- **한국어 커뮤니티 — 네이버 카페**: <https://cafe.naver.com/egtools>
  사용법 질문, 활용 팁 공유, 버그 제보 등 한국어 소통 공간입니다.
- **GitHub [Discussions](https://github.com/EGTools/EGToolsPP/discussions)**: 질문·아이디어·피드백 (영어/기타 언어 포함)
- 버그 신고는 [Issues](https://github.com/EGTools/EGToolsPP/issues)를 이용해 주세요.

## 설치

관리자 권한이 필요 없습니다. Excel 내장 대화상자로 등록합니다:

1. [Releases](https://github.com/EGTools/EGToolsPP/releases)에서 Excel 비트수에 맞는 `.xll` 다운로드
2. Excel → **파일 → 옵션 → 추가 기능**
3. 관리: **Excel 추가 기능** 선택 → **이동...**
4. **찾아보기...** → `EGTools++64.xll`(또는 `32`) 선택 → **확인**

제거는 같은 대화상자에서 체크를 해제하면 됩니다.

## 빌드

요구 사항: **Visual Studio 2022** (C++ 데스크톱 워크로드: MSVC v143, C++ ATL, CMake) + 번들 vcpkg

```powershell
# Developer PowerShell for VS 2022 에서
pwsh tools/build_xloil.ps1 -Platform x64   # 의존 정적 라이브러리 빌드 (최초 1회, x86은 -Platform Win32)
cmake --preset x64
cmake --build --preset x64-release         # 산출물: dist/EGTools++64.xll
```

- 정적 링크(`/MT`)라 VC++ 재배포 패키지가 필요 없습니다.
- `tools/build_xloil.ps1`은 개발 빌드 전용 스크립트로, 의존 소스를 받아 빌드하는 특성상
  일부 백신이 오탐할 수 있습니다(개발 PC에서만 예외 등록 권장). 배포 `.xll`과는 무관합니다.

## 프로젝트 구조

```
src/            추가기능 소스 (core / functions / intellisense / ribbon)
manual/         다국어 매뉴얼 (함수·리본 도움말, 6개 언어)
patches/        xlOil 정적 링크·호환성 패치
cmake/          빌드 스크립트 모듈
tools/          개발 빌드 도구
resources/      리소스·아이콘·다국어 카탈로그·엔트리포인트 정의
```

## 기반 및 라이선스

- **EGTools++는 [Apache License 2.0](LICENSE)으로 배포됩니다.**
- 베이스 프레임워크: [xlOil](https://gitlab.com/stevecu/xloil) (Apache-2.0) — 정적 링크 및
  구버전 호환 패치 적용(`patches/` 참조)
- 실행 파일 압축: UPX
