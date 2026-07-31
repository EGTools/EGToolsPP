# EGTools++

구형 Excel(2010+)에서 최신 배열·조회·텍스트 함수와 업무 리본을 제공하는
**단일 self-contained `.xll` 추가기능** (Windows 전용). 다국어(ko·en·zh-CN·zh-TW·ja·es) 기본 지원.

> 산출물: `EGTools++64.xll` (x64) / `EGTools++32.xll` (x86) — 무런타임 단독 실행.

## 현재 상태

**Phase 0 — 기술 스파이크 / 타당성 검증** 진행 중.
설계·결정·로드맵의 단일 진실 소스(SoT)는 [`plan/`](plan/) 폴더입니다.

| 문서 | 내용 |
|------|------|
| [plan/00_README_계획개요.md](plan/00_README_계획개요.md) | 개요·결정 로그 D1~D8 |
| [plan/04_아키텍처설계.md](plan/04_아키텍처설계.md) | 구조·빌드·Spill/IntelliSense/i18n |
| [plan/06_로드맵_및_리스크.md](plan/06_로드맵_및_리스크.md) | Phase 0~6, 리스크 |
| [plan/07_개발착수_핸드오프.md](plan/07_개발착수_핸드오프.md) | 착수 가이드 |

## 빌드 (개요)

- Visual Studio 2022 (C++ 데스크톱: MSVC v143, C++ ATL, CMake), 번들 vcpkg
- 정적 링크(`/MT`) + xlOil 정적 → 단일 `.xll` (재배포 런타임 불필요)

```powershell
# Developer PowerShell for VS 2022 에서
pwsh tools/build_xloil.ps1 -Platform x64   # xlOil 정적 라이브러리(최초 1회)
cmake --preset x64
cmake --build --preset x64-release
```

> ℹ️ `tools/build_xloil.ps1`은 **개발 빌드 전용**(배포물 아님)입니다. 인터넷에서 의존성을
> 받아(git) 곧바로 컴파일하므로 일부 백신이 "다운로드→빌드" 연쇄를 행위 휴리스틱으로
> 오탐할 수 있습니다. **개발 PC에서만** 해당 스크립트/`tools` 폴더를 백신 예외로 등록하세요.
> (배포되는 `.xll`과는 무관합니다.)

## 설치 (현재 사용자, 관리자 불필요)

Excel 내장 대화상자로 등록합니다 (AV-안전):

> 파일 → 옵션 → 추가 기능 → 관리: **Excel 추가 기능** → **이동** → **찾아보기**
> → `dist/EGTools++64.xll` 선택 → 확인

> ⚠️ Office 자동시작 레지스트리(`Excel\Options\OPEN`)를 스크립트로 직접 쓰는 방식은
> 백신이 persistence 패턴으로 오탐할 수 있어 사용하지 않습니다. 배포 시 미서명 경고는
> 코드 서명으로 해결합니다.

## 라이선스

- 베이스 프레임워크: xlOil (Apache-2.0) — Phase 0 S2(x86) 검증 조건부.
  폴백 시 Keith Lewis `xll`(MIT) 또는 Microsoft Excel XLL SDK.
