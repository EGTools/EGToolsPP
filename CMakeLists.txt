cmake_minimum_required(VERSION 3.21)

project(EGToolsPP
  VERSION 0.8
  DESCRIPTION "Modern Excel functions for legacy Excel as a single static .xll add-in"
  LANGUAGES CXX)

# --- Standard ---------------------------------------------------------------
# C++17 to match xlOil (its headers/libs are built as stdcpp17). Revisit C++20
# for our own modules once the static-link baseline is proven (Phase 1).
set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_CXX_EXTENSIONS OFF)

list(APPEND CMAKE_MODULE_PATH "${CMAKE_SOURCE_DIR}/cmake")
include(xlOil)

if(NOT MSVC)
  message(FATAL_ERROR "EGTools++ is a Windows XLL add-in and requires MSVC (Visual Studio 2022).")
endif()

# --- Bitness-suffixed artifact name (EGTools++64 / EGTools++32) -------------
if(CMAKE_SIZEOF_VOID_P EQUAL 8)
  set(EG_BITS 64)
else()
  set(EG_BITS 32)
endif()
message(STATUS "EGTools++ target bitness: ${EG_BITS}-bit")

# --- Sources ----------------------------------------------------------------
# Phase 0 scaffolding: only AddinMain for now. Module dirs (core/spill/
# intellisense/functions/ribbon) are wired in as they are implemented.
set(EG_SOURCES
  src/AddinMain.cpp
  src/core/I18n.cpp
  src/core/Version.cpp
  src/core/Spill.cpp
  src/core/Registry.cpp
  src/functions/Spike.cpp
  src/functions/FxLookup.cpp
  src/functions/FxLogic.cpp
  src/functions/FxText.cpp
  src/functions/FxWeb.cpp
  src/functions/FxDate.cpp
  src/functions/FxInfo.cpp
  src/functions/FxAgg.cpp
  src/functions/FxGen.cpp
  src/functions/FxFilterSort.cpp
  src/functions/FxArrayShape.cpp
  src/functions/FxRegex.cpp
  src/functions/FxGroup.cpp
  src/functions/FxMath.cpp
  src/functions/FxBit.cpp
  src/functions/FxLet.cpp
  src/functions/FxImage.cpp
  src/functions/FxGoogle.cpp
  src/functions/FxEgText.cpp
  src/functions/FxEgData.cpp
  src/functions/FxEgQC.cpp
  src/functions/FxEgDate.cpp
  src/functions/FxEgJson.cpp
  src/functions/FxEgUtil.cpp
  src/functions/FxBarcode.cpp
  src/functions/FxColor.cpp
  src/functions/FxPublicApi.cpp
  src/intellisense/IntelliSense.cpp
  src/intellisense/FormulaParser.cpp
  src/intellisense/Catalog.cpp
  src/intellisense/ToolTipWindow.cpp
  src/ribbon/Ribbon.cpp
  src/ribbon/Convert.cpp
  src/ribbon/commands/Dialogs.cpp
  src/ribbon/commands/CmdCopy.cpp
  src/ribbon/commands/CmdMerge.cpp
  src/ribbon/commands/CmdPicture.cpp
  src/ribbon/commands/CmdUtil.cpp
  src/ribbon/commands/CmdCalendar.cpp
  src/ribbon/commands/CmdForm.cpp
  src/ribbon/commands/CmdApiKeys.cpp
  src/ribbon/commands/CmdSmtp.cpp
  src/ribbon/commands/DatePicker.cpp
  resources/egtools.rc
  # Module-definition file: exports the XLL entry points (xlAutoOpen, …) with
  # UNDECORATED names. Required on x86 — without it __stdcall decoration
  # (_xlAutoOpen@0) hides the entry points from Excel and RegisterXLL fails.
  # CMake passes a .def listed as a source to the MSVC linker as /DEF.
  resources/EGToolsPP.def)

add_library(EGToolsPP SHARED ${EG_SOURCES})

set_target_properties(EGToolsPP PROPERTIES
  OUTPUT_NAME "EGTools++${EG_BITS}"
  PREFIX ""
  SUFFIX ".xll")

target_include_directories(EGToolsPP PRIVATE src resources)

# CRT: static (/MT) for a self-contained, runtime-free .xll (no VC++ redist).
# Requires xlOil to be built with the static CRT too — tools/build_xloil.ps1
# writes external/xloil/Directory.Build.props to force that (plan O1). Both sides
# must match or LNK2038 (RuntimeLibrary mismatch) results.
set_property(TARGET EGToolsPP PROPERTY
  MSVC_RUNTIME_LIBRARY "MultiThreaded$<$<CONFIG:Debug>:Debug>")

target_compile_definitions(EGToolsPP PRIVATE
  WIN32_LEAN_AND_MEAN
  NOMINMAX
  UNICODE _UNICODE
  # 버전 major.minor — build 번호는 gen/BuildNum.h(매 빌드 자동 증가)가 제공하고,
  # EG_VERSION_STR("major.minor.build")도 그 헤더가 조합한다. 리본 About(C++)과
  # VERSIONINFO 리소스(RC, resources/egtools.rc)가 함께 사용.
  EG_VER_MAJOR=${PROJECT_VERSION_MAJOR}
  EG_VER_MINOR=${PROJECT_VERSION_MINOR}
  # No spdlog: logging was never initialised anywhere (no file/sink setup), so
  # the XLO_* macros only dragged dead spdlog code into the .xll (-133KB; the
  # bundled fmt that Throw.h really uses stays). Must match the xlOil lib build
  # (patch: xloil-0.22.1-no-spdlog + Directory.Build.props, see plan/17).
  XLOIL_NO_SPDLOG)

# UTF-8 source/exec charset — sources and i18n strings are UTF-8 (avoids C4819
# on non-UTF-8 system code pages such as cp949). C++ only — the RC compiler
# does not accept /utf-8.
target_compile_options(EGToolsPP PRIVATE $<$<COMPILE_LANGUAGE:CXX>:/utf-8>)

# Release: optimize for size (/O1), not speed — UDF hot paths are dominated by
# Excel's own marshalling, while binary size is a distribution goal (single
# self-contained .xll). /Gw puts data in COMDATs so /OPT:REF can drop unused
# globals; /Zc:inline drops unreferenced inline functions at compile time.
# 실측(2026-08-01, 100k행 TEXTNUMSORT/MVLOOKUP×2/REGEXTEST 20k/UNIQUE 풀재계산):
# /O2는 1319ms→1286ms(~2.5%, 편차 수준)에 크기 +52KB(+7%) — /O1 유지 결정.
target_compile_options(EGToolsPP PRIVATE
  $<$<AND:$<COMPILE_LANGUAGE:CXX>,$<CONFIG:Release>>:/O1 /Ob1 /Gw /Zc:inline>)
target_link_options(EGToolsPP PRIVATE
  $<$<CONFIG:Release>:/OPT:REF /OPT:ICF>)

# --- 빌드 번호 (major.minor.build) ------------------------------------------
# build_number.txt(저장소 루트)가 현재 빌드 번호를 보관한다. 없으면 master
# 커밋 수로 초기화. 매 빌드 전에 cmake/bump_build.cmake가 각 빌드 트리의
# 스탬프(build_stamp.txt)와 비교해 "같은 아키텍처의 새 라운드"일 때만 +1 —
# 그래서 이어서 빌드하는 x64/x86 쌍은 같은 번호를 갖는다(상세는 스크립트 주석).
set(EG_BUILD_NUM_FILE ${CMAKE_SOURCE_DIR}/build_number.txt)
set(EG_GEN_DIR ${CMAKE_BINARY_DIR}/gen)
if(NOT EXISTS ${EG_BUILD_NUM_FILE})
  execute_process(COMMAND git rev-list --count master
                  WORKING_DIRECTORY ${CMAKE_SOURCE_DIR}
                  OUTPUT_VARIABLE EG_GIT_COUNT
                  OUTPUT_STRIP_TRAILING_WHITESPACE ERROR_QUIET)
  if(NOT EG_GIT_COUNT MATCHES "^[0-9]+$")
    set(EG_GIT_COUNT 0)
  endif()
  file(WRITE ${EG_BUILD_NUM_FILE} "${EG_GIT_COUNT}")
endif()
# 구성 시점에 초기 헤더 생성(스탬프 없음 = 소비 없음) — 첫 컴파일 전에 필요.
execute_process(COMMAND ${CMAKE_COMMAND}
  -DTXT=${EG_BUILD_NUM_FILE} -DHDR=${EG_GEN_DIR}/BuildNum.h
  -P ${CMAKE_SOURCE_DIR}/cmake/bump_build.cmake)

add_custom_target(eg_bump
  COMMAND ${CMAKE_COMMAND}
    -DTXT=${EG_BUILD_NUM_FILE} -DHDR=${EG_GEN_DIR}/BuildNum.h
    -DSTAMP=${CMAKE_BINARY_DIR}/build_stamp.txt
    -P ${CMAKE_SOURCE_DIR}/cmake/bump_build.cmake
  BYPRODUCTS ${EG_GEN_DIR}/BuildNum.h
  COMMENT "Bump build number (build_number.txt)")
add_dependencies(EGToolsPP eg_bump)
target_include_directories(EGToolsPP PRIVATE ${EG_GEN_DIR})

# Link xlOil static libraries (see cmake/xlOil.cmake).
eg_link_xloil(EGToolsPP)

# JSON parsing for the embedded i18n catalogs (header-only).
find_package(nlohmann_json CONFIG REQUIRED)
target_link_libraries(EGToolsPP PRIVATE nlohmann_json::nlohmann_json)

# Barcode encode/decode (BARCODE/QRCODE/…/READBARCODE) — static zxing-cpp.
find_package(ZXing CONFIG REQUIRED)
target_link_libraries(EGToolsPP PRIVATE ZXing::ZXing)

# Console self-test that round-trips the barcode core without Excel
# (encodes sample BMPs which tests/ then decodes via READBARCODE).
# Not part of ALL and never shipped.
add_executable(eg_barcode_selftest EXCLUDE_FROM_ALL tools/barcode_selftest.cpp)
target_include_directories(eg_barcode_selftest PRIVATE src)
target_link_libraries(eg_barcode_selftest PRIVATE ZXing::ZXing)
set_property(TARGET eg_barcode_selftest PROPERTY
  MSVC_RUNTIME_LIBRARY "MultiThreaded$<$<CONFIG:Debug>:Debug>")

# Numeric-aggregation core self-test (AggCore.h, no Excel/xlOil) — results are
# diffed against native Excel formulas by tests. Never shipped.
add_executable(eg_agg_selftest EXCLUDE_FROM_ALL tools/aggregate_selftest.cpp)
target_include_directories(eg_agg_selftest PRIVATE src)
set_property(TARGET eg_agg_selftest PROPERTY
  MSVC_RUNTIME_LIBRARY "MultiThreaded$<$<CONFIG:Debug>:Debug>")

# --- Place artifact into dist/ (Release: UPX-packed by default) --------------
# Deployment default = UPX-compressed dist (~-65%, decided 2026-07-31 after the
# AV/function verification in plan/16 §4A; revisit only if a problem shows up).
# The build-dir .xll stays uncompressed for debugging/map analysis.
# tools/upx.exe is vendored in the repo so both multiboot OSes (plan/12) pack
# identically without a per-OS install; PATH is the fallback.
find_program(EG_UPX upx HINTS "${CMAKE_SOURCE_DIR}/tools")
add_custom_command(TARGET EGToolsPP POST_BUILD
  COMMAND ${CMAKE_COMMAND}
          -DCFG=$<CONFIG> "-DUPX=${EG_UPX}"
          "-DSRC=$<TARGET_FILE:EGToolsPP>"
          "-DDST=${CMAKE_SOURCE_DIR}/dist/$<TARGET_FILE_NAME:EGToolsPP>"
          -P "${CMAKE_SOURCE_DIR}/cmake/pack_dist.cmake"
  COMMENT "Packing $<TARGET_FILE_NAME:EGToolsPP> into dist/ (UPX on Release)")
