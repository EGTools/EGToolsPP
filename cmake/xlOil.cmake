# xlOil.cmake — link a target against prebuilt xlOil static libraries.
#
# xlOil itself is an MSBuild project (no CMake). We build its static libs once
# (msbuild xlOil.sln /p:Configuration=ReleaseStatic /p:Platform=<x64|Win32>),
# then this module wires the produced .lib files + headers + defines into our
# CMake target, replicating libs/xlOilStaticLib.props.
#
# Prerequisite: run tools/build_xloil.ps1 for the matching platform first.

function(eg_link_xloil target)
  set(XLOIL_ROOT "${CMAKE_SOURCE_DIR}/external/xloil")
  if(NOT EXISTS "${XLOIL_ROOT}/include/xlOil/xlOil.h")
    message(FATAL_ERROR
      "xlOil not found at ${XLOIL_ROOT}. Clone it and run tools/build_xloil.ps1.")
  endif()

  # MSVC/xlOil platform folder + Excel SDK lib subdir per bitness.
  if(CMAKE_SIZEOF_VOID_P EQUAL 8)
    set(XLOIL_PLAT "x64")
    set(XLOIL_SDK_LIB "${XLOIL_ROOT}/external/Excel2013SDK/LIB/x64")
  else()
    set(XLOIL_PLAT "Win32")
    set(XLOIL_SDK_LIB "${XLOIL_ROOT}/external/Excel2013SDK/LIB/WIn32")
  endif()

  set(XLOIL_STATIC_DIR "${XLOIL_ROOT}/build/${XLOIL_PLAT}/ReleaseStatic")
  set(XLOIL_RELEASE_DIR "${XLOIL_ROOT}/build/${XLOIL_PLAT}/Release")

  if(NOT EXISTS "${XLOIL_STATIC_DIR}/xlOil-XLL.lib")
    message(FATAL_ERROR
      "xlOil static libs missing for ${XLOIL_PLAT} (${XLOIL_STATIC_DIR}). "
      "Build them: tools/build_xloil.ps1 -Platform ${XLOIL_PLAT}")
  endif()

  # spdlog include stays even under XLOIL_NO_SPDLOG: xlOil's Throw.h/Events.h
  # use the fmt library BUNDLED inside spdlog for exception-message formatting
  # (that part is genuinely used; only the logging machinery is stripped).
  target_include_directories(${target} PRIVATE
    "${XLOIL_ROOT}/include"
    "${XLOIL_ROOT}/src"
    "${XLOIL_ROOT}/external/boost-1.67"
    "${XLOIL_ROOT}/external/spdlog/include"
    "${XLOIL_ROOT}/external/tomlplusplus/include"
    "${XLOIL_ROOT}/external")

  # SPDLOG_COMPILED_LIB kept for the bundled fmt (see include note above);
  # XLOIL_NO_SPDLOG (CMakeLists.txt) still strips all logging call sites.
  target_compile_definitions(${target} PRIVATE
    XLOIL_STATIC_LIB
    SPDLOG_COMPILED_LIB
    _SILENCE_CXX17_CODECVT_HEADER_DEPRECATION_WARNING
    _SILENCE_STDEXT_ARR_ITERS_DEPRECATION_WARNING)

  # xlOil headers are C++17.
  target_compile_features(${target} PRIVATE cxx_std_17)

  # Mirror TestAddin's link set: xlOil core (ReleaseStatic) + externals (Release)
  # + Excel C API + Active Accessibility (used by xlOil IntelliSense/COM).
  # spdlog.lib stays linked but ONLY its bundled-fmt objects get pulled in:
  # XLOIL_NO_SPDLOG removes every logging reference, so the linker drops the
  # spdlog logging machinery (sinks/registry/logger) automatically.
  target_link_libraries(${target} PRIVATE
    "${XLOIL_STATIC_DIR}/xlOil-XLL.lib"
    "${XLOIL_STATIC_DIR}/xlOil-COM.lib"
    "${XLOIL_STATIC_DIR}/xlOil-Dynamic.lib"
    "${XLOIL_RELEASE_DIR}/xlOilHelpers.lib"
    "${XLOIL_RELEASE_DIR}/asmjit.lib"
    "${XLOIL_RELEASE_DIR}/spdlog.lib"
    "${XLOIL_RELEASE_DIR}/rdcfswatcher.lib"
    "${XLOIL_SDK_LIB}/xlcall32.lib"
    Oleacc.lib)
endfunction()
