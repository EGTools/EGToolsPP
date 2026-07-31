# pack_dist.cmake — POST_BUILD helper: place the .xll into dist/.
#
# Release + upx available  -> dist gets the UPX-packed .xll (deployment default,
#                             decided 2026-07-31; see plan/16 §4A).
# otherwise                -> plain copy, with a warning when upx is missing so
#                             a machine without UPX doesn't silently ship an
#                             uncompressed artifact (install: winget install UPX.UPX).
#
# Inputs (-D): CFG=$<CONFIG>  UPX=<upx path or "">  SRC=<built .xll>  DST=<dist .xll>
#
# The build-dir binary (SRC) is never packed — keep it raw for debugging and
# linker-map analysis; only the dist copy is compressed.

get_filename_component(_dist_dir "${DST}" DIRECTORY)
file(MAKE_DIRECTORY "${_dist_dir}")

set(_packed FALSE)
if(CFG STREQUAL "Release")
  if(UPX)
    file(REMOVE "${DST}")
    execute_process(
      COMMAND "${UPX}" --best -q -q "${SRC}" -o "${DST}"
      RESULT_VARIABLE _upx_rc)
    if(_upx_rc EQUAL 0)
      set(_packed TRUE)
      message(STATUS "dist: UPX-packed ${DST}")
    else()
      message(WARNING "upx failed (rc=${_upx_rc}) — falling back to plain copy")
    endif()
  else()
    message(WARNING
      "upx not found — dist gets an UNCOMPRESSED .xll. The deployment default "
      "is packed (plan/16 §4A): winget install UPX.UPX")
  endif()
endif()

if(NOT _packed)
  execute_process(COMMAND "${CMAKE_COMMAND}" -E copy_if_different "${SRC}" "${DST}")
endif()
