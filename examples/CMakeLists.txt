file(GLOB Demo_cppfiles "Demo*.cpp")

foreach(Demo_file ${Demo_cppfiles})
  #WE for without extension
  get_filename_component(Demo_exe "${Demo_file}" NAME_WE)
  add_executable(${Demo_exe} ${Demo_file})
  target_link_libraries(${Demo_exe} PRIVATE OpenXLSX-static)
  target_compile_definitions(${Demo_exe} PRIVATE OPENXLSX_STATIC_DEFINE)
endforeach()

