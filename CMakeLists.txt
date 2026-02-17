cmake_minimum_required(VERSION 3.20)
project(naiv4vibe_thumbnail_provider LANGUAGES CXX)

include(FetchContent)

if(NOT WIN32)
  message(FATAL_ERROR "This project only supports Windows.")
endif()

add_library(naiv4vibe_thumbnail_provider SHARED
  src/Naiv4VibeThumbnailProvider.def
  src/dllmain.cpp
  src/ThumbnailProvider.cpp
)

target_include_directories(naiv4vibe_thumbnail_provider PRIVATE src)
target_compile_features(naiv4vibe_thumbnail_provider PRIVATE cxx_std_20)
target_compile_definitions(naiv4vibe_thumbnail_provider PRIVATE UNICODE _UNICODE NOMINMAX)

target_link_libraries(naiv4vibe_thumbnail_provider PRIVATE
  ole32
  shlwapi
  windowscodecs
  crypt32
)

set_target_properties(naiv4vibe_thumbnail_provider PROPERTIES
  OUTPUT_NAME "Naiv4VibeThumbnailProvider"
)


find_package(nlohmann_json CONFIG QUIET)

if(NOT nlohmann_json_FOUND)
  message(STATUS "nlohmann_json not found locally, falling back to FetchContent")
  FetchContent_Declare(
    nlohmann_json
    URL
      https://github.com/nlohmann/json/releases/download/v3.11.3/json.tar.xz
      https://codeload.github.com/nlohmann/json/tar.gz/refs/tags/v3.11.3
  )
  FetchContent_MakeAvailable(nlohmann_json)
endif()

target_link_libraries(naiv4vibe_thumbnail_provider PRIVATE nlohmann_json::nlohmann_json)
