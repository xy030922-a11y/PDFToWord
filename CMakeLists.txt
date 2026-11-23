cmake_minimum_required(VERSION 3.21)
project(PdfToWord LANGUAGES CXX)

if(NOT WIN32 OR NOT MSVC)
    message(FATAL_ERROR "Build PdfToWord on Windows with MSVC and the Windows 10 or 11 SDK.")
endif()

set(CMAKE_CXX_STANDARD 20)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

find_package(Qt6 6.5 REQUIRED COMPONENTS Widgets Concurrent)
qt_add_executable(PdfToWord WIN32
    src/main.cpp
    src/converter.cpp
    src/converter.h
)
target_compile_options(PdfToWord PRIVATE /utf-8)
target_compile_definitions(PdfToWord PRIVATE NOMINMAX WIN32_LEAN_AND_MEAN)
target_link_libraries(PdfToWord PRIVATE
    Qt6::Widgets
    Qt6::Concurrent
    ole32
    oleaut32
    runtimeobject
    urlmon
)
