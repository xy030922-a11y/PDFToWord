cmake_minimum_required(VERSION 3.21)
project(PdfToWord LANGUAGES CXX)

if(NOT WIN32 OR NOT MSVC)
    message(FATAL_ERROR "Build PdfToWord on Windows with MSVC and the Windows 10 or 11 SDK.")
endif()

# app.rc embeds app.ico so Windows Explorer shows the application icon.
enable_language(RC)

set(CMAKE_CXX_STANDARD 20)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

# PDFium extracts text directly from the PDF, so editable output no longer needs
# Microsoft Word. Point PDFIUM_PACKAGE_DIR at another copy of the prebuilt
# package if third_party/pdfium-win-x64 is not used (see README.md).
set(PDFIUM_PACKAGE_DIR "${CMAKE_CURRENT_SOURCE_DIR}/third_party/pdfium-win-x64" CACHE PATH
    "Directory containing PDFiumConfig.cmake, include/, lib/ and bin/")
find_package(PDFium REQUIRED CONFIG PATHS "${PDFIUM_PACKAGE_DIR}" NO_DEFAULT_PATH)

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
    pdfium
    delayimp
    ole32
    oleaut32
    runtimeobject
    shcore
    urlmon
)
# The DLL is only needed by "editable text" mode. Delay-loading it keeps the
# default image mode usable even if pdfium.dll was not deployed.
target_link_options(PdfToWord PRIVATE "/DELAYLOAD:pdfium.dll")
add_custom_command(TARGET PdfToWord POST_BUILD
    COMMAND ${CMAKE_COMMAND} -E copy_if_different
        "$<TARGET_FILE:pdfium>" "$<TARGET_FILE_DIR:PdfToWord>/pdfium.dll"
    COMMENT "Copying pdfium.dll next to PdfToWord.exe"
    VERBATIM
)
