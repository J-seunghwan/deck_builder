cmake_minimum_required(VERSION 3.19)
project(Deck_Builder LANGUAGES CXX)

find_package(Qt6 6.5 REQUIRED COMPONENTS Core Widgets LinguistTools Sql Network)

qt_standard_project_setup()

set(PROJECT_VERSION "1.1.1")
add_definitions(-DAPP_VERSION="${PROJECT_VERSION}")

qt_add_executable(Deck_Builder
    WIN32 MACOSX_BUNDLE
    main.cpp
    mainwindow.cpp
    mainwindow.h
    mainwindow.ui
    character_slot_widget.ui
    characterslotwidget.h characterslotwidget.cpp
    updatemanager.h updatemanager.cpp
)

# Updater executable
add_executable(updater WIN32 updater_main.cpp)
target_link_libraries(updater PRIVATE Qt::Core)

qt_add_translations(
    TARGETS Deck_Builder
    TS_FILES Deck_Builder_ko_KR.ts
)

target_link_libraries(Deck_Builder
    PRIVATE
        Qt::Core
        Qt::Widgets
        Qt6::Sql
        Qt6::Network
)

include(GNUInstallDirs)

install(TARGETS Deck_Builder
    BUNDLE  DESTINATION .
    RUNTIME DESTINATION ${CMAKE_INSTALL_BINDIR}
    LIBRARY DESTINATION ${CMAKE_INSTALL_LIBDIR}
)

qt_generate_deploy_app_script(
    TARGET Deck_Builder
    OUTPUT_SCRIPT deploy_script
    NO_UNSUPPORTED_PLATFORM_ERROR
)
install(SCRIPT ${deploy_script})
