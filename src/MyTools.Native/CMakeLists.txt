cmake_minimum_required(VERSION 3.24)

project(MyToolsNative LANGUAGES C CXX)

string(REPLACE "/" "\\\\" APP_ICON_RC_PATH
    "${CMAKE_CURRENT_SOURCE_DIR}/../MyTools/Resources/AppIcon.ico")
configure_file(
    resources/app.rc.in
    ${CMAKE_CURRENT_BINARY_DIR}/generated/app.rc
    @ONLY
)

add_executable(MyToolsNative WIN32
    app/main.cpp
    app/app_context.cpp
    app/crash_handler.cpp
    app/single_instance.cpp
    app/tray_host.cpp
    ccore/json/mt_json_schema.c
    modules/codex_profiles/codex_profile_module.cpp
    modules/module_registry.cpp
    services/config_store.cpp
    services/codex_profile_backup_service.cpp
    services/codex_profile_box_service.cpp
    services/codex_profile_diff_service.cpp
    services/codex_profile_edit_service.cpp
    services/codex_profile_export_service.cpp
    services/codex_profile_import_service.cpp
    services/codex_profile_switch_service.cpp
    services/file_scanner.cpp
    services/file_system.cpp
    services/hotkey_service.cpp
    services/logger.cpp
    services/network_service.cpp
    services/process_runner.cpp
    services/secret_store_dpapi.cpp
    services/task_runner.cpp
    ui/main_window.cpp
    ui/modal_dialogs.cpp
    ui/renderer_d2d.cpp
    ${CMAKE_CURRENT_BINARY_DIR}/generated/app.rc
)

target_compile_features(MyToolsNative PRIVATE cxx_std_20)

set_property(TARGET MyToolsNative PROPERTY C_STANDARD 11)
set_property(TARGET MyToolsNative PROPERTY C_STANDARD_REQUIRED ON)

target_include_directories(MyToolsNative PRIVATE
    ${CMAKE_CURRENT_SOURCE_DIR}
    ${CMAKE_CURRENT_SOURCE_DIR}/ccore/include
)

target_compile_definitions(MyToolsNative PRIVATE
    UNICODE
    _UNICODE
    NOMINMAX
    WIN32_LEAN_AND_MEAN
    WINVER=0x0A00
    _WIN32_WINNT=0x0A00
)

if(MSVC)
    set_property(TARGET MyToolsNative PROPERTY
        MSVC_RUNTIME_LIBRARY "MultiThreaded$<$<CONFIG:Debug>:Debug>")

    target_compile_options(MyToolsNative PRIVATE
        /W4
        /utf-8
        /MP
        $<$<COMPILE_LANGUAGE:CXX>:/permissive->
        $<$<COMPILE_LANGUAGE:CXX>:/EHsc>
    )
endif()

target_link_libraries(MyToolsNative PRIVATE
    advapi32
    bcrypt
    comdlg32
    crypt32
    d2d1
    dwrite
    gdi32
    ole32
    shell32
    user32
    ws2_32
)
