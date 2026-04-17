cmake_minimum_required(VERSION 3.20)
project(native_drag_probe LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

add_executable(native_drag_probe main.cpp)
target_compile_definitions(native_drag_probe PRIVATE UNICODE _UNICODE NOMINMAX WIN32_LEAN_AND_MEAN)
