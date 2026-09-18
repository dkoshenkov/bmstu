cmake_minimum_required(VERSION 3.16.3)
project(l4)

set(CMAKE_CXX_STANDARD 14)

include_directories(src/include)
include_directories(src/include/hw)
SET(CMAKE_CXX_FLAGS "-pthread -fPIC -pedantic-errors -Wall -Wextra -Wpedantic -Wcast-align -Wcast-qual -Wctor-dtor-privacy -Wduplicated-branches -Wduplicated-cond -Wextra-semi -Wfloat-equal -Wlogical-op -Wnon-virtual-dtor -Wold-style-cast -Woverloaded-virtual -Wredundant-decls -Wsign-promo")

SET(CMAKE_CXX_FLAGS_RELEASE        "-O3")
SET(CMAKE_CXX_FLAGS_RELWITHDEBINFO "-O0 -g --coverage")
SET(CMAKE_CXX_FLAGS_DEBUG          "-O0 -g")

add_subdirectory(src/stressgen)
add_subdirectory(src/lab)
add_subdirectory(src/tp)