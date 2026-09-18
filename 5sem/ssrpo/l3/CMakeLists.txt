cmake_minimum_required(VERSION 3.16.3)
project(l3)

set(CMAKE_CXX_STANDARD 14)

include_directories(src/include)
include_directories(src/include/hw)

add_executable(l3
        src/include/hw/l2_ApplicationLayer.hpp
        src/include/hw/l3_DomainLayer.hpp
        src/include/hw/l4_InfrastructureLayer.hpp
        src/lab/l1_UserInterface.cpp
        src/lab/l2_ApplicationLayer.cpp
        src/lab/l3_DomainLayer.cpp
        src/lab/l4_InfrastructureLayer.cpp
        configure
)

SET(CMAKE_CXX_FLAGS_RELEASE        "-O3")
SET(CMAKE_CXX_FLAGS_RELWITHDEBINFO "-O0 -g --coverage")
SET(CMAKE_CXX_FLAGS_DEBUG          "-O0 -g")

add_subdirectory(src/stressgen)
add_subdirectory(src/lab)

