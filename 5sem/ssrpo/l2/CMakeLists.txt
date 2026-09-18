cmake_minimum_required(VERSION 3.29)
project(l2)

set(CMAKE_CXX_STANDARD 14)

include_directories(include)
include_directories(include/hw)

add_executable(l2
        include/hw/l2_ApplicationLayer.hpp
        include/hw/l3_DomainLayer.hpp
        include/hw/l4_InfrastructureLayer.hpp
        src/l1_UserInterface.cpp
        src/l2_ApplicationLayer.cpp
        src/l3_DomainLayer.cpp
        src/l4_InfrastructureLayer.cpp
        configure
)
