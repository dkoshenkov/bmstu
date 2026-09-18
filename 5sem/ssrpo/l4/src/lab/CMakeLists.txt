cmake_minimum_required(VERSION 3.16.3 FATAL_ERROR)

project(lab)

add_executable(${PROJECT_NAME}
        l1_UserInterface.cpp
        l2_ApplicationLayer.cpp
        l3_DomainLayer.cpp
        l4_InfrastructureLayer.cpp
)

set_target_properties(${PROJECT_NAME} PROPERTIES
        CXX_STANDARD 17
        CXX_STANDARD_REQUIRED YES
        CXX_EXTENSIONS YES
)
target_include_directories(${PROJECT_NAME} PUBLIC ${CMAKE_SOURCE_DIR}/src/include)
target_link_libraries(${PROJECT_NAME} tp)

set_target_properties(${PROJECT_NAME} PROPERTIES RUNTIME_OUTPUT_DIRECTORY ${CMAKE_SOURCE_DIR}/bin)

