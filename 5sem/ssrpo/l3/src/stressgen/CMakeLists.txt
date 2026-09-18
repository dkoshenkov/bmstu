cmake_minimum_required(VERSION 3.16.3 FATAL_ERROR)

project(stressgen)

add_executable(${PROJECT_NAME}
        stressgen.cpp
    )

set_target_properties(${PROJECT_NAME} PROPERTIES
    CXX_STANDARD 17
    CXX_STANDARD_REQUIRED YES
    CXX_EXTENSIONS YES
)

set_target_properties(${PROJECT_NAME} PROPERTIES RUNTIME_OUTPUT_DIRECTORY ${CMAKE_SOURCE_DIR}/bin)

