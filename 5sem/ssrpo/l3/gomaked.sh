#!/bin/bash
#
cmake -Bbuild/debug -H. -DCMAKE_BUILD_TYPE=Debug
cmake --build build/debug --config Debug -j 6
