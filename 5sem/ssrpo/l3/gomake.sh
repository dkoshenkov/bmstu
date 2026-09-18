#!/bin/bash

cmake -Bbuild/release -H. -DCMAKE_BUILD_TYPE=Release
cmake --build build/release --config Release -j 6
