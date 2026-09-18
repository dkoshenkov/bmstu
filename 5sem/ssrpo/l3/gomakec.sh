#!/bin/bash

cmake -Bbuild/rwdi -H. -DCMAKE_BUILD_TYPE=RelWithDebInfo
cmake --build build/rwdi --config RelWithDebInfo -j 6
