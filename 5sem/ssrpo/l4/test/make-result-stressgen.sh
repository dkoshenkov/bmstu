#!/bin/bash

echo "Start" > test/stressgen.out

for file in test/source/stressgen/*.test
do
    echo "=== ${file} ===" >> test/stressgen.out
    cat ${file} >> test/stressgen.out
    echo "--- Test --->" >> test/stressgen.out
    bin/stressgen < ${file} >> test/stressgen.out
done

exit 0
