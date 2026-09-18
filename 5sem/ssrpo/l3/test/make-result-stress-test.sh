#!/bin/bash

echo "Start" > test/stress-test.out

for file in test/source/stress-test-templates/*.test
do
    echo "=== ${file} ===" >> test/stress-test.out
    cat ${file} >> test/stress-test.out
    echo "--- Test --->" >> test/stress-test.out
    bin/stressgen < ${file} > test/stress-test-tmp.out
    bin/lab < test/stress-test-tmp.out >> test/stress-test.out
done

exit 0
