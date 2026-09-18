#!/bin/bash

echo "Start" > test/lab.out

for file in test/source/lab/*.test
do
    echo "=== ${file} ===" >> test/lab.out
    cat ${file} >> test/lab.out
    echo "--- Test --->" >> test/lab.out
    bin/lab < ${file} >> test/lab.out
done

exit 0
