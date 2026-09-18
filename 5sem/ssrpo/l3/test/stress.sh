#!/bin/bash

if [ -f lab.data ]
then
  rm lab.data
  echo "Файл lab.data удалён"
fi

./test/make-result-stress-test.sh
# diff test/samples/stress-test.sample test/stress-test.out

if [ $? -ne 0 ]
then
  echo -e "\033[1mОшибка при выполнении теста test/make-result-stress-test\033[0m"
  exit 1
fi

