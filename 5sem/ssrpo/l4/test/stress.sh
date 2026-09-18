#!/bin/bash

if [ -f lab.data ]
then
  rm lab.data
  echo "Файл lab.data удалён"
fi

./test/make-result-stress-test.sh
# grep "499:" test/stress-test.out > /dev/null

if [ $? -ne 0 ]
then
  echo -e "\033[1mОшибка при выполнении теста test/make-result-stress-test\033[0m"
  exit 1
fi

