#!/bin/bash

if [ -f lab.data ]
then
  rm lab.data
  echo "Файл lab.data удалён"
fi

./test/make-result-lab.sh
diff test/samples/lab.sample test/lab.out

if [ $? -ne 0 ]
then
  echo -e "\033[1mОшибка при выполнении теста test/make-result-lab\033[0m"
  exit 1
fi

./test/make-result-stressgen.sh
# diff test/samples/stressgen.sample test/stressgen.out

if [ $? -ne 0 ]
then
  echo -e "\033[1mОшибка при выполнении теста test/make-result-stressgen\033[0m"
  exit 1
fi

