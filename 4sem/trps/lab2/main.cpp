#include <iostream>
#include <cstdlib>
#include <malloc.h>
#include <chrono>
#include <cmath>

void qsortRecursive(double *mas, int size);
void random_data(double* mas[], int size);
double** new_memory(int size);
void free_mem(double** mas);
void out(double** mas, int size);



void func1(){
    int N , i, j, k, l;
    printf("Введите количество элементов : ");
    scanf("%d",&N);

    auto start = std::chrono::high_resolution_clock::now();
    double** a = (double**)malloc(sizeof(int*) * N);
    for(i = 0; i < N; i++)
        a[i] = (double*)malloc(sizeof(int*) * N);
    for (i = 0; i<N; i++){
        for (j = 0; j<N; j++){
            a[i][j] = - 100 + rand()%(100 + 100 + 1);
        }
    }

    auto start_sort = std::chrono::high_resolution_clock::now();
    for (i = 1; i<N; i+=2){
        for (j = 1; j<N; j++){
            for (k =0; k<N-1; k++){
                if (a[i][k]>a[i][k+1]){
                    float x = a[i][k];
                    a[i][k] = a[i][k+1];
                    a[i][k+1] = x;
                }
            }
        }
    }
    auto end_sort = std::chrono::high_resolution_clock::now();


    // Освобождение памяти
    for (int i = 0; i < N; i++) {
        free(a[i]); // Освободить память для каждой строки массива
    }


    free(a); // Освободить память для указателей на строки

// Установить указатель на NULL
    a = NULL;
    auto end = std::chrono::high_resolution_clock::now();
    double total_time = static_cast<double>(std::chrono::duration_cast<std::chrono::nanoseconds>(end - start).count()) ;
    double sort_time = static_cast<double>(std::chrono::duration_cast<std::chrono::nanoseconds>(end_sort - start_sort).count());
    std::cout << "runtime func 1 = " << std::endl; // время работы программы
    std::cout << (total_time / std::pow(10, 6)) << ' ' << (sort_time / std::pow(10, 6)) << std::endl;
}

// ошибка 1 - нужно объявлять переменные максимально близко к их использованию (i j k)
// ошибка 2 - нет проверки на ввод отрицательных чисел
// ошибка 3 - использование c-функций (printf scanf malloc)
// ошибка 4 - нет очистки памяти
// l - неиспользуется
// ошибка 6 - уменьшение точности float x

// улучшение -- можно выделять память под одномерный массив и делить на столбцы
// улучшение -- изменен алгоритм с bubble sort на quick sort
void func2(){
    int size;
    std::cout << "Введите количество элементов: ";
    std::cin >> size;
    if (size <= 0){
        std::cout << "ERROR!!! size>0!!!" << std::endl;
        return;
    }
    auto start = std::chrono::high_resolution_clock::now();
    auto m = new_memory(size);

    random_data(m, size);
    auto start_sort = std::chrono::high_resolution_clock::now();
    for (int i = 1; i < size; i+=2) {
        qsortRecursive(m[i], size);
    }
    auto end_sort = std::chrono::high_resolution_clock::now();

    free_mem(m);
    auto end = std::chrono::high_resolution_clock::now();
    double total_time = static_cast<double>(std::chrono::duration_cast<std::chrono::nanoseconds>(end - start).count()) ;
    double sort_time = static_cast<double>(std::chrono::duration_cast<std::chrono::nanoseconds>(end_sort - start_sort).count());

    std::cout << "runtime func 2 = " << std::endl; // время работы программы
    std::cout << (total_time / std::pow(10, 6)) << ' ' << (sort_time / std::pow(10, 6)) << std::endl;
}


int main(){
    func1();
//    func2();
    return 0;
}
void out(double** mas, int size){
    for (int i = 0; i<size; i++){
        for (int j = 0; j<size; j++){
            std::cout <<"    " << mas[i][j];
        }
        std::cout << std::endl;
    }
}

void free_mem(double** mas){
    delete[] mas[0];
    delete[] mas;
}


// new_memory эффективное выделение памяти в стиле С++
double** new_memory(int size){
    auto ** m = new double*[size];
    m[0] = new double[size*size];
    for (auto i = 1; i < size; ++i){
        m[i] = m[i-1] + size;
    }

    return m;
}
void random_data(double* mas[], int size){
    for (int i = 0; i<size; i++){
        for (int j = 0; j<size; j++){
            mas[i][j] = - 100 + rand()%(100 + 100 + 1);
        }
    }
}
void input(double* mas[], int size){
    for (int i = 0; i<size; i++){
        for (int j = 0; j<size; j++){
            std::cin >>mas[i][j];
        }
    }
}
void qsortRecursive(double *mas, int size) {
    //Указатели в начало и в конец массива
    int i = 0;
    int j = size - 1;

    //Центральный элемент массива
    double mid = mas[size / 2];

    //Делим массив
    do {
        //Пробегаем элементы, ищем те, которые нужно перекинуть в другую часть
        //В левой части массива пропускаем(оставляем на месте) элементы, которые меньше центрального
        while(mas[i] < mid) {
            i++;
        }
        //В правой части пропускаем элементы, которые больше центрального
        while(mas[j] > mid) {
            j--;
        }

        //Меняем элементы местами
        if (i <= j) {
            auto tmp = mas[i];
            mas[i] = mas[j];
            mas[j] = tmp;

            i++;
            j--;
        }
    } while (i <= j);


    //Рекурсивные вызовы, если осталось, что сортировать
    if(j > 0) {
        //"Левый кусок"
        qsortRecursive(mas, j + 1);
    }
    if (i < size) {
        //"Првый кусок"
        qsortRecursive(&mas[i], size - i);
    }
}
