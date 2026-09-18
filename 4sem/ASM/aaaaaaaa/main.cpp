#include <iostream>

extern "C" int sum(int a, int b); // Объявление extern "C" 
                                   // для отключения name mangling

int main() {
    int result = sum(5, 3);
    std::cout << "Сумма: " << result << std::endl;
    return 0;
}