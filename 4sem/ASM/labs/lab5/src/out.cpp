
#include <iostream>

extern "C" void print_result(char *res) {
    std::cout << "results of the execution: " << res << std::endl;
}
