#include <iostream>
#include <string>

void run(std::istream &in, std::ostream &out);

int main()
{
    run(std::cin, std::cout);
}

extern "C" char *change_order(const char *input, size_t num);


void run(std::istream &in, std::ostream &out)
{
    std::string input;
    std::cout << "Input str:\n";
    std::getline(std::cin, input);
    size_t n;

    std::cout << "enter `num word `to run: ";
    int _n;
    std::cin >> _n;
    while (_n > 0)
    {
        input = change_order(input.c_str(), _n);

        std::cout <<"f ret: " << input << std::endl;

        std::cout << "enter `num word `to run: ";
        std::cin >> _n;
    }
}