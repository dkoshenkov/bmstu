#include <cassert>
#include <iostream>
#include <sstream>
#include <string>

#include "../include/hw/l2_ApplicationLayer.hpp"
#include "../include/hw/l3_DomainLayer.hpp"
#include "../include/hw/l4_InfrastructureLayer.hpp"

class Console : public InputOutputInterface {
public:
    explicit Console(std::istream &input = std::cin, std::ostream &out = std::cout): _in(input), _out(out) {
    }

    std::ostream &output(const std::string &output_data) const override {
        _out << output_data << std::endl;

        return _out;
    }

    std::vector<std::string> input() const override {
        std::vector<std::string> result;
        std::string line, argument;
        std::getline(_in, line);

        std::istringstream arguments(line);
        while (arguments >> argument) {
            result.push_back(argument);
        }

        return result;
    }

private:
    std::istream &_in;
    std::ostream &_out;
};

int main() {
    DataBase db;
    Console konsole;

    Application app(konsole, db);

    if (!app.run()) {
        return 1;
    }

    return 0;
}
