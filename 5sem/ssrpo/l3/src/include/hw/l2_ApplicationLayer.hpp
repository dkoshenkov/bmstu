#pragma once

#include <atomic>
#include <string>
#include <vector>

#include "l3_DomainLayer.hpp"


typedef std::string arg;
typedef std::vector<arg> args;
typedef std::vector<std::vector<arg> > decompized_args;


class InputOutputInterface {
public:
    virtual ~InputOutputInterface() = default;

    virtual args input() const = 0;

    virtual std::ostream &output(const arg &output_data) const = 0;
};


class Application {
public:
    Application() = delete;

    Application(const Application &) = delete;

    Application &operator=(const Application &) = delete;

    Application(InputOutputInterface &io, DataBase &db) : _io(io), _db(db), _product_manager(db), _person_manager(db),
                                                          _order_manager(db) {
    }

    bool run();

    static std::atomic<int> order_counter;

private:
    decompized_args split_by_spaces(const arg &input_string);

    InputOutputInterface &_io;
    DataBase &_db;
    ProductManager _product_manager;
    PersonManager _person_manager;
    OrderManager _order_manager;
};
