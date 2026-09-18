#pragma once

#include <atomic>
#include <string>
#include <utility>
#include <vector>
#include <tp/Task_interface.h>

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


enum Command {
    NEW_PRODUCT,
    DELETE_PRODUCT,
    UPDATE_PRODUCT,
    VIEW_ALL_PRODUCTS,
    NEW_CLIENT,
    DELETE_CLIENT,
    UPDATE_CLIENT,
    VIEW_ALL_CLIENTS,
    NEW_ORDER,
    DELETE_ORDER,
    ADD_PRODUCT_TO_ORDER,
    VIEW_ALL_ORDERS,
    DEFAULT
};


class Application : public tp::Task_interface {
public:
    Application() = delete;

    Application(const Application &) = delete;

    Application &operator=(const Application &) = delete;

    Application(InputOutputInterface &io,
                DataBase &db,
                Command command,
                std::string product_name = "",
                unsigned int product_cost = 0,
                unsigned int product_count = 0,
                std::string client_first_name = "",
                std::string client_last_name = "",
                std::string client_email = "",
                std::string customer_name = "",
                std::string delivery_address = "",
                int order_id = 0) : _io(io), _db(db),
                                    _product_manager(&db),
                                    _person_manager(&db),
                                    _order_manager(&db),
                                    _command(command), _product_name(std::move(product_name)),
                                    _product_cost(product_cost),
                                    _product_count(product_count), _client_first_name(std::move(client_first_name)),
                                    _client_last_name(std::move(client_last_name)),
                                    _client_email(std::move(client_email)),
                                    _customer_name(std::move(customer_name)),
                                    _delivery_address(std::move(delivery_address)),
                                    _order_id(order_id) {
    }

    Application(InputOutputInterface &io, DataBase &db, Command command,
                std::string client_first_name, std::string client_last_name, std::string client_email) : Application(
        io, db, command, "", 0, 0, std::move(client_first_name), std::move(client_last_name), std::move(client_email)) {
    }

    Application(InputOutputInterface &io, DataBase &db, Command command,
                std::string customer_name, std::string delivery_address) : Application(
        io, db, command, "", 0, 0, "", "", "",
        std::move(customer_name), std::move(delivery_address)) {
    }

    Application(InputOutputInterface &io, DataBase &db, Command command, int order_id): Application(
        io, db, command, "", 0, 0, "", "", "",
        "", "", order_id) {
    }

    static std::atomic<int> order_counter;

    void work() override;

private:
    decompized_args split_by_spaces(const arg &input_string);

    InputOutputInterface &_io;
    DataBase &_db;
    ProductManager _product_manager;
    PersonManager _person_manager;
    OrderManager _order_manager;
    Command _command;
    std::string _product_name;
    unsigned int _product_cost;
    unsigned int _product_count;

    std::string _client_first_name;
    std::string _client_last_name;
    std::string _client_email;

    std::string _customer_name;
    std::string _delivery_address;

    int _order_id;
};
