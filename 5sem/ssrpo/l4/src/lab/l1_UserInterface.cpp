#include <cassert>
#include <iostream>
#include <sstream>
#include <string>
#include <tp/ThreadPool.h>

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


void run(Console &konsole, DataBase &db, int number_of_threads = -1) {
    tp::ThreadPool tp(number_of_threads);
    tp.start();

    konsole.output("для окончания e/end");
    konsole.output("для создания продукта p/new <ИМЯ ПРОДУКТА> <ЦЕНА ПРОДУКТА> <КОЛИЧЕСТВО>");
    konsole.output("для удаления продукта p/remove <ИМЯ ПРОДУКТА>");
    konsole.output("для обновления продукта p/update <ИМЯ ПРОДУКТА> <НОВАЯ ЦЕНА> <НОВОЕ КОЛИЧЕСТВО>");
    konsole.output("для вывода всех продуктов p/view_all");
    konsole.output("для создания клиента c/new <ИМЯ> <ФАМИЛИЯ> <EMAIL>");
    konsole.output("для удаления клиента c/remove <ИМЯ> <ФАМИЛИЯ>");
    konsole.output("для вывода всех клиентов c/view_all");
    konsole.output("для создания заказа o/new_order <ИМЯ КЛИЕНТА> <ФАМИЛИЯ КЛИЕНТА> <АДРЕС ДОСТАВКИ>");
    konsole.output("для добавления продукта в заказ o/add_product <НОМЕР ЗАКАЗА> <НАЗВАНИЕ ПРОДУКТА> <ЦЕНА>");
    konsole.output("для вывода всех заказов o/view_all_orders");
    konsole.output("для удаления заказа o/remove <НОМЕР ЗАКАЗА>");

    while (true) {
        auto user_input = konsole.input();

        if (user_input.empty()) {
            konsole.output("Нет введённых данных. Завершаем работу...");
            break;
        }

        if (user_input[0] == "e" || user_input[0] == "end") {
            konsole.output("Ввод завершен");
            break;
        }

        // Работа с продуктами
        if (user_input[0] == "p/new") {
            if (user_input.size() < 4) {
                konsole.output("Недостаточно данных для создания продукта.");
                continue;
            }
            const std::string &name = user_input[1];
            try {
                unsigned int cost = std::stoul(user_input[2]);
                unsigned int count = std::stoul(user_input[3]);
                tp.submit(new Application(konsole, db, NEW_PRODUCT, name, cost, count));
            } catch (const std::exception &e) {
                konsole.output("Ошибка при создании продукта: " + std::string(e.what()));
            }
        } else if (user_input[0] == "p/remove") {
            if (user_input.size() < 2) {
                konsole.output("Недостаточно данных для удаления продукта.");
                continue;
            }
            tp.submit(new Application(konsole, db, DELETE_PRODUCT, user_input[1]));
        } else if (user_input[0] == "p/update") {
            if (user_input.size() < 4) {
                konsole.output("Недостаточно данных для обновления продукта.");
                continue;
            }
            try {
                const std::string &name = user_input[1];
                unsigned int new_cost = std::stoul(user_input[2]);
                unsigned int new_count = std::stoul(user_input[3]);
                tp.submit(new Application(konsole, db, UPDATE_PRODUCT, name, new_cost, new_count));
            } catch (const std::exception &e) {
                konsole.output("Ошибка при обновлении продукта: " + std::string(e.what()));
            }
        } else if (user_input[0] == "p/view_all") {
            tp.submit(new Application(konsole, db, VIEW_ALL_PRODUCTS));
        }

        // Работа с клиентами
        else if (user_input[0] == "c/new") {
            if (user_input.size() < 4) {
                konsole.output("Недостаточно данных для создания клиента.");
                continue;
            }
            tp.submit(new Application(konsole, db, NEW_CLIENT, user_input[1], user_input[2], user_input[3]));
        } else if (user_input[0] == "c/remove") {
            if (user_input.size() < 3) {
                konsole.output("Недостаточно данных для удаления клиента.");
                continue;
            }
            tp.submit(new Application(konsole, db, DELETE_CLIENT, user_input[1], user_input[2]));
        } else if (user_input[0] == "c/view_all") {
            tp.submit(new Application(konsole, db, VIEW_ALL_CLIENTS));
        } else if (user_input[0] == "c/update") {
            if (user_input.size() < 4) {
                konsole.output("Недостаточно данных для обновления клиента.");
                continue;
            }
            tp.submit(new Application(konsole, db, UPDATE_CLIENT, user_input[1], user_input[2], user_input[3]));
        }


        // Работа с заказами
        else if (user_input[0] == "o/new_order") {
            if (user_input.size() < 4) {
                konsole.output("Недостаточно данных для создания заказа.");
                continue;
            }
            std::string customer_name = user_input[1] + " " + user_input[2];
            tp.submit(new Application(konsole, db, NEW_ORDER, customer_name, user_input[3]));
        } else if (user_input[0] == "o/add_product") {
            if (user_input.size() < 4) {
                konsole.output("Недостаточно данных для добавления продукта в заказ.");
                continue;
            }
            try {
                int order_id = std::stoi(user_input[1]); // Преобразование order_id
                const std::string &product_name = user_input[2];
                unsigned int product_cost = std::stoul(user_input[3]); // Преобразование стоимости продукта

                tp.submit(new Application(konsole, db, ADD_PRODUCT_TO_ORDER,
                                          product_name, product_cost, 0, "", "", "",
                                          "", "", order_id)); // Используем конструктор с order_id
            } catch (const std::invalid_argument &e) {
                konsole.output("Ошибка: некорректный ID заказа или стоимость продукта.");
            } catch (const std::out_of_range &e) {
                konsole.output("Ошибка: значение выходит за допустимые пределы.");
            }
        } else if (user_input[0] == "o/view_all_orders") {
            tp.submit(new Application(konsole, db, VIEW_ALL_ORDERS));
        } else if (user_input[0] == "o/remove") {
            if (user_input.size() < 2) {
                konsole.output("Недостаточно данных для удаления заказа.");
                continue;
            }
            try {
                unsigned int order_id = std::stoul(user_input[1]);
                tp.submit(new Application(konsole, db, DELETE_ORDER, order_id));
            } catch (const std::exception &e) {
                konsole.output("Ошибка при удалении заказа: " + std::string(e.what()));
            }
        } else {
            konsole.output("Неизвестная команда: " + user_input[0]);
        }
    }
}


int main() {
    DataBase db;
    Console konsole;

    Application app(konsole, db, DEFAULT);

    run(konsole, db);

    return 0;
}
