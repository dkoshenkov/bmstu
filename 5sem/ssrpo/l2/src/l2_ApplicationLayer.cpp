#include "../include/hw/l3_DomainLayer.hpp"
#include <sstream>
#include "../include/hw/l2_ApplicationLayer.hpp"
#include <stdexcept>     // std::invalid_argument, std::out_of_range

constexpr int PRODUCT_NAME_NUM = 1;
constexpr int PRODUCT_COST_NUM = 2;
constexpr int PRODUCT_COUNT_NUM = 3;

std::atomic<int> Application::order_counter{0};

bool Application::run() {
    _io.output("для окончания e/end");
    _io.output("для создания продукта p/new <ИМЯ ПРОДУКТА> <ЦЕНА ПРОДУКТА> <КОЛИЧЕСТВО>");
    _io.output("для удаления продукта p/remove <ИМЯ ПРОДУКТА>");
    _io.output("для обновления продукта p/update <ИМЯ ПРОДУКТА> <НОВАЯ ЦЕНА> <НОВОЕ КОЛИЧЕСТВО>");
    _io.output("для обновления продукта p/view_all");
    _io.output("для создания клиента c/new <ИМЯ> <ФАМИЛИЯ> <EMAIL>");
    _io.output("для удаления клиента c/remove <ИМЯ ФАМИЛИЯ>");
    _io.output("для вывода всех клиентов c/view_all");
    _io.output("для создания заказа o/new_order <ИМЯ КЛИЕНТА> <ФАМИЛИЯ КЛИЕНТА> <АДРЕС ДОСТАВКИ>");
    _io.output("для добавления продукта в заказ o/add_product <НОМЕР ЗАКАЗА> <НАЗВАНИЕ ПРОДУКТА> <ЦЕНА>");
    _io.output("для вывода всех заказов o/view_all_orders");
    _io.output("для удаления заказа o/remove order_id");

    while (true) {
        auto user_input = _io.input();

        if (user_input.empty()) {
            _io.output("Нет введённых данных. Завершаем работу...");
            break;
        }

        // командам "e" или "end" выходим
        if (user_input[0] == "e" || user_input[0] == "end") {
            _io.output("Ввод завершен");
            break;
        }

        // ---------- работа с продуктами ----------
        if (user_input[0] == "p/new") {
            if (user_input.size() < 4) {
                _io.output("Недостаточно данных для создания продукта.");
                continue;
            }
            const std::string& name = user_input[1];
            try {
                unsigned int cost = std::stoul(user_input[2]);
                unsigned int count = std::stoul(user_input[3]);

                Product product(name, cost, count);
                _product_manager.add_product(product);

                _io.output("Продукт успешно добавлен.");
            }
            catch (const std::invalid_argument& e) {
                _io.output("Ошибка: некорректное число в команде p/new.");
            } catch (const std::out_of_range& e) {
                _io.output("Ошибка: число выходит за допустимые пределы.");
            }
        }
        else if (user_input[0] == "p/remove") {
            if (user_input.size() < 2) {
                _io.output("Недостаточно данных для удаления продукта.");
                continue;
            }
            const std::string& name = user_input[1];
            _product_manager.delete_product(name);
            _io.output("Продукт успешно удалён.");
        }
        else if (user_input[0] == "p/update") {
            if (user_input.size() < 4) {
                _io.output("Недостаточно данных для обновления продукта.");
                continue;
            }
            const std::string& name = user_input[1];
            try {
                unsigned int new_cost = std::stoul(user_input[2]);
                unsigned int new_count = std::stoul(user_input[3]);

                Product product(name, new_cost, new_count);
                _product_manager.update_product(product);
                _io.output("Продукт успешно обновлён.");
            }
            catch (const std::invalid_argument& e) {
                _io.output("Ошибка: некорректное число в команде p/update.");
            } catch (const std::out_of_range& e) {
                _io.output("Ошибка: число выходит за допустимые пределы.");
            }
        }
        else if (user_input[0] == "p/view_all") {
            auto products = _product_manager.get_all_products();
            for (auto & product : products) {
                auto cost = std::to_string(product->product_cost());
                _io.output(product->product_name() + '/' + cost);
            }
        }
        // ---------- работа с клиентами ----------

        else if (user_input[0] == "c/new") {
            if (user_input.size() < 4) {
                _io.output("Недостаточно данных для создания клиента.");
                continue;
            }

            const std::string& first_name = user_input[1];
            const std::string& last_name = user_input[2];
            const std::string& email = user_input[3];
            Person person(first_name, last_name, email);
            _person_manager.add_person(person);
            _io.output("Клиент успешно добавлен: " + first_name + " " + last_name + " " + email);
        }
        else if (user_input[0] == "c/remove") {
            if (user_input.size() < 3) {
                _io.output("Недостаточно данных для удаления клиента.");
                continue;
            }
            std::string full_name = user_input[1] + " " + user_input[2];
            _person_manager.delete_person(full_name);
            _io.output("Клиент " + full_name + " успешно удалён.");
        }
        else if (user_input[0] == "c/view_all") {
            auto all_persons = _person_manager.get_all_persons();
            if (all_persons.empty()) {
                _io.output("Нет доступных клиентов.");
                continue;
            }
            for (const auto& person : all_persons) {
                std::stringstream ss;
                ss << "customer: " << person->get_first_name()
                    << " " << person->get_last_name()
                    << " " << person->get_email();
                _io.output(ss.str());
            }

            // ---------- работа с заказами ----------
        }
        else if (user_input[0] == "o/new_order") {
            // Нужно минимум 4 аргумента: "o/new_order", "<Имя>", "<Фамилия>", "<Адрес>"
            if (user_input.size() < 4) {
                _io.output("Недостаточно данных для создания заказа.");
                continue;
            }
            std::string customer_name = user_input[1] + " " + user_input[2];
            const std::string& delivery_address = user_input[3];
            if (auto customer = _person_manager.get_person(customer_name)) {
                Order order(customer, delivery_address);
                int current_order_id = Application::order_counter.fetch_add(1);
                _order_manager.add_order(order, current_order_id);
                _io.output("Заказ успешно создан (ID: " + std::to_string(current_order_id) + ").");
            }
            else {
                _io.output("Клиент не найден: " + customer_name);
            }
        }
        else if (user_input[0] == "o/add_product") {
            if (user_input.size() < 4) {
                _io.output("Недостаточно данных для добавления продукта в заказ.");
                continue;
            }
            try {
                int order_id = std::stoi(user_input[1]);
                const std::string& product_name = user_input[2];
                unsigned int product_price = std::stoul(user_input[3]);

                if (auto order = _order_manager.get_order(order_id)) {
                    auto product = std::make_shared<Product>(product_name, product_price, 1);
                    order->add_product(product);
                    _io.output("Продукт добавлен в заказ " + std::to_string(order_id) + ".");
                }
                else {
                    _io.output("Заказ с ID " + std::to_string(order_id) + " не найден.");
                }
            }
            catch (const std::invalid_argument& e) {
                _io.output("Ошибка: некорректный номер заказа или цена продукта.");
            } catch (const std::out_of_range& e) {
                _io.output("Ошибка: число выходит за допустимые пределы.");
            }
        }
        else if (user_input[0] == "o/view_all_orders") {
            auto all_orders = _order_manager.get_all_orders();
            if (all_orders.empty()) {
                _io.output("Нет доступных заказов.");
                continue;
            }
            for (const auto& order : all_orders) {
                _io.output(order->get_info());
            }
            // ---------- нераспознанная команда ----------
        }
        else if (user_input[0] == "o/remove") {
            auto order_id = std::stoi(user_input[1]);
            _order_manager.delete_order(order_id);
            _io.output("заказ удален с ID: " + user_input[1]);
        }

        else {
            _io.output("Недопустимая команда '" + user_input[0] + "'");
        }
    }

    return true;
}
