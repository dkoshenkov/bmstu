#include "../include/hw/l3_DomainLayer.hpp"
#include <sstream>
#include "../include/hw/l2_ApplicationLayer.hpp"
#include <stdexcept>     // std::invalid_argument, std::out_of_range
#include <tp/ThreadPool.h>

constexpr int PRODUCT_NAME_NUM = 1;
constexpr int PRODUCT_COST_NUM = 2;
constexpr int PRODUCT_COUNT_NUM = 3;

std::atomic<int> Application::order_counter{0};

void Application::work() {
    switch (_command) {
        case NEW_PRODUCT: {
            Product new_product(_product_name, _product_cost, _product_count);
            _product_manager.add_product(new_product);
            _io.output("Продукт успешно добавлен.");
            break;
        }

        case DELETE_PRODUCT: {
            _product_manager.delete_product(_product_name);
            _io.output("Продукт успешно удалён.");
            break;
        }

        case UPDATE_PRODUCT: {
            Product update_product(_product_name, _product_cost, _product_count);
            _product_manager.update_product(update_product);
            _io.output("Продукт успешно обновлён.");
            break;
        }

        case VIEW_ALL_PRODUCTS: {
            auto products = _product_manager.get_all_products();
            for (auto &product: products) {
                auto cost = std::to_string(product->product_cost());
                _io.output(product->product_name() + '/' + cost);
            }
            break;
        }

        case NEW_CLIENT: {
            Person person(_client_first_name, _client_last_name, _client_email);
            _person_manager.add_person(person);
            _io.output(
                "Клиент успешно добавлен: " + _client_first_name + " " + _client_last_name + " " + _client_email);
            break;
        }

        case DELETE_CLIENT: {
            std::string full_name = _client_first_name + " " + _client_last_name;
            _person_manager.delete_person(full_name);
            _io.output("Клиент " + full_name + " успешно удалён.");
            break;
        }

        case UPDATE_CLIENT: {
            std::string full_name = _client_first_name + " " + _client_last_name;
            auto client = _person_manager.get_person(full_name);
            if (!client) {
                _io.output("Клиент " + full_name + " не найден.");
                break;
            }
            Person updated_person(_client_first_name, _client_last_name, _client_email);
            _person_manager.add_person(updated_person); // Перезаписываем клиента
            _io.output("Клиент успешно обновлён: " + full_name + " " + _client_email);
            break;
        }

        case VIEW_ALL_CLIENTS: {
            auto all_persons = _person_manager.get_all_persons();
            if (all_persons.empty()) {
                _io.output("Нет доступных клиентов.");
            }
            for (const auto &person: all_persons) {
                std::stringstream ss;
                ss << "Клиент: " << person->get_first_name()
                        << " " << person->get_last_name()
                        << " " << person->get_email();
                _io.output(ss.str());
            }
            break;
        }

        case NEW_ORDER: {
            if (auto customer = _person_manager.get_person(_customer_name)) {
                Order order(customer, _delivery_address);
                int current_order_id = order_counter.fetch_add(1);
                _order_manager.add_order(order, current_order_id);
                _io.output("Заказ успешно создан (ID: " + std::to_string(current_order_id) + ").");
            } else {
                _io.output("Клиент не найден: " + _customer_name);
            }
            break;
        }

        case ADD_PRODUCT_TO_ORDER: {
            auto order = _order_manager.get_order(_order_id);
            if (!order) {
                _io.output("Заказ с ID " + std::to_string(_order_id) + " не найден.");
                break;
            }
            auto product = _product_manager.get_all_products();
            auto it = std::find_if(product.begin(), product.end(),
                                   [this](const std::shared_ptr<Product> &p) {
                                       return p->product_name() == _product_name;
                                   });
            if (it == product.end()) {
                _io.output("Продукт " + _product_name + " не найден.");
                break;
            }
            order->add_product(*it);
            _io.output(
                "Продукт " + _product_name + " успешно добавлен в заказ с ID " + std::to_string(_order_id) + ".");
            break;
        }

        case DELETE_ORDER: {
            _order_manager.delete_order(_order_id);
            _io.output("Заказ удалён с ID: " + std::to_string(_order_id));
            break;
        }

        case VIEW_ALL_ORDERS: {
            auto orders = _order_manager.get_all_orders();
            for (auto &order: orders) {
                _io.output(order->get_info());
            }
            break;
        }

        default:
            _io.output("Неизвестная команда.");
            break;
    }
}
