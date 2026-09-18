
#include "../include/hw/l3_DomainLayer.hpp"
#include <sstream>
#include "../include/hw/l2_ApplicationLayer.hpp"

const int PRODUCT_NAME_NUM = 1;
const int PRODUCT_COST_NUM = 2;
const int PRODUCT_COUNT_NUM = 3;

bool Application::run() {
    _io.output("для окончания e/end");
    _io.output("для создания n/new <ИМЯ ПРОДУКТА> <ЦЕНА ПРОДУКТА> <КОЛИЧЕСТВО>");
    _io.output("для удаления продукта r/remove <ИМЯ ПРОДУКТА>");
    _io.output("для поставки продукта a/add <ИМЯ ПРОДУКТА> <НОВАЯ ЦЕНА> <КОЛИЧЕСТВО>");
    _io.output("для покупки b/buy <ИМЯ ПРОДУКТА> <КОЛИЧЕСТВО>");
    _io.output("для вывода информации о продукте v/view <ИМЯ ПРОДУКТА>");
    _io.output("для вывода всех продуктов va/view_all");

    while (true){
        auto user_input = _io.input();

        if (user_input.empty()) {
            _io.output("Нет введённых данных.");
            break;
        }

        if (user_input[0] == "e" || user_input[0] == "end"){
            _io.output("Ввод завершен");
            break;
        }

        else if (user_input[0] == "n" || user_input[0] == "new") {
            if (user_input.size() < 4) {
                return false;
            }

            std::string name = user_input[PRODUCT_NAME_NUM];
            unsigned int cost = std::stoul(user_input[PRODUCT_COST_NUM]);
            unsigned int count = std::stoul(user_input[PRODUCT_COUNT_NUM]);

            _pm.add_product({name, cost, count});

            _io.output("Продукт успешно добавлен");
        }

        else if (user_input[0] == "r" || user_input[0] == "remove") {
            _pm.delete_product(user_input[PRODUCT_NAME_NUM]);
        }

        else if (user_input[0] == "a" || user_input[0] == "add") {
            std::string name = user_input[1];
            unsigned int cost = std::stoul(user_input[2]);
            unsigned int count = std::stoul(user_input[3]);
            _pm.update_product({name, cost, count});
        }

        else if (user_input[0] == "v" || user_input[0] == "view"){
            std::string name = user_input[1];
            auto product = _pm.get_product(name);
            std::stringstream ss;
            ss << "Имя: " << product->product_name() << " Цена: " << product->product_cost() << " Количество " << product->count();
            _io.output(ss.str());
        }

        else if (user_input[0] == "va" || user_input[0] == "view_all") {
            std::vector<std::shared_ptr<Product>> all_products = _pm.get_all_products();
               if (all_products.empty()) {
                   _io.output("Нет доступных продуктов.");
                   continue;
               }

               for (const auto& product_ptr : all_products) {
                   if (product_ptr) {
                       std::stringstream ss;
                       ss << "Имя: " << product_ptr->product_name() << " Цена: " << product_ptr->product_cost() << "Количество " << product_ptr->count();
                       _io.output(ss.str());
                   }
               }
        }

        else {
            _io.output("Недопустимая команда '" + user_input[0] + "'");
            return false;
        }
    }

    return true;
}
