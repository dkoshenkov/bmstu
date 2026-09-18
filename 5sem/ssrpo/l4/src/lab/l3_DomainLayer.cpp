#include "../include/hw/l3_DomainLayer.hpp"
#include <chrono>
#include <memory>
#include <sstream>

bool Product::invariant() const {
    std::lock_guard locker(_product_mutex);
    return !_name.empty() && _created_at <= std::chrono::system_clock::now();
}

void Product::set_cost(unsigned int new_cost) {
    std::lock_guard locker(_product_mutex);
    this->_cost = new_cost;
}

void Product::set_count(int new_count) {
    std::lock_guard locker(_product_mutex);
    this->_count = new_count;
}

const std::string &Product::product_name() const {
    std::lock_guard locker(_product_mutex);
    return _name;
}

unsigned int Product::product_cost() const {
    std::lock_guard locker(_product_mutex);
    return _cost;
}

unsigned int Product::count() const {
    std::lock_guard locker(_product_mutex);
    return _count;
}

void Product::set_created_at(
    std::chrono::time_point<std::chrono::system_clock> time) {
    std::lock_guard locker(_product_mutex);
    _created_at = time;
}

void ProductManager::add_product(Product &product) const {
    _db.add_item(product.product_name(), std::make_shared<Product>(product));
}

void ProductManager::delete_product(std::string product_name) const {
    _db.remove_item(product_name);
}

products ProductManager::get_all_products() const {
    std::lock_guard locker(_product_manager_mutex);
    db_elements all_elements = _db.get_all();
    std::vector<std::shared_ptr<Product> > _products;
    _products.reserve(all_elements.size());

    for (const auto &[key, item_ptr]: all_elements) {
        if (std::shared_ptr<Product> product_ptr = std::dynamic_pointer_cast<Product>(item_ptr)) {
            _products.emplace_back(product_ptr);
        }
    }

    return _products;
}

void ProductManager::update_product(const Product &product) const {
    std::lock_guard locker(_product_manager_mutex);

    auto item_ptr = _db.get_item(product.product_name());
    auto item = std::static_pointer_cast<Product>(item_ptr);
    if (item) {
        item->set_cost(product.product_cost());
        item->set_count(product.count());
        _db.update_item(product.product_name(), item);
    }
}

void Person::set_created_at(std::chrono::time_point<std::chrono::system_clock> time) {
    std::lock_guard locker(_person_mutex);

    created_at = time;
}

const std::string &Person::get_first_name() const {
    std::lock_guard locker(_person_mutex);
    return first_name;
}

const std::string &Person::get_last_name() const {
    std::lock_guard locker(_person_mutex);

    return last_name;
}

const std::string &Person::get_email() const {
    std::lock_guard locker(_person_mutex);

    return email;
}


void Order::add_product(const std::shared_ptr<Product> &product) {
    std::lock_guard locker(_order_mutex);

    products.push_back(product);
}

std::string Order::get_info() const {
    std::lock_guard locker(_order_mutex);

    std::stringstream ss;

    ss << "customer: " + this->customer->get_first_name()
            + " " + this->customer->get_last_name()
            + " " + this->customer->get_email()
            << std::endl;

    ss << "products: " << products.size() << std::endl;
    for (const auto &product: products) {
        ss << "  name:" << product->product_name() << std::endl;
        ss << "  cost:" << product->product_cost() << std::endl;
    }
    ss << "delivery address: " << this->delivery_address << std::endl;

    auto t_c = std::chrono::system_clock::to_time_t(this->created_at);

    struct tm *tm_info = localtime(&t_c);

    char date_str[11];
    strftime(date_str, sizeof(date_str), "%Y-%m-%d", tm_info);
    ss << "date: " << date_str << std::endl;

    return ss.str();
}


void Order::set_created_at(std::chrono::time_point<std::chrono::system_clock> time) {
    std::lock_guard locker(_order_mutex);

    created_at = time;
}


void PersonManager::add_person(const Person &person) const {
    auto person_ptr = std::make_shared<Person>(person);
    _db.add_item(person.get_first_name() + " " + person.get_last_name(), person_ptr);
}


void PersonManager::delete_person(const std::string &full_name) const {
    _db.remove_item(full_name);
}

std::shared_ptr<Person> PersonManager::get_person(const std::string &full_name) const {
    std::lock_guard locker(_person_manager_mutex);

    return std::dynamic_pointer_cast<Person>(_db.get_item(full_name));
}

std::vector<std::shared_ptr<Person> > PersonManager::get_all_persons() const {
    db_elements all_elements = _db.get_all();
    std::lock_guard locker(_person_manager_mutex);
    std::vector<std::shared_ptr<Person> > persons;

    for (const auto &[key, value]: all_elements) {
        auto person = std::dynamic_pointer_cast<Person>(value);
        if (person) {
            persons.push_back(person);
        }
    }
    return persons;
}

void OrderManager::add_order(const Order &order, int order_id) const {
    auto order_ptr = std::make_shared<Order>(order);
    _db.add_item("order_" + std::to_string(order_id), order_ptr);
}

void OrderManager::delete_order(int order_id) const {
    _db.remove_item("order_" + std::to_string(order_id));
}

std::shared_ptr<Order> OrderManager::get_order(int order_id) const {
    std::lock_guard locker(_order_manager_mutex);
    return std::dynamic_pointer_cast<Order>(_db.get_item("order_" + std::to_string(order_id)));
}

std::vector<std::shared_ptr<Order> > OrderManager::get_all_orders() const {
    db_elements all_elements = _db.get_all();
    std::lock_guard locker(_order_manager_mutex);

    std::vector<std::shared_ptr<Order> > orders;
    for (const auto &[key, value]: all_elements) {
        auto order = std::dynamic_pointer_cast<Order>(value);
        if (order) {
            orders.push_back(order);
        }
    }
    return orders;
}
