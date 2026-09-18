#include "../include/hw/l3_DomainLayer.hpp"
#include <chrono>
#include <memory>

bool Product::invariant() const {
    return _name != "" && _created_at <= std::chrono::system_clock::now();
}

void Product::set_cost(int new_cost) {
    this->_cost = new_cost;
}

void Product::set_count(int new_count) {
    this->_count = new_count;
}

const std::string& Product::product_name() const{
    return _name;
}

const unsigned int Product::product_cost() const {
    return _cost;
}

unsigned int Product::count() const{
    return _count;
}

void Product::set_created_at(std::chrono::time_point<std::chrono::system_clock> time = std::chrono::system_clock::now()) {
    _created_at = time;
}

void ProductManager::add_product(Product product) {
    _db.add_item(product.product_name(), std::make_shared<Product>(product));
}

void ProductManager::delete_product(std::string product_name) {
    _db.remove_item(product_name);
}

products ProductManager::get_all_products() {
    db_elements all_elements = _db.get_all();
    std::vector<std::shared_ptr<Product>> _products;
    _products.reserve(all_elements.size());

    for (const auto& [key, item_ptr] : all_elements) {
        std::shared_ptr<Product> product_ptr = std::dynamic_pointer_cast<Product>(item_ptr);
        if (product_ptr) {
            _products.emplace_back(product_ptr);
        }
    }

    return _products;
}

void ProductManager::update_product(Product product) {
    auto item_ptr = _db.get_item(product.product_name());
    auto item = std::static_pointer_cast<Product>(item_ptr);
    item->set_cost(item->product_cost() + product.product_cost());
    item->set_count(item->count() + product.count());
    _db.update_item(product.product_name(), item);
}

std::shared_ptr<Product> ProductManager::get_product(std::string product_name) {
    auto item = _db.get_item(product_name);
    return std::static_pointer_cast<Product>(item);
}
