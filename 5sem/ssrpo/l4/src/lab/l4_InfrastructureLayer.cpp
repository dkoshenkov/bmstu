#include "../include/hw/l4_InfrastructureLayer.hpp"
#include <memory>
#include <mutex>

void DataBase::add_item(const std::string &key, item_ptr item_to_insert) {
    std::lock_guard<std::mutex> lock(db_mutex);
    item_to_insert->set_created_at(std::chrono::system_clock::now());
    data[key] = item_to_insert;
}

/**
    * Удаляет элемент из базы данных по ключу.
    * @param {std::string} key - Ключ элемента, который нужно удалить.
    */
void DataBase::remove_item(const std::string &key) {
    std::lock_guard<std::mutex> lock(db_mutex);
    data.erase(key);
}

/**
    * Обновляет элемент в базе данных, если он существует.
    * @param {std::string} key - Ключ элемента, который нужно обновить.
    * @param {std::shared_ptr<AppendableInterface>} item - Новый элемент для обновления.
    */
void DataBase::update_item(const std::string &key, item_ptr item) {
    std::lock_guard<std::mutex> lock(db_mutex);
    if (data.find(key) != data.end()) {
        data[key] = item;
    }
}

/**
    * Получает элемент из базы данных по ключу.
    * @param {std::string} key - Ключ элемента для получения.
    * @returns {std::shared_ptr<AppendableInterface>} - Элемент, если найден, иначе nullptr.
    */
item_ptr DataBase::get_item(const std::string &key) {
    std::lock_guard<std::mutex> lock(db_mutex);
    if (data.find(key) != data.end()) {
        return data[key];
    }

    return nullptr;
}

/**
    * Получает все элементы базы данных.
    * @returns {db_elements} - Контейнер со всеми элементами базы данных.
    */
db_elements DataBase::get_all() {
    std::lock_guard<std::mutex> lock(db_mutex);
    db_elements elements;
    for (const auto &[key, value]: data) {
        elements.push_back({key, value});
    }

    return elements;
}
