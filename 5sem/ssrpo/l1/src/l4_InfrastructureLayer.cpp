
#include "../include/hw/l4_InfrastructureLayer.hpp"
#include <memory>

void DataBase::add_item(const std::string& key, item_ptr item_to_insert){
    data[key] = item_to_insert;
}

/**
    * Удаляет элемент из базы данных по ключу.
    * @param {std::string} key - Ключ элемента, который нужно удалить.
    */
void DataBase::remove_item(const std::string& key){
    data.erase(key);
}

/**
    * Обновляет элемент в базе данных, если он существует.
    * @param {std::string} key - Ключ элемента, который нужно обновить.
    * @param {std::shared_ptr<AppendableInterface>} item - Новый элемент для обновления.
    */
void DataBase::update_item(const std::string& key, item_ptr item_to_insert){
    data[key] = item_to_insert;
}

/**
    * Получает элемент из базы данных по ключу.
    * @param {std::string} key - Ключ элемента для получения.
    * @returns {std::shared_ptr<AppendableInterface>} - Элемент, если найден, иначе nullptr.
    */
item_ptr DataBase::get_item(const std::string& key) {
    auto db_item = data.at(key);
    return db_item;
}

/**
    * Получает все элементы базы данных.
    * @returns {db_elements} - Контейнер со всеми элементами базы данных.
    */
db_elements DataBase::get_all() const {
    db_elements elements;
    std::copy(data.begin(), data.end(), std::back_inserter(elements));

    return elements;
}
