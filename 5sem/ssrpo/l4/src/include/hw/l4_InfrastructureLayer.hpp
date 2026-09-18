#pragma once

#include <chrono>
#include <unordered_map>
#include <memory>
#include <mutex>
#include <vector>
#include <string>

class AppendableInterface {
public:
    virtual ~AppendableInterface() = default;

    virtual void set_created_at(std::chrono::time_point<std::chrono::system_clock> time) = 0;
};

/**
  * Значение в БД
  */
typedef std::shared_ptr<AppendableInterface> item_ptr;


/**
  * Представление базы данных - key-value map
  */
typedef std::unordered_map<std::string, item_ptr> database;

/**
  * Тип для вывода всех элементов из БД
  */
typedef std::vector<std::pair<std::string, item_ptr> > db_elements;

/**
  * Класс для управления элементами базы данных.
  */
class DataBase {
    database data;
    std::mutex db_mutex;

public:
    ~DataBase() = default;

    DataBase() = default;

    DataBase(const DataBase &db) : data(db.data) {
    }

    /**
      * Добавляет элемент в базу данных.
      * @param {std::string} key - Ключ элемента.
      * @param {std::shared_ptr<AppendableInterface>} item - Элемент, который добавляется в базу данных.
      */
    void add_item(const std::string &key, item_ptr item);

    /**
        * Удаляет элемент из базы данных по ключу.
        * @param {std::string} key - Ключ элемента, который нужно удалить.
        */
    void remove_item(const std::string &key);

    /**
        * Обновляет элемент в базе данных, если он существует.
        * @param {std::string} key - Ключ элемента, который нужно обновить.
        * @param {std::shared_ptr<AppendableInterface>} item - Новый элемент для обновления.
        */
    void update_item(const std::string &key, item_ptr item);

    /**
        * Получает элемент из базы данных по ключу.
        * @param {std::string} key - Ключ элемента для получения.
        * @returns {std::shared_ptr<AppendableInterface>} - Элемент, если найден, иначе nullptr.
        */
    item_ptr get_item(const std::string &key);

    /**
        * Получает все элементы базы данных.
        * @returns {db_elements} - Контейнер со всеми элементами базы данных.
        */
    db_elements get_all();
};
