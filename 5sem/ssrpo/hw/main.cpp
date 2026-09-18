#include <cassert>
#include <iomanip>
#include <iostream>
#include <memory>
#include <utility>
#include <vector>
#include <string>

// Интерфейс, описывающий предмет
class ItemInterface {
public:
    virtual ~ItemInterface() = default;

    /**
     * Функция, возвращающая название предмета
     *
     * @return Название предмета
     */
    virtual std::string get_name() const = 0;

    /**
     * Функция, возвращающая стоимость предмета
     *
     * @return Стоимость предмета
     */
    virtual double get_cost() const = 0;

    /**
     * Функция, проверяющая наличие предмета
     *
     * @return true, если предмет присутствует, иначе false
     */
    virtual bool is_present() const = 0;
};

// Интерфейс для описания коллекции
class CollectionInterface {
public:
    virtual ~CollectionInterface() = default;

    /**
     * Функция, возвращающая украденные предметы
     *
     * @return Вектор украденных предметов
     */
    virtual std::vector<std::shared_ptr<ItemInterface>> get_stolen_items() const = 0;

    /**
     * Функция, возвращающая общую стоимость украденных предметов
     *
     * @return Общая стоимость украденных предметов
     */
    virtual double get_stolen_items_cost_amount() const = 0;
};

// Интерфейс для описания отчета
class ReportInterface {
public:
    virtual ~ReportInterface() = default;

    /**
     * Функция, создающая отчет и возвращающая сумму страховой выплаты
     *
     * @param collection Коллекция предметов
     * @return Сумма страховой выплаты
     */
    virtual double make_report(const CollectionInterface* collection) const = 0;

    /**
     * Функция, создающая отчет по коллекции и выводящая его в поток
     *
     * @param out Поток вывода
     * @param collection Коллекция предметов
     */
    virtual void make_report(std::ostream& out, const CollectionInterface* collection) const = 0;
};

// Реализация интерфейса предмета в виде монеты
class Coin final : public ItemInterface {
public:
    Coin(std::string name, const double cost, const bool is_present = true) : _name(std::move(name)),
        _cost(cost),
        _is_present(is_present) {
        invariant();
    }

    /**
     * Функция, возвращающая название монеты
     *
     * @return Название монеты
     */
    std::string get_name() const override {
        return _name;
    }

    /**
     * Функция, возвращающая стоимость монеты
     *
     * @return Стоимость монеты
     */
    double get_cost() const override {
        return _cost;
    }

    /**
     * Функция, проверяющая наличие монеты
     *
     * @return true, если монета присутствует, иначе false
     */
    bool is_present() const override {
        return _is_present;
    }

private:
    std::string _name;
    double _cost;
    bool _is_present;

    /**
     * Инвариант для проверки корректности стоимости монеты
     */
    void invariant() const {
        assert(_cost > 0);
    }
};

// Класс, представляющий коллекцию монет
class CoinCollection final : public CollectionInterface {
public:
    CoinCollection() = delete;

    explicit CoinCollection(const std::vector<Coin>& items) : _coins(items) {}

    /**
     * Функция, возвращающая общую стоимость украденных монет
     *
     * @return Общая стоимость украденных монет
     */
    double get_stolen_items_cost_amount() const override {
        double total_cost = 0;
        for (const auto& coin : _coins) {
            if (!coin.is_present()) {
                total_cost += coin.get_cost();
            }
        }
        return total_cost;
    }

    /**
     * Функция, возвращающая украденные монеты
     *
     * @return Вектор украденных монет
     */
    std::vector<std::shared_ptr<ItemInterface>> get_stolen_items() const override {
        std::vector<std::shared_ptr<ItemInterface>> result;

        for (const auto& coin : _coins) {
            if (!coin.is_present()) {
                result.push_back(std::make_shared<Coin>(coin));
            }
        }

        return result;
    }

private:
    std::vector<Coin> _coins;
};

// Класс для управления отчетами
class ReportManager final : public ReportInterface {
public:
    ReportManager() = default;

    explicit ReportManager(const CollectionInterface* collection) : _cost(collection->get_stolen_items_cost_amount()) {}

    /**
     * Функция, создающая отчет по коллекции и выводящая его в поток
     *
     * @param out Поток вывода
     * @param collection Коллекция предметов
     */
    void make_report(std::ostream& out, const CollectionInterface* collection) const override {
        out << "--------------------------------\n";
        out << "Collection REPORT" << std::endl;
        out << std::fixed << std::setprecision(2) << make_report(collection) << std::endl;
        out << std::endl;
        out << "stolen:" << std::endl;
        for (const auto& item : collection->get_stolen_items()) {
            out << item->get_name() << " " << std::fixed << std::setprecision(2) << item->get_cost() <<
                std::endl;
        }
        out << "--------------------------------\n";
    }

    /**
     * Функция, создающая отчет и возвращающая сумму страховой выплаты
     *
     * @param collection Коллекция предметов
     * @return Сумма страховой выплаты
     */
    double make_report(const CollectionInterface* collection) const override {
        return collection->get_stolen_items_cost_amount();
    }

private:
    double _cost{};
};

int main() {
    const std::vector<Coin> items = {
        {"coin1", 100, false},
        {"coin7", 100, true},
        {"coin2", 200},
        {"coin3", 123.4, false},
        {"coin4", 21},
        {"coin5", 10009, false},
        {"coin6", 2323}
    };

    const std::vector<Coin> items2 = {
        {"coin8", 500, false},
        {"coin9", 150, false},
        {"coin10", 300},
        {"coin11", 234.5, false},
        {"coin12", 45},
        {"coin13", 10050.1, false},
        {"coin14", 2500.99, false}
    };

    const std::vector<Coin> items3 = {
        {"coin15", 800, true},
        {"coin16", 200, true},
        {"coin17", 400},
        {"coin18", 345.6, false},
        {"coin19", 67},
        {"coin20", 10080, false},
        {"coin21", 2800}
    };

    const std::vector<Coin> items4 = {
        {"coin22", 1200, false},
        {"coin23", 250},
        {"coin24", 600},
        {"coin25", 456.7, false},
        {"coin26", 89},
        {"coin27", 10110, false},
        {"coin28", 3100, false}
    };

    const CoinCollection collection(items);
    const CoinCollection collection2(items2);
    const CoinCollection collection3(items3);
    const CoinCollection collection4(items4);

    const ReportManager report_manager;
    report_manager.make_report(std::cout, &collection);
    report_manager.make_report(std::cout, &collection2);
    report_manager.make_report(std::cout, &collection3);
    report_manager.make_report(std::cout, &collection4);

    return 0;
}
