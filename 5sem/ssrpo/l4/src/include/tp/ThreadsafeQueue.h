/**
 * @file ThreadsafeQueue.h
 * @author Michael Fetisov (fetisov.michael@bmstu.ru)
 * @brief 
 * @version 0.1
 * @date 2022-08-26
 * 
 * @copyright Copyright (c) 2022
 * 
 * Anthony Williams. C++ Concurrency in Action. Second Edition
 * 
 */

#ifndef threadsafe_queue_H
#define threadsafe_queue_H

#include <queue>
#include <mutex>
#include <condition_variable>

namespace tp
{

/**
 * @brief Потокобезопасная очередь
 * 
 * @tparam T Тип данных, который планируется хранить в очереди
 */
template <typename T>
class ThreadsafeQueue
{
    std::queue<T>           _queue;
    mutable std::mutex      _queue_mutex;
    std::condition_variable _queue_waiting_condition;

public:

    /**
     * @brief Конструктор по умолчанию
     * 
     */
    ThreadsafeQueue() = default;

    /**
     * @brief Конструктор копирования
     * 
     * @param other Копируемый объект очереди
     * 
     */
    ThreadsafeQueue(ThreadsafeQueue const& other)
    {
        std::lock_guard<std::mutex> locker(other._queue_mutex);
        _queue = other._queue;
    }

    /**
     * @brief Метод размещения объекта в очереди
     * 
     * @param new_value Объект для размещения
     * 
     */
    void push(T new_value)
    {
        std::lock_guard<std::mutex> locker(_queue_mutex);
        _queue.push(std::move(new_value));
        _queue_waiting_condition.notify_one();
    }

    /**
     * @brief Получение очередного объекта с ожиданием
     * 
     * @param value Ссылка на участок памяти, в который будет помещён объект из очереди
     * 
     * @details Данный метод будет бесконечно ожидать, пока в очередь не появится объект.
     * 
     * @attention Необходимо иметь в виду возможное "зависание" данного метода, 
     * если по ходу выполнения программы не предусмотрено размещение объектов в очереди, 
     * но данный метод вызван.
     * 
     */
    void wait_and_pop(T& value)
    {
        std::unique_lock<std::mutex> locker(_queue_mutex);
        _queue_waiting_condition.wait(locker,[this]{return !_queue.empty();});
        value = std::move(_queue.front());
        _queue.pop();
    }

    /**
     * @brief Получение очередного объекта без ожидания
     * 
     * @param value Ссылка на участок памяти, в который будет помещён объект из очереди
     * 
     * @return true В очереди был хотя бы один объект, который размещается в указанной памяти.
     * @return false Очередь оказалось пустой.
     * 
     */
    bool try_pop(T& value)
    {
        std::lock_guard<std::mutex> locker(_queue_mutex);
    
        if (_queue.empty())
            return false;
    
        value = std::move(_queue.front());
        _queue.pop();
        return true;
    }

    /**
     * @brief Возвращает признак пустой очереди на момент вызова
     * 
     * @return true Очередь пуста.
     * @return false В очереди есть объекты.
     * 
     * @attention Необходимо учитывать, что использование данного метода не гарантирует 
     * определённое им состояние очереди на момент использования полученного данным методом результата.
     * 
     */
    bool empty() const
    {
        std::lock_guard<std::mutex> locker(_queue_mutex);
        return _queue.empty();
    }

    /**
     * @brief Возвращает количество объектов в очереди на момент вызова
     * 
     * @return size_t Количество объектов в очереди
     * 
     * @attention Необходимо учитывать, что использование данного метода не гарантирует 
     * определённое им состояние очереди на момент использования полученного данным методом результата.
     * 
     */
    size_t size() const
    {
        std::lock_guard<std::mutex> locker(_queue_mutex);
        return _queue.size();
    }
};

}

#endif
