/**
 * @file ThreadPool.cpp
 * @author Michael Fetisov (fetisov.michael@bmstu.ru)
 * @brief 
 * @version 0.1
 * @date 2022-08-26
 * 
 * @copyright Copyright (c) 2022
 * 
 * Модифицированная версия пула потоков из книги 
 * Anthony Williams. C++ Concurrency in Action. Second Edition
 * 
 * Оптимизация работы (предотвращение нагрузки процессора) пока очередь пуста
 * с использованием std::condition_variable.
 * 
 */

#include "tp/ThreadPool.h"

#include <cassert>

using namespace tp;

ThreadPool::ThreadPool(int number_of_threads)
    : _number_of_threads(number_of_threads)
      , _necessary_to_stop(false) {
    if (_number_of_threads < 0)
        _number_of_threads = std::thread::hardware_concurrency();
}

ThreadPool::~ThreadPool() {
    _necessary_to_stop = true;
    _waiting_condition.notify_all();

    for (size_t i = 0; i < _threads.size(); ++i)
        _threads[i].join();
}

void ThreadPool::submit(Task_interface *task) {
    if (_number_of_threads == 0) {
        task->work();
        delete task;
    } else {
        _task_queue.push(task);
        _waiting_condition.notify_one();
    }
}

void ThreadPool::start()
try {
    if (_number_of_threads == 0)
        return;

    _threads.reserve(_number_of_threads);

    for (int i = 0; i < _number_of_threads; ++i)
        _threads.push_back(std::thread(&ThreadPool::worker, this));

    _waiting_condition.notify_all();
} catch (...) {
    _necessary_to_stop = true;
    _waiting_condition.notify_all();
    throw;
}

void ThreadPool::worker() {
    while (!_necessary_to_stop || !_task_queue.empty()) {
        if (_task_queue.empty()) {
            std::unique_lock<std::mutex> locker(_waiting_mutex);
            _waiting_condition.wait(locker, [this] { return !_task_queue.empty() || _necessary_to_stop; });
        }

        Task_interface *task = nullptr;

        if (_task_queue.try_pop(task)) {
            assert(task);
            task->work();
            delete task;
        }
    }
}
