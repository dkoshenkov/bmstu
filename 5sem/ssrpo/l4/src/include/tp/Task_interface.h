/**
 * @file Task_interface.hpp
 * @author Michael Fetisov (fetisov.michael@bmstu.ru)
 * @brief 
 * @version 0.1
 * @date 2022-08-26
 * 
 * @copyright Copyright (c) 2022
 * 
 */

#ifndef task_interface_H
#define task_interface_H

namespace tp
{

/**
 * @brief Интерфейс задачи для пула потока
 * 
 * @details
 * Для работы с пулом потоков tp::ThreadPool, необходимо создать реализацию данного интерфейса, 
 * переопределив метод work. 
 * 
 * См. описание tp::ThreadPool.
 */
class Task_interface
{
public:
    virtual ~Task_interface() = default;

    /**
     * @brief Метод вызывается пулом потоков tp::ThreadPool, 
     * когда доходит очередь до выполнения переданной в пул потоков реализации интерфейса 
     * tp::Task_interface
     * 
     */
    virtual void work() = 0;
};

}

#endif
