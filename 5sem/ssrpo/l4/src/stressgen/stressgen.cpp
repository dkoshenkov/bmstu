/**
 * @file stressgen.cpp
 * @author Fetisov Michael (fetisov.michael@bmstu.com)
 * @brief Программа командной строки для генерации набора команд для стресс-теста
 * лабораторных работ к курсу СТРПО.
 * @version 0.1
 * @date 2022-11-30
 *
 * @copyright Copyright (c) 2022
 *
 * На вход (стандартный поток std::cin) программа принимает шаблон,
 * состоящий из строк, которые просто копируются в результирующий поток (стандартный поток std::cout),
 * и групп генерации.
 *
 * Группы генерации начинаются со строки "#n" и заканчиваются строкой "#",
 * где n определяет количество повторений группы.
 *
 * Возможны вложенные группы.
 *
 * Управляющий символ '#' может быть заменён на другой посредством параметра --control-character.
 *
 * В строках команд можно использовать следующие макрозамены:
 *
 * * #i - заменяется на номер повторения строки внутри группы
 * * #N - заменяется на случайное целое число из диапазона от 1 до 1000 включительно
 * * #W - заменяется на случайный набор латинских букв (строчных и прописных)
 * в количестве от 1 до 10 (включительно).
 *
 * Макрозамены можно использовать в управляющей строки групп генерации.
 *
 */

#include <iostream>
#include <string>
#include <sstream>
#include <vector>
#include <memory>
#include <cstdlib>
#include <cassert>

class Generator_interface {
public:
    virtual ~Generator_interface() = default;

    virtual std::string generate() const = 0;
};

class LineNoGenerator : public Generator_interface {
    std::string _line_no;

public:
    LineNoGenerator(int line_no) : _line_no(std::to_string(line_no)) {
    }

    virtual std::string generate() const override { return _line_no; }
};

class NumberGenerator : public Generator_interface {
public:
    virtual std::string generate() const override {
        return std::to_string(std::rand() % 1000 + 1);
    }
};

class WordGenerator : public Generator_interface {
public:
    virtual std::string generate() const override {
        const std::string letters = "QWERTYUIOPASDFGHJKLZXCVBNMqwertyuiopasdfghjklzxcvbnm";
        std::string str;
        int word_length = std::rand() % 10 + 1;

        for (int i = 0; i < word_length; ++i)
            str += letters.substr(std::rand() % letters.size(), 1);

        return str;
    }
};

void replaceAll(std::string &str, const std::string &from, const Generator_interface &generator) {
    if (from.empty())
        return;

    size_t start_pos = 0;
    while ((start_pos = str.find(from, start_pos)) != std::string::npos) {
        std::string to = generator.generate();
        str.replace(start_pos, from.length(), to);
        start_pos += to.length();
    }
}

std::string generateLine(std::string str, int line_no, const std::string &control_character) {
    std::vector<std::pair<std::string, std::shared_ptr<Generator_interface> > > replaces
    {
        {control_character + "i", std::make_shared<LineNoGenerator>(line_no)},
        {control_character + "N", std::make_shared<NumberGenerator>()},
        {control_character + "W", std::make_shared<WordGenerator>()},
    };

    for (auto &[from,gen]: replaces)
        replaceAll(str, from, *gen);

    return str;
}

std::vector<std::string> processGroup(std::string command_line, const std::string &control_character,
                                      long long default_repetitions) {
    long long repetitions = default_repetitions;
    if (command_line.size() > control_character.size())
        repetitions = std::stoll(generateLine(command_line, 0, control_character).substr(1));

    std::vector<std::string> group;

    for (std::string line; std::getline(std::cin, line);) {
        if (line.substr(0, control_character.size()) == control_character) {
            if (line.size() == control_character.size())
                break;

            std::vector<std::string> processed_group = processGroup(line, control_character, default_repetitions);

            for (std::string s: processed_group)
                group.push_back(s);

            continue;
        }
        group.push_back(line);
    }

    std::vector<std::string> processed_group;

    for (long long i = 0; i < repetitions; ++i)
        for (std::string s: group)
            processed_group.push_back(generateLine(s, i, control_character));

    return processed_group;
}

int main(int argc, char *argv[]) {
    std::string control_character = "#";
    long long default_repetitions = 100;
    bool parameters_error = false;

    std::vector<std::string> arguments(argv + 1, argv + argc);

    if (arguments.empty());
    else if (arguments.size() != 2)
        parameters_error = true;
    else if (arguments[0] != "--control-character")
        parameters_error = true;
    else
        control_character = arguments[1][0];

    if (parameters_error) {
        std::cerr << "Error in the program launch line" << std::endl;
        return 1;
    }

    for (std::string line; std::getline(std::cin, line);) {
        if (line.substr(0, control_character.size()) == control_character) {
            std::vector<std::string> processed_group = processGroup(line, control_character, default_repetitions);

            for (std::string s: processed_group)
                std::cout << s << std::endl;

            continue;
        }

        std::cout << line << std::endl;
    }

    std::cout << std::endl;

    return 0;
}
