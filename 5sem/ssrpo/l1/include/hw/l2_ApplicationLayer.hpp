#pragma once

#include <string>
#include <vector>

#include "l3_DomainLayer.hpp"


typedef std::string arg;
typedef std::vector<arg> args;
typedef std::vector<std::vector<arg>> decompized_args;


class InputOutputInterface {
public:
  virtual ~InputOutputInterface() = default;

  virtual args input() const = 0;
  virtual std::ostream &output(const arg& output_data) const = 0;
};

class Application {
public:
  Application() = delete;
  Application(const Application &) = delete;
  Application &operator=(const Application &) = delete;

  Application(const InputOutputInterface &io, const ProductManager &pm) : _io(io), _pm(pm) {}

  bool run();
private:
    decompized_args split_by_spaces(const arg &input_string);
    const InputOutputInterface& _io;
    ProductManager _pm;
};
