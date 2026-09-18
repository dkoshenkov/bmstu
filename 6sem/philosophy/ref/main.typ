#import "@local/gost732-2017:0.3.0": *
#import "@local/bmstu:0.1.1": *

#show: гост732-2017


#страница(image("титул.jpg", height: 100%), номер: нет)

#include "введение.typ"

#содержание()

#include "наполнение.typ"


#bibliography("liba.yml")
