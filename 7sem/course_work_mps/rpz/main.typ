#import "@local/gost732-2017:0.4.2": *
#import "@local/bmstu:0.3.0": *

#show: гост732-2017

#страница(image("titul.jpeg", height: 100%), номер: нет)
#страница(image("tak.png", height: 100%), номер: нет)
#содержание()

#include "sections/0-intro.typ"
#include "sections/1-design.typ"
#include "sections/2-tech.typ"
#include "sections/3-conclusion.typ"

#bibliography("bib.yaml")

#include "sections/5-A.typ"
#include "sections/5-B.typ"

#include "sections/5-C.typ"
#include "sections/5-D.typ"