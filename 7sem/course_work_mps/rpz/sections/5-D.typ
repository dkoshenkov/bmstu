#import "@local/gost732-2017:0.4.2": *
#import "@local/bmstu:0.3.0": *


#show: приложение.with(буква: "Г", содержание: [ Исходный текст программы ])

#let files = (
  "main.c",
)

#for file in files {
  листинг(raw(read("../code/" + file)))[ Содержимое файла #file ] 
}