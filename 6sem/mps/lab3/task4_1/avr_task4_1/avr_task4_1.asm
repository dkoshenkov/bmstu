.include "m8515def.inc" ; файл определений для ATmega8515
.def temp = r16 ;временный регистр
.equ START = 0 ;0-й бит порта PA
.equ STOP = 1 ;1-й бит порта PA
;***Макросы для настройки таймера/счетчика
.equ PRESCALER_64 = ((1<<CS01)|(1<<CS00))
.equ TIMER_SETTINGS = ((1<<WGM01)|(1<<COM00)|PRESCALER_64)
.equ VALUE = 109 ; частота 0,5 кГц
;***Таблица векторов прерываний
.org $000
 rjmp INIT
;***Инициализация МК
INIT:
	ldi temp,high(RAMEND) ;установка
	out SPH,temp ; указателя стека
	ldi temp,low(RAMEND) ; на последнюю
	out SPL,temp ; ячейку ОП
	ser temp ;инициализация выводов
	out DDRB,temp ; порта PB на вывод
	out PORTB,temp ;подать высокий сигнал
	clr temp ;инициализация
	out DDRA,temp ; порта PA на ввод
	ldi temp,0b00000011 ;включение подтягивающих
	out PORTA,temp ; резисторов порта PA
	ldi temp,VALUE ;установка конечного
	out OCR0,temp ; значения счета
test_start:
	sbic PINA,START ;проверка состояния
	rjmp test_stop ; кнопки START
	ldi temp,TIMER_SETTINGS
	out TCCR0,temp ; включение таймера
wait_0:
	sbis PINA,START ;проверка отпускания
	rjmp wait_0 ; кнопки
test_stop:
	sbic PINA,STOP ;проверка состояния
	rjmp test_start ; кнопки stop
	clr temp
	out TCCR0,temp ;выключение таймера
wait_1:
	sbis PINA,STOP ;проверка отпускания
	rjmp wait_1 ; кнопки
	rjmp test_start
