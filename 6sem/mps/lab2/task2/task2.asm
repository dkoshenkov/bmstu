.include "m8515def.inc" ;файл определений для ATmega8515
.def temp = r16 ;временный регистр
.equ led = 0 ;0-о бит порта PB
.equ sw0 = 2 ;2-й бит порта PD
.equ sw1 = 3 ;3-й бит порта PD
.org $000
	;***Таблица векторов прерываний, начиная с адреса $000***
	rjmp INIT ;обработка сброса
	rjmp led_on1 ;обработка запроса INT0
	rjmp led_on2 ;обработка запроса INT1


;***Инициализация SP, портов, регистра маски***
INIT:
	ldi temp,high(RAMEND) ;Установка
	out SPH,temp ; указателя стека
	ldi temp,low(RAMEND) ; на последнюю
	out SPL,temp ; ячейку ОП

	ser temp ;инициализация выводов
	out DDRB,temp ; порта PB на вывод
	out PORTB,temp ;погасить светодиод

	clr temp ;инициализация
	out DDRD,temp ; порта PD на ввод
	ldi temp,0b00001100 ;включение подтягивающих
	out PORTD,temp ; резисторов порта PD

	ldi temp,((1<<INT0)|(1<<INT1))
	out GICR,temp ; разрешение прерываний в регистре маски GICR

	ldi temp,0 ;установка обработки прерываний
	out MCUCR,temp ; по низкому уровню

	sei ;глобальное разрешение прерываний

loop:
	nop ;режим ожидания
	rjmp loop

led_on1: ;обработчик прерывания INT0 (вкл. на 1 с)
	cbi PORTB,led

	rcall delay1
	sbi PORTB,led

wait_0:
	sbis pind,sw0 ;ожидание отпускания кнопки
	rjmp wait_0
	reti ;выход из прерывания, установка I=1

led_on2: ;обработчик прерывания INT1 (вкл. на 2 с)
	cbi PORTB,led
	rcall delay2 ; *-- здесь будет 3 адреса
	sbi PORTB,led

wait_1:
	sbis pind,sw1 ;ожидание отпускания кнопки
	rjmp wait_1
	reti ;выход из прерывания, установка I=1



delay1: ; задержка 1с
	ldi r16,23
	d3:  ldi r17,236
		d1: ldi r18,245 ; loading init value to reg
			d2: dec r18 ; reg value--;    
				brne d2 ; if != 0 => transition
			dec r17
			brne d1
		dec r16
		brne d3
	ret

delay2: ; ? задержка 2с
	rcall delay1 ; после вызова на стеке будет 2 адреса возврата
	rcall delay1
	
	ret
