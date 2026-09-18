.include "m8515def.inc"
.def reg_led = r20      ; Состояние светодиодов
.def temp = r19
.def direction = r22    ; Направление движения: 0-вправо, 1-влево

.def temp1 = r16 ;????????? ??????? 1
.def temp2 = r17 ;????????? ??????? 2

.equ START = 0
.equ STOP = 1

.macro reload_timer ;??????
	 ldi temp1,high(48239) ; ??? ???????? ?
	 ldi temp2,low(48239) ; ??????
	 out TCNT1H,temp1 ; ????????? ????????
	 out TCNT1L,temp2 ; ??? ?????? ???????
.endmacro

.org $000
  	rjmp INIT

.org $0006
	 rjmp T1_OVF ;

INIT:
	ldi temp, high(RAMEND)
	out SPH, temp
	ldi temp, low(RAMEND)
	out SPL, temp

	; Инициализация портов
	ldi reg_led, 0b00000011 ; Начинаем с двух включенных светодиодов (биты 0 и 1)
	ser temp
	out DDRB, temp          ; Порт B на выход

	clr temp
	out DDRD, temp          ; Порт D на вход
	ldi temp, 0x03
	out PORTD, temp         ; Подтяжка на кнопки

	ldi direction, 0        ; Начальное направление – вправо

	ldi temp1,(1<<TOIE1) ;?????????? ??????????
	out TIMSK,temp1 ; ??????? ?? ????????????

	ser temp
	out PIND, temp

	reload_timer
	ldi temp1, ((1<<CS11)|(1<<CS10))  ; выбор делителя 64
	out TCCR1B, temp1
	
	sei
	

WAITSTART:
 	sbic PIND, START        ; Ждем нажатия кнопки START
  	rjmp WAITSTART 

LOOP:
	com reg_led             ; Инвертируем состояние светодиодов перед выводом!

	out PORTB, reg_led      ; Выводим состояние на порт B
	com reg_led             ; Возвращаем обратно для правильного расчета

	sbis PIND, STOP         ; Проверка кнопки STOP
	rjmp WAITSTART 
	rcall DELAY_MS          ; Задержка 300 мс

	; Проверка направления
	tst direction
	breq MOVE_RIGHT
	rjmp MOVE_LEFT

MOVE_RIGHT:
	lsl reg_led             ; Сдвигаем оба бита влево
	cpi reg_led, 0b11000000 ; Проверяем, достигли ли позиции (6,7)
	brne LOOP
	ldi direction, 1        ; Меняем направление на лево

	sbis PIND, STOP         ; Проверка кнопки STOP
	rjmp WAITSTART 
	rjmp LOOP

MOVE_LEFT:
	lsr reg_led             ; Сдвигаем оба бита вправо
	cpi reg_led, 0b00000011 ; Проверяем, достигли ли позиции (0,1)
	brne LOOP
	clr direction           ; Меняем направление на право

	sbis PIND, STOP         ; Проверка кнопки STOP
	rjmp WAITSTART 
	rjmp LOOP

; Процедура задержки 300 мс
DELAY_MS:
	clt

WAIT_END:
	brtc WAIT_END

	ret

; обработчик прерывания переполнения таймера
T1_OVF:

	reload_timer
	ldi temp1, ((1<<CS11)|(1<<CS10))  ; выбор делителя 64
	out TCCR1B, temp1

	set

	reti
