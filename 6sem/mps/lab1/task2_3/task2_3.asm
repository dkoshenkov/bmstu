.include "m8515def.inc"
.def reg_led = r20      ; Состояние светодиодов
.def temp = r19
.def direction = r22    ; Направление движения: 0-вправо, 1-влево

.equ START = 0
.equ STOP = 1

.org $000
  rjmp INIT

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
  ldi r16,9
d3: ldi r17,202
d1: ldi r18,221
d2: dec r18
  brne d2
  dec r17
  brne d1
  dec r16
  brne d3
  ret
