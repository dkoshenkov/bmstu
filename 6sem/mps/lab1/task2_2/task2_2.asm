.include "m8515def.inc"
.def reg_led = r20
.def temp = r19 
.def store = r21
.def direction = r22

.equ START = 0 ;
.equ STOP = 1 ;
.equ DELAY = 300 ; 300 мс задержка

.org $000
  rjmp INIT

INIT:
  ldi reg_led,0x03   ; Начальное положение двух включенных светодиодов (биты 0 и 1)
  ser temp
  out DDRB,temp      ; Порт B на выход

  clr temp
  out DDRD,temp      ; Порт D на вход
  ldi temp,0x03
  out PORTD,temp     ; Подтяжка на кнопки

  clr direction      ; Устанавливаем начальное направление движения (вправо)

WAITSTART:
  sbic PIND,START    ; Ждем нажатия кнопки START
  rjmp WAITSTART 

LOOP:
  out PORTB,reg_led  ; Вывод состояния на порт B

  rcall DELAY_MS     ; Задержка 300 мс

  sbic PIND,STOP     ; Проверка нажатия кнопки STOP
  rjmp WAITSTART 

  ; Проверяем направление движения
  tst direction      
  breq MOVE_RIGHT   ; Если direction = 0, двигаем вправо
  rjmp MOVE_LEFT    ; Если direction = 1, двигаем влево

MOVE_RIGHT:
  lsl reg_led        ; Сдвигаем на 1 бит влево
  lsl reg_led
  cpi reg_led,0xC0   ; Достигли предела (биты 6 и 7)?
  brne LOOP
  ldi direction,1    ; Меняем направление на левое
  rjmp LOOP

MOVE_LEFT:
  lsr reg_led        ; Сдвигаем на 1 бит вправо
  lsr reg_led
  cpi reg_led,0x03   ; Достигли предела (биты 0 и 1)?
  brne LOOP
  clr direction      ; Меняем направление на правое
  rjmp LOOP

; Процедура задержки 300 мс
DELAY_MS:
  ldi r16,12
d3: ldi r17,230
d1: ldi r18,245
d2: dec r18
  brne d2
  dec r17
  brne d1
  dec r16
  brne d3
  ret
