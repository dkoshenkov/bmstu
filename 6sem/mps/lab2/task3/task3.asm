.include "m8515def.inc"

.def reg_led   = r20
.def temp      = r19
.def direction = r22
.def job       = r21          

.equ START = 0
.equ STOP  = 1


.org 0x000              
    rjmp INIT           ;  RESET
    nop     
    rjmp START_ISR      ;  INT1 – запуск  кнопка PIND3
    nop    
    nop    
    nop     
    nop
    nop
    nop
    nop
    nop
    nop
    nop
    rjmp STOP_ISR       ; INT2 – остановка кнопка PINE0


INIT:
    ; Инициализация указателя стека
    ldi temp, high(RAMEND)
    out SPH, temp
    ldi temp, low(RAMEND)
    out SPL, temp
   
    ; Настраиваем порт B как выход (светодиоды)
    ser temp
    out DDRB, temp  

    ; Настраиваем порт D как вход (INT1 на PD3)
    clr temp
    out DDRD, temp  

    ; Настраиваем порт E как вход (INT2 на PE0) и включаем подтяжку
    clr temp
    out DDRE, temp      
    ldi temp, (1<<PE0)

    ; Разрешаем внешние прерывания INT1 и INT2
    ldi temp, (1<<INT1)|(1<<INT2)
    out GICR, temp

    ; Конфигурируем INT1 и INT2 на срабатывание по низкому уровню
    clr temp
    out MCUCR, temp     ; INT1 (PD3) – по низкому уровню
    out EMCUCR, temp    ; INT2 (PE0) – по низкому уровню

	ser temp
	out PIND, temp
	out PINE, temp

    sei                 ; Глобальное разрешение прерываний    
    
    ; Инициализация переменных для переключения светодиодов
    ldi reg_led, 0b00000011
    ldi direction, 0    
    clr job              

;******************** ОСНОВНОЙ ЦИКЛ *************************
LOOP:
    sbrs job, 0         ; если бит 0 в job установлен (то есть, разрешён запуск),
    rjmp LOOP           ; иначе остаёмся в цикле ожидания

    ; Переключение светодиодов по заданному алгоритму
    com reg_led
    out PORTB, reg_led
    com reg_led          

    rcall DELAY_MS        

    tst direction
    breq MOVE_RIGHT
    rjmp MOVE_LEFT

MOVE_RIGHT:
    lsl reg_led
    cpi reg_led, 0b11000000
    brne LOOP
    ldi direction, 1
    rjmp LOOP

MOVE_LEFT:
    lsr reg_led
    cpi reg_led, 0b00000011
    brne LOOP
    clr direction
    rjmp LOOP

;******************** ОБРАБОТЧИКИ ПРЕРЫВАНИЙ ******************

; Обработчик INT1 – запуск (кнопка на PD3)
START_ISR:
    ldi job, 1
wait_start_release:
	sbis PIND, PD3
	rjmp wait_start_release	
    reti

; Обработчик INT2 – остановка (кнопка на PE0)
STOP_ISR:
    clr job
wait_stop_release:
    sbis PINE, PE0      ; ждём отпускания кнопки на PE0
    rjmp wait_stop_release
    reti

;********************** ЗАДЕРЖКА ******************************
DELAY_MS:
    ldi r16,9
del3:
    ldi r17,202
del2:
    ldi r18,221
del1:
    dec r18
    brne del1
    dec r17
    brne del2
    dec r16
    brne del3
    ret




