.include "m8515def.inc"
.def temp = r16
.equ led = 0     
.equ sw0 = 0         ;  нопка SW0 Ц идентификаци€ на PA0
.equ sw1 = 1         ;  нопка SW1 Ц идентификаци€ на PA1


.org 0x000
    rjmp INIT           
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
    nop 
    nop 
    rjmp button_ISR      ; INT2 Ц запрос кнопок



INIT:
    ; »нициализаци€ указател€ стека
    ldi temp, high(RAMEND)
    out SPH, temp
    ldi temp, low(RAMEND)
    out SPL, temp

    ; Ќастройка порта B: LED (PB0) как выход, начальное состо€ние Ц гас (1)
    ldi temp, 1<<led
    out DDRB, temp
    ldi temp, 0xFF       ; предполагаем, что LED гаснет при установке лог. 1
    out PORTB, temp

    ; Ќастройка порта A дл€ идентификации кнопок:
    ; PA0 Ц SW0, PA1 Ц SW1, входы с включЄнными внутренними подт€жками
    clr temp
    out DDRA, temp       ; порт A как вход
    ldi temp, 0b00000011 ; подт€жки на PA0 и PA1
    out PORTA, temp

    ; Ќастройка порта E дл€ INT2 (например, INT2 на PE0)
    clr temp
    out DDRE, temp       ; порт E как вход
    ldi temp, (1<<PE0)   ; включение внутренней подт€жки на PE0
    out PORTE, temp

    ; –азрешение прерывани€ INT2
    ldi temp, (1<<INT2)
    out GICR, temp

    ; Ќастройка INT2 на срабатывание по низкому уровню (активный при замыкании на землю)
    clr temp
    out EMCUCR, temp

    ; √лобальное разрешение прерываний
    sei

loop:
    rjmp loop


; ќбработчик прерывани€ INT2 (объединЄнный дл€ двух кнопок)
button_ISR:
    in  temp, PINA
    sbic PINA, sw0       ; если бит PA0 очищен (0), кнопка SW0 нажата
    rjmp check_sw1       ; если PA0 не опущен (бит = 1), проверить PA1
    rjmp button_sw0      ; если PA0 нажата, перейти к обработке SW0

check_sw1:
    sbic PINA, sw1       ; если PA1 = 0, кнопка SW1 нажата
    rjmp wait_release    ; если ни одна кнопка не нажата, выходим
    rjmp button_sw1      ; если PA1 нажата, перейти к обработке SW1

button_sw0:
    ; ќбработка нажати€ кнопки SW0: LED включЄн на 1 с
    cbi PORTB, led       ; включить LED (активное состо€ние Ц 0)
    rcall delay1
    sbi PORTB, led       ; выключить LED (возврат в состо€ние 1)
    rjmp wait_release

button_sw1:
    cbi PORTB, led
    rcall delay2
    sbi PORTB, led
    rjmp wait_release

wait_release:
wait_rel_loop:
    in temp, PINA
    sbis PINA, sw0       ; если PA0 = 1 (отпущена), продолжаем проверку
    rjmp wait_rel_loop   ; иначе Ц ждЄм
    sbis PINA, sw1
    rjmp wait_rel_loop
    reti                 ; выход из прерывани€

; задержки
delay1:
    ldi r16, 23
d3:
    ldi r17, 236
d1:
    ldi r18, 245
d2:
    dec r18
    brne d2
    dec r17
    brne d1
    dec r16
    brne d3
    ret

delay2:
    rcall delay1
    rcall delay1
    ret
