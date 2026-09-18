.include "m8515def.inc"

.def temp = r16

.equ led = 0    ; светодиод на PB0
.equ btn0 = 0   ; кнопка идентификации нажатия на PA0
.equ btn1 = 1   ; кнопка идентификации нажатия на PA1

.org $000
    rjmp INIT        ; RESET
    reti             ; INT0 не используется
    reti           ; INT1 не используется
    rjmp INT2_ISR   ; INT2 (PE0) общий обработчик прерывания

INIT:
    ldi temp, high(RAMEND)
    out SPH, temp
    ldi temp, low(RAMEND)
    out SPL, temp

    ser temp
    out DDRB, temp    ; Порт B - выход (светодиод)

    clr temp
    out DDRE, temp    ; Порт E - вход INT2
    out DDRA, temp    ; Порт A - входы (идентификация кнопок)

    ; подтяжка PA0, PA1 и INT2 (PE0)
    ldi temp, (1<<PE0)
    out PORTE, temp
    ldi temp, (1<<btn0)|(1<<btn1)
    out PORTA, temp

    ; Разрешаем внешнее прерывание INT2
    ldi temp, (1<<INT2)
    out GICR, temp

    ; INT2 по низкому уровню
    clr temp
    out EMCUCR, temp

    sei                ; глобальное разрешение прерываний

loop:
    rjmp loop          ; Бесконечный цикл ожидания прерываний

; Обработчик INT2
INT2_ISR:
    sbic PINA, btn0
    rjmp check_btn1

    ; обработка кнопки на PA0 (1 секунда)
    cbi PORTB, led
    rcall delay_1s
    sbi PORTB, led

wait_release0:
    sbis PINA, btn0
    rjmp wait_release0
    reti

check_btn1:
    sbic PINA, btn1
    reti ; если ни одна кнопка не нажата – ложное срабатывание

    ; обработка кнопки на PA1 (2 секунды)
    cbi PORTB, led
    rcall delay_2s
    sbi PORTB, led

wait_release1:
    sbis PINA, btn1
    rjmp wait_release1
    reti

; Подпрограмма задержки 1 секунда (на частоте 8МГц)
delay_1s:
    ldi r17, 16
dloop1:
    ldi r18, 244
d_loop2:
    ldi r19, 250
d_loop:
    dec r19
    brne d1
    dec r18
    brne d2
    dec r17
    brne dloop1
    ret

d1: nop
    nop
    nop
    nop
    rjmp d2
d2: nop
    rjmp d_loop
d_loop: rjmp d_loop

; Проще использовать точную задержку, например:
delay_1s:
    ldi r20, 31
L1: ldi r21, 255
L2: ldi r22, 255
L3: dec r22
    brne L2
    dec r21
    brne L2
    dec r20
    brne delay_1s
    ret

; Задержка 2 секунды
delay_2s:
    rcall delay_1s
    rcall delay_1s
    ret

