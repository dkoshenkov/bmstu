.include "m8515def.inc" 
.def temp = r16 
.equ led = 0 
.equ start_1q = 0
.equ start_3q = 1 
.equ stop = 2

.equ PRESCALER_1 = (1<<CS10)
.equ TIMER_SETTINGS = ((1<<COM1A1)|(1<<COM1A0)|(1<<WGM10)|(1<<WGM11)) ; WRGM 11 is for 10-bit mode (page 121)
.equ VALUE_1Q = 256
.equ VALUE_3Q = 768


.macro pwm_on
 ldi temp,TIMER_SETTINGS
 out TCCR1A,temp
 ldi temp,PRESCALER_1
 out TCCR1B,temp
.endmacro

.macro pvm_1q
 ldi temp,high(VALUE_1Q) 
 out OCR1AH,temp
 ldi temp,low(VALUE_1Q)
 out OCR1AL,temp
.endmacro

.macro pvm_3q
 ldi temp,high(VALUE_3Q) 
 out OCR1AH,temp
 ldi temp,low(VALUE_3Q)
 out OCR1AL,temp
.endmacro

.macro pwm_off
 clr temp
 out TCCR1A,temp
 out TCCR1B,temp
.endmacro

.org $000
 rjmp INIT

INIT:
 ldi temp,high(RAMEND)
 out SPH,temp 
 ldi temp,low(RAMEND)
 out SPL,temp 
 ser temp 
 out DDRD,temp 
 out PORTD,temp 
 clr temp 
 out DDRA,temp 
 ldi temp,0b00000111 
 out PORTA,temp 

test_on_1q:
 sbic PINA,start_1q 
 rjmp test_on_3q 
 pwm_on
 pvm_1q 
wait_0:
 sbis PINA,start_1q 
 rjmp wait_0 

test_on_3q:
 sbic PINA,start_3q 
 rjmp test_off 
 pwm_on
 pvm_3q 
wait_2:
 sbis PINA,start_3q 
 rjmp wait_2

test_off:
 sbic PINA,stop 
 rjmp test_on_1q 
 pwm_off
 ser temp
 out PORTD,temp 
wait_1:
 sbis PINA,stop 
 rjmp wait_1 
 rjmp test_on_1q
