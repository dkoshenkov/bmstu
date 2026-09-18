.include "m8515def.inc"
.def reg_led = r20
.def temp = r19 
.equ START = 0 ;
.equ STOP = 1 ;
.org $000
  rjmp INIT

INIT:
  ldi reg_led,0xFE 
  ser temp
  out DDRB,temp 
  sec ;C=1
  clr temp
  out DDRD,temp 
  ldi temp,0x03
  out PORTD,temp 

WAITSTART:
  sbic PIND,START 
  rjmp WAITSTART 

LOOP:
  out PORTB,reg_led ;

  ldi r16,12
d3:  ldi r17,230
d1: ldi r18,245 ; loading init value to reg
d2: dec r18 ; reg value--;    
  brne d2 ; if != 0 => transition
  dec r17
  brne d1
  dec r16
  brne d3
  sbic PIND,STOP    ;
  rjmp CONTINUE ;
  rjmp WAITSTART ;
CONTINUE:
  rol reg_led ;
  rjmp LOOP ;
