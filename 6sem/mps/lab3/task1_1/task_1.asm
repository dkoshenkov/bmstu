.include "m8515def.inc" 
.def temp = r16 

.org $000
	rjmp INIT 
.org $007
	rjmp T0_OVF

INIT:
	ldi temp,low(RAMEND) 
	out SPL,temp 
	ldi temp,high(RAMEND) 
	out SPH,temp 
	clr temp 
	out DDRB,temp
	ldi temp,(1<<PB0) 
	out PORTB,temp
	ser temp 
	out DDRD,temp 
	out PORTD,temp 

	ldi temp,(1<<SE) 
	out MCUCR,temp 

	ldi temp,0x02 
	out TIMSK,temp 
	ldi temp,0x07
	out TCCR0,temp
	sei 
	ldi temp,0xFC
	out TCNT0,temp 
LOOP:
	sleep 
	nop 
	rjmp LOOP

T0_OVF:
	clr temp
	out PORTD,temp 
	rcall DELAY
	ser temp
	out PORTD,temp 
	ldi temp,0xFC 
	out TCNT0,temp 
	reti

DELAY:
	ldi r19,6
	ldi r20,255
	ldi r21,255
dd:
	dec r21
	brne dd
	
	dec r20
	brne dd
	dec r19
	brne dd
	ret
