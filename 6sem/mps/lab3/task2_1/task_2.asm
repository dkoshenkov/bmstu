.include "m8515def.inc" 
.def temp1 = r16 
.def temp2 = r17 
.equ led = 0 
.equ start = 0 
.equ stop = 1

.macro reload_timer 
	ldi temp1,high(7880) 
	ldi temp2,low(7880)
	out TCNT1H,temp1 
	out TCNT1L,temp2
.endmacro

.org $0000
	rjmp INIT 
.org $0006
	rjmp T1_OVF 

INIT:
	ldi temp1,high(RAMEND)
	out SPH,temp1 
	ldi temp1,low(RAMEND) 
	out SPL,temp1 
	ser temp1
	out DDRB,temp1 
	out PORTB,temp1 
	clr temp1
	out DDRA,temp1 
	ldi temp1,0b00000011 
	out PORTA,temp1
	reload_timer 
	ldi temp1,(1<<TOIE1) 
	out TIMSK,temp1 

	set 
	sei 
test_start:
	sbic PINA,start
	rjmp test_stop
	ldi temp1,((1<<CS11)|(1<<CS10))
	out TCCR1B,temp1 
wait_0:
	sbis PINA,start 
	rjmp wait_0
test_stop:
	sbic PINA,stop
	rjmp test_start
	clr temp1
	out TCCR1B,temp1 
wait_1:
	sbis PINA,stop
	rjmp wait_1
	rjmp test_start

T1_OVF:
	clr temp1
	out TCCR1B,temp1
	brts switch_on 
	sbi PORTB,led
	set
	rjmp set_timer
switch_on:
	cbi PORTB,led 
	clt
set_timer:
	reload_timer 
	ldi temp1,((1<<CS11)|(1<<CS10))
	out TCCR1B,temp1
	reti
