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

  ser temp
  out DDRB,temp 

  clr temp
  out DDRE, temp
  ldi temp, 0x01
  out PORTE, temp

  ser temp 
  out DDRD,temp 
  out PORTD,temp 

  
  ldi temp,0x02 
  out TIMSK,temp 

  ldi temp,0x07 
  out TCCR0,temp 

  sei
  ldi temp,0xFC 
  out TCNT0,temp 


LOOP:
  sbic PINE, PE0
  rjmp LOOP
  rcall SHORT_DELAY

LOGIC:
  sbis PINE, PE0
  rjmp LOGIC
  rcall SHORT_DELAY

  cbi PORTB,0
  sbi PORTB,0

 
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

SHORT_DELAY:
  ldi r20,50
  ldi r21,50

inner_short_delay:
  dec r21
  brne inner_short_delay
  
  dec r20
  brne inner_short_delay
  
  
  ret

