.include "m8515def.inc"

.def reg_led = r20     
.def temp = r19        
.def cnt = r21  
.def i = r22    

.org $000
  rjmp INIT

INIT:
  ldi reg_led,0xFF
  ser temp
  out DDRB,temp 
  clr temp
  out DDRD,temp 
  ldi temp,0xFF
  out PORTD,temp
  clc
	
LOOP:
  out PORTB, reg_led  

  clr cnt       
  in temp, PORTD
  ldi i, 8
  
  LOOPER:
  	ROL temp
	brcs END

	inc cnt

	END:
	dec i	
	cpi i, 0
	brne LOOPER

  
  cpi cnt, 2
  brne LAMP_OFF        
  breq LAMP_ON        
       
  rjmp LOOP          

LAMP_OFF:
  ldi reg_led, 0b11111111  
  rjmp LOOP

LAMP_ON:
  ldi reg_led, 0b11101111 
  rjmp LOOP
