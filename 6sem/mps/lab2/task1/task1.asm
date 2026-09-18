.include "m8515def.inc" 

.def temp = r16 

.equ led = 0 ;0-й бит порта PB
.equ sw0 = 0 ;0-й бит порта PA
.equ sw1 = 1 ;1-й бит порта PA

.org $000
	rjmp INIT ;обработка сброса


INIT:
	ldi temp,high(RAMEND) ;установка
	out SPH,temp ; указател€ стека
	ldi temp,low(RAMEND) ; на последнюю
	out SPL,temp ; €чейку ќѕ
	ser temp ; иниц вызовов
	out DDRB,temp ; порта PB на вывод
	out PORTB,temp ; погасить LED
	clr temp ; инициализаци€
	out DDRA,temp ; порта PA на ввод
	ldi temp,0b00000011 ; включение подт€гивающих
	out PORTA,temp ; резисторов порта PA

test_sw0:
	sbic PINA,sw0 ; проверка состо€ни€
	rjmp test_sw1 ; кнопки sw0
	cbi PORTB,led ; вкл LED
	rcall delay1
	sbi PORTB,led ; выкл LED
		
wait_0:
	sbis PINA,sw0 ; проверка отпускани€
	rjmp wait_0 ; кнопки
		
test_sw1:
	sbic PINA,sw1 ; проверка состо€ни€
	rjmp test_sw0 ; кнопки sw1
	cbi PORTB,led ; вкд LED
	rcall delay2
	sbi PORTB,led ; выкл LED

wait_1:
	sbis PINA,sw1 ; проверк отпускани€
	rjmp wait_1 ; кнопки
	rjmp test_sw0

delay1: ; задержка 1с
	ldi r16,23
	d3:  ldi r17,236
		d1: ldi r18,245 ; 
			d2: dec r18 ;    
				brne d2 ; 
			dec r17
			brne d1
		dec r16
		brne d3
	ret

delay2: ; задержка 2с
	rcall delay1 ; после вызова на стеке будет 2 адреса возврата
	rcall delay1
	
	ret
