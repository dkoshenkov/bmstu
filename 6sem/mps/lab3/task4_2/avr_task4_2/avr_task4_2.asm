.include "m8515def.inc" ; файл определений дл€ ATmega8515

.def temp = r16 ;временный регистр

.equ G = 0 ;0-й бит порта PA
.equ D = 1 ;1-й бит порта PA
.equ A = 2 ;1-й бит порта PA
.equ F = 3 ;1-й бит порта PA


;***ћакросы дл€ настройки таймера/счетчика
.equ PRESCALER_64 = ((1<<CS01)|(1<<CS00))

.equ TIMER_SETTINGS = ((1<<WGM01)|(1<<COM00)|PRESCALER_64)

.equ VALUE_G = 146 ; частота 0,5 к√ц
.equ VALUE_D = 195 ; частота 0,5 к√ц
.equ VALUE_A = 130 ; частота 0,5 к√ц
.equ VALUE_F = 164 ; частота 0,5 к√ц


.macro timer_on
	sbi  PORTB, 0 
	ldi temp,TIMER_SETTINGS 
    out TCCR0,temp 
.endmacro 

.macro timer_off
	clr temp 
	out TCCR0,temp
.endmacro 



;***“аблица векторов прерываний

.org $000
	 rjmp INIT
;***»нициализаци€ ћ 

INIT:
	ldi temp,high(RAMEND) ;установка
	out SPH,temp ; указател€ стека
	ldi temp,low(RAMEND) ; на последнюю
	out SPL,temp ; €чейку ќѕ
	ser temp ;инициализаци€ выводов

	out DDRB,temp ; порта PB на вывод
	cbi PORTB, 0
	;out PORTB,temp ;подать высокий сигнал

	clr temp ;инициализаци€
	out DDRA,temp ; порта PA на ввод
	ldi temp,0b00001111 ;включение подт€гивающих
	out PORTA,temp ; резисторов порта PA
	

START:
	timer_off
check_G:
	sbic PINA, G
	rjmp check_D
; ***** загрузка нужного коэф дл€ ноты G 196√ц  коэф 146
	ldi temp,VALUE_G ;установка конечного
	out OCR0,temp ; значени€ счета
	timer_on

	play_G:
		sbis PINA, G
		rjmp play_G
	
	timer_off


check_D:
	sbic PINA, D
	rjmp check_A
; ***** загрузка коэф дл€ ноты D 146,8√ц  -- 195
	ldi temp,VALUE_D ;установка конечного
	out OCR0,temp ; значени€ счета
	
	timer_on
	
	play_D:
		sbis PINA, D
		rjmp play_D
	timer_off
check_A:
	sbic PINA, A
	rjmp check_F
; ***** загрузка коэф дл€ ноты A 220√ц  -- 130
	ldi temp,VALUE_A ;установка конечного
	out OCR0,temp ; значени€ счета
	

	timer_on

	play_A:
		sbis PINA, A
		rjmp play_A

	timer_off
check_F:
	sbic PINA, F
	rjmp START
	; ***** загрузка коэф дл€ ноты A 174,6√ц  -- 164
	ldi temp,VALUE_F ;установка конечного
	out OCR0,temp ; значени€ счета

	timer_on

	play_F:
	    sbis PINA, F
	    rjmp play_F
	    
	    timer_off
	    rjmp START
