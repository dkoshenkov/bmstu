#include <avr/io.h>
#include <avr/interrupt.h>
#include <stdbool.h>
#define BUTTON_START 0

const unsigned int ubrrValue = 23;

#define DATA_LENGTH 3
const unsigned char data[DATA_LENGTH] = {65, 86, 82};

uint8_t counter = 0;
bool flag = false;


ISR(USART_UDRE_vect) {
   if (flag)
      UDR = data[counter++];
}

int main() {

   UBRRH = (unsigned char)(ubrrValue>>8);
   UBRRL = (unsigned char) ubrrValue;

   UCSRB = (1<<TXEN) | (1<<UDRIE);
   UCSRC = (1<<URSEL) | (3<<UCSZ0);

   PORTB = (1<<BUTTON_START);

   sei();
   
 while (1) {
	 if (!(PINB & (1<<BUTTON_START))) {
	 	while (!(PINB & (1 << BUTTON_START)))
		;
		
		flag = true;
		while (counter < 3)
		;

		counter = 0;
		
	 } 


	 flag = false;
	 
 }
 return 0;
}
