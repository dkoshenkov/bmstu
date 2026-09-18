#include <avr/io.h>
#include <avr/interrupt.h>

#define BUTTON_SHOW 0

const unsigned int ubrrValue = 23;

#define DATA_LENGTH 3
unsigned char data[DATA_LENGTH] = { 0 };

uint8_t receivedBytes = 0;

ISR(USART_RX_vect) {
   if (receivedBytes < DATA_LENGTH) {
      data[receivedBytes++] = UDR;
   }
}

int main() {
   UBRRH = (unsigned char)(ubrrValue>>8);
   UBRRL = (unsigned char)ubrrValue;

   UCSRB = (1<<RXEN) | (1<<RXCIE);

   UCSRC = (1<<URSEL) | (3<<UCSZ0);

   PORTB = (1<<BUTTON_SHOW);

   DDRC = 0xFF;

   PORTC = 0xFF;

   sei();
   uint8_t i = 0; 

   while (1) {
      if (!(PINB & (1<<BUTTON_SHOW))) {
	 while (!(PINB & (1<<BUTTON_SHOW)))
	 ;

	 PORTC = ~data[i];
	 i = (i + 1) % DATA_LENGTH;
      }
   }
   
   return 0;
}

