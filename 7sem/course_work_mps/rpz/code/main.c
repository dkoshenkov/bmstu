#include <avr/io.h>
#include <avr/interrupt.h>
#include <avr/eeprom.h> 
#include <avr/pgmspace.h> 
#include <util/delay.h>
#include <stdlib.h>
#include <string.h>

// Моторы (PC0, PC1)
#define MOTOR_PORT PORTC
#define MOTOR_DDR  DDRC
#define IN1_PIN    PC0
#define IN2_PIN    PC1

// ШИМ (PB0)
#define PWM_PORT   PORTB
#define PWM_DDR    DDRB
#define PWM_PIN    PB0

// Датчик DHT (PD7)
#define DHT_PORT   PORTD
#define DHT_DDR    DDRD
#define DHT_PIN    PD7
#define DHT_PIN_IN PIND

// КНОПКИ (PORTA)
#define BTN_PORT   PORTA
#define BTN_PIN    PINA
#define BTN_DDR    DDRA
#define BTN_AUTO   PA0
#define BTN_TOGGLE PA1
#define BTN_UP     PA2  
#define BTN_DOWN   PA3 

#define INT0_PIN   PD2

// LCD ЭКРАН (PC2-PC7)
#define LCD_PORT   PORTC
#define LCD_DDR    DDRC
#define LCD_RS     PC2 
#define LCD_E      PC3

// eeprom
#define EE_MAGIC_VAL 0xAC 

uint8_t EEMEM ee_magic;
uint8_t EEMEM ee_mode_auto;
uint8_t EEMEM ee_fan_on;
uint8_t EEMEM ee_temp_high;
uint8_t EEMEM ee_temp_low;
uint8_t EEMEM ee_fan_speed; 

// переменные
volatile uint8_t fan_on = 0;       
volatile uint8_t mode_auto = 1;    
volatile uint8_t temp_c = 0;
volatile uint8_t humidity = 0;

volatile uint8_t temp_high = 23; 
volatile uint8_t temp_low = 21;  
volatile uint8_t fan_speed = 255; 

// Буфер UART
volatile char rx_buffer[20];
volatile uint8_t rx_pos = 0;
volatile uint8_t cmd_ready = 0;

volatile uint8_t update_lcd_needed = 1;

// --- ФУНКЦИИ LCD ---
void lcd_pulse(void) {
    LCD_PORT |= (1 << LCD_E);
    _delay_us(5);
    LCD_PORT &= ~(1 << LCD_E);
    _delay_us(50);
}

void lcd_byte(uint8_t val, uint8_t is_data) {
    if (is_data) LCD_PORT |= (1 << LCD_RS);
    else         LCD_PORT &= ~(1 << LCD_RS);

    uint8_t temp = (LCD_PORT & 0x0F); 
    temp |= (val & 0xF0);             
    LCD_PORT = temp;                  
    lcd_pulse();

    temp = (LCD_PORT & 0x0F);
    temp |= ((val << 4) & 0xF0);      
    LCD_PORT = temp;
    lcd_pulse();
}

void lcd_cmd(uint8_t cmd) {
    lcd_byte(cmd, 0);
    if (cmd == 0x01 || cmd == 0x02) _delay_ms(2);
}

void lcd_char(char data) { lcd_byte(data, 1); }

void lcd_str_P(const char *str) {
    while (pgm_read_byte(str)) lcd_char(pgm_read_byte(str++));
}

void lcd_str(const char *str) { while (*str) lcd_char(*str++); }

void lcd_num(int val) {
    char buf[10];
    itoa(val, buf, 10); 
    lcd_str(buf);
}

void lcd_init(void) {
    LCD_DDR |= (1 << LCD_RS) | (1 << LCD_E) | 0xF0; 
    _delay_ms(20);
    LCD_PORT &= ~(1 << LCD_RS);
    
    LCD_PORT = (LCD_PORT & 0x0F) | 0x30; lcd_pulse(); _delay_ms(5);
    LCD_PORT = (LCD_PORT & 0x0F) | 0x30; lcd_pulse(); _delay_us(200);
    LCD_PORT = (LCD_PORT & 0x0F) | 0x30; lcd_pulse(); _delay_us(200);
    LCD_PORT = (LCD_PORT & 0x0F) | 0x20; lcd_pulse(); _delay_ms(5); 
    
    lcd_cmd(0x28); 
    lcd_cmd(0x0C); 
    lcd_cmd(0x06); 
    lcd_cmd(0x01); 
}

void lcd_gotoxy(uint8_t x, uint8_t y) {
    uint8_t addr = (y == 0) ? 0x80 : 0xC0;
    lcd_cmd(addr + x);
}
void load_config(void) {
    uint8_t magic = eeprom_read_byte(&ee_magic);
    if (magic != EE_MAGIC_VAL) {
        eeprom_update_byte(&ee_mode_auto, 1); 
        eeprom_update_byte(&ee_fan_on, 0);    
        eeprom_update_byte(&ee_temp_high, 23);
        eeprom_update_byte(&ee_temp_low, 21);
        eeprom_update_byte(&ee_fan_speed, 255);
        eeprom_update_byte(&ee_magic, EE_MAGIC_VAL);
    }
    mode_auto = eeprom_read_byte(&ee_mode_auto);
    fan_on = eeprom_read_byte(&ee_fan_on);
    temp_high = eeprom_read_byte(&ee_temp_high);
    temp_low = eeprom_read_byte(&ee_temp_low);
    fan_speed = eeprom_read_byte(&ee_fan_speed);
}

// --- UART ---
void uart_init_fixed(void) {
    UBRRH = 0; UBRRL = 51; // 9600
    UCSRB = (1 << RXCIE) | (1 << RXEN) | (1 << TXEN);
    UCSRC = (1 << URSEL) | (1 << UCSZ1) | (1 << UCSZ0);
}

void uart_send_char(char c) { while (!(UCSRA & (1 << UDRE))); UDR = c; }
void uart_send_str(const char *s) { while (*s) uart_send_char(*s++); }
void uart_send_str_P(const char *s) {while (pgm_read_byte(s)) uart_send_char(pgm_read_byte(s++));}
void uart_send_num(int val) {
    char buf[10];
    itoa(val, buf, 10);
    uart_send_str(buf);
}

ISR(USART_RX_vect) {
    char data = UDR;
    if (data == '\r' || data == '\n') { 
        rx_buffer[rx_pos] = 0; cmd_ready = 1; rx_pos = 0; 
    } else { 
        if (rx_pos < 18) rx_buffer[rx_pos++] = data; 
    }
}

// --- УПРАВЛЕНИЕ МОТОРОМ ---
void set_motor(uint8_t state) {
    fan_on = state;
    eeprom_update_byte(&ee_fan_on, state);
    if (state) {
        MOTOR_PORT |= (1 << IN1_PIN); 
        MOTOR_PORT &= ~(1 << IN2_PIN);
        OCR0 = fan_speed;
    } else {
        MOTOR_PORT &= ~((1 << IN1_PIN) | (1 << IN2_PIN)); 
        OCR0 = 0;
    }
    update_lcd_needed = 1;
}

// --- ОБРАБОТЧИК ПРЕРЫВАНИЯ КНОПОК INT0 ---
void init_int0(void) {
    DDRD &= ~(1 << INT0_PIN);
    PORTD |= (1 << INT0_PIN);

    MCUCR |= (1 << ISC01);
    MCUCR &= ~(1 << ISC00);

    GICR |= (1 << INT0);
}

ISR(INT0_vect) {
    _delay_ms(30); 

    // Кнопка AUTO
    if (!(BTN_PIN & (1 << BTN_AUTO))) {
        mode_auto = 1;
        eeprom_update_byte(&ee_mode_auto, 1);
        uart_send_str_P(PSTR("INT: AUTO Mode\r\n"));
        update_lcd_needed = 1;
    }
    // Кнопка TOGGLE (Manual ON/OFF)
    else if (!(BTN_PIN & (1 << BTN_TOGGLE))) {
        mode_auto = 0; 
        eeprom_update_byte(&ee_mode_auto, 0);
        set_motor(!fan_on);
        uart_send_str_P(PSTR("INT: Toggle Man\r\n"));
        update_lcd_needed = 1;
    }
    // Кнопка UP
    else if (!(BTN_PIN & (1 << BTN_UP))) {
        int new_speed = fan_speed + 25;
        if (new_speed >= 255) new_speed = 255;
        fan_speed = (uint8_t)new_speed;
        eeprom_update_byte(&ee_fan_speed, fan_speed);
        if (fan_on) OCR0 = fan_speed;
        
        uart_send_str_P(PSTR("INT: Speed UP: ")); uart_send_num(fan_speed); uart_send_str_P(PSTR("\r\n"));
        update_lcd_needed = 1;
    }
    // Кнопка DOWN
    else if (!(BTN_PIN & (1 << BTN_DOWN))) {
        int new_speed = fan_speed - 25;
        if (new_speed <= 0) new_speed = 0;
        fan_speed = (uint8_t)new_speed;
        eeprom_update_byte(&ee_fan_speed, fan_speed);
        if (fan_on) OCR0 = fan_speed;
        
        uart_send_str_P(PSTR("INT: Speed DOWN: ")); uart_send_num(fan_speed); uart_send_str_P(PSTR("\r\n"));
        update_lcd_needed = 1;
    }

    while(!(PIND & (1 << INT0_PIN))); 
}

// --- DHT ДАТЧИК ---
uint8_t dht_read(void) {
    uint8_t bits[5]; uint8_t cnt = 7; uint8_t idx = 0;
    for(int k=0; k<5; k++) bits[k] = 0;

    DHT_DDR |= (1 << DHT_PIN); DHT_PORT &= ~(1 << DHT_PIN); 
    _delay_ms(20); DHT_PORT |= (1 << DHT_PIN); DHT_DDR &= ~(1 << DHT_PIN); 
    
    _delay_us(40); if ((DHT_PIN_IN & (1 << DHT_PIN))) return 1; 
    _delay_us(80); if (!(DHT_PIN_IN & (1 << DHT_PIN))) return 2; 
    _delay_us(80);

    for (int i = 0; i < 40; i++) {
        uint8_t timeout = 0;
        while(!(DHT_PIN_IN & (1 << DHT_PIN))) { _delay_us(1); if (++timeout > 100) return 3; }
        _delay_us(30); 
        if (DHT_PIN_IN & (1 << DHT_PIN)) {
            if (idx < 5) bits[idx] |= (1 << cnt); 
            timeout = 0;
            while(DHT_PIN_IN & (1 << DHT_PIN)) { _delay_us(1); if (++timeout > 100) return 4; }
        }
        if (cnt == 0) { cnt = 7; idx++; } else { cnt--; }
    }
    if ((uint8_t)(bits[0] + bits[1] + bits[2] + bits[3]) == bits[4]) {
        humidity = bits[0]; temp_c = bits[2]; return 0; 
    }
    return 5; 
}

void update_lcd_info(void) {
    lcd_gotoxy(0, 0);
    lcd_str_P(PSTR("T:")); lcd_num(temp_c); lcd_char(0xDF); lcd_str_P(PSTR("C H:")); lcd_num(humidity); 
    lcd_str_P(PSTR("% ")); lcd_char(mode_auto ? 'A' : 'M');

    lcd_gotoxy(0, 1);
    if (fan_on) {
        lcd_str_P(PSTR("ON  Spd:")); lcd_num(fan_speed); lcd_str_P(PSTR("    "));
    } else {
        lcd_str_P(PSTR("OFF H:")); lcd_num(temp_high); lcd_str_P(PSTR(" L:")); lcd_num(temp_low);
    }
}

int main(void) {
    MOTOR_DDR |= (1 << IN1_PIN) | (1 << IN2_PIN); 
    PWM_DDR |= (1 << PWM_PIN); 
    
    TCCR0 = (1 << WGM00) | (1 << WGM01) | (1 << COM01) | (1 << CS01); 
    OCR0 = 0;

    BTN_DDR &= ~((1 << BTN_AUTO) | (1 << BTN_TOGGLE) | (1 << BTN_UP) | (1 << BTN_DOWN)); 
    BTN_PORT |= (1 << BTN_AUTO) | (1 << BTN_TOGGLE) | (1 << BTN_UP) | (1 << BTN_DOWN);

    uart_init_fixed();
    load_config();
    lcd_init();   
    
    set_motor(fan_on); 
    
    init_int0();

    sei();

    uart_send_str_P(PSTR("System Ready (INT0 Mode)\r\n"));
    lcd_cmd(0x01); lcd_gotoxy(0, 0); lcd_str_P(PSTR("System Ready"));
    _delay_ms(1000); lcd_cmd(0x01);

    uint8_t timer_ticks = 0;

    while (1) {
        if (cmd_ready) {
            _delay_ms(5);
            char *cmd = (char*)rx_buffer;
            
            if (cmd[0] == '1') { 
                mode_auto = 0; eeprom_update_byte(&ee_mode_auto, 0); 
                set_motor(1); 
                uart_send_str_P(PSTR("Manual ON\r\n"));
            }
            else if (cmd[0] == '0') { 
                mode_auto = 0; eeprom_update_byte(&ee_mode_auto, 0); 
                set_motor(0); 
                uart_send_str_P(PSTR("Manual OFF\r\n"));
            }
            else if (cmd[0] == 'A' || cmd[0] == 'a') { 
                mode_auto = 1; eeprom_update_byte(&ee_mode_auto, 1); 
                uart_send_str_P(PSTR("Mode: AUTO\r\n"));
                update_lcd_needed = 1; 
            }
            else if (cmd[0] == 'M' || cmd[0] == 'm') { 
                mode_auto = 0; eeprom_update_byte(&ee_mode_auto, 0); 
                uart_send_str_P(PSTR("Mode: MANUAL\r\n"));
                update_lcd_needed = 1; 
            }
            else if (cmd[0] == 'V' || cmd[0] == 'v') { 
                int val = atoi(&cmd[1]); 
                if (val >= 0 && val <= 255) { 
                    fan_speed = val; 
                    eeprom_update_byte(&ee_fan_speed, val); 
                    if(fan_on) OCR0 = fan_speed; 
                    update_lcd_needed = 1;
                } 
            }
            else if (cmd[0] == 'H' || cmd[0] == 'h') { 
                int val = atoi(&cmd[1]); 
                if (val > temp_low && val < 99) { 
                    temp_high = val; eeprom_update_byte(&ee_temp_high, val); 
                    update_lcd_needed = 1; 
                }
            }
            else if (cmd[0] == 'L' || cmd[0] == 'l') { 
                int val = atoi(&cmd[1]); 
                if (val > 0 && val < temp_high) { 
                    temp_low = val; eeprom_update_byte(&ee_temp_low, val); 
                    update_lcd_needed = 1; 
                }
            }
            cmd_ready = 0;
        }

        _delay_ms(100);
        timer_ticks++;
        if (timer_ticks >= 20) {
            timer_ticks = 0;
            if (dht_read() == 0) {
                update_lcd_needed = 1;
                if (mode_auto) {
                    if (temp_c >= temp_high && !fan_on) set_motor(1);
                    else if (temp_c <= temp_low && fan_on) set_motor(0);
                }
            }
        }

        if (update_lcd_needed) {
            update_lcd_info();
            update_lcd_needed = 0;
        }
    }
}