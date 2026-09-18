Вот схемы алгоритмов на языке **Mermaid**, составленные на основе предоставленного кода. Вы можете вставить этот код в любой редактор, поддерживающий Mermaid (например, Notion, Obsidian, GitHub или онлайн-редактор Mermaid Live).

### 1. Схема основного цикла (Main Loop)
Эта схема описывает логику внутри `while(1)`: обработку команд UART, периодический опрос датчика, логику термостата и обновление экрана.

```mermaid
flowchart TD
    Start([Начало Main]) --> Init[Инициализация портов, UART, LCD, прерываний]
    Init --> LoadCfg[Загрузка настроек из EEPROM]
    LoadCfg --> Loop{Бесконечный\nцикл while 1}
    
    Loop --> CheckUART{Флаг cmd_ready\nКоманда UART?}
    
    %% Ветка обработки UART
    CheckUART -- Да --> ParseCmd[Парсинг команды]
    ParseCmd --> TypeCmd{Тип команды?}
    TypeCmd -- 1/0 --> ManualSet[Ручн. вкл/выкл мотора]
    TypeCmd -- A/M --> SetMode[Смена режима Auto/Manual]
    TypeCmd -- V --> SetSpeed[Установка скорости]
    TypeCmd -- H/L --> SetThresh[Уст. порогов темп.]
    ManualSet & SetMode & SetSpeed & SetThresh --> ResetFlag[Сброс cmd_ready]
    ResetFlag --> DelayLoop
    CheckUART -- Нет --> DelayLoop
    
    %% Ветка таймера и DHT
    DelayLoop[Задержка 100мс\nТики таймера ++] --> CheckTimer{Тики >= 20?\n~2 сек}
    
    CheckTimer -- Да --> ResetTimer[Сброс тиков]
    ResetTimer --> ReadDHT[Вызов dht_read]
    ReadDHT --> DHTSuccess{Успешно?}
    
    DHTSuccess -- Да --> FlagLCD[Флаг update_lcd_needed = 1]
    FlagLCD --> CheckAuto{Режим Auto?}
    
    CheckAuto -- Да --> LogicH{Temp >= High\nИ вент выкл?}
    LogicH -- Да --> FanOn[Включить вентилятор]
    LogicH -- Нет --> LogicL{Temp <= Low\nИ вент вкл?}
    LogicL -- Да --> FanOff[Выключить вентилятор]
    LogicL -- Нет --> CheckLCD
    FanOn & FanOff --> CheckLCD
    
    CheckAuto -- Нет --> CheckLCD
    DHTSuccess -- Нет --> CheckLCD
    
    CheckTimer -- Нет --> CheckLCD
    
    %% Обновление экрана
    CheckLCD{Нужно обновить\nLCD?}
    CheckLCD -- Да --> UpdateScreen[Вывод Temp, Hum, Mode, Fan на LCD]
    UpdateScreen --> ClearLCDFlag[Сброс update_lcd_needed]
    ClearLCDFlag --> Loop
    CheckLCD -- Нет --> Loop
```

---

### 2. Алгоритм проверки кнопок (ISR INT0)
Так как кнопки обрабатываются через внешнее прерывание `ISR(INT0_vect)`, алгоритм запускается аппаратно при нажатии.

```mermaid
flowchart TD
    Start([Прерывание ISR INT0]) --> Debounce[Задержка 30мс\nАнтидребезг]
    
    Debounce --> CheckAuto{Нажата BTN_AUTO?\nPA0 == 0}
    
    CheckAuto -- Да --> SetAuto[Режим = AUTO\nСохранить в EEPROM]
    SetAuto --> SendUART1[UART: AUTO Mode]
    SendUART1 --> SetLCD[Флаг update_lcd = 1]
    SetLCD --> End([Конец ISR])
    
    CheckAuto -- Нет --> CheckTog{Нажата BTN_TOGGLE?\nPA1 == 0}
    
    CheckTog -- Да --> SetMan[Режим = MANUAL\nИнверсия Fan On/Off]
    SetMan --> SaveMan[Сохранить состояние]
    SaveMan --> SendUART2[UART: Toggle Man]
    SendUART2 --> SetLCD
    
    CheckTog -- Нет --> CheckUp{Нажата BTN_UP?\nPA2 == 0}
    
    CheckUp -- Да --> IncSpd[Скорость + 25\nLimit 255]
    IncSpd --> ApplySpd1[Обновить OCR0\nСохранить в EEPROM]
    ApplySpd1 --> SendUART3[UART: Speed UP]
    SendUART3 --> SetLCD
    
    CheckUp -- Нет --> CheckDown{Нажата BTN_DOWN?\nPA3 == 0}
    
    CheckDown -- Да --> DecSpd[Скорость - 25\nLimit 0]
    DecSpd --> ApplySpd2[Обновить OCR0\nСохранить в EEPROM]
    ApplySpd2 --> SendUART4[UART: Speed DOWN]
    SendUART4 --> SetLCD
    
    CheckDown -- Нет --> End
```

---

### 3. Алгоритм UART (Прием данных)
Здесь показан процесс приема байта в прерывании и формирование строки команды.

```mermaid
flowchart TD
    Start([Прерывание ISR USART_RX]) --> ReadUDR[Чтение регистра UDR\nchar data]
    
    ReadUDR --> CheckTerm{Символ\nCR или LF?\n r или n}
    
    CheckTerm -- Да --> Terminate[Добавить \0 в буфер]
    Terminate --> SetFlag[cmd_ready = 1]
    SetFlag --> ResetPos[rx_pos = 0]
    ResetPos --> End([Конец ISR])
    
    CheckTerm -- Нет --> CheckOvr{Буфер полон?\nrx_pos < 18}
    CheckOvr -- Нет --> End
    CheckOvr -- Да --> Store[rx_buffer = data\nrx_pos++]
    Store --> End
```

---

### 4. Алгоритм работы с DHT11
Функция `dht_read`. Описывает процесс "рукопожатия" с датчиком и чтение 40 бит данных.

```mermaid
flowchart TD
    Start([Старт dht_read]) --> ResetArr[Обнуление массива bits]
    
    %% Импульс запроса
    ResetArr --> SendStart[Пин Output -> 0\nЖдать 20 мс]
    SendStart --> PullUp[Пин Output -> 1\nПин Input]
    
    %% Ожидание ответа
    PullUp --> CheckResp1{Датчик ответил 0?\n40 мкс}
    CheckResp1 -- Нет --> Err1[Ошибка 1]
    CheckResp1 -- Да --> CheckResp2{Датчик отпустил 1?\n80 мкс}
    CheckResp2 -- Нет --> Err2[Ошибка 2]
    
    %% Чтение данных
    CheckResp2 -- Да --> InitLoop[i = 0\nЦикл чтения 40 бит]
    InitLoop --> WaitLow[Ждать пока 0\nТаймаут]
    WaitLow --> WaitHigh[Ждать пока 1\nТаймаут]
    WaitHigh --> Measure[Задержка 30 мкс\nПроверка уровня]
    
    Measure --> Level{Уровень?}
    Level -- High (1) --> BitOne[Записать 1 в байт]
    Level -- Low (0) --> BitZero[Оставить 0]
    
    BitOne & BitZero --> NextBit[Индекс бита++]
    NextBit --> LoopCond{i < 40?}
    LoopCond -- Да --> WaitLow
    
    %% Проверка контрольной суммы
    LoopCond -- Нет --> Checksum{Сумма 4 байт\n== 5 байт?}
    Checksum -- Да --> SaveGlob[Сохранить Temp и Hum]
    SaveGlob --> Success([Возврат 0: OK])
    
    Checksum -- Нет --> ErrCS[Возврат 5: Ошибка КС]
    
    Err1 & Err2 --> ErrorExit([Возврат кода ошибки])
```