section .data
    text db "abcd efgh ijkl mnop qrst uvwx",0
    N dq 2
    space db " ",0
    newline db 10,0
    original_text_msg db "Исходный текст: ",0
    reversed_text_msg db "Перевернутый текст: ",0

section .bss
    result resb 100
    _word resb 20

section .text
    global _start

_start:
    mov rsi, text       ; rsi - указатель на текущий символ в text
    mov rdi, result      ; rdi - указатель на текущую позицию в result
    mov rcx, 1          ; rcx - счетчик слов
    xor rdx, rdx        ; rdx - индекс символа в слове

.loop:
    mov al, [rsi]      ; al - текущий символ

    cmp al, 0          ; Проверяем на конец строки
    je .end_loop

    cmp al, ' '        ; Проверяем на пробел
    je .space_found

    mov [_word+rdx], al ; Добавляем букву к текущему слову
    inc rdx             ; Увеличиваем индекс символа в слове
    jmp .next_char

.space_found:
    cmp rcx, [N]     ; Проверяем, нужно ли переворачивать слово
    jne .add_word

    ; Переворачиваем слово
    mov rbx, rdx       ; rbx - индекс символа в слове (в обратном порядке)
    dec rbx

.reverse_word:
    mov al, [_word+rbx]
    mov [rdi], al
    inc rdi
    dec rbx
    cmp rbx, 0
    jge .reverse_word

    jmp .add_space

.add_word:
    ; Добавляем слово без изменений
    mov rbx, 0
.add_word_loop:
    mov al, [_word+rbx]
    mov [rdi], al
    inc rdi
    inc rbx
    cmp rbx, rdx
    jl .add_word_loop

.add_space:
    mov al, ' '
    mov [rdi], al
    inc rdi
    inc rcx             ; Увеличиваем счетчик слов
    xor rdx, rdx        ; Обнуляем индекс символа в слове

.next_char:
    inc rsi             ; Переходим к следующему символу
    jmp .loop

.end_loop:
    ; Обработка последнего слова
    cmp rcx, [N]
    jne .add_last_word

    ; Переворачиваем последнее слово
    mov rbx, rdx
    dec rbx

.reverse_last_word:
    mov al, [_word+rbx]
    mov [rdi], al
    inc rdi
    dec rbx
    cmp rbx, 0
    jge .reverse_last_word

    jmp .print_result

.add_last_word:
    ; Добавляем последнее слово без изменений
    mov rbx, 0
.add_last_word_loop:
    mov al, [_word+rbx]
    mov [rdi], al
    inc rdi
    inc rbx
    cmp rbx, rdx
    jl .add_last_word_loop

.print_result:
    ; Выводим исходный текст
    mov rax, 1          ; write
    mov rdi, 1          ; stdout
    mov rsi, original_text_msg
    mov rdx, 17
    syscall

    mov rax, 1          ; write
    mov rdi, 1          ; stdout
    mov rsi, text
    mov rdx, 26
    syscall

    ; Выводим перенос строки
    mov rax, 1          ; write
    mov rdi, 1          ; stdout
    mov rsi, newline
    mov rdx, 1
    syscall

    ; Выводим перевернутый текст
    mov rax, 1          ; write
    mov rdi, 1          ; stdout
    mov rsi, reversed_text_msg
    mov rdx, 20
    syscall

    mov rax, 1          ; write
    mov rdi, 1          ; stdout
    mov rsi, result
    mov rdx, 26
    syscall

    ; Выходим из программы
    mov rax, 60         ; exit
    xor rdi, rdi        ; exit code 0
    syscall