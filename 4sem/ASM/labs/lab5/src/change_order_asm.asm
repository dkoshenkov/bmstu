global change_order
extern print_result

section .data
    text times 255 db 0
    len db 0
    N dq 0
    space db " "


section .bss
    result resb 255
    _word resb 30

section .text
    global _start

change_order:
    push rbp
    mov rbp, rsp
    lea rsp, [rsp]
    mov qword [N], rsi
    push rax
    push rbx
    push rcx
    push rdx

    xor rax, rax
    xor rbx, rbx
    xor rcx, rcx
    xor rdx, rdx

    push rsi
    push rdi

    cld
    mov rcx, 255
    mov al, 0xA ; помещаем Enter для сравнения
    repne scasb ; ищем Enter в строке
    mov rax, 255
    sub rax, rcx ; 255 - число, оставшеейся в ECX до Enter
    mov rcx, rax ; длина строки + 1
    dec rcx
    ; mov BYTE[InBuf+RCX-1], ' ' ; добавление пробела в конец
    mov byte [len], cl
    pop rsi
    lea rdi, [text]
    rep movsb
    pop rdi

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
    lea rdi, [result]
    call print_result

    pop rdx
    pop rcx
    pop rbx
    pop rax
    mov rsp, rbp
    pop rbp
    mov rax, result
    ret