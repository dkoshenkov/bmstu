section .data
    text db '1234 1234 1234 1234 1234 1234', 0
    text_len equ $-text
    N       equ 3
    res_len dq 0

section .bss
    result resb 255

section .text
    global _start

_start:
    lea rsi, [text]
    lea rdi, [result]

    xor rcx, rcx 
    xor rdx, rdx

    next_word:
        call count


end:
    ; Выводим результат
    mov rax, 1             ; Системный вызов для write
    mov rdi, 1             ; Файловый дескриптор stdout
    mov rsi, text        ; Указатель на буфер результата
    mov rdx, text_len   ; Длина текста
    syscall                ; Выполняем системный вызов

    mov eax, 1             ; Системный вызов 1 (exit)
    xor ebx, ebx          ; Код возврата 0
    int 0x80


