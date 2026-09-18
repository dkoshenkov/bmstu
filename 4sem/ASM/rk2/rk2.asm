%define EOF_ch   0Ah
%define space_ch 20h

section .data ; сегмент инициализированных переменных
    word_len db 0
    count    db 0
    strLen   db 64
    char     db 0
    outputLen  db 0
    space_s db 20h

section .bss ; сегмент неинициализированных переменных
    Input resb    64 ; буфер для вводимой строки
    OutBuf dw  10
    OutLen equ $-OutBuf

section .text ; сегмент кода
    global _start
_start:
    ; ввод строки
    mov eax, 3
    mov ebx, 0
    mov ecx, Input
    mov edx, strLen
    int 80h

    ;итерац
    movzx ecx, byte [strLen]
    lea ebx, [Input]

cycl:
    cmp byte[ebx], EOF_ch
    je _space
    cmp byte[ebx], space_ch
    je _space

    cmp byte[word_len], 1
    je _if_2nd
    jg _proc
    jmp _end_if

    _if_2nd:
        mov al, byte [ebx]
        mov byte[char], al
        dec bl
        mov al, byte [ebx]
        inc bl
        cmp byte[char], al
        jne _end_if
        inc byte[count]

        jmp _end_if
    _proc:
        mov al, byte [ebx]
        cmp byte[char], al
        jne _end_if
        inc byte[count]
    
    _end_if:
        inc byte[word_len]
        jmp _end_cycl_block
    
    _space:
        cmp byte[word_len], 0
        je _end_cycl_block

        push ecx
        push ebx

        mov esi, OutBuf
        movzx eax,byte [count]
        call IntToStr

        mov [outputLen], eax
        mov eax, 4
        mov ebx, 1
        mov ecx, OutBuf
        mov edx, [outputLen]
        int 80h

        mov eax, 4
        mov ebx, 1
        mov ecx, space_s
        mov edx, 1
        int 80h

        mov byte[count], 0
        mov byte[word_len], 0

        pop ebx
        pop ecx
    _end_cycl_block:
        inc ebx
        dec ecx
        jnz cycl

_exit:
    ; exit
    mov eax, 1       ; системная функция 1 (exit)
    xor ebx, ebx     ; код возврата 0    
    int 80h          ; вызов системной функции


%include "lib32.asm"