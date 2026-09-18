section .data ; сегмент инициализированных переменных
div_zero_err db "ERROR! Div by 0", 10
   err_len equ $-div_zero_err
section .bss ; сегмент неинициализированных переменных
    A resq 1
    K resq 1
    R resq 1

    OutBuf resq 10 ; буфер вывода
    OutLen resq 1
   
    InBuf resd 10 ; буфер для вводимой строки
    LenIn equ $-InBuf

section .text ; сегмент кода
    global _start

; Функция ввода
section .text
   global input_func
   input_func:
      ; Чтение
      mov  rax, 0
      mov  rdi, 0
      mov  rsi, InBuf
      mov  rdx, LenIn
      syscall
      ; Перевод из строки в число
      mov  rsi, InBuf
      call StrToInt64
      cmp  rbx, 0     ; проверка кода ошибки
      jne  input_func ; при преобразовании обнаружена ошибка

      ret
;Функция вывода
section .text
   global output_func
   output_func:
      mov  rsi,      OutBuf ; загрузка адреса буфера вывода
      call IntToStr64
      mov  [OutLen], rax

      mov rax, 1        ; 
      mov rdi, 1        ; 
      mov rsi, OutBuf   ; 
      mov rdx, [OutLen] ; 
      syscall; 

      ret

_start:
    ; input a k r
    call input_func
    mov  [A], rax

    call input_func
    mov  [K], rax

    call input_func
    mov  [R], rax
    
    mov ebx, 80000000h
    and ebx, eax
    cmp ebx, 0
jl minus
    mov [R], rax
    jmp exitMin
minus:
    mov rbx, 4294967296
    sub rax, rbx
    mov [R], rax

exitMin:

    mov  rbx, [K]
    mov  rax, [A]
    imul rbx

    cmp rax, 5
    jle else
;if 
    ; rax >= 5
    mov  rax, [K]
    sub  rax, 5
    cqo
    imul rax
    mov rbx, [R]
    cmp  rbx, 0
    je  zero_err

    idiv dword[R]
    jmp   end_if_else

    else:             ; rax < 5
        mov rax, 8
        sub rax, [A]

end_if_else:
    call    output_func
exit:
    mov     rax, 60     ; системная функция 60 (exit)
    xor     rdi, rdi    ; return code 0    
    syscall             ; вызов системной функции
    
zero_err:
    mov     eax, 1            ; системная функция 1 (write)
    mov     edi, 1            ; дескриптор файла stdout=1
    mov     esi, div_zero_err ; адрес выводимой строки
    mov     edx, err_len      ; длина строки
    syscall                   ; вызов системной функции
    jmp     exit

%include "../../lib.asm"