%define aa -1

section .data ; сегмент инициализированных переменных
   div_zero_err db  "ERROR! Div by 0", 10
   err_len      equ $-div_zero_err

   _test db "TEST", 10
   _test_len equ $-_test

   sign db 1
   mask dq 0xFFFFFFFFFFFFFFFF

   of_err           db  "ERROR! Overflow", 10
   err_overflow_len equ $-of_err
section .bss ; сегмент неинициализированных переменных
   A resq 1
   B resq 1
   C resq 1
   D resq 1

   OutBuf resq 10 ; буфер вывода
   OutLen resq 1
   
   InBuf resd 10 ; буфер для вводимой строки
   LenIn equ $-InBuf
   

; Функция ввода
section .text
   global _TEST
   _TEST:
      push rax

      mov rax, 1 ; системная функция 1 (write)
      mov     rdi, 1            ; дескриптор файла stdout=1
      mov     rsi, _test ; адрес выводимой строки
      mov     rdx, _test_len      ; длина строки
      syscall                   ; вызов системной функции

      pop rax
      ret

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
      mov  [OutLen], eax

      mov rax, 1        ; 
      mov rdi, 1        ; 
      mov rsi, OutBuf   ; 
      mov rdx, [OutLen] ; 
      syscall; 

      ret



section .text ; сегмент кода
   global _start
        
;    a = b * ( c - d ) - c / ( d - 1 )
_start:
   ; Вводим переменные 
   call input_func
   mov  [B], rax

   call input_func
   mov  [C], rax

   call input_func
   mov  [D], rax

    mov ebx, 80000000h
    and ebx, eax
    cmp ebx, 0
jl minus
    mov [D], rax
    jmp exitMin
minus:
    mov rbx, 4294967296
    sub rax, rbx
    mov [D], rax

exitMin:

   mov rbx, [D]
   dec rbx
   cmp rbx, 0
   je  zero_err
  
   mov  rax, [C]
   
   idiv rbx

   mov [A], rax
   ; call output_func

   mov rbx, [C] ; c
   sub rbx, [D] ; c - d
   jo  of_err

   mov  rax, [B]
   imul rbx      ; b * ( c - d )
   jo   of_err

   sub rax, [A] ; b * ( c - d ) - c / ( d - 1 ) -- ответ
   
   call output_func
exit:
   mov rax, 60  ; 
   xor rdi, rdi ; 
   syscall; 

zero_err:
    mov     eax, 1            ; системная функция 1 (write)
    mov     edi, 1            ; дескриптор файла stdout=1
    mov     esi, div_zero_err ; адрес выводимой строки
    mov     edx, err_len      ; длина строки
    syscall                   ; вызов системной функции
    jmp     exit

overflow_err:
    mov     eax, 1            ; системная функция 1 (write)
    mov     edi, 1            ; дескриптор файла stdout=1
    mov     esi, div_zero_err ; адрес выводимой строки
    mov     edx, err_len      ; длина строки
    syscall                   ; вызов системной функции
    jmp     exit

%include "../../lib.asm"