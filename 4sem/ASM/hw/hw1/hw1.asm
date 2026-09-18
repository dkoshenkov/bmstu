%include "../../lib.asm"

%define big_p   50h
%define litle_p 70h
%define space_ch 20h

section  .data
inputStrLen     db 255
P_counter       db  0
letters_counter db  0

section .bss
    OutBuf  resq 10 ; буфер вывода
    OutLen  resq 1
    InBuf resb 256
    


section .text
   global output_func
   output_func:
    push rcx
      mov  rsi,      OutBuf ; загрузка адреса буфера вывода
      call IntToStr64
      mov  [OutLen], rax

      mov rax, 1        ; 
      mov rdi, 1        ; 
      mov rsi, OutBuf   ; 
      mov rdx, [OutLen] ; 
      syscall; 
    pop rcx
      ret

section .text
    global input
    input:
        mov rax, 0
        mov rdi, 0
        mov rsi, InBuf
        mov rdx, inputStrLen
        syscall

        ret


section .text
    global _start

_start:
    call input
    

    lea rdi, [InBuf] ; загружаем адрес строки в edi
; загружаем размер буфера ввода
    mov rcx, [inputStrLen ]
    mov al,  0Ah   ; загружаем 0Ah для поиска в буфере ввода
    repne scasb; ищем код Enter в строке
    mov rax, 255
;
    sub rax, rcx ; вычитаем из размера буфера остаток cx
    mov rcx, rax ; полученная разница – длина строки +1 
    ;добавление пробела в конец строки  

    mov [inputStrLen], rcx
    add rcx, InBuf
    dec rcx
    mov byte[rcx], space_ch
    ; mov byte[cl+InBuf-1], ' '

    lea r8, [InBuf]
    mov rcx, [inputStrLen]
    xor rax, rax
    inc rcx
    xor r9, r9
    dec r8

range1:
    dec rcx
    jrcxz exit
    inc r8
    
    cmp byte[r8], space_ch
    je if_space

    inc r9 ; увеличивает количество букв
    cmp byte[r8], litle_p
    je its_p
    cmp byte[r8], big_p
    je its_p

    jmp range1
    ; rax - количество букв p
its_p:
    inc rax
    jmp range1

    ; цикл по буквам
    ; получаем букву по индексу
    ; если ' ' 
        ; если и letters_counter > 0 тогда выводим
        ; чистим
    ; иначе сравниваем с р

    ; выход
if_space:
    ; обнуляем
    cmp r9, 0
    je range1

    call output_func
    xor rax, rax
    xor r9, r9
    jmp range1


exit:
   ; exit
   mov     rax, 60  ; системная функция 60 (exit)
   xor     rdi, rdi ; return code 0    
   syscall          ; вызов системной функции