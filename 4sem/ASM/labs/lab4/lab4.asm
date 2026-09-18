%define MATRIX_SIZE 25
%define ROW_COUNT   5
%define STR_COUNT   5
%define LEN         40

section .data ; сегмент инициализированных переменных
    newline db 30
   ind dw 0
   matrix times MATRIX_SIZE dq 0 ; матрица 25 символов 
   InBuf  times 32  db 0         ; буфер для вводимой строки
   LenIn equ $-InBuf

   err_num     db  "Only numbers and spaces can be entered", 10
   err_num_len equ $-err_num

   err_line     db  "Each line should have exactly 5 numbers divided by spaces", 10
   err_line_len equ $-err_line

   line_sum times 5 dq 0
   sorok db 40

section .bss ; сегмент неинициализированных переменных
   OutBuf resq 10 ; буфер вывода
   OutLen resq 1

; Функция ввода
section .text
   global input_funcToI
   input_funcToI:
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

      mov ebx, 80000000h
      and ebx, eax
      cmp ebx, 0
      jl minus
      jmp __exit
   minus:
      mov rbx, 4294967296
      sub rax, rbx
   __exit:   
      ret

section .text
   global input_func
   input_func:
      mov rax, 0
      mov rdi, 0
      mov rsi, InBuf
      mov rdx, LenIn
      syscall

      ret

;Функция вывода
section .text
   global output_func
   output_func:
      push rdi
      push rax
      push rsi
      push rdx
      push rcx
      mov  rsi, OutBuf ; загрузка адреса буфера вывода
      call IntToStr64
      mov  [OutLen], rax

      mov rax, 1        ; 
      mov rdi, 1        ; 
      mov rsi, OutBuf   ; 
      mov rdx, [OutLen] ; 
      syscall; 
      pop rcx
      pop rdx
      pop rsi
      pop rax
      pop rdi
      ret

; ввод матрицы    
section .text
   global matrix_input
   matrix_input:
      read_line: 
         push rcx
         push rdi

         call input_func

         pop rdi
         mov rcx, rax
         xor rdx, rdx
         xor r10, r10 ; r10 - индекс числа в строке
      process_line:
         cmp byte[InBuf + rdx], 10
         je  process_number

         cmp byte[InBuf + rdx], ' '
         jne next

         mov byte[InBuf + rdx], 10

         cmp r10, rdx
         jne process_number
         jmp next

      process_number:
         push rdx

         call StrToInt64
         cmp  rbx, 0
         jne  error_num

         cmp rdi, MATRIX_SIZE
         jge error_line       ; обработка ошибки

         mov [matrix + 8 * rdi], rax ; сохраняем число в матрице
         inc rdi

         pop rdx
         mov r10, rdx
         inc r10
         lea rsi, [InBuf + r10]

      next:
         inc  rdx
         loop process_line

         pop rcx

         mov  rax, ROW_COUNT
         sub  rax, rcx
         inc  rax
         push rdx
         mov  rdx, ROW_COUNT
         mul  rdx
         pop  rdx

         cmp rdi, rax
         jne error_line

         dec rcx
         cmp rcx, 0
         jne read_line
      ret

section .text
   global matrix_out
   matrix_out:
      mov rcx, ROW_COUNT
      xor rsi, rsi
      lea rdi, [matrix]

   .print_loop:
      push rcx
      mov rcx, STR_COUNT
      .print_row:
         mov rax, qword[rdi]
         call output_func ; вызываем функцию вывода числа
         add rdi, 8
         loop .print_row

      ; print a newline at the end of the row
      push rdi
      push rax
      push rsi
      push rdx
      mov rax, 1
      mov rdi, 1
      mov rsi, newline
      mov rdx, 1
      syscall
      pop rdx
      pop rsi
      pop rax
      pop rdi
      pop rcx
      loop .print_loop

      ret

section .text
   global proc_row
    proc_row:
      ; в r9 - нужный индекс
      ; в r10 - матрица
      ; в r11 - сумм
      push rcx
      push rax
      xor r11, r11
      mov rcx, STR_COUNT
             
      _sum_str:
         cmp byte[r10], 0
         jle .skip
         add r11, qword[r10]
         .skip:
         add r10, 8
         loop _sum_str

         .done:
         xor rax, rax
         mov rax, r9
         imul byte[sorok]
         
         mov qword[matrix + rax+r9*8], r11

         pop rax
         pop rcx
         ret

section .text ; сегмент кода
   global _start

_start:
   mov rcx, ROW_COUNT
   xor rdi, rdi

   call matrix_input

   xor rax, rax

   mov rcx, STR_COUNT
   lea r10, [matrix]
   xor r9,  r9

.loop_rows:
   call proc_row ; вызов функции для подсчета суммы элементов строки
    inc r9       ; номер строки/элемента в строке

    loop .loop_rows

   lea r10, [matrix]
   call matrix_out
exit:
   ; exit
   mov     rax, 60  ; системная функция 60 (exit)
   xor     rdi, rdi ; return code 0    
   syscall          ; вызов системной функции

error_line:
    mov rax, 1
    mov rdi, 1
    mov rsi, err_line
    mov rdx, err_line_len
    syscall
    jmp exit

error_num:
    mov rax, 1
    mov rdi, 1
    mov rsi, err_num
    mov rdx, err_num_len
    syscall
    jmp exit


%include "../../lib.asm"