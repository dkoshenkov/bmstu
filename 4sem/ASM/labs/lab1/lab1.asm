section .data              ; сегмент инициализированных переменных 
   f1 dw 65535
   f2 dd 65535
section .text         ; сегмент кода
   global  _start
_start:
   add word[f1], 1
   add dword[f2], 1
   
   ; exit
   mov     rax, 60       ; системная функция 60 (exit)
   xor     rdi, rdi      ; return code 0    
   syscall               ; вызов системной функции


