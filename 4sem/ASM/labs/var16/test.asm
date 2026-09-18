    ; в каждом столбце найти произвеление первых 
    ; двух и последних двух
    ; элементов проверить равны ли и вывести соотв сообщение
    
    
    
    section .data     
    OutputStrLen dq 0
    OutputStr dq 0

   msgOk db "OK", 10
   msgNO db "NO", 10
   lenMsg equ $-msgNO


    A db 1, 2, 3, 4, 5,1
      db 7, 8, 9, 10, 11, 1
      db 13,14,15,16,17,18
      db 1, 20, 21, 22, 23, 1
      db 7, 26, 27, 28, 29, 1
   
   h db 0
  section .bss    
    ;output buf
    OutBuf resb 10
    lenOutBuf equ $-OutBuf


  section .text    
  global _start
  
_start:
   mov ebx, 0
   mov rcx, 6

cycle1:
   push rcx
   mov al, byte[ebx + 0 + A]
   mov dl, byte[ebx + 6 + A]
   imul dl
   mov [h], al

   mov al, [ebx + 18 + A]
   mov dl, [ebx + 24 + A]
   imul dl

   mov dl, [h]
   push rbx
   cmp dl, al
   je ok
   mov rax, 1
   mov rbx, 1
   mov rsi, msgNO
   mov rdx, lenMsg
   
   syscall
   jmp exitIF

ok:
   mov rax, 1
   mov rbx, 1
   mov rsi, msgOk
   mov rdx, lenMsg
   syscall
   jmp exitIF


exitIF:
   pop rbx
   inc ebx
   pop rcx
   loop cycle1
 ;exit
    mov rax, 60
    xor rdi, rdi
    syscall

