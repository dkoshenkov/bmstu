global sum

section .text
sum:
    push rbp          
    mov  rbp, rsp     
    sub  rsp, 8       ; Выравниваем стек по границе 16 байт

    mov rax, rdi     
    add rax, rsi              

    add  rsp, 8       ; Восстанавливаем стек
    mov rsp, rbp      
    pop rbp          
    ret                ; Возвращаем управление