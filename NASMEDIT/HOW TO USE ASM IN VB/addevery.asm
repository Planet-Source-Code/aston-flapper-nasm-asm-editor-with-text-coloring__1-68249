%DEFINE P1 [ebp+8]              ;Define variable.
%DEFINE P2 [ebp+12]             ;Define variable.
%DEFINE P3 [ebp+16]             ;Define variable.
%DEFINE P4 [ebp+20]             ;Define variable.
[BITS 32]                       ;Indicate that registers are 32 bits.
PUSH ebp                        ;Push stack into ebp.
MOV ebp,esp                     ;ebp = esp
PUSH ebx                        ;Push stack into ebx.
PUSH esi                        ;Push stack into esi.
PUSH edi                        ;Push stack into edi.

MOV ebx, P1                     ;ebx = P1
MOV eax, 1                      ;eax = 1
MOV ecx, 1                      ;ecx = 1
start:                       
ADD ecx, 1                      ;ecx = ecx + 1
ADD eax, ecx                    ;eax = eax + ecx
CMP ecx, ebx                    ;Compare ecx and ebx.
JNE start                       ;If ecx <> ebx Then goto start

POP edi                         ;Pop edi back into stack.
POP esi                         ;Pop esi back into stack.
POP ebx                         ;Pop ebx back into stack.
MOV esp, ebp                    ;esp = ebp
POP ebp                         ;Pop ebp back into stack.
RET 16                          ;Return 16





