; Merge.asm by Robert Rayment  July 2003

;'............................................................
; VB Struc
; Public Type ASMType
;   W As Long
;   H As Long
;   T As Long
;   L As Long
;   iza As Long
;   ptrmemPic As Long
;   
;   ptrmemBack As Long
;   maxw As Long
;   maxh As Long
;   TransColor As Long
; End Type
;'............................................................

;res = CallWindowProc(ptrMC, ptrStruc, 2&, 3&, 4&)
;                             8         12  16  20


%macro movab 2      ;name & num of parameters
  push dword %2     ;2nd param
  pop dword %1      ;1st param
%endmacro           ;use  movab %1,%2
;Allows eg  movab bmW,[ebx+4]

;Define names to match VB code
%define W         [ebp-4]
%define H         [ebp-8]
%define T         [ebp-12]
%define L         [ebp-16]
%define iza       [ebp-20]
%define ptrmemPic [ebp-24]
%define ptrmemBack [ebp-28]
%define maxw      [ebp-32]
%define maxh      [ebp-36]
%define TransColor  [ebp-40]

; Some variables
%define iybac    [ebp-44]
%define ixbac    [ebp-48]    
%define picCul   [ebp-52]    
%define mbacCul  [ebp-56]    
%define ix       [ebp-60]
%define iy       [ebp-64] 

[bits 32]

;----------------------------

    push ebp
    mov ebp,esp
    sub esp,64
    push edi
    push esi
    push ebx

;----------------------------

    ;Fill structure
    mov ebx,[ebp+8]
    movab W,          [ebx]
    movab H,          [ebx+4]
    movab T,          [ebx+8]
    movab L,          [ebx+12]
    movab iza,        [ebx+16]
    movab ptrmemPic,  [ebx+20]
    movab ptrmemBack,  [ebx+24]
    movab maxw,       [ebx+28]
    movab maxh,       [ebx+32]
    movab TransColor,   [ebx+36]
;----------------------------
    ; Set up alpha multiplier in mmx
    pxor mm6,mm6    ; mm6=0
    mov eax,iza     ; alpha
    mov ebx,eax
    shl ebx,16
    add eax,ebx
    movd mm7,eax
    movq mm6,mm7
    punpckldq mm7,mm6   ; mm7 4 word multiplier
    
    pxor mm6,mm6    ; mm6=0

; For iy = H To 1 Step -1   ' picFull(i).Height

    mov ecx,H
For_iy:
    mov iy,ecx

;    iybac = iy + T    ' picFull(i).Top
;    
;    If iybac >= 1 Then
;    If iybac <= maxh Then

    mov eax,T
    add eax,ecx
    mov iybac,eax
    cmp eax,1
    jl near Nex_iy
    cmp eax,maxh
    jg near Nex_iy

    push ecx

;       For ix = W To 1 Step -1   ' picFull(i).Width
    mov ecx,W
For_ix:
    mov ix,ecx
;          ixbac = ix + L    ' picFull(i).Left
;          
;          If ixbac >= 1 Then
;          If ixbac <= maxw Then
    mov eax,L
    add eax,ecx
    mov ixbac,eax
    cmp eax,1
    jl near Nex_ix      
    cmp eax,maxw
    jg near Nex_ix
;--------------------------------------
;             picCul = memPic(ix, iy)
;             
    mov esi,ptrmemPic
    mov eax,iy
    dec eax
    mov ebx,W
    mul ebx
    mov ebx,ix
    dec ebx
    add eax,ebx
    shl eax,2   ; x4
    add esi,eax
;             If picCul <> TransColor Then
    mov eax,[esi]
    cmp eax,TransColor
    je Nex_ix

    mov picCul,eax
    
;                mbacCul = memBack(ixbac, iybac)
    mov edi,ptrmemBack
    mov eax,iybac
    dec eax
    mov ebx,maxw
    mul ebx
    mov ebx,ixbac
    dec ebx
    add eax,ebx
    shl eax,2   ; x4
    add edi,eax
    
;                ' Cross fade, iza = 0 to 256
;                bacblue = iza * (bpicblue - bacblue) \ 256 + bacblue
;                bacgreen = iza * (bpicgreen - bacgreen) \ 256 + bacgreen
;                bacred = iza * (bpicred - bacred) \ 256 + bacred
    
    movd mm0,[esi]
    movd mm1,[edi]
    punpcklbw mm0,mm6
    punpcklbw mm1,mm6
    psubw mm0,mm1       ; pic BGRA - bac BGRA
    pmullw mm0,mm7      ; * iza
    psrlw mm0,8         ; \ 256
    paddb mm0,mm1       ; + bac BGRA
    packuswb mm0,mm6
    movd [edi],mm0

;--------------------------------------
Nex_ix:
    dec ecx
    jnz near For_ix
    
    pop ecx

Nex_iy:
    dec ecx
    jnz near For_iy
;----------------------------
mov eax,ecx
GETOUT:
    emms
    pop ebx
    pop esi
    pop edi
    mov esp,ebp
    pop ebp

    ret 16

;========================================
