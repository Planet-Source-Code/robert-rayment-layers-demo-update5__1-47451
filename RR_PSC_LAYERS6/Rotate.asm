; Rotate.asm  by  Robert Rayment June 2003

;VB
;Public Type ASMRot   ' Input
;   Wsrc As Long      ' memPic width (picFull(PicNum)
;   Hsrc As Long
;   Wdes As Long      ' memBack width
;   Hdes As Long
;   ptrmemPic As Long
;   ptrmemBack As Long
;   ztheta As Single  ' rotation angle (rad)
;   TransColor As Long
;                     ' Output:
;   ptop As Long      ' min rect after rotation
;   pleft As Long     ' to be extracted from memBack
;   pright As Long    ' to new picFull(PicNum)
;   pbottom As Long
;End Type

;Public ASMRotation As ASMRot

;res = CallWindowProc(ptrMC2, ptrStruc2, 2&, 3&, ptrMC2)
;                             8          12  16  20


%macro movab 2      ;name & num of parameters
  push dword %2     ;2nd param
  pop dword %1      ;1st param
%endmacro           ;use  movab %1,%2
;Allows eg  movab bmW,[ebx+4]

;Define names to match VB code
%Define Wsrc    [ebp-4]
%Define Hsrc    [ebp-8]
%Define Wdes    [ebp-12]
%Define Hdes    [ebp-16]
%define ptrmemPic   [ebp-20]
%define ptrmemBack  [ebp-24]
%define ztheta      [ebp-28]
%define TransColor  [ebp-32]

%define  ptop       [ebp-36]
%define  pleft      [ebp-40]
%define  pright     [ebp-44]
%define  pbottom    [ebp-48]

%define zcos        [ebp-52]
%define zsin        [ebp-56]

%define iy      [ebp-60]
%define ix      [ebp-64]
%define ixcsrc  [ebp-68]
%define iycsrc  [ebp-72]
%define ixcdes  [ebp-76]
%define iycdes  [ebp-80]
%define xs      [ebp-84]
%define ys      [ebp-88]
%define ixs     [ebp-92]
%define iys     [ebp-96]

%define xsf     [ebp-100]
%define ysf     [ebp-104]

%define cul0    [ebp-108]
%define cul1    [ebp-112]
%define culr    [ebp-116]

%define m256  [ebp-120]
%define half  [ebp-124]

[bits 32]

;----------------------------

    push ebp
    mov ebp,esp
    sub esp,124
    push edi
    push esi
    push ebx

;----------------------------
    ;Fill structure
    mov ebx,[ebp+8]
    movab Wsrc,       [ebx]
    movab Hsrc,       [ebx+4]
    movab Wdes,       [ebx+8]
    movab Hdes,       [ebx+12]
    movab ptrmemPic,  [ebx+16]
    movab ptrmemBack, [ebx+20]
    movab ztheta,     [ebx+24]
    movab TransColor, [ebx+28]
    movab ptop,       [ebx+32]
    movab pleft,      [ebx+36]
    movab pright,     [ebx+40]
    movab pbottom,    [ebx+44]
;------------------------------------- 

;    ; Long data pick up - need pointer to mcode bytes
    mov ebx,[ebp+20]    ; ebx=ptrMC2
    mov eax,Valu        ; offset @ valu
    add ebx,eax         ; ebx->256
    mov eax,[ebx]       ; eax=256
    mov m256,eax

    mov ebx,[ebp+20]    ; ebx=ptrMC2
    mov eax,Valu2       ; offset @ valu2
    add ebx,eax         ; ebx->0.5
    mov eax,[ebx]       ; eax=0.5 pattern
    mov half,eax

;   JMP GETOUT

    ; Pre-fill memBack with TransCul
    mov eax,Wdes
    mov ebx,Hdes
    mul ebx
    mov ecx,eax         ; ecx=size in longs of memBack
    mov eax,TransColor
    mov edi,ptrmemBack
    rep stosd
    
    
    fld dword ztheta
    fsincos
    fstp dword zcos
    fstp dword zsin

	fild dword Wsrc
	fld1
	faddp st1			
	fld dword half
	fmulp st1
	fstp dword ixcsrc		; (Wsrc+1)/2

	fild dword Hsrc
	fld1
	faddp st1			
	fld dword half
	fmulp st1
	fstp dword iycsrc		; (Hsrc+1)/2

	fild dword Wdes
	fld1
	faddp st1			
	fld dword half
	fmulp st1
	fstp dword ixcdes		; (Wdes+1)/2

	fild dword Hdes
	fld1
	faddp st1			
	fld dword half
	fmulp st1
	fstp dword iycdes		; (Hdes+1)/2

	; Dest coords
	; iy 1->Hdes
	; ix 1->Wdes

	; From des coords find if src coords
	; are in the range iys 1->Hsrc, ixs 1-Wsrc

    mov ecx,Hdes
Foriy:
    mov iy,ecx
    push ecx
    
    mov ecx,Wdes
Forix:
    mov ix,ecx
    push ecx
;----------------------------
;   Source coords for rotated point
;   xs = ixcsrc + (ix - ixcdes) * zCos + (iy - iycdes) * zSin

    fld dword ixcsrc

    fild dword ix
    fld dword ixcdes  
    fsubp st1           ; st1-st0 = ix-ixcdes
    fld dword zcos
    fmulp st1           ; st0 = (ix - ixcdes) * zCos

    fild dword iy
    fld dword iycdes  
    fsubp st1           ; st1-st0 = iy-iycdes
    fld dword zsin
    fmulp st1           ; st0 = (iy - iycdes) * zSin

    faddp st1			; (ix-ixcdes)*zCos + (iy-iycdes)*zSin
    faddp st1			; + ixcsrc
    fst dword xs		; xs
    fld dword half
    fsubp st1
    fistp dword ixs		; Int(xs)


	; Test if ixs in range
;   ' Make sure source available from Picture2
;   If ixs >= 1 Then
;   If ixs <= Wsrc Then

    mov eax,ixs
    cmp eax,1
    jl near Nexix
    cmp eax,Wsrc
    jg near Nexix

;   ys = iycsrc + (iy - iycdes) * zCos - (ix - ixcdes) * zSin
    
    fld dword iycsrc
    
	fild dword iy
    fld dword iycdes  
    fsubp st1           ; st1-st0 = iy-iycdes
    fld dword zcos
    fmulp st1           ; st0 = (iy - iycdes) * zCos

    fild dword ix
    fld dword ixcdes  
    fsubp st1           ; st1-st0 = ix-ixcdes
    fld dword zsin
    fmulp st1           ; st0 = (ix - ixcdes) * zSin
    
	fsubp st1			; (iy-iycdes)*zCos - (ix-ixcdes)*zSin
    faddp st1			; + iycsrc
    fst dword ys		; ys
    fld dword half
    fsubp st1
    fistp dword iys		; Int(ys)

    ; Test if iys in range
;   If iys >= 1 Then
;   If iys <= Hsrc Then
;   < Hsrc because looking at iys & iys+1 !!!

    mov eax,iys
    cmp eax,1
    jl near Nexix
    cmp eax,Hsrc
    jg near Nexix


	; InRange
    ; Get scale factors xsf=xs-ixs, ysf=ys-iys
	; xfs = xs - ixs 'Int(xs)
	; yfs = ys - iys 'Int(ys)

    fld dword xs
    fild dword ixs
    fsubp st1       	; xs-ixs
    fild dword m256
    fmulp st1
    fld dword half
    fsubp st1
    fistp dword xsf 	; Int(xsf * 256)
    
    fld dword ys
    fild dword iys
    fsubp st1       	; ys-iys
    fild dword m256
    fmulp st1
    fld dword half
    fsubp st1
	fistp dword ysf 	; Int(ysf * 256)

    pxor mm6,mm6
    mov eax,xsf
    
	mov ebx,eax
    shl ebx,16
    add eax,ebx
    movd mm7,eax
    movq mm6,mm7
    punpckldq mm7,mm6   ; mm7=xsf xsf xsf xsf
    pxor mm6,mm6
    
    ; GET SOURCE COLORS
	mov esi,ptrmemPic	; Source mem
    mov eax,iys
    dec eax
    mov ebx,Wsrc
    mul ebx
    mov ebx,ixs
    dec ebx
    add eax,ebx
    shl eax,2
    add esi,eax     ; esi-> LongCul @ ixs,iys
	mov eax,[esi]
	cmp eax,TransColor
    je near Nexix

	; Get weighted colors along x-axis @ iys
	; cul0 = xsf*256 * (bgra2 - bgra1)\256 + bgra1
    
	mov eax,ixs
	cmp eax,Wsrc
	jl X1			; OK to do ixs & ixs+1
	movd mm0,[esi]	; Use RH color
    jmp X2

X1:
	movd mm0,[esi+4]    ; bgra2
	movd mm1,[esi]      ; bgra1
    punpcklbw mm0,mm6
    punpcklbw mm1,mm6
    psubw mm0,mm1       ; bgra2-bgra1
    pmullw mm0,mm7      ; *xsf*256
    psrlw mm0,8         ; \256
    paddb mm0,mm1       ; + bgra1
    packuswb mm0,mm6
X2:    
	movd cul0,mm0
    
    
	mov eax,iys
	cmp eax,Hsrc
	jl Y1				; OK to use iys & iys+1
	movd mm0,[esi]		; Use top color
	jmp Y2
Y1:
	mov eax,Wsrc
    shl eax,2
    add esi,eax     ; esi-> LongCul @ ixs,iys+1

	; Get weighted colors along x-axis @ iys+1
	; cul1 = xsf*256 * (bgra2 - bgra1)\256 + bgra1

	movd mm0,[esi+4]    ; bgra2	X
    movd mm1,[esi]      ; bgra1
    punpcklbw mm0,mm6
    punpcklbw mm1,mm6
    psubw mm0,mm1       ; bgra2-bgra1
    pmullw mm0,mm7      ; *xsf*256
    psrlw mm0,8         ; \256
    paddb mm0,mm1       ; + bgra1
    packuswb mm0,mm6
 Y2:
	movd cul1,mm0

    pxor mm6,mm6
    mov eax,ysf
    
	mov ebx,eax
    shl ebx,16
    add eax,ebx
    movd mm7,eax
    movq mm6,mm7
    punpckldq mm7,mm6 ;mm7=ysf ysf ysf ysf
    pxor mm6,mm6

	; Get weighted colors along y-axis 
	; culr = 256*ysf * (cul1 - cul0)\256 + cul0

    movd mm0,cul1      ; cul1
    movd mm1,cul0      ; cul0
    punpcklbw mm0,mm6
    punpcklbw mm1,mm6
    psubw mm0,mm1       ; cul1-cul0
    pmullw mm0,mm7      ; *ysf*256
    psrlw mm0,8         ; \256
    paddb mm0,mm1       ; + cul0 mm0=res color 
    packuswb mm0,mm6
    movd culr,mm0

    mov eax,culr
    cmp eax,TransColor
    je Nexix

    ; Put AA color into des
	mov edi,ptrmemBack
    mov eax,iy
    dec eax
    mov ebx,Wdes
    mul ebx
    mov ebx,ix
    dec ebx
    add eax,ebx
    shl eax,2
    add edi,eax     ; edi-> LongCul @ ix,iy
    movd [edi],mm0

    ;----------------------------
Nexix:
    emms
    pop ecx
    dec ecx
    jnz near Forix
Nexiy:
    emms
    pop ecx
    dec ecx
    jnz near Foriy
    
;----------------------------------
;----------------------------------

;   Find smallest rectangle
;   ptop,pleft,pright,pbottom

    xor eax,eax
    mov ptop,eax
    mov pbottom, eax

    mov ecx,Hdes
ForiyTB:
    mov iy,ecx
    push ecx

    mov ecx,Wdes
ForixTB:
    mov ix,ecx
    ;----------------------------

    mov edi,ptrmemBack
    mov eax,iy
    dec eax
    mov ebx,Wdes
    mul ebx
    mov ebx,ix
    dec ebx
    add eax,ebx
    shl eax,2
    add edi,eax     ; edi-> LongCul @ ix,iy
    mov eax,[edi]

    cmp eax,TransColor
    je TestBot
    mov eax,ptop
    cmp eax,0
    jne TestBot		; ptop already found
        mov eax,iy		; ptop at 1st
        mov ptop,eax	; non-trans color

TestBot:
;;;;;;  
    mov edi,ptrmemBack
    mov eax,Hdes
    sub eax,iy			; Hdes-iy  (0->Hdes-1)
    mov ebx,Wdes
    mul ebx
    mov ebx,ix
    dec ebx
    add eax,ebx
    shl eax,2
    add edi,eax     ; edi-> LongCul @ ix,Hdes-iy-1
    mov eax,[edi]

    cmp eax,TransColor
    je NexixTB
    mov eax,pbottom
    cmp eax,0
    jne NexixTB
        mov eax,Hdes
        sub eax,iy
        mov pbottom,eax

    ;----------------------------
NexixTB:
    dec ecx
    jnz near ForixTB
NexiyTB:
    
    pop ecx

    mov eax,ptop
    cmp eax,0
    je goon
    mov eax,pbottom
    cmp eax,0
    jne FindLR

goon:
        
    dec ecx
    jnz near ForiyTB

FindLR:

    ; Return ptop & pbottom
	mov ebx,[ebp+8]
    movab [ebx+32],ptop      
    movab [ebx+44],pbottom
    xor eax,eax
    mov pleft,eax
    mov pright,eax
    
    mov ecx,Wdes
ForixLR:
    mov ix,ecx
    push ecx

    mov ecx,Hdes
ForiyLR:
    mov iy,ecx
    mov edi,ptrmemBack
    mov eax,iy
    dec eax
    mov ebx,Wdes
    mul ebx
    mov ebx,ix
    dec ebx
    add eax,ebx
    shl eax,2
    add edi,eax     ; edi-> LongCul @ ix,iy
    mov eax,[edi]

    cmp eax,TransColor
    je TestLeft
    mov eax,pright
    cmp eax,0
    jne TestLeft
        mov eax,ix
        mov pright,eax

TestLeft:
;;;;;;  
    mov edi,ptrmemBack
    mov eax,iy
    dec eax
    mov ebx,Wdes
    mul ebx
    mov ebx,Wdes
    sub ebx,ix		; Wdes-ix ( 0-> Wdes-1)
    add eax,ebx
    shl eax,2
    add edi,eax     ; edi-> LongCul @ ix,Hdes-iy-1
    mov eax,[edi]

    cmp eax,TransColor
    je NexiyLR
    mov eax,pleft
    cmp eax,0
    jne NexiyLR
        mov eax,Wdes
        sub eax,ix
        mov pleft,eax

    ;----------------------------
NexiyLR:
    dec ecx
    jnz near ForiyLR
NexixLR:
    
    pop ecx

    mov eax,pleft
    cmp eax,0
    je goon2
    mov eax,pright
    cmp eax,0
    jne FillLR

goon2:
        
    dec ecx
    jnz near ForixLR

FillLR:
	; Return pleft & pright
    mov ebx,[ebp+8]
    movab [ebx+36],pleft
    movab [ebx+40],pright

;-------------------------------------    
GETOUT:
    emms
    pop ebx
    pop esi
    pop edi
    mov esp,ebp
    pop ebp

    ret 16


;----------------------------------
[SECTION .data]
Valu dd 256
Valu2 dd 0.5

