; FEffects.asm by Robert Rayment  July 2003

;With ASMEffects
;   .Wsrc = Wsrc
;   .Hsrc = Hsrc
;   .ptrmemPic = ptrmemPic     ' SOURCE: memPic
;   .ptrmemBack = ptrmemBack   ' DEST:   memBack
;   .FilterParam = FilterParam
;   .TransColor = TransColor
;End With
;'response = CallWindowProc(ptrMC3, ptrStruc3, ptrMC3, [Seed], EffectsType)
;                             ~ line number
;'EffectsType=0 Invert        128
;'EffectsType=1 Sharp-Soft    171
;'EffectsType=2 Dark-Bright   413

;'EffectsType=3 Red +/-       440
;'EffectsType=4 Green +/-     440
;'EffectsType=5 blue +/-      440
;'EffectsType=6 Diffuse       527
;'EffectsType=7 Relief        659
;'EffectsType=8 Metallic      659

;'EffectsType=9 HShade 
;'EffectsType=10 VShade

%macro movab 2      ;name & num of parameters
  push dword %2     ;2nd param
  pop dword %1      ;1st param
%endmacro           ;use  movab %1,%2
;Allows eg  movab bmW,[ebx+4]

;Define names to match VB code
%define W            [ebp-4]
%define H            [ebp-8]
%define ptrmemPic    [ebp-12]
%define ptrmemBack   [ebp-16]
%define TransColor   [ebp-24]
%define ZeroFilterParam    [ebp-28]
%define FilterParam        [ebp-32]
%define ZeroFP   [ebp-36]
%define FP       [ebp-40]    
%define ZerodF   [ebp-44]
%define dF       [ebp-48]    
%define BWidth   [ebp-52]

; For Diffuse
;%define Seed     [ebp-56]
;%define arand    [ebp-60]
;%define ix       [ebp-64]
;%define iy       [ebp-68]
;%define incx     [ebp-72]
;%define incy     [ebp-76]
;%define a255     [ebp-80]
;%define ahalf    [ebp-84]
; For Relief
;%define FourFPS  [ebp-88]
; For HShade & VShade
;%define im   [ebp-92]
;%define ic   [ebp-96]
;%define iB   [ebp-100]
;%define iG   [ebp-104]
;%define iR   [ebp-108]
;%define iBGR   [ebp-112]
;%define a256   [ebp-116] 




[bits 32]

;----------------------------

    push ebp
    mov ebp,esp
    sub esp,116
    push edi
    push esi
    push ebx

;----------------------------

    ;Fill structure
    mov ebx,[ebp+8]
    movab W,           [ebx]
    movab H,           [ebx+4]
    movab ptrmemPic,   [ebx+8]
    movab ptrmemBack,  [ebx+12]
    movab FilterParam, [ebx+16]
    movab TransColor,  [ebx+20]
;----------------------------

    ; Zero dwords above for psrlw mm# FilterPram & FP
    xor eax,eax
    mov ZeroFilterParam,eax
    mov ZeroFP,eax
    mov ZerodF,eax
    
    mov eax,W
    shl eax,2
    mov BWidth,eax      ; Width in bytes
    
    mov eax,[ebp+20]    ; Pick up Opcode
    cmp eax,0
    jne Test1
    call Invert
    jmp GETOUT
Test1:
    cmp eax,1
    jne Test2
    Call Soft_Sharp
    jmp GETOUT
Test2:
    cmp eax,2
    jne Test345
    Call Dark_Bright
    jmp GETOUT
Test345:
    cmp eax,5   ; 3,4,5  RGB
    jg Test6
    Call ColorShift
    jmp GETOUT
Test6;
    cmp eax,6
    jne Test78
    Call Diffuse
    jmp GETOUT
Test78:
    cmp eax,8
    jg Test9
    Call Relief ; & Metallic
Test9:
    cmp eax,9
    jg Test10
    Call HShade
    jmp GETOUT
Test10:
    cmp eax,10
    jg GETOUT
    call VShade
    
GETOUT:
    emms
    pop ebx
    pop esi
    pop edi
    mov esp,ebp
    pop ebp

    ret 16
;========================================

Invert:
    
    ;mov eax,0FFh ; 0 0 0 F
    
    mov eax,FilterParam
    mov edx,0FFh ; 0 0 0 F
    shl edx,8    ; 0 0 F 0
    or eax,edx   ; 0 0 F F
    mov edx,0FFh ; 0 0 0 F
    shl edx,16   ; 0 F 0 0
    or eax,edx   ; 0 FF FF FF
    
    movd mm2,eax
    movq mm3,mm2

    mov esi,ptrmemPic   ; Src
    mov edi,ptrmemBack  ; Des
    
    xor edx,edx
    mov eax,H
    mov ebx,W
    mul ebx
    mov ecx,eax
ForInv:
    mov eax,[esi]
    cmp eax,TransColor
    je NexInv

    movd mm1,eax            ; | |ARGB|
    psubusb mm2,mm1         ; | |FF|FF|FF|FF| -  | |ARGB|
    movd[edi],mm2

NexInv:
    movq mm2,mm3
    add esi,4
    add edi,4
    dec ecx
    jnz near ForInv

RET

;========================================
Soft_Sharp:

    mov eax,FilterParam
    cmp eax,5
    jge near Softness

Sharpness:
    ; FilterParam         1-4
    ; dF=2^FilterParam   \ 2, 4, 8,16
    ; FP=dF+8            * 10,12,16,24
    
    mov ecx,FilterParam ; 1,2,3,4
    mov eax,1
    shl eax,CL          ; \dF = 2^FilterParam = 2,4,8,16
    add eax,8           ; \dF = eax = 10,12,16,24
    mov edx,eax
    shl edx,16
    add eax,edx         ; eax = 10|10, 12|12, 16|16, or 24|24
    movd mm4,eax        ; mm4 = | | | | |10|10|10|10|
    movq mm5,mm4        ; mm5 = | | | | |10|10|10|10|
    punpckldq mm4,mm5   ; Multiplier mm4 =|  FP|  FP|  FP|  FP|

    mov eax,0FFFFFFh
    movd mm5,eax        ; mm5 to mask out Alpha
    
    mov esi,ptrmemPic   ; Src
    mov edi,ptrmemBack  ; Des

    mov ebx, BWidth

    mov ecx,1
iySharp:
    push ecx
    mov ecx,1
ixSharp:
    push ecx

    mov eax,[esi+ebx+4]
    cmp eax,TransColor
    je near NexixSharp   ; No action on TransColor

    Call Sum8       ; mm0=|SA|SumR|SumG|SumB|
                    ; also mm7=0
    
    movd mm1,[esi+ebx+4]    ; memPic(x,y) | |ARGB|
    punpcklbw mm1,mm7       ; mm1=|A|R|G|B|
    
    ; mm4 = FP FP FP FP
    ; mm5 = mask to GET RGB

    pmullw mm1,mm4          ; mm1=|FP*A|FP*R|FP*G|FP*B|
    psubusw mm1,mm0         ; mm1= FPARGB - SumARGB
    
    psrlw mm1,FilterParam   ; \dF   NB ZeroFilterParam must = 0
    
    packuswb mm1,mm7        ; | |ARGB|          
    pand mm1,mm5            ; Mask out alpha

    movd[edi+ebx+4],mm1     ;  -> memBack Stalling sequence??
                            ; would need to unroll loop
                            ; to take 2 pixel in more mms

NexixSharp:
    add esi,4
    add edi,4

    pop ecx
    inc ecx
    mov eax,W
    dec eax
    cmp ecx,eax     ;ix-(W-1)
    jl near ixSharp
NexiySharp:
    add edi,8
    add esi,8

    pop ecx
    inc ecx
    mov eax,H
    dec eax
    cmp ecx,eax     ;iy-(H-1)
    jl near iySharp
    

RET
;----------------------------

Softness:
    
    ; Sum 8, \8, \2 fixed
    
    mov eax,0FFFFFFh
    movd mm5,eax        ; mm5 to mask out Alpha
    
    mov esi,ptrmemPic   ; Src
    mov edi,ptrmemBack  ; Des

    mov ebx, BWidth

    mov ecx,1
iySoft:
    push ecx
mov edx,ecx
    mov ecx,1
ixSoft:
    push ecx

    mov eax,[esi+ebx+4]
    and eax,0FFFFFFh
    cmp eax,TransColor
    je near NexixSoft   ; No action on TransColor

    Call Sum8       ; Uses esi mm0=|SA|SumR|SumG|SumB|
                    ; also mm7=0
    
    psrlw mm0,3     ; \8     mm0 = AVA AvR AvG AvB

    movd mm1,[esi+ebx+4]    ; memPic(x,y) | |ARGB|
    punpcklbw mm1,mm7       ; mm1=|A|R|G|B|
    
    paddw mm1,mm0           ; mm1= FPARGB + AvARGB
    
    psrlw mm1,1             ; \2
    
    packuswb mm1,mm7        ; | |ARGB|          

    ; mm5 = mask to GET RGB

    pand mm1,mm5              ; Mask out alpha

    movd[edi+ebx+4],mm1     ; -> memBack Stalling sequence??
                            ; would need to unroll loop
                            ; to take 2 pixel in more mms

NexixSoft:
    add esi,4
    add edi,4

    pop ecx
    inc ecx
    mov eax,W
    dec eax
    cmp ecx,eax     ;ix-(W-1)
    jl near ixSoft
NexiySoft:
    add esi,8
    add edi,8

    pop ecx
    inc ecx
    mov eax,H
    dec eax
    cmp ecx,eax     ;iy-(H-1)
    jl near iySoft

    ;==============================
    ; Further smoothing
    ; FilterParam= 5-8 FilterParam-4  1-4 Loop 
    
    mov eax,FilterParam
    sub eax,4       ; 1-4 Loop

    cmp eax,1
    je near SoftRET
    ;--------------------
    
    mov ecx,eax
ForNN:
    push ecx

    mov esi,ptrmemBack

    mov ebx, BWidth

    mov ecx,1
iySoftNN:
    push ecx
    mov ecx,1
ixSoftNN:
    push ecx

    mov eax,[esi+ebx+4]
    and eax,0FFFFFFh
    cmp eax,TransColor
    je near NexixSoftNN   ; No action at TransColor
    ;;;;;;;;;;;;;;;;;;

    Call Sum8       ; Uses esi mm0=|SA|SumR|SumG|SumB|
                    ; also mm7=0
    
    psrlw mm0,3     ; \8     mm0 = AVA AvR AvG AvB

    movd mm1,[esi+ebx+4]    ; memBack(x,y) | |ARGB|
    punpcklbw mm1,mm7       ; mm1=|A|R|G|B|
    
    paddw mm1,mm0           ; mm1= FPARGB + AvARGB
    
    psrlw mm1,1             ; \2
    
    packuswb mm1,mm7        ; | |ARGB|          

    ; mm5 = mask to GET RGB

    pand mm1,mm5            ; Mask out alpha

    movd[esi+ebx+4],mm1     ;  -> memBack Stalling sequence??
                            ; would need to unroll loop
                            ; to take 2 pixel in more mms

;;;;;;;;;;;;;;;;;;
NexixSoftNN:
    add esi,4

    pop ecx
    inc ecx
    mov eax,W
    dec eax
    cmp ecx,eax     ;ix-(W-1)
    jl near ixSoftNN
NexiySoftNN:
    add esi,8

    pop ecx
    inc ecx
    mov eax,H
    dec eax
    cmp ecx,eax     ;iy-(H-1)
    jl near iySoftNN

    ;====================

NexNN:
    pop ecx
    dec ecx
    jnz near ForNN
    ;====================

SoftRET:

RET
;========================================

Dark_Bright:

    mov eax,FilterParam

    cmp eax,0
    je near ColorRET

    jg AddColor
SubColor:
    ; Get ABS eax & subtract
    neg eax
    js SubColor

AddColor:
    pxor mm2,mm2

    mov edx,eax ; 0 0 0 F
    shl edx,8   ; 0 0 F 0
    add eax,edx ; 0 0 F F
    mov edx,eax ; 0 0 F F
    shl edx,16  ; F F 0 0
    add eax,edx ; F F F F
    movd mm2,eax
    jmp near DoColor

;====================================================

ColorShift:

    ; eax=3,4 or 5 RGB

    mov ecx,FilterParam

    cmp ecx,0
    je near ColorRET
    jg MoreColor
LessColor:
    ; Get ABS eax & subtract
    neg ecx
    js LessColor

MoreColor:
    pxor mm2,mm2
    movd mm2,ecx    ; mm2 = | | | | |0|0|0|F| 

    cmp eax,3
    jg TestG

    psllq mm2,16    ; mm2 = | | | | |0|F|0|0| Red
    jmp DoColor
TestG:
    cmp eax,4
    jg DoColor

    psllq mm2,8     ; mm2 = | | | | |0|0|F|0| Green

    ;eax=5  B       ; mm2 = | | | | |0|0|0|F| Blue

DoColor:

    mov esi,ptrmemPic
    mov edi,ptrmemBack  ; Des
    
    xor edx,edx
    mov eax,H
    mov ebx,W
    mul ebx
    mov ecx,eax

ForRGB:
    push ecx
    ;---------    
    mov eax,FilterParam
    cmp eax,0
    jg IncRGB
DecRGB: ; RRB - F
    mov eax,[esi]
    cmp eax,TransColor
    je NexRGB
    
    movd mm1,eax            ; | |ARGB|
    psubusb mm1,mm2         ; eg Red | |ARGB| - | |0F00|
    movd[edi],mm1
    jmp NexRGB

IncRGB: ; RGB + F
    mov eax,[esi]
    cmp eax,TransColor
    je NexRGB

    movd mm1,[esi]          ; | |ARGB|
    paddusb mm1,mm2         ; eg Red | |ARGB| + | |0F00|
    movd[edi],mm1

    ;---------    
NexRGB:
    add esi,4
    add edi,4
    pop ecx
    dec ecx
    jnz near ForRGB
ColorRET:
RET
;====================================================

%define Seed     [ebp-56]
%define arand    [ebp-60]
%define ix       [ebp-64]
%define iy       [ebp-68]
%define incx     [ebp-72]
%define incy     [ebp-76]
%define a255     [ebp-80]
%define ahalf    [ebp-84]

Diffuse:

    mov eax,[ebp+16]
    mov Seed,eax
    
    mov eax,255
    mov a255,eax
    
    mov ebx,[ebp+12]    ; ebx=ptrMC3
    mov eax,Valu
    add ebx,eax
    mov eax,[ebx]
    mov ahalf,eax

    mov eax,W
    dec eax
    mov ix,eax
    
    mov ecx,H
ForDiffY:
    push ecx
    mov iy,ecx


    mov ecx,W
ForDiffX:
    push ecx
    mov ix,ecx
    
    ; Test if TransColor
    mov edi,ptrmemBack
    mov eax,iy
    dec eax
    mov ebx,W
    mul ebx
    mov ebx,ix
    dec ebx
    add eax,ebx
    shl eax,2
    add edi,eax

    mov eax,[edi]
    cmp eax,TransColor
    je near NexDiffX
    

    call RANDOMIZE      ; aran & eax = 0-255
    fild dword arand
    fild dword a255
    fdivp st1           ; rnd 0-1
    fld dword ahalf
    fsubp st1           ; rnd-.5
    
    fild dword FilterParam
    fmulp st1

    fistp dword incy

    ; Test if 1->H
    mov eax,incy
    mov ebx,iy
    add eax,ebx     ;iy+incy
    
    cmp eax,1
    jge TincyGTW
    xor eax,eax
    jmp SetincyZero
TincyGTW:
    cmp eax,H
    jle Safeincy
SetincyZero:
    xor eax,eax
    mov incy,eax
Safeincy:    

    call RANDOMIZE      ; aran & eax = 0-255
    fild dword arand
    fild dword a255
    fdivp st1           ; rnd 0-1
    fld dword ahalf
    fsubp st1            ; rnd-.5
    fild dword FilterParam
    fmulp st1
    fistp dword incx
    
    ;Test incx 1->W
    mov eax,incx
    mov ebx,ix
    add eax,ebx     ;ix+incx
    
    cmp eax,1
    jge TincxGTW
    jmp SetincxZero
TincxGTW:
    cmp eax,W
    jle Safeincx
SetincxZero:
    xor eax,eax
    mov incx,eax
Safeincx:    

    ; edi->memBack(ix,iy)
    push edi
    pop esi ; esi->memBack(ix,iy)

    mov eax,incy
    mov ebx,W
    imul ebx
    mov ebx,4
    imul ebx            ; 4*W*incy
    add esi,eax
    mov eax,incx
    imul ebx            ; 4*incx
    add esi,eax     ; esi-> memBack(ix+incx,iy+incy)
    
    mov eax,[esi]
    mov [edi],eax
    ;movsd       ; [esi]->[edi

NexDiffX:
    pop ecx
    dec ecx
    jnz near ForDiffX
NexDiffY:
    pop ecx
    dec ecx
    jnz near ForDiffY

RET
;====================================================

%define FourFPS  [ebp-88]

Relief:     ;Metallic

    mov eax,W
    shl eax,2       ; 4W
    mov BWidth,eax

    mov eax,FilterParam
    mov edx,eax ; 0 0 0 F
    shl edx,8   ; 0 0 F 0
    add eax,edx ; 0 0 F F
    mov edx,eax ; 0 0 F F
    shl edx,16  ; F F 0 0
    add eax,edx ; F F F F
    mov dword FourFPS,eax

    mov ecx,H
    dec ecx
ForRELY:
    push ecx
    mov iy,ecx

    mov ecx,W
    dec ecx
ForRELX:
    push ecx
    mov ix,ecx

    ; Test if TransColor
    mov esi,ptrmemPic
    mov eax,iy
    dec eax
    mov ebx,W
    mul ebx
    mov ebx,ix
    dec ebx
    add eax,ebx
    shl eax,2
    add esi,eax

    mov eax,[esi]
    cmp eax,TransColor
    je near NexRELX

    mov ebx,BWidth
    Call SumRELIEF  ; In: esi->  memPic(ix,iy), ebx=4W
                    ; mm0 = (memPic(ix,iy) + RELSUM [A% R% G% B%])\2
                    ; used mm0->mm6 also mm7=0  esi @ TL

    mov eax,dword FourFPS    
    movd mm1,eax
    punpcklbw mm1,mm7

    paddsw mm0,mm1      ; mm0 = (memPic(ix,iy) + RELSUM [A% R% G% B%])\2
                        ;       + [FilterParams]
    packuswb mm0,mm7    ; | |ARGB|  

    mov edi,ptrmemBack
    mov eax,iy
    dec eax
    mov ebx,W
    mul ebx
    mov ebx,ix
    dec ebx
    add eax,ebx
    shl eax,2
    add edi,eax     ; edi-> memBack(ix,iy)

    movd[edi],mm0

NexRELX:
    pop ecx
    dec ecx
    cmp ecx,1
    ja near ForRELX

NexRELY:
    pop ecx
    dec ecx
    cmp ecx,1
    ja near ForRELY


RET
;====================================================

%define im   [ebp-92]
%define ic   [ebp-96]
%define iB   [ebp-100]
%define iG   [ebp-104]
%define iR   [ebp-108]
%define iBGR   [ebp-112]
%define a256   [ebp-116] 


HShade: ;9
    mov eax,FilterParam
    shl eax,2
HIsAbs:
    neg eax
    js HIsAbs
    cmp eax,255
    jle HStim
    mov eax,255
HStim:    
    mov im,eax      ; im = Abs(4*FilterParam)
	;------
    mov eax,256
	mov a256,eax
	;------
    mov ecx,H
HFORiys:
    mov iy,ecx 
    push ecx
    
    mov ecx,W
HFORixs:
    push ecx
    mov ix,ecx
    ;--------------------------
    ; Calc ic = Sgn(FilterParam) * (512 * ixs/maxw -256)
    mov eax,ix
	mov ebx,512
	mul ebx
	mov ebx,W
	;xor edx,edx
	div ebx
	mov ebx,256
	sub eax,ebx
	mov edx,FilterParam
	cmp edx,0
	jge HPos
	neg eax
HPos:
	cmp eax,256
	jle HTZero
	mov eax,256		; >256 make = 256
	jmp HSTic
HTZero:
	cmp eax,0
	jge HSTic
	xor eax,eax		; < 0 make = 0
HSTic:
	mov ic,eax
	;------

    Call GetNewBGR

    ;--------------------------
HNexixs:
    pop ecx
    dec ecx
    jnz near HFORixs
HNexiys:
    pop ecx
    dec ecx
    jnz near HFORiys

RET
;====================================================

VShade: ;10
    mov eax,FilterParam
    shl eax,2
VIsAbs:
    neg eax
    js VIsAbs
    cmp eax, 255
    jle VStim
    mov eax,255
VStim:    
    mov im,eax      ; im = Abs(4*FilterParam)
    
    mov eax,256
	mov a256,eax
	;------
    mov ecx,W
VFORixs:
    mov ix,ecx 
    push ecx
    
    mov ecx,H
VFORiys:
    push ecx
    mov iy,ecx
    ;--------------------------
    ; Calc ic
    mov eax,iy
	mov ebx,512
	mul ebx
	mov ebx,H
	xor edx,edx
	div ebx
	mov ebx,256
	sub eax,ebx
	mov edx,FilterParam
	cmp edx,0
	jge VPos	
	neg eax
VPos:
	cmp eax,256
	jle VTZero
	mov eax,dword 256
	jmp VSTic
VTZero:
	cmp eax,0
	jge VSTic
	xor eax,eax
VSTic:
	mov ic,eax
	;------

    Call GetNewBGR

    ;--------------------------
VNexiys:
    pop ecx
    dec ecx
    jnz near VFORiys
VNexixs:
    pop ecx
    dec ecx
    jnz near VFORixs
RET
;====================================================

GetNewBGR:  ;Used by HShade & VShade
	mov edi,ptrmemBack
    mov eax,iy
    dec eax
    mov ebx,W
    mul ebx
    mov ebx,ix
    dec ebx
    add eax,ebx
    shl eax,2
    add edi,eax     ; edi-> memBack(ix,iy)
	;------
	mov ecx,[edi]	; ARGB  xx xx ah al

	xor edx,edx
	mov DL,CL		; iB
	mov iBGR,edx

	call GetBGR
	mov iB,eax

	xor edx,edx
	mov DL,CH		; iG
	mov iBGR,edx

	call GetBGR
	mov iG,eax

	bswap ecx
	xor edx,edx
	mov DL,CH		; iR
	mov iBGR,edx

	call GetBGR
	mov iR,eax
;((
	mov eax,iR
	shl eax,8
	or eax,iG
	shl eax,8
	or eax,iB

	cmp eax,TransColor
	jne STNew
	cmp eax,0
	jg GThanZ
	inc eax			; <= 0 inc blue (lob) by 1
	jmp STNew
GThanZ:
	dec eax			; > 0  dec blue (lob) by 1
STNew:
	
	mov[edi],eax

RET
;====================================================

GetBGR:		; In: edx = B,G or R Out: eax New B,G or R
	
	fild dword ic
	fild dword iBGR
	fsubp st1			; ic-iB,G,R
	fild dword im
	fmulp st1			; im*(ic-iB,G,R)
	fild dword a256
	fdivp st1			; im*(ic-iB,G,R)\256
	fistp dword iBGR	; iBGR = im*(ic-iB,G,R)
	mov eax,iBGR
	add eax,edx			; im*(ic-iB,G,R)\256 + iB,G,R
	cmp eax,0
	jge HVT255
	xor eax,eax
	jmp HVSTibgr
HVT255:
	cmp eax,255
	jle HVSTibgr
	mov eax,255
HVSTibgr:

RET
;====================================================

SumRELIEF:  ; esi->memPic(ix,iy)
            ; ebx = 4W

    pxor mm7,mm7
    
    movd mm0,[esi]

    mov eax,4
    sub esi,eax
    sub esi,ebx     ; esi-> -1,-1

    mov eax,[ebp+20]    ; Pick up Opcode
    cmp eax,7
    je REL
    ;je Engrave


    ; Metallic
    movd mm1,[esi]
    movd mm2,[esi+4]
    movd mm3,[esi+12]
    movd mm4,[esi+20]
    movd mm5,[esi+28]
    movd mm6,[esi+32]
    jmp UNPACK

REL:
    ; Relief
    movd mm1,[esi]
    movd mm2,[esi+4]
    movd mm3,[esi+ebx]
    movd mm4,[esi+ebx+8]
    movd mm5,[esi+2*ebx+4]
    movd mm6,[esi+2*ebx+8]
    

UNPACK:

    punpcklbw mm0,mm7
    punpcklbw mm1,mm7
    punpcklbw mm2,mm7
    punpcklbw mm3,mm7
    punpcklbw mm4,mm7
    punpcklbw mm5,mm7
    punpcklbw mm6,mm7

    psllw mm1,1         ; *2
    psllw mm6,1         ; *2

    paddsw mm0,mm1      ; +2 
    paddsw mm0,mm2      ; +2 +1
    paddsw mm0,mm3      ; +2 +1 +1 
    psubsw mm0,mm4      ; +2 +1 +1 -1
    psubsw mm0,mm5      ; +2 +1 +1 -1 -1
    psubsw mm0,mm6      ; +2 +1 +1 -1 -1 -2
    psraw mm0,1         ; \2

    ;mm0 = A% R% G% B% lob

RET

Engrave:    ;Not used

    movd mm1,[esi+8]
    movd mm2,[esi+ebx+8]
    movd mm3,[esi+2*ebx+8]
    movd mm4,[esi]
    movd mm5,[esi+ebx]
    movd mm6,[esi+2*ebx]
    
    punpcklbw mm0,mm7
    punpcklbw mm1,mm7
    punpcklbw mm2,mm7
    punpcklbw mm3,mm7
    punpcklbw mm4,mm7
    punpcklbw mm5,mm7
    punpcklbw mm6,mm7

    paddsw mm0,mm1      ; +1 
    paddsw mm0,mm2      ; +1 +1
    paddsw mm0,mm3      ; +1 +1 +1 
    psubsw mm0,mm4      ; +1 +1 +1 -1 
    psubsw mm0,mm5      ; +1 +1 +1 -1 -1
    psubsw mm0,mm6      ; +1 +1 +1 -1 -1 -1
    psraw mm0,1         ; \2

RET
;====================================================

;Used by Sharpness & Softness
Sum8:   ; In esi->BotLeft, ebx BWidth
        ; Out: mm0 = A% R% G% B% lob
        ; Called by Sharpness & Softness
    pxor mm7,mm7

    ; Row 1
    movd mm0,[esi]
    movd mm1,[esi+4]
    movd mm2,[esi+8]
    punpcklbw mm0,mm7
    punpcklbw mm1,mm7
    punpcklbw mm2,mm7
    paddusw mm0,mm1
    paddusw mm0,mm2
    
    ; Row 2
    movd mm1,[esi+ebx]
    movd mm2,[esi+ebx+8]
    punpcklbw mm1,mm7
    punpcklbw mm2,mm7
    paddusw mm0,mm1
    paddusw mm0,mm2
    
    
    ; Row 3
    movd mm1,[esi+2*ebx]
    movd mm2,[esi+2*ebx+4]
    movd mm3,[esi+2*ebx+8]
    punpcklbw mm1,mm7
    punpcklbw mm2,mm7
    punpcklbw mm3,mm7
    paddusw mm0,mm1
    paddusw mm0,mm2
    paddusw mm0,mm3 ; mm0 = A% R% G% B% lob
RET
;====================================================
;============================================================
RANDOMIZE:      ; Out: aL & Seed = rand(0-255)
    mov eax,011813h     ; 71699 prime 
    imul DWORD Seed
    add eax, 0AB209h    ; 700937 prime
    rcr eax,1           ; leaving out gives vertical lines plus
                        ; faint horizontal ones, tartan

    ;----------------------------------------
    ;jc ok              ; these 2 have little effect
    ;rol eax,1          ;
ok:                     ;
    ;----------------------------------------
    
    ;----------------------------------------
    ;dec eax            ; these produce vert lines
    ;inc eax            ; & with fsin marble arches
    ;----------------------------------------

    mov Seed,eax    ; save seed
    and eax,255
    mov arand,eax    ; arand = rnd(0-255)
RET

;============================================================
[SECTION .data]
Valu dd 0.5

