
;--------- macros for SEH in Win32 ----------

; the following macros are defined here:
; .try/.tryf        : start of try block (.except / .finally)
; .exceptfilter     : exception filter 
; .except           : exception block
; .finally          : finally block
; .endtry/.endtryf  : end of try block (.except / .finally)
;
; @sehprologue      : defines prologue with SEH
; @sehepilogue      ; defines epilogue with SEH
;
; _exception_code   ; usage : mov eax,_exception_code()
; _exception_info   ; usage : mov eax,_exception_info()
;
;   Note:
;   - for every .try block three dwords are defined in section .text$2  
;   - .try blocks can be nested
;   - if you jump into a .try block, the code is not guarded
;   - always define all parts (.try, .exceptfilter, .except, .endtry)
;     or (.tryf, .finally, .endtryf)                                   
;   _ on entries .exceptfilter, _except and _finally registers 
;     ebx,esi,edi are undefined. In consequence, if you want this
;     registers in no case to be changed by your function (as in C), include 
;     "uses ebx,esi,edi" for all procs containing .try blocks.
;   - Do not use "ret" inside an .exceptfilter block. And do not
;     jump out of such a block. Reason: .exceptfilter is called from 
;     the exception handler. The stack is not what you might expect.
;     EBP is correct, so access to locals is ok. At the end of the filter 
;     is a hidden retn to return to the handler. If you jump out of the 
;     block, no termination code (.finally) can be executed and registers
;     saved by "uses " will be lost.   
;   - _exception_code and _exception_info work only in _exceptfilter
;     and _except. _exception_code gets the ExceptionCode from the
;     EXCEPTION_RECORD (defined in windows.inc), _exception_info gets
;     a pointer to structure EXCEPTION_INFO_PTRS (see below)
;   - restriction: currently for every function there is a
;     limit of 10 try blocks. To overcome you will need to change
;     macros .try and .endtry/.endtryf    
;
;   thanks to Matt Pietrek and his article: 
;   "A crash course on the Depths of Win32 Structured Exception Handling"
;   in MSJ January 1997
;--------------------------------------------------------------------
;
; example how to use: 
;
;       OPTION PROLOGUE: @sehprologue
;       OPTION EPILOGUE: @sehepilogue
;
;       testseh proc uses ebx esi edi
;
;        .try
;            mov eax,0
;            mov byte ptr [eax],1	;causes exception
;        .exceptfilter	
;                         ; look for possible return values in windows.inc
;            mov eax,EXCEPTION_EXECUTE_HANDLER	;set rc for except handler
;        .except
;            ... do any cleanup
;        .endtry   
;
;        ret
;
;       testseh endp
;
;       OPTION PROLOGUE: prologuedef
;       OPTION EPILOGUE: epiloguedef
;--------------------------------------------------------------------

	includelib except.lib

	option dotname

__except_list equ <[0]>				;address of previous structure in TIB
externdef c _except_handler3:proc	;same name as in VC!

;--------------------------------------------------------------------

?trysuffix textequ <_>
?tryroot textequ @CatStr( ?try, %?trysuffix )
?tryindex = -1
?tryroot = ?tryindex

;* prologue for SEH activation

@sehprologue macro procname,flag,parmbyte,localbyte,reglist,userparms
local scopebeg

	push ebp						;[ebp+00]
	mov ebp,esp
	push -1							;[ebp-04] trylevel
	push offset scopebeg			;[ebp-08]
	push OFFSET _except_handler3	;[ebp-12]
	xor edx,edx
	db 64h
	push ds:[edx]					;[ebp-16]
	db 64h
	mov ds:[edx], esp				;[ebp-16]
	add esp,-localbyte - 3 * sizeof dword

	for r,reglist
		push r
	endm
	mov [ebp-24],esp				;save "normal" esp

?tryroot textequ @CatStr( ?try, %?trysuffix )
?tryindex = -1
?tryroot = ?tryindex

.text$2 segment dword public 'CODE'
if @Model ne 7
DGROUP group .text$2
endif
scopebeg:
.text$2 ends

	exitm %localbyte+ 4 * sizeof dword + 3 * sizeof dword
endm

;* epilogue for SEH activation

@sehepilogue macro procname,flag,parmbyte,localbyte,reglist,userparms

	mov ecx,[ebp-16]
	db 64h
	mov DWORD PTR ds:__except_list, ecx
	for r,reglist
		pop r
	endm
	mov esp, ebp
	pop ebp
	ret parmbyte
endm

;--------------------------------------------------------------------

;*** since _try blockes can be nested, for every level
;*** a variable ?tryroot is needed to save the index on .try and
;*** to be restored on .endtry 

__try macro xx
local lbl1,lbl2

?trysuffix textequ @CatStr( %?trysuffix,%?tryindex+1 )
lbl1 textequ @CatStr( ?testexc, %?trysuffix )
lbl2 textequ @CatStr( ?execexc, %?trysuffix )

.text$2 segment
	dd ?tryroot
ifb <xx>
	dd lbl1
else
	dd 0
endif
	dd lbl2
.text$2 ends

?tryroot textequ @CatStr( ?try, %?trysuffix )
?tryindex = ?tryindex + 1
?tryroot = ?tryindex

	mov dword ptr [ebp-4],?tryindex
	endm

;*** .try macro

.try macro
	__try
	endm

;*** .endtry macro

.endtry macro
local lbl

lbl textequ @CatStr( ?tryexit, %?trysuffix )
?trysuffix textequ @SubStr(%?trysuffix,0,@SizeStr(%?trysuffix) - 1)
?tryroot textequ @CatStr( ?try, %?trysuffix )

lbl:
	mov dword ptr [ebp-4],?tryroot
	endm


;*** .exceptfilter macro

.exceptfilter macro
local lbl1,lbl2

lbl1 textequ @CatStr( ?testexc, %?trysuffix )
lbl2 textequ @CatStr( ?tryexit, %?trysuffix )

	jmp lbl2	;; this jmp terminates preceeding _try
lbl1:
	endm

;*** .except macro: just restores esp

.except macro
local lbl
	retn		;; this ret is for preceeding _exceptfilter	
lbl textequ @CatStr( ?execexc, %?trysuffix )
lbl:
	mov esp,[ebp-24]
	endm

;---------------------------------------------------

.tryf  macro
	__try 1
	endm

;*** .finally macro

.finally macro
local lbl1,lbl2

lbl1 textequ @CatStr( ?execexc, %?trysuffix )
lbl2 textequ @CatStr( ?tryexit, %?trysuffix )

	push lbl2
lbl1:
	endm

;*** .endtryf macro

.endtryf macro
	retn
	.endtry
	endm

;*** the following structure is dynamically generated by
;*** exception handler code on the stack. it can be accessed in
;*** _exceptfilter and _except by _exception_info()

ifndef EXCEPTION_POINTERS
;--- PEXCEPTION_RECORD, PCONTEXT, EXCEPTION_POINTERS
;--- usually defined in WINNT.INC

PEXCEPTION_RECORD typedef ptr EXCEPTION_RECORD
PCONTEXT          typedef ptr CONTEXT

EXCEPTION_POINTERS struct
ExceptionRecord      PEXCEPTION_RECORD ?
ContextRecord        PCONTEXT          ?
EXCEPTION_POINTERS ends
endif

_exception_info macro 
	exitm <[ebp-20]>
	endm

_exception_code macro 
	mov eax, _exception_info()
	mov eax,[eax].EXCEPTION_POINTERS.ExceptionRecord
	exitm <[eax].EXCEPTION_RECORD.ExceptionCode>
	endm

;*** end of file

