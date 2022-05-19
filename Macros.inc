
;--- macros:
;--- CStr()
;--- L()
;--- CStrW()
;--- smalloc, sfree, sreload
;--- return
;--- @mov
;--- MEMBER

ifndef CStr

;--- CStr() define a string in .CONST
;--- or in .CONST$2 if .CONST is the current section

CStr macro text:VARARG
local sym
ifidni @CurSeg,<CONST>
CONST$2 segment FLAT public 'CONST'
else
CONST segment FLAT public 'CONST'
endif
ifidni <text>,<"">
    sym db 0
else
    sym db text,0
endif
@CurSeg ends
	exitm <offset sym>
	endm

endif

ifndef L

;---- L() defines a wide string
;---- usage: StringName dw L(stringvalue)
if 0
L macro y:req
local x
x		textequ <>
		forc chr$,<y>
x		CatStr x,<'&chr$'>,<,>
		endm
x		CatStr x,<0>
		exitm <x>
		endm

endif	;L()
else
L macro parms:VARARG
local wstr,i,c,tstr
    wstr textequ <>
    i = 0
    for parm,<parms>
      c SubStr <parm>,1,1
      ifidn c,<">
        tstr SubStr <parm>,2,@SizeStr(parm)-2    
%       forc chr$,<tstr>
          if i
            wstr CatStr wstr,<,>
          endif
          wstr CatStr wstr,<'&chr$'>
          i = i + 1      
        endm
      else  
        if i
          wstr CatStr wstr,<,>
        endif
        wstr CatStr wstr,<parm>
      endif
    endm
    exitm <wstr>
endm	
endif

;--- CStrW defines a wide string in .CONST
;--- returns a pointer to that string

CStrW macro text:VARARG
local sym
ifidni @CurSeg,<CONST>
CONST$2 segment FLAT public 'CONST'
else
CONST segment FLAT public 'CONST'
endif
	align 2
sym dw text,0
@CurSeg ends
	exitm <offset sym>
	endm


ifndef smalloc

;--- the smalloc + sfree macros are used to alloc space
;--- on the stack.
;--- usage: "smalloc register,numBytes" and "sfree"

smalloc macro reg:req,bytes:req	;alloc space on the stack (local only)
local ?bytes
?bytes = (bytes + 3) and 0fffffffch
	sub esp,?bytes
	mov reg,esp
	push ?bytes+4
	endm
sreload macro reg:req,index		;reload address of stack items
ifnb <index>
	mov reg,esp
	repeat index
	add reg,[reg]
	endm
	add reg,4
else
	lea reg,[esp+4]
endif
	endm
sfree macro
	add esp,[esp]
	endm

endif	;smalloc


;--- return: return a value in eax

return macro x
ifnb <x>
 if (OPATTR x) and 4	;;constant?
  if x
	mov eax,x
  else
	xor eax,eax
  endif
 else
	mov eax,x
 endif
endif
	ret
	endm

;--- simple macro for a 3 byte move, used i.e.: @mov ecx, 3

@mov macro x, y
	push y
	pop x
	endm

ifndef MEMBER

;--- MEMBER: create member names
;--- requires at least one other definition in program itself.
;--- assume ebx is used to access this_. so define in program
;--- "_this	textequ <[ebx].CClassName>"
;--- "MEMBER VarName"
;--- so access to member 
;--- "[ebx].CClassName.VarName" simplifies to "m_VarName"

MEMBER	macro names:VARARG
local x
	for y,<names>
x	CatStr <_this.>,<y>
m_&y	textequ x
	endm
	endm

endif	;MEMBER



