
;--- define macros for debugging output
;--- DebugOutNoLF: printf format OutputDebugStringA
;--- DebugOut: printf format OutputDebugStringA

ifndef _DEBUG
DEBUG = 0
else
DEBUG = 1
endif

DebugOutNoLF Macro x:req,y:VARARG
local sym,ii
if DEBUG
	.const
sym db x,0
	.code
	pushad
  ifdef DEBUGPREFIX
	invoke OutputDebugString,DEBUGPREFIX
	popad
	pushad
  endif
  ifnb <y>
	sub esp,256
	ii = 0
	for parname,<y>
	  ii = ii + 4
	endm
	push esp
	invoke wsprintf,[esp+ii+4],addr sym,y
	pop esp
	invoke OutputDebugString,esp
	add esp,256
  else
	invoke OutputDebugString,addr sym
  endif
	popad
endif
	endm

DebugOut macro x:req,y:VARARG
ifnb <y>
	DebugOutNoLF <x,0dh,0ah>,y
else
	DebugOutNoLF <x,0dh,0ah>
endif
	endm

