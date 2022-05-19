
;*** implements a malloc spy
;*** IMPORTANT: keep this file independant from COMView

	.386
	.model flat,stdcall
	option casemap :none   ; case sensitive
	option proc:private

ifdef _DEBUG

	.nolist
	.nocref
WIN32_LEAN_AND_MEAN	equ 1
INCL_OLE2			equ 1
	include windows.inc
	include commctrl.inc
	include objidl.inc
	include windowsx.inc

	include debugout.inc
	include macros.inc
	.list
	.cref

malloc	proto dwBytes:DWORD
free	proto pvoid:ptr byte

CMallocSpy struct
MallocSpy				IMallocSpy <>
dwRefCount				dd ?
CMallocSpy ends

	MEMBER MallocSpy, dwRefCount

	.const

CMallocSpyVtbl label dword
	IUnknownVtbl {QueryInterface, AddRef, Release}
    dd PreAlloc
	dd PostAlloc
	dd PreFree
	dd PostFree
	dd PreRealloc
	dd PostRealloc
	dd PreGetSize
	dd PostGetSize
	dd PreDidAlloc
	dd PostDidAlloc
	dd PreHeapMinimize
	dd PostHeapMinimize

iftab label dword
	dd IID_IUnknown				, CMallocSpy.MallocSpy
	dd IID_IMallocSpy			, CMallocSpy.MallocSpy
NUMIFENTRIES equ ($ - offset iftab) / (4 * 2)

g_szNull	db 0

__this	textequ <ebx>
_this	textequ <[__this].CMallocSpy>
thisarg equ <this_:LPMALLOC>

	.code

Create@CMallocSpy proc public uses __this

	DebugOut "Create@CMallocSpy"

	invoke malloc, sizeof CMallocSpy
	.if (!eax)
		ret
	.endif
	mov __this, eax
	mov m_MallocSpy.lpVtbl,		offset CMallocSpyVtbl
	mov m_dwRefCount, 1
	return __this

Create@CMallocSpy endp

Destroy@CMallocSpy proc uses __this this_:ptr CMallocSpy

	DebugOut "Destroy@CMallocSpy enter"
	mov __this,this_
	invoke free, __this
	DebugOut "Destroy@CMallocSpy exit"
	ret
Destroy@CMallocSpy endp

;--------------------------------------------------------------
;--- interface IUnknown
;--------------------------------------------------------------

AddRef proto :ptr CMallocSpy

QueryInterface proc uses esi edi __this this_:ptr CMallocSpy, riid:ptr IID, ppReturn:ptr ptr

ifdef _DEBUG
local	wszIID[40]:word
local	szKey[128]:byte
local	dwSize:DWORD
local	hKey:HANDLE
endif

	mov __this,this_

	mov esi, offset iftab
	mov edx, NUMIFENTRIES
	.while (edx)
		lodsd
		xchg eax, esi
		mov edi, riid
		mov ecx, 4
		repz cmpsd
		xchg eax, esi
		.break .if (ZERO?)
		add esi, sizeof DWORD
		dec edx
	.endw 

	.if (edx)
		invoke AddRef, __this
		lodsd
		lea edx, [__this+eax]
		mov eax, S_OK
	.else
		xor edx, edx
		mov eax, E_NOINTERFACE
	.endif
	mov ecx, ppReturn
	mov [ecx], edx

ifdef _DEBUG
;--------------------- print the name of the interface we have just been queried
	push eax
	invoke StringFromGUID2,riid, addr wszIID,40
	invoke wsprintf, addr szKey, CStr("Interface\%s"), addr wszIID
	invoke RegOpenKeyEx, HKEY_CLASSES_ROOT, addr szKey, 0, KEY_READ, addr hKey 
	.if (eax == ERROR_SUCCESS)
		mov dwSize, sizeof szKey
		invoke RegQueryValueEx,hKey,addr g_szNull,NULL,NULL,addr szKey,addr dwSize
		invoke RegCloseKey, hKey
	.else
		mov szKey,0
	.endif
	pop eax
	DebugOut "CMallocSpy::QueryInterface(%S[%s])=%X", addr wszIID, addr szKey, eax
endif
	ret

QueryInterface endp


AddRef proc uses __this this_:ptr CMallocSpy

	mov __this,this_
	inc m_dwRefCount
	mov eax, m_dwRefCount
	ret

AddRef endp

Release proc uses __this this_:ptr CMallocSpy

	mov __this,this_
	dec m_dwRefCount

	mov eax, m_dwRefCount
	.if (eax == 0)
		invoke Destroy@CMallocSpy, __this
		xor eax,eax
	.endif
	ret

Release endp

PreAlloc proc thisarg, cbRequest:dword
	DebugOut "PreAlloc called"
	return cbRequest
PreAlloc endp

PostAlloc proc thisarg, pActual:ptr 
	DebugOut "PostAlloc called"
	return pActual
PostAlloc endp

PreFree proc thisarg, pRequest:ptr, fSpyed:BOOL
	DebugOut "PreFree called"
	ret
PreFree endp

PostFree proc thisarg, fSpyed:BOOL
	DebugOut "PostFree called"
	ret
PostFree endp

PreRealloc proc thisarg, pRequest:ptr, cbRequest:dword,ppNewRequest:ptr ptr, fSpyed:BOOL
	return cbRequest
PreRealloc endp

PostRealloc proc thisarg, pActual:ptr, fSpyed:BOOL
	return pActual
PostRealloc endp

PreGetSize proc thisarg, pRequest:ptr, fSpyed:BOOL
	return pRequest
PreGetSize endp

PostGetSize proc thisarg, cbActual:DWORD, fSpyed:BOOL
	return cbActual
PostGetSize endp

PreDidAlloc proc thisarg, pRequest:ptr, fSpyed:BOOL
	return pRequest
PreDidAlloc endp

PostDidAlloc proc thisarg, pRequest:ptr, fSpyed:BOOL, fActual:DWORD
	return fActual
PostDidAlloc endp

PreHeapMinimize proc thisarg
	ret
PreHeapMinimize endp

PostHeapMinimize proc thisarg
	ret
PostHeapMinimize endp

endif

	end
