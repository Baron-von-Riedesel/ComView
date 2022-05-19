
	.386
	.model flat,stdcall
	option casemap:none
	option proc:private

INSIDE_CDROPSOURCE equ 1

	include COMView.inc
	include classes.inc
	include debugout.inc

CDropSource struct
_IDropSource	IDropSource  <>
dwRefCount		DWORD ?
CDropSource ends

QueryInterface proto :ptr CDropSource, :REFIID, :ptr LPUNKNOWN
AddRef	proto :ptr CDropSource
Release proto :ptr CDropSource

	.const

vtblIDropSource label dword
	IUnknownVtbl {QueryInterface, AddRef, Release}
    dd      offset QueryContinueDrag
    dd      offset GiveFeedback

iftab label dword
	dd IID_IUnknown				, 0
	dd IID_IDropSource			, CDropSource._IDropSource
NUMIFENTRIES textequ %($ - offset iftab) / (4 * 2)

	.code

__this	textequ <ebx>
_this	textequ <[__this].CDropSource>

	MEMBER _IDropSource, dwRefCount

;//**********************************************************************
;// Implementation of the IDropSource interface
;//**********************************************************************
;-----------------------------------------------------------------------------
;Constructor
;
Create@CDropSource proc public uses __this

	DebugOut "Create@CDropSource"

	invoke malloc, sizeof CDropSource
	.if (eax == 0)
		ret
	.endif

	mov __this, eax
	mov m__IDropSource.lpVtbl, offset vtblIDropSource
	mov m_dwRefCount, 1
	return __this
	align 4

Create@CDropSource endp

;-----------------------------------------------------------------------------

QueryInterface proc this_:ptr CDropSource, riid:REFIID, ppReturn:ptr LPUNKNOWN

	DebugOut "CDropSource::QueryInterface"
	invoke IsInterfaceSupported, riid, offset iftab, NUMIFENTRIES,	this_, ppReturn
	ret
	align 4

QueryInterface endp

AddRef proc this_:ptr CDropSource

	mov ecx, this_
	inc [ecx].CDropSource.dwRefCount
	mov eax, [ecx].CDropSource.dwRefCount
	ret
	align 4

AddRef endp

Release proc uses __this this_:ptr CDropSource

	mov __this, this_
	dec m_dwRefCount
	mov eax, m_dwRefCount
	.if (!eax)
		invoke free, __this
		DebugOut "CDropSource destroyed"
		xor eax, eax
	.endif
	ret
	align 4
Release endp

;-----------------------------------------------------------------------------
;//  Determines whether to continue a drag operation or cancel it.

QueryContinueDrag proc this_:ptr CDropSource, fEsc:DWORD, grfKeyState:DWORD

	mov eax, fEsc
	test eax, eax
	jz @F
	mov eax, DRAGDROP_S_CANCEL	 ;to stop the drag
	jmp exit
@@:
	mov eax, grfKeyState
	and eax, MK_LBUTTON or MK_RBUTTON
	test eax,eax
	jnz @F
	mov eax, DRAGDROP_S_DROP	;to drop the data where it is
	jmp exit
@@:
	mov eax, S_OK
exit:
	DebugOut "CDropSource::QueryContinueDrag=%X", eax
	ret
	align 4
QueryContinueDrag endp

;//  Provides cursor feedback to the user

GiveFeedback proc this_:ptr CDropSource, dwEffect:DWORD

	mov eax, DRAGDROP_S_USEDEFAULTCURSORS
	DebugOut "CDropSource::GiveFeedback=%X", eax
	ret
	align 4

GiveFeedback endp

	end

