
	.386
	.model flat,stdcall
	option casemap:none
	option proc:private

INSIDE_CENUMFORMATETC equ 1

	include COMView.inc
	include classes.inc
	include debugout.inc

CEnumFORMATETC struct
_IEnumFORMATETC	IEnumFORMATETC  <>
dwRefCount	DWORD ?
pUnkRef		LPUNKNOWN ?		; IUnknown for ref counting
iCur		dd ?			; Current element
cntFE		dd ?			; Number of FORMATETCs in us
prgFE		dd ?			; Source of FORMATETCs
CEnumFORMATETC ends

QueryInterface proto :ptr CEnumFORMATETC, :REFIID, :ptr LPUNKNOWN
AddRef	proto :ptr CEnumFORMATETC
Release proto :ptr CEnumFORMATETC

	.const

;--- vtable

vtblIEnumFORMATETC label dword
	IUnknownVtbl {QueryInterface, AddRef, Release}
    dd Next, Skip, Reset, Clone

iftab label dword
	dd IID_IUnknown				, 0
	dd IID_IEnumFORMATETC		, CEnumFORMATETC._IEnumFORMATETC
NUMIFENTRIES textequ %($ - offset iftab) / (4 * 2)

	.code

__this	textequ <ebx>
_this	textequ <[__this].CEnumFORMATETC>

	MEMBER _IEnumFORMATETC, dwRefCount, pUnkRef, iCur, cntFE, prgFE


Create@CEnumFORMATETC proc public uses __this esi edi pUnkRef:LPUNKNOWN, cntFE:DWORD, prgFE:ptr FORMATETC

	invoke malloc, sizeof CEnumFORMATETC
	.if (!eax)
		ret
	.endif
	mov __this, eax
	mov m__IEnumFORMATETC.lpVtbl, offset vtblIEnumFORMATETC
	mov m_dwRefCount, 1

	mov eax,pUnkRef
	mov m_pUnkRef,eax

	mov m_iCur, 0

	mov eax, cntFE
	mov m_cntFE,eax

;--------------------------------- Allocate the array of FORMATETC
	mov ecx,SIZEOF FORMATETC
	imul ecx
	mov edi,eax 				;number of FORMATETC * SIZEOF FORMATETC
	invoke malloc, eax
	.if (!eax)
		invoke free, __this
		return 0
	.endif
	mov m_prgFE,eax
;--------------------------------- initialise the array of FORMATETC
	mov ecx, edi
	mov edi,eax
	mov esi, prgFE
	rep movsb

	invoke vf(m_pUnkRef, IUnknown, AddRef)

	return __this
	align 4

Create@CEnumFORMATETC endp


;--- IUnknown methods


QueryInterface proc this_:ptr CEnumFORMATETC, riid:REFIID, ppReturn:ptr LPUNKNOWN

	invoke IsInterfaceSupported, riid, offset iftab, NUMIFENTRIES,  this_, ppReturn
	ret

QueryInterface endp


AddRef proc this_:ptr CEnumFORMATETC

	mov ecx, this_
	inc [ecx].CEnumFORMATETC.dwRefCount
	return [ecx].CEnumFORMATETC.dwRefCount

AddRef endp


Release proc uses __this this_:ptr CEnumFORMATETC

	mov __this, this_
	dec m_dwRefCount
	mov eax, m_dwRefCount
	.if (!eax)
;---------------------------- free object data & object itself
		invoke vf(m_pUnkRef, IUnknown, Release)
		invoke free, m_prgFE
		invoke free, __this
		xor eax, eax
	.endif
	ret

Release endp


;--- Returns the next element in the enumeration


Next proc uses __this esi edi, this_:ptr CEnumFORMATETC, cntFE:DWORD, pFE:ptr FORMATETC, pulFE:ptr DWORD

Local       cntReturn:DWORD

		mov __this,this_

		mov eax, pulFE
		.if (eax)
			mov DWORD PTR [eax],0
		.endif

		.if ((!m_prgFE) || (!pFE))
			jmp error
		.endif

		mov eax, m_iCur
		mov ecx, m_cntFE
		cmp eax,ecx
		jnb error
		mov cntReturn, 0
		mov edi, pFE
		mov esi, m_prgFE
		mov eax, m_iCur
		xor edx,edx
		mov ecx,SIZEOF FORMATETC
		imul ecx
		add esi,eax

		mov edx, m_cntFE
		.WHILE ((m_iCur < edx) && (cntFE > 0))
			mov ecx,SIZEOF FORMATETC
			rep movsb
			inc m_iCur
			inc cntReturn
			dec cntFE
		.ENDW
		mov eax, pulFE
		.if (eax)
			mov ecx, cntReturn
			sub ecx, cntFE
			mov DWORD PTR [eax],ecx
		.endif
		return S_OK
error:
		return S_FALSE
		align 4

Next endp

;--- Skips the next n elements in the enumeration

Skip proc uses __this this_:ptr CEnumFORMATETC, cSkip:DWORD

		mov __this, this_
		mov eax, m_prgFE
		or eax,eax
		jz error
		mov eax, m_iCur
		add eax, cSkip
		cmp eax, m_cntFE
		jnb error
		mov m_iCur, eax
		return S_OK
error:
		return S_FALSE

Skip endp

;--- Resets the current element index in the enumeration to zero

Reset proc this_:ptr CEnumFORMATETC
		mov eax, this_
		mov [eax].CEnumFORMATETC.iCur,0
		return S_OK
Reset endp

;--- Returns another IEnumFORMATETC with the same state as ourselves

Clone proc this_:ptr CEnumFORMATETC, ppEnum:ptr ptr IEnumFORMATETC
		return E_FAIL
Clone endp

	end

