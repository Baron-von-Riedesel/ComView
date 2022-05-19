
;*** classes CConnection ***

	.386
	.model flat,stdcall
	option casemap :none   ; case sensitive
	option proc:private

	include COMView.inc
INSIDE_CCONNECTION equ 1
	include classes.inc
	include rsrc.inc
	include CEditDlg.inc

SafeRelease	proto :LPUNKNOWN

BEGIN_CLASS CConnection
pConnectionPoint	LPCONNECTIONPOINT	?
pEventSink			pCEventSink			?
iid					IID					<>
dwCookie			DWORD				?
pObjectItem			pCObjectItem		?
END_CLASS

__this	textequ <ebx>
_this	textequ <[__this].CConnection>

	MEMBER pConnectionPoint, pObjectItem, iid, pEventSink, dwCookie

	.code

Create@CConnection proc public uses __this pObjectItem:ptr CObjectItem, riid:ptr IID

	invoke malloc, sizeof CConnection
	.if (!eax)
		jmp exit
	.endif
	mov __this, eax
	pushad
	mov esi, riid
	lea edi, m_iid
	movsd
	movsd
	movsd
	movsd
	popad
	mov eax, pObjectItem
	mov m_pObjectItem, eax
exit:
	return __this

Create@CConnection endp

IsEqualGUID@CConnection proc public this_:ptr CConnection, riid:REFIID
	mov ecx, this_
	invoke IsEqualGUID, addr [ecx].CConnection.iid, riid
	ret
IsEqualGUID@CConnection endp

ifdef @StackBase
	option stackbase:ebp
endif
	option prologue:@sehprologue
	option epilogue:@sehepilogue

Disconnect@CConnection proc public uses esi edi __this  this_:ptr CConnection, hWnd:HWND

local hr:DWORD

	nop
	.try

	mov __this, this_
	mov hr, S_OK
	xor esi, esi
	xchg esi, m_pConnectionPoint
	.if (esi)
		xor edi, edi
		xchg edi, m_dwCookie
		.if (edi)
			invoke vf(esi, IConnectionPoint, Unadvise), edi
			mov hr, eax
		.endif
		invoke vf(esi, IUnknown, Release)
	.endif

	.exceptfilter
		mov __this,this_	;reload this register
		mov eax, _exception_info()
		invoke DisplayExceptionInfo, hWnd, eax, CStr("Disconnect"), EXCEPTION_EXECUTE_HANDLER
	.except
		mov __this,this_	;reload this register
		mov hr, E_UNEXPECTED
	.endtry

	invoke SafeRelease, m_pEventSink
	mov m_pEventSink, NULL

	return hr

Disconnect@CConnection endp


Connect@CConnection proc public uses esi edi __this this_:ptr CConnection, pUnknown:LPUNKNOWN, hWnd:HWND, ppszError:ptr LPSTR

local	hr:DWORD
local	pConnectionPointContainer:LPCONNECTIONPOINTCONTAINER

	mov pConnectionPointContainer, NULL

	.try

	mov __this, this_
	invoke vf(pUnknown, IUnknown, QueryInterface),
			addr IID_IConnectionPointContainer, addr pConnectionPointContainer
	mov hr, eax
	.if (eax == S_OK)
		invoke Disconnect@CConnection, __this, hWnd
		invoke vf(pConnectionPointContainer, IConnectionPointContainer, FindConnectionPoint),
				addr m_iid, addr m_pConnectionPoint
		mov hr, eax
		.if (eax == S_OK)
			invoke Create@CEventSink, addr m_iid, m_pObjectItem, NULL
			mov m_pEventSink, eax
			invoke vf(m_pConnectionPoint, IConnectionPoint, Advise),
					m_pEventSink, addr m_dwCookie
			.if (eax != S_OK)
				mov hr, eax
				invoke SafeRelease, m_pConnectionPoint
				mov m_pConnectionPoint, NULL
				mov ecx, CStr("IConnectionPoint::Advise failed[%X]")
			.endif
		.else
			mov ecx, CStr("FindConnectionPoint failed[%X]")
		.endif
	.else
		mov ecx, CStr("QueryInterface(IConnectionPointContainer) failed[%X]")
	.endif

	.exceptfilter
		mov __this,this_	;reload this register
		mov eax, _exception_info()
		invoke DisplayExceptionInfo, hWnd, eax, CStr("Connect"), EXCEPTION_EXECUTE_HANDLER
	.except
		mov __this,this_	;reload this register
		mov hr, E_UNEXPECTED
		mov ecx, CStr("Failure [%X]")
	.endtry

	.if (hr != S_OK)
		mov edx, ppszError
		mov [edx], ecx
	.endif

	invoke SafeRelease, pConnectionPointContainer

	return hr

Connect@CConnection endp

	option prologue: prologuedef
	option epilogue: epiloguedef
ifdef @StackBase
	option stackbase:esp
endif

Destroy@CConnection proc public uses __this this_:ptr CConnection
	mov __this, this_
	invoke Disconnect@CConnection, __this, NULL
	invoke free, __this
	ret
Destroy@CConnection endp

	end
