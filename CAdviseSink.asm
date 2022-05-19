

;*** definition of class CAdviseSink

	.386
	.model flat,stdcall
	option casemap :none   ; case sensitive
	option proc:private

	include COMView.inc
INSIDE_CADVISESINK equ 1
	include classes.inc
	include rsrc.inc
	include debugout.inc

?ADVISESINKEX	equ 1

BEGIN_CLASS CAdviseSink
AdviseSink		IAdviseSink <>
dwRefCount		dd ?
END_CLASS

__this	textequ <ebx>
_this	textequ <[__this].CAdviseSink>

	MEMBER AdviseSink
	MEMBER dwRefCount

;--- private methods

Destroy@CAdviseSink proto :ptr CAdviseSink

	.data

	.const

;*** vtbl of interface IDispatch

CAdviseSinkVtbl label IAdviseSinkVtbl
	IUnknownVtbl {QueryInterface, AddRef, Release}
	dd offset OnDataChange
	dd offset OnViewChange
	dd offset OnRename
	dd offset OnSave
	dd offset OnClose
if ?ADVISESINKEX
	dd offset OnViewStatusChange
endif

;*** table of supported interfaces

iftab label dword
	dd IID_IUnknown				, 0
	dd IID_IAdviseSink			, CAdviseSink.AdviseSink
if ?ADVISESINKEX
	dd IID_IAdviseSinkEx		, CAdviseSink.AdviseSink
endif
;NUMIFENTRIES equ ($ - offset iftab) / (4 * 2)
NUMIFENTRIES textequ %($ - offset iftab) / (4 * 2)

g_szAdviseSink	db "IAdviseSink",0
externdef g_szContainer:BYTE

	.code

Display proc pszString:LPSTR
	invoke printf@CLogWindow, CStr("%s_%s::%s",10),
		addr g_szContainer, addr g_szAdviseSink, pszString
	ret
	align 4
Display endp

;--- IAdviseSink members

OnDataChange proc this_:ptr CAdviseSink, pFormatetc:ptr FORMATETC, pStgmed:ptr STGMEDIUM
	invoke Display, CStr("OnDataChange") 
	return S_OK
	align 4
OnDataChange endp

OnViewChange proc this_:ptr CAdviseSink, dwAspect:DWORD, lindex:DWORD
	invoke Display, CStr("OnViewChange")
	return S_OK
	align 4
OnViewChange endp

OnRename proc this_:ptr CAdviseSink, pmk:ptr IMoniker
	invoke Display, CStr("OnRename")
	return S_OK
	align 4
OnRename endp

OnSave proc this_:ptr CAdviseSink
	invoke Display, CStr("OnSave")
	return S_OK
	align 4
OnSave endp

OnClose proc this_:ptr CAdviseSink
	invoke Display, CStr("OnClose")
	return S_OK
	align 4
OnClose endp

OnViewStatusChange proc this_:ptr CAdviseSink, dwViewStatus:DWORD
	invoke printf@CLogWindow, CStr("%s_%s%s::%s",10),
		addr g_szContainer, addr g_szAdviseSink, CStr("Ex"), CStr("OnViewStatusChange")
	return S_OK
	align 4
OnViewStatusChange endp

;--------------------------------------------------------------
;--- interface IUnknown
;--------------------------------------------------------------

AddRef proto :ptr CAdviseSink

QueryInterface proc uses __this this_:ptr CAdviseSink, riid:ptr IID, ppReturn:ptr ptr

;;local	szIID[40]:byte
local	wszIID[40]:word
local	szKey[128]:byte
local	dwSize:DWORD
local	hKey:HANDLE

	mov __this,this_
	invoke IsInterfaceSupported, riid, offset iftab, NUMIFENTRIES,  this_, ppReturn

	.if (g_bLogActive && g_bDispQueryIFCalls)
;--------------------- print the name of the interface we have just been queried
	    push eax
		invoke StringFromGUID2,riid, addr wszIID,40
;;		invoke WideCharToMultiByte, CP_ACP, 0, addr wszIID, -1, addr szIID, 40, NULL, NULL
		invoke wsprintf, addr szKey, CStr("%s\%S"),addr g_szInterface, addr wszIID
		invoke RegOpenKeyEx, HKEY_CLASSES_ROOT, addr szKey, 0, KEY_READ, addr hKey 
		.if (eax == ERROR_SUCCESS)
			mov dwSize, sizeof szKey
			invoke RegQueryValueEx,hKey,addr g_szNull,NULL,NULL,addr szKey,addr dwSize
			invoke RegCloseKey, hKey
		.else
			mov szKey,0
		.endif
		pop eax
		push eax
;;		DebugOut "CAdviseSink_IUnknown::QueryInterface(%S[%s])=%X", addr wszIID, addr szKey, eax
		invoke printf@CLogWindow, CStr("CAdviseSink_IUnknown::QueryInterface(%S[%s])=%X",10),
					addr wszIID, addr szKey, eax
		pop eax
	.endif
	ret
	align 4

QueryInterface endp


AddRef proc uses __this this_:ptr CAdviseSink

	mov __this,this_
	inc m_dwRefCount
	mov eax, m_dwRefCount
	ret
	align 4

AddRef endp

Release proc uses __this this_:ptr CAdviseSink

	mov __this,this_
	dec m_dwRefCount

	mov eax, m_dwRefCount
	.if (eax == 0)
		invoke Destroy@CAdviseSink, __this
		xor eax,eax
	.endif
	ret
	align 4

Release endp


Destroy@CAdviseSink proc uses __this this_:ptr CAdviseSink

	DebugOut "Destroy@CAdviseSink"
	mov __this,this_
	invoke free, __this
	ret
	align 4
Destroy@CAdviseSink endp


Create@CAdviseSink proc public uses __this

	DebugOut "Create@CAdviseSink"
	invoke malloc, sizeof CAdviseSink
	.if (!eax)
		ret
	.endif

	mov __this, eax
	mov m_dwRefCount, 1
	mov m_AdviseSink.lpVtbl,	offset CAdviseSinkVtbl
	return __this
	align 4

Create@CAdviseSink endp


	end
