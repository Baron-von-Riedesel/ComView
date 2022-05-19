

;*** definition of class CPropertyBag

	.386
	.model flat,stdcall
	option casemap :none   ; case sensitive
	option proc:private

	include COMView.inc
INSIDE_CPROPERTYBAG equ 1
	include classes.inc
	include rsrc.inc
	include debugout.inc

BEGIN_CLASS CPropertyBag
PropertyBag				IPropertyBag <>
dwRefCount				dd ?
pPropertyStorage		LPPROPERTYSTORAGE ?
END_CLASS

__this	textequ <ebx>
_this	textequ <[__this].CPropertyBag>

	MEMBER PropertyBag
	MEMBER dwRefCount, pPropertyStorage

Create@CErrorLog proto

;--- private methods

Destroy@CPropertyBag proto :ptr CPropertyBag

	.data

	.const

;*** vtbl of interface IDispatch

CPropertyBagVtbl label IPropertyBagVtbl
	IUnknownVtbl {QueryInterface, AddRef, Release}
	dd Read, Write

;*** table of supported interfaces

iftabPropertyBag label dword
	dd IID_IUnknown				, 0
	dd IID_IPropertyBag			, CPropertyBag.PropertyBag
NUMIFENTRIESPROPBAG textequ %($ - offset iftabPropertyBag) / (4 * 2)

g_szPropertyBag	db "IPropertyBag",0
g_szPropertyStorage	db "IPropertyStorage",0
externdef g_szContainer:BYTE

	.code

if 0
Display	proc pszString:LPSTR
	invoke printf@CLogWindow, CStr("%s_%s::%s",10),
		addr g_szContainer, addr g_szPropertyBag, pszString
	ret
Display endp
endif

;--- IPropertyBag members

;--- ReadMultiple/WriteMultiple for VT_DISPATCH/VT_UNKNOWN will fail
;--- because no implementation (Stand-alone / Compound File) supports these types

Read proc uses __this this_:ptr CPropertyBag, pszPropName:LPOLESTR, pVar:ptr VARIANT, pErrorLog:LPERRORLOG

local	hr:DWORD
local	vt2:PROPVARIANT
local	vt:VARTYPE
local	propspec:PROPSPEC
local	excepinfo:EXCEPINFO

		mov __this, this_
		mov eax, pszPropName
		mov propspec.lpwstr, eax
		mov propspec.ulKind, PRSPEC_LPWSTR
		mov ecx,pVar
		mov cx, [ecx].VARIANT.vt
		mov vt,cx
		.if ((cx == VT_DISPATCH) || (cx == VT_UNKNOWN))
			invoke VariantInit, addr vt2
			mov vt2.vt, VT_STREAMED_OBJECT
			mov vt2.pStream, NULL
			invoke vf(m_pPropertyStorage, IPropertyStorage, ReadMultiple), 1, addr propspec, addr vt2
			.if (eax == S_FALSE)
				mov eax, E_INVALIDARG
			.endif
			mov hr, eax
			.if (eax == S_OK)
				.if (vt == VT_DISPATCH)
					mov ecx, offset IID_IDispatch
				.else
					mov ecx, offset IID_IUnknown
				.endif
				mov edx, pVar
				invoke OleLoadFromStream, vt2.pStream, ecx, addr [edx].VARIANT.pdispVal
				mov hr, eax
				.if ((eax != S_OK) && g_bDispContainerCalls)
					invoke printf@CLogWindow, CStr("%s_%s(%S): OleLoadFromStream failed [%X]",10),
						addr g_szContainer, addr g_szPropertyBag, pszPropName, eax
				.endif
				invoke VariantClear, addr vt2
			.elseif (g_bDispContainerCalls)
				invoke printf@CLogWindow, CStr("%s_%s(%S): call %s::ReadMultiple failed [%X]",10),
					addr g_szContainer, addr g_szPropertyBag, pszPropName, addr g_szPropertyStorage, eax
			.endif
		.else
			invoke vf(m_pPropertyStorage, IPropertyStorage, ReadMultiple), 1, addr propspec, pVar
			.if (eax == S_FALSE)
				mov eax, E_INVALIDARG
			.endif
			mov hr, eax
			.if ((eax != S_OK) && g_bDispContainerCalls)
				invoke printf@CLogWindow, CStr("%s_%s(%S): call %s::ReadMultiple failed [%X]",10),
					addr g_szContainer, addr g_szPropertyBag, pszPropName, addr g_szPropertyStorage, eax
			.endif
		.endif
		.if ((hr != S_OK) && pErrorLog)
			invoke ZeroMemory, addr excepinfo, sizeof EXCEPINFO
			mov eax, hr
			mov excepinfo.scode, eax
			invoke vf(pErrorLog, IErrorLog, AddError), pszPropName, addr excepinfo
		.endif
		DebugOut "IPropertyBag::Read(%S) returns %X", pszPropName, hr
		return hr
		align 4
Read endp

Write proc uses __this this_:ptr CPropertyBag, pszPropName:LPOLESTR, pVar:ptr VARIANT

local	hr:DWORD
local	pStream:LPSTREAM
local	pPersistStream:LPPERSISTSTREAM
local	clsid:CLSID
local	wszCLSID[40]:word
local	propspec:PROPSPEC
local	vt2:PROPVARIANT

		mov __this, this_
		mov eax, pszPropName
		mov propspec.lpwstr, eax
		mov propspec.ulKind, PRSPEC_LPWSTR
		mov ecx,pVar
		.if (([ecx].VARIANT.vt == VT_DISPATCH) || ([ecx].VARIANT.vt == VT_UNKNOWN))
			invoke vf([ecx].VARIANT.pdispVal, IUnknown, QueryInterface), addr IID_IPersistStream, addr pPersistStream
			.if (eax == S_OK)
				invoke CreateStreamOnHGlobal, NULL, TRUE, addr pStream
				mov hr, eax
				.if (eax == S_OK)
					invoke OleSaveToStream, pPersistStream, pStream
					mov hr, eax
					.if (eax == S_OK)
						invoke vf(pStream, IStream, Seek), g_dqNull, STREAM_SEEK_SET, NULL
						invoke VariantInit, addr vt2
						mov vt2.vt, VT_STREAMED_OBJECT
						mov eax, pStream
						mov vt2.pStream, eax
						invoke vf(m_pPropertyStorage, IPropertyStorage, WriteMultiple), 1, addr propspec, addr vt2, 2
						mov hr, eax
						.if ((eax != S_OK) && g_bDispContainerCalls)
							invoke printf@CLogWindow, CStr("%s_%s: call %s::WriteMultiple(%S) failed [%X]",10),
								addr g_szContainer, addr g_szPropertyBag, addr g_szPropertyStorage, pszPropName, eax
						.endif
					.elseif (g_bDispContainerCalls)
						invoke printf@CLogWindow, CStr("%s_%s(%S): OleSaveToStream failed [%X]",10),
							addr g_szContainer, addr g_szPropertyBag, pszPropName, eax
					.endif
					invoke vf(pStream, IUnknown, Release)
				.endif
				invoke vf(pPersistStream, IUnknown, Release)
			.elseif (g_bDispContainerCalls)
				invoke printf@CLogWindow, CStr("%s_%s(%S): QueryInterface(IID_IPersistStream) failed [%X]",10),
					addr g_szContainer, addr g_szPropertyBag, pszPropName, eax
			.endif
		.else
			invoke vf(m_pPropertyStorage, IPropertyStorage, WriteMultiple), 1, addr propspec, ecx, 2
			mov hr, eax
			.if ((eax != S_OK) && g_bDispContainerCalls)
				invoke printf@CLogWindow, CStr("%s_%s(%S): call %s::WriteMultiple failed [%X]",10),
					addr g_szContainer, addr g_szPropertyBag, pszPropName, addr g_szPropertyStorage, eax
			.endif
		.endif
		DebugOut "IPropertyBag::Write(%S) returns %X", pszPropName, hr
		return hr
		align 4
Write endp

;--------------------------------------------------------------
;--- interface IUnknown
;--------------------------------------------------------------

AddRef proto :ptr CPropertyBag

QueryInterface proc uses esi edi __this this_:ptr CPropertyBag, riid:ptr IID, ppReturn:ptr ptr

local	wszIID[40]:word
local	szKey[128]:byte
local	dwSize:DWORD
local	hKey:HANDLE

    mov __this,this_
	invoke IsInterfaceSupported, riid, offset iftabPropertyBag, NUMIFENTRIESPROPBAG,  this_, ppReturn

	.if (g_bLogActive && g_bDispQueryIFCalls)
;--------------------- print the name of the interface we have just been queried
	    push eax
		invoke StringFromGUID2,riid, addr wszIID,40
		invoke wsprintf, addr szKey, CStr("%s\%S"), addr g_szInterface, addr wszIID
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
;;		DebugOut "CPropertyBag_IUnknown::QueryInterface(%S[%s])=%X", addr wszIID, addr szKey, eax
		invoke printf@CLogWindow, CStr("CPropertyBag_IUnknown::QueryInterface(%S[%s])=%X",10),
					addr wszIID, addr szKey, eax
		pop eax
	.endif
	ret
	align 4

QueryInterface endp


AddRef proc uses __this this_:ptr CPropertyBag

	mov __this,this_
	inc m_dwRefCount
	mov eax, m_dwRefCount
	ret
	align 4

AddRef endp

Release proc uses __this this_:ptr CPropertyBag

	mov __this,this_
	dec m_dwRefCount

	mov eax, m_dwRefCount
	.if (eax == 0)
		invoke Destroy@CPropertyBag, __this
		xor eax,eax
	.endif
	ret
	align 4

Release endp


Destroy@CPropertyBag proc uses __this this_:ptr CPropertyBag

	DebugOut "Destroy@CPropertyBag"
    mov __this,this_
	.if (m_pPropertyStorage)
		invoke vf(m_pPropertyStorage, IUnknown, Release)
	.endif
	invoke free, __this
	ret
	align 4
Destroy@CPropertyBag endp


Create@CPropertyBag proc public uses __this, pStorage:LPSTORAGE, pFmtId:ptr FMTID, bCreateAlways:BOOL, ppPropertyBag:ptr LPPROPERTYBAG

local	pPropertySetStorage:LPPROPERTYSETSTORAGE

	DebugOut "Create@CPropertyBag"
	mov ecx, ppPropertyBag
	mov dword ptr [ecx], NULL
	invoke malloc, sizeof CPropertyBag
	.if (!eax)
		return E_OUTOFMEMORY
	.endif

	mov __this, eax
	mov m_dwRefCount, 1
	mov m_PropertyBag.lpVtbl, offset CPropertyBagVtbl
	invoke vf(pStorage, IUnknown, QueryInterface), addr IID_IPropertySetStorage, addr pPropertySetStorage
	.if (eax == S_OK)
		invoke vf(pPropertySetStorage, IPropertySetStorage, Open), pFmtId,\
			STGM_READWRITE or STGM_SHARE_EXCLUSIVE, addr m_pPropertyStorage 
		.if ((eax != S_OK) && (bCreateAlways))
			invoke vf(pPropertySetStorage, IPropertySetStorage, Create), pFmtId,\
				NULL, PROPSETFLAG_NONSIMPLE, STGM_READWRITE or STGM_SHARE_EXCLUSIVE,\
				addr m_pPropertyStorage
		.endif
		push eax
		invoke vf(pPropertySetStorage, IUnknown, Release)
		pop eax
if 1
	.else
		invoke InitIProp
		.if (g_pfnStgOpenPropStg)
			invoke g_pfnStgOpenPropStg, pStorage, pFmtId, PROPSETFLAG_DEFAULT, NULL,\
				addr m_pPropertyStorage
			.if ((eax != S_OK) && (bCreateAlways))
				invoke g_pfnStgCreatePropStg, pStorage, pFmtId, NULL, PROPSETFLAG_DEFAULT, NULL,\
					addr m_pPropertyStorage
			.endif
		.endif
endif
	.endif
	.if (!m_pPropertyStorage)
		push eax
		invoke Destroy@CPropertyBag, __this
		pop eax
		ret
	.endif
	mov ecx, ppPropertyBag
	mov [ecx], __this
	return S_OK
	align 4

Create@CPropertyBag endp

	end
