

;*** definition of class CEventSink
;*** CContainer implements a simple IDispatch, so being able
;*** to receive event calls from host

	.386
	.model flat,stdcall
	option casemap :none   ; case sensitive
	option proc:private

	include COMView.inc
INSIDE_CEVENTSINK equ 1
	include classes.inc
	include rsrc.inc
	include debugout.inc
	include msdatasrc.inc



BEGIN_CLASS CEventSink
Dispatch				IDispatch <>
PropertyNotifySink		IPropertyNotifySink <>
_DataSourceListener		DataSourceListener <>
dwRefCount				dd ?
pTypeInfo				LPTYPEINFO ?
pObjectItem				pCObjectItem ?
iid						IID <>
szType					db 32 dup (?)
END_CLASS

__this	textequ <ebx>
_this	textequ <[__this].CEventSink>

	MEMBER Dispatch, PropertyNotifySink, _DataSourceListener
	MEMBER dwRefCount
	MEMBER pTypeInfo, iid, pObjectItem, szType

;--- private methods

Destroy@CEventSink proto :ptr CEventSink

	.data

	.const

;*** vtbl of interface IDispatch

CDispatchVtbl label IDispatchVtbl
	IUnknownVtbl {QueryInterface, AddRef, Release}
	dd		offset GetTypeInfoCount
	dd		offset GetTypeInfo
	dd		offset GetIDsOfNames
	dd		offset Invoke_

CPropertyNotifySinkVtbl label IPropertyNotifySinkVtbl
	IUnknownVtbl {QueryInterface_, AddRef_, Release_}
	dd		offset OnChanged_
	dd		offset OnRequestEdit_

CDataSourceListenerVtbl label DataSourceListenerVtbl
	IUnknownVtbl {QueryInterface_3, AddRef_3, Release_3}
	dd		offset dataMemberChanged
	dd		offset dataMemberAdded
	dd		offset dataMemberRemoved

;*** table of supported interfaces

iftab label dword
	dd IID_IUnknown				, CEventSink.Dispatch
	dd IID_IDispatch			, CEventSink.Dispatch
	dd IID_IPropertyNotifySink	, CEventSink.PropertyNotifySink
	dd IID_DataSourceListener	, CEventSink._DataSourceListener
NUMIFENTRIES textequ %($ - offset iftab) / (4 * 2)

externdef IID_DataSourceListener:IID

IID_DataSourceListener sIID_DataSourceListener

	.code

@MakeIUnknownStubs CEventSink.PropertyNotifySink
@MakeIUnknownStubs CEventSink._DataSourceListener, 3

AssembleStdDispIdStr proc dispId:DISPID, pText:LPSTR

		invoke GetStdDispIdStr, dispId
		.if (eax)
			push eax
			invoke lstrcpy, pText, CStr("DISPID_")
			pop	eax
			invoke lstrcat, pText, eax
        .else
        	invoke lstrlen, pText
            add eax, pText
            invoke wsprintf, eax, CStr("#0x%X"), dispId
		.endif
		ret

AssembleStdDispIdStr endp


;--- IPropertyNotifySink members

OnChanged_:
	sub dword ptr [esp+4],CEventSink.PropertyNotifySink
OnChanged proc this@:ptr CEventSink, dispID:DISPID

local szName[32]:byte

	.if (g_bLogActive)
		mov szName, 0
		.if (dispID < 0)
			invoke AssembleStdDispIdStr, dispID, addr szName
		.endif
		invoke printf@CLogWindow, CStr("IPropertyNotifySink::OnChanged( %d %s)",10), dispID, addr szName
	.endif
	mov eax, this@
	.if ([eax].CEventSink.pObjectItem)
		invoke vf([eax].CEventSink.pObjectItem, IObjectItem, GetPropDlg)
		.if (eax)
			invoke OnChanged@CPropertiesDlg, eax, dispID
		.endif
	.endif
	return S_OK

OnChanged endp

OnRequestEdit_:
	sub dword ptr [esp+4],CEventSink.PropertyNotifySink
OnRequestEdit proc this@:ptr CEventSink, dispID:DISPID

local szName[32]:byte

	.if (g_bLogActive)
		mov szName, 0
		.if (dispID < 0)
			invoke AssembleStdDispIdStr, dispID, addr szName
		.endif
		invoke printf@CLogWindow, CStr("IPropertyNotifySink::OnRequestEdit( %d %s)",10), dispID, addr szName
	.endif
	return S_OK

OnRequestEdit endp


;--- DataSourceListener members


dataMemberChanged proc this@:ptr CEventSink, dataMember:BSTR
	invoke printf@CLogWindow, CStr("dataMemberChanged",10)
	invoke SysFreeString, dataMember
	return S_OK
dataMemberChanged endp

dataMemberAdded proc this@:ptr CEventSink, dataMember:BSTR
	invoke printf@CLogWindow, CStr("dataMemberAdded",10)
	invoke SysFreeString, dataMember
	return S_OK
dataMemberAdded endp

dataMemberRemoved proc this@:ptr CEventSink, dataMember:BSTR
	invoke printf@CLogWindow, CStr("dataMemberRemoved",10)
	invoke SysFreeString, dataMember
	return S_OK
dataMemberRemoved endp


;--------------------------------------------------------------
;--- interface IUnknown
;--------------------------------------------------------------

AddRef proto :ptr CEventSink

QueryInterface proc uses esi edi __this this_:ptr CEventSink, riid:ptr IID, ppReturn:ptr ptr

local	wszIID[40]:word
local	szKey[128]:byte
local	dwSize:DWORD
local	hKey:HANDLE

    mov __this,this_

	invoke IsInterfaceSupported, riid, offset iftab, NUMIFENTRIES, this_, ppReturn
	.if (eax != S_OK)
		mov edi, riid
		lea esi, m_iid
		mov ecx, 4
		repz cmpsd
		.if (ZERO?)
			invoke AddRef, __this
			mov ecx, ppReturn
			mov [ecx], __this
			mov eax, S_OK
		.endif	
	.endif

	.if (g_bLogActive && g_bDispQueryIFCalls)
;--------------------- print the name of the interface we have just been queried
	    push eax
		invoke StringFromGUID2,riid, addr wszIID,40
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
;;		DebugOut "CEventSink_IUnknown::QueryInterface(%S[%s])=%X", addr wszIID, addr szKey, eax
		invoke printf@CLogWindow, CStr("CEventSink_IUnknown::QueryInterface(%S[%s])=%X",10),
					addr wszIID, addr szKey, eax
		pop eax
	.endif
	ret

QueryInterface endp


AddRef proc uses __this this_:ptr CEventSink

	mov __this,this_
	inc m_dwRefCount
	mov eax, m_dwRefCount
	ret

AddRef endp

Release proc uses __this this_:ptr CEventSink

	mov __this,this_
	dec m_dwRefCount

	mov eax, m_dwRefCount
	.if (eax == 0)
		invoke Destroy@CEventSink, __this
		xor eax,eax
	.endif
	ret

Release endp

;--------------------------------------------------------------
;--- interface IDispatch
;--------------------------------------------------------------

GetTypeInfoCount proc this_:ptr CEventSink, pctinfo:ptr DWORD
	DebugOut "IDispatch::GetTypeInfoCount"
    return E_NOTIMPL
GetTypeInfoCount endp

GetTypeInfo proc this_:ptr CEventSink, iTInfo:DWORD, lcid:LCID, ppTInfo:ptr LPTYPEINFO
	DebugOut "IDispatch::GetTypeInfo"
    return E_NOTIMPL
GetTypeInfo endp

GetIDsOfNames proc this_:ptr CEventSink, riid:ptr IID, rgszNames:DWORD, cNames:DWORD, lcid:LCID, rgDispId:ptr DISPID
	DebugOut "IDispatch::GetIDsOfNames"
    return DISP_E_UNKNOWNINTERFACE
GetIDsOfNames endp


GetDispType proc uses esi edi wFlags:DWORD, pszType:LPSTR

		.const
DispTypeTab label dword
		dd DISPATCH_METHOD, CStr("Method")
		dd DISPATCH_PROPERTYGET, CStr("PropertyGet")
		dd DISPATCH_PROPERTYPUT, CStr("PropertyPut")
		dd DISPATCH_PROPERTYPUTREF, CStr("PropertyPutRef")
NUMDISPTYPES equ ($ - DispTypeTab) / 8
		.code

		mov edi, pszType
		mov byte ptr [edi],0
		mov esi, offset DispTypeTab
		mov ecx, NUMDISPTYPES
		.while (ecx)
			push ecx
			lodsd
			mov edx, eax
			lodsd
			mov ecx, eax
			.if (edx & wFlags)
				.if (edi != pszType)
					mov ax, "| "
					stosw
					stosb
				.endif
				invoke lstrcpy, edi, ecx
				invoke lstrlen, edi
				add edi, eax
			.endif
			pop ecx
			dec ecx
		.endw
		ret
GetDispType endp

Invoke_ proc uses __this this_:ptr CEventSink, dispIdMember:DISPID, riid:ptr IID,
			lcid:LCID, wFlags:DWORD, pDispParams:ptr DISPPARAMS,
			pVarResult:ptr VARIANT, pExcepInfo:DWORD, puArgErr:ptr DWORD

local   szText[256]:byte
local   szName[128]:byte
local   szType[64]:byte
local   szParams[128]:byte
local	szIID[40]:byte
local	wszIID[40]:word
local   dwNumNames:dword
local   bstr:BSTR


	mov __this,this_

	.if (g_bLogActive)

	    mov szName,0
		.if (dispIdMember >= 0)
			.if (m_pTypeInfo != 0)
		        invoke vf(m_pTypeInfo, ITypeInfo, GetNames), dispIdMember, addr bstr, 1, addr dwNumNames
			    .if ((eax == S_OK) && (dwNumNames > 0))
					invoke WideCharToMultiByte,CP_ACP,0,bstr,-1,addr szName,sizeof szName,0,0
					invoke SysFreeString, bstr
		        .endif
			.endif
		.else
			invoke AssembleStdDispIdStr, dispIdMember, addr szName
		.endif

		invoke GetDispType, wFlags, addr szType

		mov szIID,0
		.if (riid)
			push edi
			mov edi,riid
			mov ecx,4
			xor eax,eax
			repz scasd
			pop edi
			.if (!ZERO?)
				invoke StringFromGUID2, riid, addr wszIID, 40
				invoke WideCharToMultiByte, CP_ACP, 0, addr wszIID, -1, addr szIID, sizeof szIID,0,0 
			.endif
		.endif

		mov szParams,0
		.if (pDispParams)
			mov word ptr szParams,'('
			mov eax,pDispParams
			mov ecx,[eax].DISPPARAMS.cArgs
			mov edx,[eax].DISPPARAMS.rgvarg
			.while (ecx)
				push ecx
				push edx
				invoke GetArgument, edx, addr szParams
				pop edx
				pop ecx
				add edx,sizeof VARIANT
				dec ecx
				.if (ecx)
					pushad
					invoke lstrcat, addr szParams, CStr(",")
					popad
				.endif
			.endw
			invoke lstrcat, addr szParams, CStr(29h)
		.endif
		.if (m_szType)
			lea ecx, m_szType
		.else
			mov ecx, CStr("IDispatch")
		.endif

	    invoke printf@CLogWindow,
			CStr("CEventSink_%s::Invoke %s, %s, ID=%d, %s%s",10),
				ecx, addr szIID, addr szType, dispIdMember, addr szName, addr szParams

	.endif
if 0
	mov eax,pDispParams
	mov ecx, [eax].DISPPARAMS.cArgs
	mov edx, [eax].DISPPARAMS.rgvarg
	.while (ecx)
		push edx
		push ecx
		invoke VariantClear, edx
		pop ecx
		pop edx
		dec ecx
		add edx, sizeof VARIANT
	.endw
endif
	return S_OK
    align 4
Invoke_ endp


Destroy@CEventSink proc uses __this this_:ptr CEventSink

	DebugOut "Destroy@CEventSink enter"
    mov __this,this_
	.if (m_pTypeInfo)
		invoke vf(m_pTypeInfo, IUnknown, Release)
	.endif
	invoke vf(m_pObjectItem, IObjectItem, Release)
	invoke free, __this
    invoke printf@CLogWindow,
		CStr("--- event sink %X destroyed",10), __this
	DebugOut "Destroy@CEventSink exit"
	ret
Destroy@CEventSink endp


Create@CEventSink proc public uses __this riid:REFIID, pObjectItem:pCObjectItem, pTypeInfo:LPTYPEINFO

local pUnknown:LPUNKNOWN
local pDispatch:LPDISPATCH
local pTypeLib:LPTYPELIB
local bstr:BSTR
local dwIndex:DWORD
local wszIID[40]:WORD

	DebugOut "Create@CEventSink"
	invoke malloc, sizeof CEventSink
	.if (!eax)
		ret
	.endif

	mov __this, eax
	mov m_dwRefCount, 1
	mov m_Dispatch.lpVtbl,			offset CDispatchVtbl
	mov m_PropertyNotifySink.lpVtbl, offset CPropertyNotifySinkVtbl
	mov m__DataSourceListener.lpVtbl, offset CDataSourceListenerVtbl
	pushad
	lea edi, m_iid
	mov esi, riid
	mov ecx, 4
	rep movsd
	popad
	mov eax, pObjectItem
	mov m_pObjectItem, eax
	invoke vf(eax, IObjectItem, AddRef)
	mov eax, pTypeInfo
	mov m_pTypeInfo, eax
	.if (eax)
		invoke vf(m_pTypeInfo, IUnknown, AddRef)
	.else
		invoke GetUnknown@CObjectItem, pObjectItem
		mov pUnknown, eax
		invoke vf(pUnknown, IUnknown, QueryInterface), addr IID_IDispatch, addr pDispatch
		.if (eax == S_OK)
			invoke vf(pDispatch, IDispatch, GetTypeInfo), 0, g_LCID, addr pTypeInfo
			.if (eax == S_OK)
				invoke vf(pTypeInfo, ITypeInfo, GetContainingTypeLib), addr pTypeLib, addr dwIndex
				.if (eax == S_OK)
					invoke vf(pTypeLib, ITypeLib, GetTypeInfoOfGuid), riid, addr m_pTypeInfo
					invoke vf(pTypeLib, ITypeLib, Release)
				.endif
				invoke vf(pTypeInfo, ITypeInfo, Release)
			.endif
			invoke vf(pDispatch, IUnknown, Release)
		.endif
	.endif
	.if (m_pTypeInfo)
		invoke vf(m_pTypeInfo, ITypeInfo, GetDocumentation), MEMBERID_NIL, addr bstr, NULL, NULL, NULL
		.if (eax == S_OK)
			invoke WideCharToMultiByte, CP_ACP, 0, bstr, -1, addr m_szType, sizeof CEventSink.szType, NULL, NULL
			invoke SysFreeString, bstr
		.endif
	.endif
   	invoke StringFromGUID2, riid, addr wszIID, LENGTHOF wszIID
    invoke printf@CLogWindow,
		CStr("--- event sink %X created for %S, pTypeInfo=%X",10), __this, addr wszIID, m_pTypeInfo
	return __this

Create@CEventSink endp


	end
