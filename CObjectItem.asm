
;*** classes CObjectItem ***

	.386
	.model flat,stdcall
	option casemap :none   ; case sensitive
	option proc:private

	include COMView.inc
INSIDE_COBJECTITEM equ 1
	include classes.inc
	include rsrc.inc
	include CEditDlg.inc
	include debugout.inc

?HANDSOFF	equ 1

BEGIN_CLASS CObjectItem, IObjectItem
;;ObjectItem	IObjectItem <>
dwRefCnt	DWORD ?
guid		GUID <>
pUnknown	LPUNKNOWN ?
iCount		DWORD ?				;number of items in MULTI_QI
dwRunLocks	DWORD ?				;number of run locks (OleLockRunning called)
pMQI		LPVOID ?			;MULTI_QI ptr
pObjectDlg		pCObjectDlg ?
pViewObjectDlg	pCViewObjectDlg ?
pPropertiesDlg	pCPropertiesDlg ?
pViewStorageDlg	pCViewStorageDlg ?
pStorage		LPSTORAGE ?
pMoniker		LPMONIKER ?
pStgHlp			LPUNKNOWN ?
pConnections	pCList	?
dwFlags			DWORD ?
dwCloseFlags	DWORD ?
pTypeInfo		LPTYPEINFO ?	;typeinfo of COCLASS
bLocked			BOOLEAN ?
bClosePending	BOOLEAN	?
END_CLASS

SetRunLock		proto :ptr CObjectItem, :BOOL
Destroy@CObjectItem	proto :ptr CObjectItem


__this	textequ <ebx>
_this	textequ <[__this].CObjectItem>

	MEMBER lpVtbl, dwRefCnt, guid, pUnknown, iCount, pMQI, bLocked, dwRunLocks
	MEMBER pObjectDlg, pViewObjectDlg, pPropertiesDlg, pViewStorageDlg
	MEMBER pStorage, pStgHlp, pMoniker
	MEMBER pConnections, bClosePending, dwCloseFlags, pTypeInfo

	.data

g_pObjects	pCList NULL

	.const

CObjectItemVtbl label dword
	dd AddRef, Release
	dd GetObjectDlg, SetObjectDlg
	dd GetViewObjectDlg, SetViewObjectDlg
	dd GetPropDlg, SetPropDlg
	dd GetViewStorageDlg, SetViewStorageDlg
	dd Lock_, Unlock, IsLocked
	dd GetRunLock, SetRunLock
	dd GetStorage, SetStorage
	dd GetMoniker, SetMoniker
	dd GetConnectionList
	dd ShowObjectDlg
	dd ShowPropertiesDlg
	dd ShowViewObjectDlg
	dd ShowViewStorageDlg
	dd GetFlags, SetFlags
	dd GetDisplayName
	dd SetWindowText_
	dd AddFilename
	dd Close
	dd GetCoClassTypeInfo, SetCoClassTypeInfo
	dd GetDefaultInterface
	dd GetMQI
	dd SetMQI

	.code

;--- refresh object view in main dialog if it's active

RefreshObjectView proc public

	invoke RefreshView@CMainDlg, g_pMainDlg, MODE_OBJECT
	ret
	align 4

RefreshObjectView endp


FreeMultiQI proc uses esi
	.if (m_pMQI)
		mov ecx, m_iCount
		mov esi, m_pMQI
		.while (ecx)
			push ecx
			.if ([esi].MULTI_QI.hr == S_OK)
				invoke vf([esi].MULTI_QI.pItf,IUnknown,Release)
			.endif
			add esi,sizeof MULTI_QI
			pop ecx
			dec ecx
		.endw
		invoke free, m_pMQI
		mov m_pMQI, NULL
	.endif
	ret
	align 4
FreeMultiQI endp

;--- increase reference counter

AddRef proc this_:ptr CObjectItem
	mov ecx, this_
	inc [ecx].CObjectItem.dwRefCnt
	ret
	align 4
AddRef endp

;*** release object

Release proc this_:ptr CObjectItem

	mov ecx, this_
	dec [ecx].CObjectItem.dwRefCnt
	.if (ZERO?)
		invoke Destroy@CObjectItem, ecx
	.endif
	ret
	align 4
Release endp


GetGUID@CObjectItem proc public this_:ptr CObjectItem, pGUID:ptr GUID
	pushad
	mov eax, this_
	lea esi, [eax].CObjectItem.guid
	mov edi, pGUID
	movsd
	movsd
	movsd
	movsd
	popad
	ret
	align 4
GetGUID@CObjectItem endp

GetUnknown@CObjectItem proc public this_:ptr CObjectItem
	mov eax, this_
	mov eax, [eax].CObjectItem.pUnknown
	ret
	align 4
GetUnknown@CObjectItem endp

GetMQI proc this_:ptr CObjectItem

	mov eax, this_
	mov eax,[eax].CObjectItem.pMQI
	ret
	align 4

GetMQI endp

SetMQI proc uses __this this_:ptr CObjectItem, iCount:DWORD, pMQI:ptr MULTI_QI

	mov __this, this_
	invoke FreeMultiQI
	mov ecx, iCount
	mov m_iCount, ecx
	mov ecx, pMQI
	mov m_pMQI, ecx
	ret
	align 4

SetMQI endp

ShowObjectDlg proc uses __this this_:ptr CObjectItem, hWnd:HWND

	mov __this, this_
	mov ecx, m_pObjectDlg
	.if (ecx)
		invoke RestoreAndActivateWindow, [ecx].CDlg.hWnd
	.else
		invoke Create@CObjectDlg, __this
		.if (eax)
			mov m_pObjectDlg, eax
			invoke Show@CObjectDlg, eax, hWnd
		.endif
	.endif
	ret
	align 4

ShowObjectDlg endp

ShowPropertiesDlg proc uses __this this_:ptr CObjectItem, hWnd:HWND

	mov __this, this_
	mov ecx, m_pPropertiesDlg
	.if (ecx)
		invoke RestoreAndActivateWindow, [ecx].CDlg.hWnd
	.else
		invoke Create2@CPropertiesDlg, m_pUnknown, NULL
		.if (eax)
			mov m_pPropertiesDlg, eax
			invoke Show@CPropertiesDlg, eax, NULL
		.endif
	.endif
	ret
	align 4
ShowPropertiesDlg endp


ShowViewObjectDlg proc public uses __this this_:ptr CObjectItem, hWnd:HWND, pItem:ptr CInterfaceItem

	mov __this, this_
	mov ecx, m_pViewObjectDlg
	.if (ecx)
		invoke RestoreAndActivateWindow, [ecx].CDlg.hWnd
	.else
		invoke Create@CViewObjectDlg, __this, pItem
		.if (eax)
			.if (g_bViewDlgAsTopLevelWnd)
				mov ecx, NULL
			.else
				mov ecx, hWnd
			.endif
			invoke Show@CViewObjectDlg, eax, ecx
		.endif
	.endif
	ret
	align 4
ShowViewObjectDlg endp

ShowViewStorageDlg proc uses __this this_:ptr CObjectItem, hWnd:HWND

	mov __this, this_
	mov ecx, m_pViewStorageDlg
	.if (ecx)
		invoke RestoreAndActivateWindow, [ecx].CDlg.hWnd
	.else
		.if (m_pStorage)
if ?HANDSOFF
			mov ecx, __this
else
			xor ecx, ecx
endif
			invoke Create@CViewStorageDlg, m_pStorage, NULL, ecx
			.if (eax)
				mov m_pViewStorageDlg, eax
				invoke Show@CViewStorageDlg, eax, hWnd
			.endif
		.endif
	.endif
	ret
	align 4

ShowViewStorageDlg endp

GetViewStorageDlg proc public this_:ptr CObjectItem
	mov eax, this_
	mov eax, [eax].CObjectItem.pViewStorageDlg
	ret
	align 4
GetViewStorageDlg endp

SetViewStorageDlg proc public this_:ptr CObjectItem, pViewStorageDlg:pCViewStorageDlg
	mov ecx, this_
	mov eax, pViewStorageDlg
	mov [ecx].CObjectItem.pViewStorageDlg, eax
	ret
	align 4
SetViewStorageDlg endp


IsLocked proc this_:ptr CObjectItem
	mov ecx, this_
	movzx eax, [ecx].CObjectItem.bLocked
	ret
	align 4
IsLocked endp

Lock_ proc this_:ptr CObjectItem
	mov ecx, this_
	.if (![ecx].CObjectItem.bLocked)
		inc [ecx].CObjectItem.bLocked
		inc [ecx].CObjectItem.dwRefCnt
	.endif
	ret
	align 4
Lock_ endp

Unlock proc this_:ptr CObjectItem
	mov ecx, this_
	.if ([ecx].CObjectItem.bLocked)
		dec [ecx].CObjectItem.bLocked
		invoke Release, ecx
	.endif
	ret
	align 4
Unlock endp

SetFlags proc this_:ptr CObjectItem, dwFlags:DWORD
	mov ecx, this_
	mov eax, dwFlags
	mov [ecx].CObjectItem.dwFlags, eax
	ret
	align 4
SetFlags endp

GetFlags proc this_:ptr CObjectItem
	mov ecx, this_
	mov eax, [ecx].CObjectItem.dwFlags
	ret
	align 4
GetFlags endp

GetDisplayName proc uses __this this_:ptr CObjectItem, ppwszName:ptr LPOLESTR

local	hr:DWORD
local	pBindCtx:LPBINDCTX

	mov __this, this_

	mov hr, E_FAIL
	mov ecx, ppwszName
	mov dword ptr [ecx], NULL
	.if (m_pMoniker)
		invoke CreateBindCtx, NULL, addr pBindCtx
		invoke vf(m_pMoniker, IMoniker, GetDisplayName), pBindCtx, NULL, ppwszName
		mov hr, eax
		invoke vf(pBindCtx, IUnknown, Release)
	.endif
	return hr
	align 4

GetDisplayName endp

SetViewObjectDlg proc this_:ptr CObjectItem, pViewObjectDlg:ptr CViewObjectDlg
	mov ecx, this_
	mov eax, pViewObjectDlg
	mov [ecx].CObjectItem.pViewObjectDlg, eax
	ret
	align 4
SetViewObjectDlg endp

GetViewObjectDlg proc this_:ptr CObjectItem
	mov ecx, this_
	mov eax, [ecx].CObjectItem.pViewObjectDlg
	ret
	align 4
GetViewObjectDlg endp

SetPropDlg proc this_:ptr CObjectItem, pPropertiesDlg:ptr CPropertiesDlg
	mov ecx, this_
	mov eax, pPropertiesDlg
	mov [ecx].CObjectItem.pPropertiesDlg, eax
	ret
	align 4
SetPropDlg endp

GetPropDlg proc this_:ptr CObjectItem
	mov ecx, this_
	mov eax, [ecx].CObjectItem.pPropertiesDlg
	ret
	align 4
GetPropDlg endp

SetObjectDlg proc this_:ptr CObjectItem, pObjectDlg:ptr CObjectDlg
	mov ecx, this_
	mov eax, pObjectDlg
	mov [ecx].CObjectItem.pObjectDlg, eax
	ret
	align 4
SetObjectDlg endp

GetObjectDlg proc this_:ptr CObjectItem
	mov ecx, this_
	mov eax, [ecx].CObjectItem.pObjectDlg
	ret
	align 4
GetObjectDlg endp

ReduceFileName proc uses esi edi pszFile:LPSTR, iSize:DWORD
	mov edi, pszFile
	mov ecx, iSize
	.while (ecx)
		dec ecx
		mov al, [edi+ecx]
		.break .if (al == '\')
	.endw
	lea esi, [edi+ecx]
	.if (!ecx)
		jmp done
	.elseif (ecx > 32)
		.if (byte ptr [esi] == '\')
			inc esi
		.endif
		.while (al)
			lodsb
			stosb
		.endw
	.endif
done:
	ret
	align 4
ReduceFileName endp


AddFilename proc uses __this this_:ptr CObjectItem, hWnd:HWND, bUseIPersistFile:BOOL

local	pPersistFile:LPPERSISTFILE
local	pwszFile:ptr WORD
local	szFile[MAX_PATH+4]:byte
local	szCaption[MAX_PATH+128]:byte

	mov __this, this_
	.if (!bUseIPersistFile)
		invoke GetDisplayName, __this, addr pwszFile
	.else
		mov eax, E_FAIL
	.endif
	.if (eax != S_OK)
		invoke vf(m_pUnknown, IUnknown, QueryInterface), addr IID_IPersistFile, addr pPersistFile
		.if (eax == S_OK)
			invoke vf(pPersistFile, IPersistFile, GetCurFile), addr pwszFile
			invoke vf(pPersistFile, IUnknown, Release)
		.endif
	.endif
	.if (!pwszFile)
		jmp done
	.endif
	invoke WideCharToMultiByte,CP_ACP,0,pwszFile,-1,addr szFile + 1,sizeof szFile - 1,0,0 
	invoke lstrlen, addr szFile+1
	.if (eax > 64)
		invoke ReduceFileName, addr szFile+1, eax
	.endif
	mov szFile, ' '
	invoke GetWindowText, hWnd, addr szCaption, sizeof szCaption
	invoke lstrcat, addr szCaption, addr szFile
	invoke SetWindowText, hWnd, addr szCaption
	invoke CoTaskMemFree, pwszFile
done:
	ret
	align 4

AddFilename endp

StringFromCLSID@CObjectItem proc public  uses __this this_:ptr CObjectItem, pszCLSID:LPSTR

local	wszCLSID[40]:WORD
local	szCaption[MAX_PATH+64]:byte

	mov __this, this_
	invoke StringFromGUID2, addr m_guid, addr wszCLSID, 40
	invoke WideCharToMultiByte,CP_ACP,0, addr wszCLSID,-1, pszCLSID, 40, 0, 0 
	ret
	align 4

StringFromCLSID@CObjectItem endp

SetWindowText_ proc uses __this this_:ptr CObjectItem, hWnd:HWND

local	szText[256]:byte
local	szStr[128]:byte
local	wszGUID[40]:word
local	szGUID[40]:byte

	mov __this, this_
	mov szStr, 0
	invoke GetTextFromCLSID, addr m_guid, addr szStr, sizeof szStr
	.if (!szStr)
		invoke StringFromGUID2, addr m_guid, addr wszGUID, 40
		invoke WideCharToMultiByte,CP_ACP,0,addr wszGUID,-1,addr szStr,sizeof szStr,0,0 
	.endif
	invoke GetWindowText, hWnd, addr szGUID, sizeof szGUID
	invoke wsprintf, addr szText, addr szGUID, addr szStr
	invoke SetWindowText, hWnd, addr szText
	ret
	align 4

SetWindowText_ endp

GetRunLock proc this_:ptr CObjectItem
	mov ecx, this_
	mov eax, [ecx].CObjectItem.dwRunLocks
	ret
	align 4
GetRunLock endp

SetRunLock proc uses __this this_:ptr CObjectItem, bLockRunning:BOOL

local pOleObject:LPOLEOBJECT

	mov __this, this_
	.if (bLockRunning == TRUE)
		.if (m_dwRunLocks)
			inc m_dwRunLocks
		.else
			invoke SetCursor, g_hCsrWait
			push eax
			invoke OleRun, m_pUnknown
			invoke OleLockRunning, m_pUnknown, TRUE, TRUE
			.if (eax == S_OK)
				inc m_dwRunLocks
			.endif
			pop eax
			invoke SetCursor, eax
		.endif
	.else
		.if (m_dwRunLocks)
			dec m_dwRunLocks
			.if (!m_dwRunLocks)
				invoke OleLockRunning, m_pUnknown, FALSE, TRUE
				.if (m_bClosePending)
					invoke vf(m_pUnknown, IUnknown, QueryInterface), addr IID_IOleObject, addr pOleObject
					.if (eax == S_OK)
						invoke vf(pOleObject, IOleObject, Close), m_dwCloseFlags
						invoke vf(pOleObject, IOleObject, Release)
					.endif
					mov m_bClosePending, FALSE
				.endif
			.endif
		.endif
	.endif
	ret
	align 4
SetRunLock endp

Close proc uses __this this_:ptr CObjectItem, dwFlags:DWORD

local	pOleObject:LPOLEOBJECT

	mov __this, this_
	.if (m_dwRunLocks)
		mov eax, dwFlags
		mov m_dwCloseFlags, eax
		mov m_bClosePending, TRUE
	.else
		invoke vf(m_pUnknown, IUnknown, QueryInterface), addr IID_IOleObject, addr pOleObject
		.if (eax == S_OK)
			invoke vf(pOleObject, IOleObject, Close), dwFlags
			invoke vf(pOleObject, IUnknown, Release)
		.endif
	.endif
	ret
	align 4
Close endp

GetCoClassTypeInfo proc uses __this this_:ptr CObjectItem, ppTypeInfo:ptr LPTYPEINFO

local	pProvideClassInfo:LPPROVIDECLASSINFO

	mov __this, this_
	mov edx, ppTypeInfo
	@mov dword ptr [edx], 0
	.if (!m_pTypeInfo)
		invoke vf(m_pUnknown, IUnknown, QueryInterface), addr IID_IProvideClassInfo, addr pProvideClassInfo
		.if (eax == S_OK)
			invoke vf(pProvideClassInfo, IProvideClassInfo, GetClassInfo_), addr m_pTypeInfo
			invoke vf(pProvideClassInfo, IUnknown, Release)
		.endif
	.endif
	mov edx, ppTypeInfo
	invoke ComPtrAssign, edx, m_pTypeInfo
	ret
	align 4
GetCoClassTypeInfo endp

SetCoClassTypeInfo proc this_:ptr CObjectItem, pTypeInfo:LPTYPEINFO
	mov ecx, this_
	invoke ComPtrAssign, addr [ecx].CObjectItem.pTypeInfo, pTypeInfo
	ret
	align 4
SetCoClassTypeInfo endp

GetDefaultInterface proc uses __this this_:ptr CObjectItem, bSource:BOOL, riid:REFIID

local	pTypeInfo:LPTYPEINFO
local	pTypeAttr:ptr TYPEATTR
local	hr:DWORD

	mov hr, E_FAIL
	mov __this, this_
	mov eax, m_pTypeInfo
	.if (!eax)
		invoke GetTypeInfoFromIProvideClassInfo, m_pUnknown, bSource
		.if ((!eax) && (bSource == FALSE))
			invoke GetTypeInfoFromIDispatch, m_pUnknown
		.endif
	.else
		invoke GetDefaultInterfaceFromCoClass, eax, bSource
	.endif
	.if (eax)
		mov pTypeInfo, eax
		invoke vf(pTypeInfo, ITypeInfo, GetTypeAttr), addr pTypeAttr
		.if (eax == S_OK)
			mov ecx, pTypeAttr
			invoke CopyMemory, riid, addr [ecx].TYPEATTR.guid, sizeof IID
			mov hr, S_OK
			invoke vf(pTypeInfo, ITypeInfo, ReleaseTypeAttr), pTypeAttr
		.endif
		invoke vf(pTypeInfo, IUnknown, Release)
	.endif
	return hr
	align 4

GetDefaultInterface endp

GetStorage proc this_:ptr CObjectItem
	mov ecx, this_
	mov eax, [ecx].CObjectItem.pStorage
	ret
	align 4
GetStorage endp

SetStorage proc uses __this this_:ptr CObjectItem, pStorage:LPSTORAGE

	mov __this, this_
	invoke ComPtrAssign, addr m_pStorage, pStorage
	ret
	align 4
SetStorage endp

GetMoniker proc this_:ptr CObjectItem
	mov ecx, this_
	mov eax, [ecx].CObjectItem.pMoniker
	ret
	align 4
GetMoniker endp

SetMoniker proc uses __this this_:ptr CObjectItem, pMoniker:LPMONIKER

	mov __this, this_
	invoke ComPtrAssign, addr m_pMoniker, pMoniker
	push eax
	invoke RefreshObjectView
	pop eax
	ret
	align 4
SetMoniker endp

GetConnectionList proc this_:ptr CObjectItem, bForceCreate
	mov ecx, this_
	mov eax, [ecx].CObjectItem.pConnections
	.if ((!eax) && (bForceCreate))
		invoke Create@CList, NULL
		mov ecx, this_
		mov [ecx].CObjectItem.pConnections, eax
	.endif
	ret
	align 4
GetConnectionList endp

SetStgHlp@CObjectItem proc public uses __this this_:ptr CObjectItem, pStgHlp:LPUNKNOWN

	mov __this, this_
	invoke ComPtrAssign, addr m_pStgHlp, pStgHlp
	ret
	align 4
SetStgHlp@CObjectItem endp

;--- static method Find

Find@CObjectItem proc public uses ebx pUnknown:LPUNKNOWN

	xor ebx, ebx
	.while (1)
		invoke GetItem@CList, g_pObjects, ebx
		.break .if (!eax)
		mov edx, pUnknown
		.break .if (edx == [eax].CObjectItem.pUnknown)
		inc ebx
	.endw
	ret
	align 4

Find@CObjectItem endp

;--- static method FindStorage

FindStorage@CObjectItem proc public uses ebx pStorage:LPSTORAGE

	xor ebx, ebx
	.while (1)
		invoke GetItem@CList, g_pObjects, ebx
		.break .if (!eax)
		mov edx, pStorage
		.break .if (edx == [eax].CObjectItem.pStorage)
		inc ebx
	.endw
	ret
	align 4

FindStorage@CObjectItem endp


;*** destroy a created object


Destroy@CObjectItem proc uses esi __this this_:ptr CObjectItem

	mov __this, this_

	.if (m_pConnections)
		xor esi, esi
		.while (1)
			invoke GetItem@CList, m_pConnections, esi
			.break .if (!eax)
;;			invoke Disconnect@CConnection, eax
			invoke Destroy@CConnection, eax
			inc esi
		.endw
		invoke Destroy@CList, m_pConnections
	.endif

	.while (m_dwRunLocks)
		invoke SetRunLock, __this, FALSE
	.endw

	invoke FindItem@CList, g_pObjects, __this
	.if (eax != -1)
		invoke DeleteItem@CList, g_pObjects, eax
	.endif

	invoke FreeMultiQI

	.if (m_pTypeInfo)
		invoke vf(m_pTypeInfo, IUnknown, Release)
	.endif
	.if (m_pStorage)
		invoke vf(m_pStorage, IUnknown, Release)
		DebugOut "Destroy@CObjectItem: Release(pStorage) returned %u", eax
	.endif
	.if (m_pStgHlp)
		invoke vf(m_pStgHlp, IUnknown, Release)
	.endif
	.if (m_pMoniker)
		invoke vf(m_pMoniker, IUnknown, Release)
	.endif
	.if (m_pUnknown)
		invoke SafeRelease, m_pUnknown
		DebugOut "Destroy@ObjectItem IUnknown::Release returned %X", eax
		invoke printf@CLogWindow, CStr("--- object destroyed, calling Release(%X) returned %u",10), m_pUnknown, eax
	.endif

	invoke free, __this

	invoke RefreshObjectView
	.if (g_bFreeLibs)
		invoke CoFreeUnusedLibraries
	.endif
exit:
	ret
	align 4

Destroy@CObjectItem endp


;*** create an object description
;*** currently only partly used because MULTI_QI doesnt work well in win9x


Create@CObjectItem proc public uses __this pUnknown:LPUNKNOWN, pCLSID:ptr GUID

local pPersist:LPPERSIST
local pOleObject:LPOLEOBJECT
local wszCLSID[40]:WORD
ifdef _DEBUG
local	this_:ptr CObjectItem
endif

	invoke malloc,sizeof CObjectItem
	.if (!eax)
		ret
	.endif
	mov __this, eax
ifdef _DEBUG
	mov this_, eax
endif
	mov m_lpVtbl, offset CObjectItemVtbl
	.if (pCLSID)
		pushad
		lea edi,m_guid
		mov esi,pCLSID
		movsd
		movsd
		movsd
		movsd
		popad
if 1
	.else
		invoke vf(pUnknown, IUnknown, QueryInterface), addr IID_IPersist, addr pPersist
		.if (eax != S_OK)
;------------------------------- sometimes IPersist is not supported, but IPersistXXX
			invoke vf(pUnknown, IUnknown, QueryInterface), addr IID_IPersistFile, addr pPersist
		.endif
		.if (eax == S_OK)
			invoke vf(pPersist, IPersist, GetClassID), addr m_guid
			invoke vf(pPersist, IUnknown, Release)
		.else
			invoke vf(pUnknown, IUnknown, QueryInterface), addr IID_IOleObject, addr pOleObject
			.if (eax == S_OK)
				invoke vf(pOleObject, IOleObject, GetUserClassID), addr m_guid
				invoke vf(pOleObject, IUnknown, Release)
			.endif
		.endif
endif
	.endif
	invoke StringFromGUID2, pCLSID, addr wszCLSID, LENGTHOF wszCLSID
	mov ecx,pUnknown
	mov m_pUnknown,ecx
	invoke printf@CLogWindow, CStr("--- object %X (%S) created",10), ecx, addr wszCLSID
	invoke vf(m_pUnknown, IUnknown, AddRef)
	mov m_dwRefCnt, 1
	invoke AddItem@CList, g_pObjects, __this
	invoke RefreshObjectView
	return __this
	align 4

Create@CObjectItem endp

	end
