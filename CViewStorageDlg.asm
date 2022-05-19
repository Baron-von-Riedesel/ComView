
	.486
	.model flat, stdcall
	option casemap :none
	option proc :private

	include COMView.inc
	include richedit.inc
	include statusbar.inc

INSIDE_CVIEWSTORAGEDLG equ 1
	include classes.inc
	include rsrc.inc
	include debugout.inc
	include CHexEdit.inc
	include CSplittButton.inc

?MODELESS	equ 1
MAXSTORAGE	equ 12
?HANDSOFF	equ 1
?USESTGIMPL	equ 1

ViewPropertyStorage	proto protoViewProc	


RecalcSize		proto :DWORD
WriteElement	proto :HTREEITEM
RefreshList		proto
CheckSave		proto :HTREEITEM, bReset:BOOL
OnCommand		proto wParam:WPARAM, lParam:LPARAM


BEGIN_INTERFACE IStorageHlp, IUnknown
END_INTERFACE

BEGIN_CLASS CStorageHlp
StorageHlp		IStorageHlp <>
dwRefCnt		DWORD ?
pStorages		LPSTORAGE MAXSTORAGE dup (?)
END_CLASS

	.const

CStorageHlpVtbl label dword
	dd QueryInterface
	dd AddRef
	dd Release

    .code

__this	textequ <ebx>
_this	textequ <[__this].CStorageHlp>

	MEMBER StorageHlp, dwRefCnt, pStorages

Destroy@CStorageHlp proc uses esi __this this_:ptr CStorageHlp

	DebugOut "Destroy@CStorageHlp(%X)", this_
	mov __this, this_
	mov ecx, MAXSTORAGE
	lea esi, m_pStorages
	.while (ecx)
		lodsd
		.break .if (!eax)
		push ecx
		invoke vf(eax, IUnknown, Release)
		pop ecx
		dec ecx
	.endw
	invoke free, __this
	ret
Destroy@CStorageHlp endp

QueryInterface proc this_:ptr CStorageHlp, riid:REFIID, ppObj:ptr LPUNKNOWN
	mov ecx, ppObj
	mov dword ptr [ecx],NULL
	return E_NOINTERFACE
QueryInterface endp

AddRef proc this_:ptr CStorageHlp
	mov ecx, this_
	inc [ecx].CStorageHlp.dwRefCnt
	ret
AddRef endp

Release proc this_:ptr CStorageHlp
	mov ecx, this_
	dec [ecx].CStorageHlp.dwRefCnt
	.if (ZERO?)
		invoke Destroy@CStorageHlp, ecx
	.endif
	ret
Release endp

GetName proc hWndTV:HWND, hItem:HTREEITEM, pwszName:LPWSTR, iMax:DWORD

local tvi:TVITEM
local szName[64]:byte

		mov eax, hItem
		mov tvi.hItem, eax
		lea eax,szName
		mov tvi.pszText, eax
		mov tvi.cchTextMax, sizeof szName
		mov tvi.mask_, TVIF_TEXT or TVIF_PARAM
		invoke TreeView_GetItem( hWndTV, addr tvi)
		.if (eax)
			mov eax,tvi.lParam
			.if (eax && (eax != -1))
				sub eax,2
				lea ecx,szName
				mov byte ptr [ecx+eax],0
			.endif
			invoke MultiByteToWideChar, CP_ACP, 0, addr szName, -1, pwszName, iMax
		.endif
		ret
GetName endp

ifdef @StackBase
	option stackbase:ebp
endif

Create@CStorageHlp proc uses esi edi __this pStorage:LPSTORAGE, hWndTV:HWND, hItem:HTREEITEM, pwszName:ptr WORD, iMax:DWORD, dwFlags:DWORD, ppStorageHlp:ptr ptr CStorageHlp, ppStorage:ptr LPSTORAGE

local dwESP:DWORD

		mov ecx, ppStorageHlp
		mov dword ptr [ecx], NULL
		invoke malloc, sizeof CStorageHlp
		.if (!eax)
			mov eax, E_OUTOFMEMORY
			jmp done
		.endif

		mov __this, eax
		mov m_StorageHlp.lpVtbl, offset CStorageHlpVtbl
		mov m_dwRefCnt, 1
		lea edi, m_pStorages
		mov eax, hItem
		xor esi, esi
		mov dwESP, esp
		.while (eax)
			push eax
			add edi, 4
			inc esi
			invoke TreeView_GetParent( hWndTV, eax)
		.endw

		mov dword ptr [edi], NULL
		sub edi, 4
		mov eax, pStorage
		mov [edi], eax
		invoke vf(eax, IUnknown, AddRef)

		.while (esi)
			pop ecx
			invoke GetName, hWndTV, ecx, pwszName, iMax
			.if (esi > 1)
				lea ecx, [edi-4]
				invoke vf([edi+0], IStorage, OpenStorage), pwszName,\
					NULL, dwFlags, NULL, NULL, ecx
				.break .if (eax != S_OK)
			.endif
			sub edi, 4
			dec esi
		.endw
		mov esp, dwESP
		.if (esi)
			push eax
			invoke Destroy@CStorageHlp, __this
			pop eax
		.else
			mov ecx, ppStorageHlp
			mov [ecx], __this
			mov ecx, ppStorage
			mov eax, m_pStorages
			mov [ecx], eax
			mov eax, S_OK
		.endif
done:
		DebugOut "Create@CStorageHlp(%X) = %X", pStorage, eax
		ret
		align 4

Create@CStorageHlp endp

ifdef @StackBase
	option stackbase:esp
endif

m_StorageHlp equ <>
m_dwRefCnt	equ <>
m_pStorages	equ <>

;----------------------------------------------------------

BEGIN_CLASS CViewStorageDlg, CDlg
hWndTV			HWND ?
hWndHE			HWND ?
hWndSplit		HWND ?
hWndSB			HWND ?
hItem			HANDLE ?		;item cursor is upon
pStorage		LPSTORAGE ?
pStream			LPSTREAM ?
if ?HANDSOFF	
pObjectItem		pCObjectItem ?
endif
pszFile			LPSTR ?
dwPos			DWORD ?
bHexEdHasFocus	BOOLEAN ?
bIsStream		BOOLEAN ?
bAllowEdit		BOOLEAN ?
bEditLabel		BOOLEAN ?
END_CLASS

__this	textequ <ebx>
_this	textequ <[__this].CViewStorageDlg>
thisarg	textequ <this@:ptr CViewStorageDlg>

	MEMBER hWnd, pDlgProc, hWndTV, hWndHE, pStorage, pStream
	MEMBER hItem, hWndSplit, hWndSB, bHexEdHasFocus, bIsStream, pszFile
	MEMBER bAllowEdit, bEditLabel
if ?HANDSOFF
	MEMBER pObjectItem
endif

SwitchStorage	proc bToHandsOff:BOOL

local pUnknown:LPUNKNOWN
local pOleObject:LPOLEOBJECT
local pPersistStorage:LPPERSISTSTORAGE

if ?HANDSOFF
	.if (m_pObjectItem)
		invoke GetUnknown@CObjectItem, m_pObjectItem
		mov pUnknown, eax
		invoke vf(pUnknown, IUnknown, QueryInterface), addr IID_IOleObject, addr pOleObject
		.if (eax == S_OK)
			invoke OleIsRunning, pOleObject
			.if (eax == TRUE)
				invoke vf(pOleObject, IUnknown, QueryInterface), addr IID_IPersistStorage, addr pPersistStorage
				.if (eax == S_OK)
					.if (bToHandsOff)
						invoke vf(pPersistStorage, IPersistStorage, HandsOffStorage)
					.else
						invoke vf(pPersistStorage, IPersistStorage, SaveCompleted), m_pStorage
					.endif
					invoke vf(pPersistStorage, IUnknown, Release)
				.endif
			.endif
			invoke vf(pOleObject, IUnknown, Release)
		.endif
	.endif
endif
	ret
SwitchStorage	endp

;*** process WM_COMMAND/IDM_LOADSTORAGE
;*** create an object from a Storage/Stream file

OnLoadStorage proc pStorage:LPSTORAGE, pStorageHlp:ptr CStorageHlp

local	pObjectDlg:ptr CObjectDlg
local	pObjectItem:ptr CObjectItem
local	pUnknown:LPUNKNOWN
local	statstg:STATSTG

	DebugOut "OnLoadStorage, pStorage=%X, m_pStorage=%X", pStorage, m_pStorage
	invoke FindStorage@CObjectItem, pStorage
	.if (eax)
		mov ecx, g_pMainDlg
		invoke vf(eax, IObjectItem, ShowViewObjectDlg), [ecx].CDlg.hWnd, NULL
		jmp done
	.endif
	.if (m_bIsStream)
		invoke vf(pStorage, IStream, Seek), g_dqNull, STREAM_SEEK_SET, NULL
		invoke ReadClassStm, pStorage, addr statstg.clsid
	.else
		invoke vf(pStorage, IStorage, Stat), addr statstg, STATFLAG_NONAME
		DebugOut "statstg.grfMode=%X", statstg.grfMode
	.endif
	.if (eax == S_OK)
		invoke GetCoCreateFlags@COptions
		mov ecx, eax
		invoke CoCreateInstance, addr statstg.clsid, NULL,
			ecx, addr IID_IUnknown, addr pUnknown
		.if (eax == S_OK)
			invoke Create@CObjectItem, pUnknown, addr statstg.clsid
			.if (eax)
				mov pObjectItem, eax
				invoke vf(pObjectItem, IObjectItem, SetStorage), pStorage
				invoke SetStgHlp@CObjectItem, pObjectItem, pStorageHlp
				invoke vf(pObjectItem, IObjectItem, SetViewStorageDlg), __this
				invoke vf(pObjectItem, IObjectItem, SetFlags), OBJITEMF_OPENVIEW
				invoke vf(pObjectItem, IObjectItem, ShowObjectDlg), m_hWnd
				invoke vf(pObjectItem, IObjectItem, Release)
			.endif
			invoke vf(pUnknown, IUnknown, Release)
		.else
			invoke OutputMessage, m_hWnd, eax, CStr("CoCreateInstance"), 0
		.endif
	.endif
done:
	ret

OnLoadStorage endp

;*** process WM_COMMAND/IDM_STG2FILE
;*** save storage object into a file

OnSaveStorage proc

local pStorage:LPSTORAGE
local statstg:STATSTG
local hFile:DWORD
local estrm:EDITSTREAM
local dwSize:DWORD
local szFile[MAX_PATH]:byte
local wszFile[MAX_PATH]:WORD

	invoke CheckSave, NULL, TRUE
	.if (!eax)
		jmp done
	.endif
	mov szFile, 0
	.if (!m_bIsStream)
		invoke vf(m_pStorage, IStorage, Stat), addr statstg, STATFLAG_DEFAULT
		.if (eax == S_OK)
			invoke WideCharToMultiByte, CP_ACP, 0, statstg.pwcsName, -1, addr szFile, sizeof szFile, NULL, NULL
			invoke CoTaskMemFree, statstg.pwcsName
		.endif
	.endif

	invoke MyGetFileName, m_hWnd, addr szFile, MAX_PATH, NULL, 0, 1, NULL
	.if (eax)
		.if (m_bIsStream)
			invoke CreateFile, addr szFile,GENERIC_WRITE,0,NULL,
					CREATE_ALWAYS,FILE_ATTRIBUTE_NORMAL,NULL 
			mov hFile,eax
			.if (eax != INVALID_HANDLE_VALUE)
				mov estrm.dwCookie, eax
				mov estrm.dwError,0
				mov estrm.pfnCallback,offset streamout2cb
				invoke SendMessage, m_hWndHE, EM_STREAMOUT, SF_TEXT, addr estrm
				invoke CloseHandle, hFile
			.endif
		.else
			invoke MultiByteToWideChar,CP_ACP,MB_PRECOMPOSED, addr szFile, -1, addr wszFile, MAX_PATH
			invoke StgCreateDocfile, addr wszFile,
				STGM_READWRITE or STGM_SHARE_EXCLUSIVE or STGM_TRANSACTED,
				NULL, addr pStorage
			.if (eax == S_OK) 
				invoke vf(m_pStorage, IStorage, CopyTo), NULL, NULL, NULL, pStorage
				.if (eax != S_OK)
					invoke OutputMessage, m_hWnd, eax, CStr("IStorage::CopyTo Error"), 0
					invoke vf(pStorage, IUnknown, Release)
				.else
					invoke vf(m_pStorage, IStorage, Release)
					mov eax, pStorage
					mov m_pStorage, eax
					invoke RefreshList
					invoke SendMessage, m_hWnd, WM_COMMAND, IDC_COMMIT, 0
				.endif
			.else
				invoke OutputMessage, m_hWnd, eax, CStr("StgCreateDocfile Error"), 0
			.endif
		.endif
	.endif
done:
	ret
OnSaveStorage endp

;--- command "Create Object" from context menu executed

OnLoadObject proc

local pStorage:LPSTORAGE
local pStorageHlp:ptr CStorageHlp
local pStorage2:LPSTORAGE
local hItem:HTREEITEM
local wszName[128]:word

	mov eax, m_hItem
	.if (!eax)
		invoke TreeView_GetSelection( m_hWndTV)
	.endif
	.if (eax)
		mov hItem, eax
		invoke Create@CStorageHlp, m_pStorage, m_hWndTV, hItem, addr wszName, 128,
			STGM_READWRITE or STGM_SHARE_EXCLUSIVE, addr pStorageHlp, addr pStorage
		.if (eax == S_OK)
			invoke vf(pStorage, IStorage, OpenStorage), addr wszName,
					NULL, STGM_READWRITE or STGM_SHARE_EXCLUSIVE, NULL, NULL, addr pStorage2
			.if (eax == S_OK)
				invoke OnLoadStorage, pStorage2, pStorageHlp
				invoke vf(pStorage2, IUnknown, Release)
			.else
				invoke OutputMessage, m_hWnd, eax, CStr("OpenStorage Error"), 0
			.endif
			invoke vf(pStorageHlp, IStorageHlp, Release)
		.else
			invoke OutputMessage, m_hWnd, eax, CStr("OpenStorage Error"), 0
		.endif
	.endif
	ret
OnLoadObject endp


streamincb proc uses esi __this dwCookie:ptr CViewStorageDlg, pbBuff:LPBYTE , cb:LONG , pcb:ptr LONG

local	dwRead:DWORD

	mov __this, dwCookie
	.if (!m_pStream)
		jmp exit
	.endif
	mov dwRead, NULL
	invoke vf(m_pStream, IStream, Read), pbBuff, cb, addr dwRead
	.if (FAILED(eax))
		invoke OutputMessage, m_hWnd, eax, CStr("IStream::Read"), 0
	.endif
	.if ((eax == S_OK) || dwRead)
		mov edx,pcb
		mov ecx, dwRead
		mov [edx],ecx
		.if (eax != S_OK)
			mov eax, 1
		.endif
		ret
	.endif
exit:
	return 1
	align 4

streamincb endp

streamoutcb proc uses esi __this dwCookie:DWORD, pbBuff:LPBYTE , cb:LONG , pcb:ptr LONG

local	dwWritten:DWORD

	mov __this, dwCookie
	.if (!m_pStream)
		jmp exit
	.endif
	mov dwWritten, NULL
	invoke vf(m_pStream, IStream, Write), pbBuff, cb, addr dwWritten
	.if (FAILED(eax))
		invoke OutputMessage, m_hWnd, eax, CStr("IStream::Write"), 0
	.endif
	.if ((eax == S_OK) || dwWritten)
		mov edx,pcb
		mov ecx, dwWritten
		mov [edx],ecx
		.if (eax != S_OK)
			mov eax, 1
		.endif
		ret
	.endif
exit:
	return 1
	align 4

streamoutcb endp

streamout2cb proc uses esi dwCookie:DWORD, pbBuff:LPBYTE , cb:LONG , pcb:ptr LONG

local	dwWritten:DWORD

	mov dwWritten, NULL
	invoke WriteFile, dwCookie, pbBuff, cb, addr dwWritten, NULL
	.if (!eax)
	.else
		mov edx,pcb
		mov ecx, dwWritten
		mov [edx],ecx
		return 0
	.endif
	return 1
	align 4

streamout2cb endp

;--- returns size in eax

ReadStream	proc

local estrm:EDITSTREAM
local szText[128]:byte

		invoke EnableWindow, m_hWndHE, TRUE
		mov estrm.dwCookie, __this
		mov estrm.dwError,0
		mov estrm.pfnCallback,offset streamincb
		invoke SendMessage, m_hWndHE, EM_STREAMIN, SF_TEXT, addr estrm
		
		invoke SendMessage, m_hWndHE, EM_SETMODIFY, FALSE, 0
		invoke GetDlgItem, m_hWnd, IDC_UPDATE
		invoke EnableWindow, eax, FALSE

		invoke SendMessage, m_hWndHE, HEM_GETSIZE, 0, 0
		push eax
		invoke wsprintf, addr szText, CStr("%u Bytes"), eax
		StatusBar_SetText m_hWndSB, 1, addr szText
		pop eax
		ret
ReadStream endp

;--- returns

WriteStream proc

local estrm:EDITSTREAM
local qwSize:LARGE_INTEGER

	xor eax, eax
	mov dword ptr qwSize+0, eax
	mov dword ptr qwSize+4, eax
	invoke vf(m_pStream, IStream, SetSize), qwSize
	mov estrm.dwCookie, __this
	mov estrm.dwError,0
	mov estrm.pfnCallback,offset streamoutcb
	invoke SendMessage, m_hWndHE, EM_STREAMOUT, SF_TEXT, addr estrm

	invoke SendMessage, m_hWndHE, EM_SETMODIFY, FALSE, 0
	invoke GetDlgItem, m_hWnd, IDC_UPDATE
	invoke EnableWindow, eax, FALSE

	ret

WriteStream endp

	.const

g_szPropStgNameToFmtId	db "PropStgNameToFmtId",0
if ?USESTGIMPL
g_szStgCreatePropStg	db "StgCreatePropStg",0
g_szStgOpenPropStg		db "StgOpenPropStg",0
endif

	.data

g_pfnPropStgNameToFmtId LPPROPSTGNAMETOFMT	NULL
if ?USESTGIMPL
externdef g_pfnStgCreatePropStg:LPSTGCREATEPROPSTG
externdef g_pfnStgOpenPropStg:LPSTGOPENPROPSTG
g_pfnStgCreatePropStg	LPSTGCREATEPROPSTG	NULL
g_pfnStgOpenPropStg		LPSTGOPENPROPSTG	NULL
endif

	.code

InitIProp proc public uses ebx

		.if (!g_pfnPropStgNameToFmtId)
			invoke GetModuleHandle, CStr("OLE32")
			.if (eax)
step1:
				mov ebx, eax
				invoke GetProcAddress, ebx, addr g_szPropStgNameToFmtId
				.if (!eax)
					invoke LoadLibrary, CStr("IPROP")
					.if (eax)
						jmp step1
					.else
						jmp done
					.endif
				.endif
				mov g_pfnPropStgNameToFmtId, eax
if ?USESTGIMPL
				invoke GetProcAddress, ebx, addr g_szStgCreatePropStg
				mov g_pfnStgCreatePropStg, eax
				invoke GetProcAddress, ebx, addr g_szStgOpenPropStg
				mov g_pfnStgOpenPropStg, eax
endif
			.endif
		.endif
		mov eax, 1
done:
		ret
		align 4
InitIProp endp

IsPropertyStorage proc hItem:HTREEITEM, pfmtid:ptr FMTID

local	hr:DWORD
local	tvi:TVITEM
local	szName[64]:byte
local	wszName[64]:word

		invoke InitIProp

		mov hr, E_FAIL
		mov eax, hItem
		mov tvi.hItem, eax
		mov tvi.mask_, TVIF_TEXT or TVIF_PARAM
		lea eax,szName
		mov tvi.pszText, eax
		mov tvi.cchTextMax, sizeof szName
		invoke TreeView_GetItem( m_hWndTV, addr tvi)
		.if (eax)
			mov eax,tvi.lParam
			.if (eax && (eax != -1))
				sub eax,2
				lea ecx,szName
				mov byte ptr [ecx+eax],0
			.endif
			invoke MultiByteToWideChar, CP_ACP, 0, addr szName, -1, addr wszName, 64
			.if (szName == 5)
				.if (g_pfnPropStgNameToFmtId)
					invoke g_pfnPropStgNameToFmtId, addr wszName, pfmtid
					mov hr, eax
				.else
					mov hr, E_FAIL
				.endif
			.endif
		.endif
		return hr
		align 4

IsPropertyStorage endp

	.const
PSGUID_DOCUMENTSUMMARYINFORMATION GUID {0d5cdd502h, 02e9ch, 0101bh, {093h, 97h, 008h, 000h, 02bh, 02ch, 0f9h, 0aeh}}
	.code

StartPropertyViewer	proc hItem:HTREEITEM, pfmtid:ptr FMTID

local	pPropSetStg:LPPROPERTYSETSTORAGE
local	pPropertyStorage:LPPROPERTYSTORAGE
local	pStorage:LPSTORAGE
local	pStorageHlp:ptr CStorageHlp
local	wszName[128]:word

	invoke Create@CStorageHlp, m_pStorage, m_hWndTV, hItem, addr wszName, 128,
			STGM_READ or STGM_SHARE_EXCLUSIVE, addr pStorageHlp, addr pStorage
	.if (eax == S_OK)
		invoke vf(pStorage, IUnknown, QueryInterface), addr IID_IPropertySetStorage, addr pPropSetStg
		.if (eax == S_OK)
			invoke vf(pPropSetStg, IPropertySetStorage, Open), pfmtid,
				STGM_READ or STGM_SHARE_EXCLUSIVE, addr pPropertyStorage
			.if (eax == S_OK)
				invoke IsEqualGUID, pfmtid, addr PSGUID_DOCUMENTSUMMARYINFORMATION
				.if (eax)
					mov ecx, pPropSetStg
				.else
					xor ecx, ecx
				.endif
				push ecx
				invoke ViewPropertyStorage, m_hWnd, pPropertyStorage, ecx
				pop ecx
				.if (!ecx)
					invoke vf(pPropertyStorage, IUnknown, Release)
				.endif
			.else
				invoke OutputMessage, m_hWnd, eax, CStr("IPropertySetStorage::Open"), 0
			.endif
			invoke vf(pPropSetStg, IUnknown, Release)
		.endif
		invoke vf(pStorageHlp, IStorageHlp, Release)
	.endif
	ret
	align 4

StartPropertyViewer endp

ShowContextMenu proc uses esi

local	pt:POINT
local	hPopupMenu:HMENU
local	tvht:TV_HITTESTINFO
local	tvi:TVITEM
local	fmtid:FMTID

		invoke GetCursorPos,addr tvht.pt
										; get the item below hit point
		invoke ScreenToClient, m_hWndTV, addr tvht.pt
		invoke TreeView_HitTest( m_hWndTV, addr tvht)
		.if (tvht.hItem)
			mov eax, tvht.hItem
			mov m_hItem, eax
			mov tvi.hItem, eax
			mov tvi.mask_, TVIF_PARAM
			invoke TreeView_GetItem( m_hWndTV, addr tvi)
			invoke CreatePopupMenu
			mov esi, eax
			mov ecx, MF_STRING
			.if ((tvi.lParam != 0) && (tvi.lParam != -1))
				or ecx, MF_GRAYED
			.endif
			invoke AppendMenu, esi, ecx, IDM_RENAME, CStr("&Rename")
			invoke AppendMenu, esi, MF_STRING, IDM_DELETE, CStr("&Delete")
			mov ecx, MF_STRING or MF_GRAYED
			.if (tvi.lParam && (tvi.lParam != -1))
				mov ecx, MF_STRING
			.endif
			invoke AppendMenu, esi, ecx, IDM_LOADOBJECT, CStr("&Create Object")

			mov ecx, MF_STRING or MF_GRAYED
			invoke IsPropertyStorage, tvi.hItem, addr fmtid
			mov ecx, MF_STRING or MF_GRAYED
			.if (eax == S_OK)
				mov ecx, MF_STRING
			.endif
			invoke AppendMenu, esi, ecx, IDM_VIEWPROP, CStr("&View Properties")

			invoke GetCursorPos,addr pt
			invoke TrackPopupMenu, esi, TPM_LEFTALIGN or TPM_LEFTBUTTON or TPM_RETURNCMD,
				pt.x,pt.y,0, m_hWnd, NULL
			push eax
			invoke DestroyMenu, esi
			pop eax
			.if (eax)
				.if (eax == IDM_VIEWPROP)
					invoke StartPropertyViewer, tvi.hItem, addr fmtid
				.else
					invoke SendMessage, m_hWnd, WM_COMMAND, eax, 0
				.endif
			.endif
			mov m_hItem, NULL
		.endif
		ret
		align 4

ShowContextMenu endp

CheckSave proc hItem:HTREEITEM, bReset:BOOL

		.if (!m_bIsStream)
			.if (!hItem)
				invoke TreeView_GetSelection( m_hWndTV)
				mov hItem, eax
				.if (!eax)
					jmp done
				.endif
			.endif
		.endif
		invoke SendMessage, m_hWndHE, EM_GETMODIFY, 0, 0
		.if (eax)
			invoke MessageBox, m_hWnd, CStr("Modified stream hasn't been written yet. Write now?"),\
				addr g_szWarning, MB_YESNOCANCEL or MB_DEFBUTTON3 or MB_ICONQUESTION
			.if (eax == IDYES)
				invoke WriteElement, hItem
				.if (eax != S_OK)
					return 0
				.endif
			.elseif (eax == IDNO)
				.if (bReset)
					invoke SendMessage, m_hWndHE, EM_SETMODIFY, FALSE, 0
				.endif
			.else
				return 0
			.endif
		.endif
done:
		return 1
		align 4

CheckSave endp


DestroyElement proc

local pStorage:LPSTORAGE
local pStorageHlp:ptr CStorageHlp
local hItem:HTREEITEM
local wszName[128]:word

	mov eax, m_hItem
	.if (!eax)
		invoke TreeView_GetSelection( m_hWndTV)
	.endif
	.if (eax)
		mov hItem, eax
		invoke Create@CStorageHlp, m_pStorage, m_hWndTV, hItem, addr wszName, 128,
			STGM_READWRITE or STGM_SHARE_EXCLUSIVE, addr pStorageHlp, addr pStorage
		.if (eax == S_OK)
			invoke MessageBox,m_hWnd,CStr("Are you sure?"),CStr("Delete element"),
					MB_YESNO or MB_DEFBUTTON2 or MB_ICONQUESTION
			.if (eax == IDYES)
				invoke vf(pStorage, IStorage, DestroyElement), addr wszName
				.if (eax != S_OK)
					invoke OutputMessage, m_hWnd, eax, CStr("IStorage::DestroyElement"), 0
					mov eax, S_FALSE
				.endif
			.else
				mov eax, S_FALSE
			.endif
			push eax
			invoke vf(pStorageHlp, IStorageHlp, Release)
			pop eax
			.if (eax != S_OK)
				return FALSE
			.endif
			invoke TreeView_DeleteItem( m_hWndTV, hItem)
			mov eax,TRUE
		.endif
	.endif
	ret
	align 4

DestroyElement endp


RenameElement proc hItem:HTREEITEM, pszNewName:LPSTR

local pStorage:LPSTORAGE
local pStorageHlp:ptr CStorageHlp
local hr:DWORD
local wszName[128]:word
local wszNewName[128]:word

	invoke MultiByteToWideChar, CP_ACP, 0, pszNewName, -1, addr wszNewName, 128
	invoke Create@CStorageHlp, m_pStorage, m_hWndTV, hItem, addr wszName, 128,
		STGM_READWRITE or STGM_SHARE_EXCLUSIVE, addr pStorageHlp, addr pStorage
	mov hr, eax
	.if (eax == S_OK)
		invoke vf(pStorage, IStorage, RenameElement), addr wszName, addr wszNewName
		mov hr, eax
		.if (eax != S_OK)
			invoke OutputMessage, m_hWnd, eax, CStr("IStorage::RenameElement"), 0
		.endif
		invoke vf(pStorageHlp, IStorageHlp, Release)
	.endif
	return hr
	align 4

RenameElement endp



WriteElement proc hItem:HTREEITEM

local pStorage:LPSTORAGE
local pStorageHlp:ptr CStorageHlp
local hr:HRESULT
local wszName[128]:word

	.if (!m_bIsStream)
		.if (!hItem)
			return S_OK
		.endif
		invoke Create@CStorageHlp, m_pStorage, m_hWndTV, hItem, addr wszName, 128,
				STGM_READWRITE or STGM_SHARE_EXCLUSIVE, addr pStorageHlp, addr pStorage
		mov hr, eax
		.if (eax == S_OK)
			invoke vf(pStorage, IStorage, OpenStream), addr wszName,
				NULL, STGM_WRITE or STGM_SHARE_EXCLUSIVE, NULL, addr m_pStream
			mov hr, eax
			.if (hr == S_OK)
				invoke WriteStream
				invoke vf(m_pStream, IStream, Release)
				mov m_pStream, NULL
			.else
				invoke OutputMessage, m_hWnd, eax, CStr("IStorage::OpenStream"), 0
			.endif
			invoke vf(pStorageHlp, IStorageHlp, Release)
		.else
			invoke OutputMessage, m_hWnd, eax, CStr("IStorage::OpenStorage"), 0
		.endif
	.else
		invoke vf(m_pStorage, IStream, Clone), addr m_pStream
		.if (eax == S_OK)
			invoke vf(m_pStream, IStream, Seek), g_dqNull, STREAM_SEEK_SET, NULL
			invoke WriteStream
			invoke vf(m_pStream, IStream, Release)
			mov m_pStream, NULL
		.endif
	.endif

	return hr
	align 4

WriteElement endp


RecalcSize proc uses esi dwYPos:DWORD

local rect:RECT
local rect2:RECT
local dwGripSize:DWORD
local dwWidth:DWORD

	invoke GetWindowRect, m_hWndSplit, addr rect
	mov eax, rect.bottom
	sub eax, rect.top
	mov dwGripSize, eax

	invoke GetWindowRect, m_hWndTV, addr rect
	invoke ScreenToClient, m_hWnd, addr rect
	invoke ScreenToClient, m_hWnd, addr rect.right
	mov eax, dwYPos
	mov rect.bottom, eax
	mov eax, rect.right
	sub eax, rect.left
	mov dwWidth, eax

	invoke BeginDeferWindowPos, 3
	mov esi, eax

;------------------------------------ set treeview control
	mov ecx, rect.bottom
	mov edx, rect.top
	sub ecx, edx
	.if (CARRY?)
		xor ecx, ecx
		mov rect.bottom, edx
		@mov dwGripSize, 0
	.endif
	invoke DeferWindowPos, esi, m_hWndTV, NULL, 0, 0, dwWidth, ecx, SWP_NOMOVE or SWP_NOZORDER or SWP_NOACTIVATE

	mov ecx, rect.bottom
	mov rect.top, ecx
	invoke DeferWindowPos, esi, m_hWndSplit, NULL, rect.left, rect.top, dwWidth, dwGripSize, SWP_NOZORDER or SWP_NOACTIVATE

	invoke GetWindowRect, m_hWndHE, addr rect2
	invoke ScreenToClient, m_hWnd, addr rect2.right
	mov ecx, rect.top
	add ecx, dwGripSize
	mov rect.top, ecx
	mov eax, rect2.bottom
	sub eax, ecx
	invoke DeferWindowPos, esi, m_hWndHE, NULL, rect.left, rect.top, dwWidth, eax, SWP_NOZORDER or SWP_NOACTIVATE

	invoke EndDeferWindowPos, esi
	ret
	align 4
RecalcSize endp

;--- WM_NOTIFY

OnNotify proc uses esi lParam:ptr NMTREEVIEW

local pStorage:LPSTORAGE
local pStorage2:LPSTORAGE
local pStorageHlp:ptr CStorageHlp
local statstg:STATSTG
local tvi:TVITEM
local szText[128]:byte
local wszText[128]:WORD


	xor eax, eax
	mov esi,lParam
	.if ([esi].NMHDR.idFrom == IDC_TREE1)

		.if ([esi].NMHDR.code == NM_RCLICK)

			invoke ShowContextMenu

		.elseif ([esi].NMHDR.code == TVN_KEYDOWN)

			.if ([esi].NMTVKEYDOWN.wVKey == VK_DELETE)
				invoke OnCommand, IDM_DELETE, 0
			.endif

		.elseif ([esi].NMHDR.code == TVN_DELETEITEM)

			invoke TreeView_GetCount( m_hWndTV)
			.if (eax == 1)
				invoke SendMessage, m_hWndHE, HEM_RESET, 0, 0
				invoke EnableWindow, m_hWndHE, FALSE
			.endif

		.elseif ([esi].NMHDR.code == TVN_SELCHANGING)

			invoke CheckSave, [esi].NMTREEVIEW.itemOld.hItem, TRUE
			.if (!eax)
				invoke SetWindowLong, m_hWnd, DWL_MSGRESULT, 1
				mov eax, 1
				jmp done
			.endif

		.elseif ([esi].NMHDR.code == TVN_SELCHANGED)

			invoke SendMessage, m_hWndHE, HEM_RESET, 0, 0
			invoke EnableWindow, m_hWndHE, FALSE
			StatusBar_SetText m_hWndSB, 1, addr g_szNull

			.if (![esi].NMTREEVIEW.itemNew.hItem)
				jmp done
			.endif

			.if ([esi].NMTREEVIEW.itemNew.lParam == 0)

				invoke SwitchStorage, TRUE
				invoke Create@CStorageHlp, m_pStorage, m_hWndTV, [esi].NMTREEVIEW.itemNew.hItem,\
					addr wszText, 128, STGM_READ or STGM_SHARE_EXCLUSIVE, addr pStorageHlp, addr pStorage
				.if (eax == S_OK)
					invoke vf(pStorage, IStorage, OpenStream), addr wszText,\
						NULL, STGM_READ or STGM_SHARE_EXCLUSIVE, NULL, addr m_pStream
					.if (eax == S_OK)
						invoke ReadStream
						invoke vf(m_pStream, IStream, Release)
						mov m_pStream, NULL
					.else
						invoke OutputMessage, m_hWnd, eax, CStr("IStorage::OpenStream"), 0
					.endif
					invoke vf(pStorageHlp, IStorageHlp, Release)
				.else
					invoke OutputMessage, m_hWnd, eax, CStr("IStorage::OpenStorage"), 0
				.endif
				invoke SwitchStorage, FALSE
			.else
				StatusBar_SetText m_hWndSB, 1, CStr("Storage")
			.endif

		.elseif ([esi].NMHDR.code == TVN_BEGINLABELEDIT)

			DebugOut "TVN_BEGINLABELEDIT"

			.if (m_bAllowEdit)
				mov m_bEditLabel, TRUE
			.else
				invoke SetWindowLong, m_hWnd, DWL_MSGRESULT, 1
			.endif
			mov eax, 1

		.elseif ([esi].NMHDR.code == TVN_ENDLABELEDIT)

			DebugOut "TVN_ENDLABELEDIT"
			mov m_bEditLabel, FALSE
			.if ([esi].NMTVDISPINFO.item.pszText)
				invoke RenameElement, [esi].NMTVDISPINFO.item.hItem, [esi].NMTVDISPINFO.item.pszText
				.if (eax == S_OK)
					invoke SetWindowLong, m_hWnd, DWL_MSGRESULT, 1
				.endif
			.endif
			mov eax, 1

		.endif

	.elseif ([esi].NMHDR.idFrom == IDC_SPLITBTN)

		.if ([esi].NMHDR.code == SBN_SETSIZE)
			invoke RecalcSize, [esi].SBNOTIFY.iPos
		.endif
    .endif
done:
    ret
	align 4

OnNotify endp

;--- insert one storage branch

InsertBranch    proc hParent:HANDLE, pStorage:LPSTORAGE

local	pEnumSTATSTG:LPENUMSTATSTG
local	statstg:STATSTG
local	pszError:LPSTR
local	tvi:TV_INSERTSTRUCT
local	dwFetched:DWORD
local	szName[256]:BYTE
local	szText[64]:BYTE
local	wszCLSID[40]:WORD
local	szCLSID[40]:BYTE
local	pStorage2:LPSTORAGE
local	hItem:HANDLE

	.if (pStorage == NULL)
		ret
	.endif

	mov tvi.hInsertAfter,TVI_LAST
	mov eax,hParent
	mov tvi.hParent, eax

	invoke vf(pStorage, IStorage, EnumElements), NULL, NULL, NULL, addr pEnumSTATSTG
	.if (eax == S_OK)
		.while (1)
			invoke vf(pEnumSTATSTG, IEnumSTATSTG, Next), 1, addr statstg, 0
			.break .if (eax != S_OK)
			.if (statstg.pwcsName)
				invoke WideCharToMultiByte, CP_ACP, 0, statstg.pwcsName, -1, addr szName, sizeof szName, NULL, NULL
			.else
				invoke lstrcpy, addr szName, CStr("<unnamed>")
			.endif

			.if (statstg.type_ == STGTY_STORAGE)
				invoke IsEqualGUID, addr statstg.clsid, addr IID_NULL
				.if (eax)
					mov tvi.item.lParam, -1
				.else
					invoke GetTextFromCLSID, addr statstg.clsid, addr szText, sizeof szText
					invoke StringFromGUID2, addr statstg.clsid, addr wszCLSID, sizeof wszCLSID
					invoke WideCharToMultiByte, CP_ACP, 0, addr wszCLSID, -1, addr szCLSID, sizeof szCLSID, NULL, NULL
					invoke lstrcat, addr szName, CStr(", ")
					invoke lstrlen, addr szName
					mov tvi.item.lParam,eax
					lea edx, szName
					add edx, eax
					invoke wsprintf, edx, CStr("%s, %s"), addr szCLSID, addr szText
				.endif
			.else
				mov tvi.item.lParam, 0
			.endif
			mov tvi.item.mask_,TVIF_TEXT or TVIF_PARAM
			lea eax,szName
			mov tvi.item.pszText,eax
			invoke TreeView_InsertItem( m_hWndTV, addr tvi)
			mov hItem,eax
			.if (statstg.type_ == STGTY_STORAGE)
				invoke vf(pStorage,IStorage, OpenStorage), statstg.pwcsName, NULL,\
					STGM_READ or STGM_SHARE_EXCLUSIVE, 0, 0, addr pStorage2
				.if (eax == S_OK)
					invoke InsertBranch, hItem, pStorage2
					invoke vf(pStorage2, IStorage, Release)
				.else
					invoke OutputMessage, m_hWnd, eax, CStr("IStorage::OpenStorage Error"), 0
					invoke MessageBeep, MB_OK
				.endif
			.endif
			invoke CoTaskMemFree, statstg.pwcsName
		.endw
		invoke vf(pEnumSTATSTG, IEnumSTATSTG, Release)
	.endif
	ret
	align 4

InsertBranch endp

IsStream proc

local pUnknown:LPUNKNOWN

	mov m_bIsStream, TRUE
	invoke vf(m_pStorage, IUnknown, QueryInterface), addr IID_IStorage, addr pUnknown
	.if (eax == S_OK)
		mov m_bIsStream, FALSE
		invoke vf(pUnknown, IUnknown, Release)
	.endif
	movzx eax, m_bIsStream
	ret
	align 4

IsStream endp

;--- refill treeview on WM_INITDIALOG or WM_COMMAND/IDM_REFRESH


RefreshList proc

local statstg:STATSTG
local szFileName[MAX_PATH]:byte
local szTitle[MAX_PATH+80]:byte
local wszCLSID[40]:WORD
local szCLSID[40]:BYTE

	mov szFileName, 0
	invoke IsStream
	.if (eax)
		invoke vf(m_pStorage, IStream, Seek), g_dqNull, STREAM_SEEK_SET, NULL
		invoke ReadClassStm, m_pStorage, addr statstg.clsid
	.else
		invoke vf(m_pStorage, IStorage, Stat), addr statstg, STATFLAG_DEFAULT
		.if (eax == S_OK)
			invoke WideCharToMultiByte, CP_ACP, 0, statstg.pwcsName, -1, addr szFileName, sizeof szFileName, NULL, NULL
			invoke CoTaskMemFree, statstg.pwcsName
			mov eax, S_OK
		.endif
	.endif
	.if (eax == S_OK)
		invoke StringFromGUID2, addr statstg.clsid, addr wszCLSID, sizeof wszCLSID
		invoke WideCharToMultiByte, CP_ACP, 0, addr wszCLSID, -1, addr szCLSID, sizeof szCLSID, NULL, NULL
		.if (m_bIsStream)
			mov ecx, CStr("CLSID=%s")
		.else
			mov ecx, CStr("CLSID=%s, Type=%u, Mode=%X")
		.endif
		invoke wsprintf, addr szTitle, ecx,
			addr szCLSID, statstg.type_, statstg.grfMode
		StatusBar_SetText m_hWndSB, 0, addr szTitle
	.endif

	.if ((!szFileName) && m_pszFile)
		invoke lstrcpy, addr szFileName, m_pszFile
	.endif

	invoke IsEqualGUID, addr statstg.clsid, addr IID_NULL
	.if (eax)
		invoke GetDlgItem, m_hWnd, IDM_LOADSTORAGE
		invoke EnableWindow, eax, FALSE
	.endif

	.if (m_bIsStream)
		mov ecx, CStr("Stream %s")
	.else
		mov ecx, CStr("Storage %s")
	.endif
	invoke wsprintf, addr szTitle, ecx, addr szFileName
	invoke SetWindowText, m_hWnd, addr szTitle

	invoke TreeView_SelectItem( m_hWndTV, NULL)

	invoke SetWindowRedraw( m_hWndTV, FALSE)

	invoke TreeView_DeleteAllItems( m_hWndTV)

	invoke SendMessage, m_hWndHE, HEM_RESET, 0, 0
	invoke EnableWindow, m_hWndHE, FALSE

	invoke SwitchStorage, TRUE
	.if (!m_bIsStream)
		invoke InsertBranch, NULL, m_pStorage
	.else
		invoke EnableWindow, m_hWndTV, FALSE
		invoke vf(m_pStorage, IStream, Clone), addr m_pStream
		.if (eax == S_OK)
			invoke vf(m_pStream, IStream, Seek), g_dqNull, STREAM_SEEK_SET, NULL
			invoke ReadStream
			invoke vf(m_pStream, IStream, Release)
			mov m_pStream, NULL
		.endif
	.endif

	invoke SwitchStorage, FALSE


	invoke SetWindowRedraw( m_hWndTV, TRUE)

	ret
	align 4

RefreshList endp


OnCommand proc wParam:WPARAM, lParam:LPARAM

local	charrange:CHARRANGE

	movzx eax,word ptr wParam

	.if (eax == IDCANCEL)

		.if (m_bEditLabel)
			invoke TreeView_EndEditLabelNow( m_hWndTV, TRUE)
		.else
			invoke PostMessage, m_hWnd, WM_CLOSE, 0, 0
		.endif
		mov eax, 1

	.elseif (eax == IDOK)

		movzx eax, word ptr wParam+2
		.if (m_bEditLabel && (!eax))
			invoke TreeView_EndEditLabelNow( m_hWndTV, FALSE)
		.endif

	.elseif (eax == IDM_REFRESH)

		invoke CheckSave, NULL, TRUE
		.if (eax)
			invoke RefreshList
		.endif

	.elseif (eax == IDM_SAVESTREAM)		;not used????

		invoke WriteElement, m_hItem

	.elseif (eax == IDM_DELETE)

		invoke DestroyElement
		.if (eax == NULL)
			invoke MessageBeep, MB_OK
		.endif

	.elseif (eax == IDM_RENAME)

		mov m_bAllowEdit, TRUE
		invoke TreeView_EditLabel( m_hWndTV, m_hItem)
		mov m_bAllowEdit, FALSE

	.elseif (eax == IDM_SELECTALL)

		.if (m_bHexEdHasFocus)
			mov charrange.cpMin, 0
			mov charrange.cpMax, -1
			invoke SendMessage, m_hWndHE, EM_EXSETSEL, 0, addr charrange
		.endif

	.elseif (eax == IDC_CUSTOM1)

		movzx ecx, word ptr wParam+2
		.if (ecx == EN_SETFOCUS)
			mov m_bHexEdHasFocus, TRUE
		.elseif (ecx == EN_KILLFOCUS)
			mov m_bHexEdHasFocus, FALSE
		.elseif (ecx == EN_CHANGE)
			invoke GetDlgItem, m_hWnd, IDC_UPDATE
			invoke EnableWindow, eax, TRUE
		.endif

	.elseif (eax == IDM_LOADSTORAGE)

		invoke OnLoadStorage, m_pStorage, NULL

	.elseif (eax == IDM_LOADOBJECT)

		invoke OnLoadObject

	.elseif (eax == IDM_STG2FILE)

		invoke OnSaveStorage

	.elseif (eax == IDM_OBJECTDLG)

		invoke Find@CObjectItem, m_pStorage
		.if (eax)
			invoke vf(eax, IObjectItem, GetObjectDlg)
			invoke RestoreAndActivateWindow, [eax].CDlg.hWnd
		.else
			invoke Create@CObjectItem, m_pStorage, NULL
			.if (eax)
				push eax
				invoke vf(eax, IObjectItem, ShowObjectDlg), m_hWnd
				pop eax
				invoke vf(eax, IObjectItem, Release)
			.endif
		.endif

	.elseif (eax == IDC_COMMIT)

		invoke CheckSave, NULL, FALSE
		.if (eax)
			.if (m_bIsStream)
				invoke vf(m_pStorage, IStream, Commit), STGC_ONLYIFCURRENT
			.else
				invoke vf(m_pStorage, IStorage, Commit), STGC_ONLYIFCURRENT
			.endif
			.if (eax != S_OK)
				invoke OutputMessage, m_hWnd, eax, CStr("Commit Error"), 0
			.endif
		.endif

	.elseif (eax == IDC_REVERT)

		invoke MessageBox, m_hWnd, CStr("This will discard all changes since last commit. Continue?"),\
				addr g_szWarning, MB_YESNO or MB_DEFBUTTON2 or MB_ICONQUESTION
		.if (eax == IDYES)
			.if (m_bIsStream)
				invoke vf(m_pStorage, IStream, Revert)
			.else
				invoke vf(m_pStorage, IStorage, Revert)
			.endif
			.if (eax != S_OK)
				invoke OutputMessage, m_hWnd, eax, CStr("Revert Error"), 0
			.else
				invoke RefreshList
			.endif
		.endif

	.elseif (eax == IDC_UPDATE)

		invoke TreeView_GetSelection( m_hWndTV)
		.if (eax)
			invoke WriteElement, eax
		.endif

	.else
		xor eax, eax
	.endif
	ret

	align 4

OnCommand endp


OnInitDialog proc

local rect:RECT
local dwWidth[2]:DWORD

	invoke GetDlgItem, m_hWnd, IDC_TREE1
	mov m_hWndTV,eax

	invoke GetDlgItem, m_hWnd, IDC_STATUSBAR
	mov m_hWndSB,eax

	invoke GetDlgItem, m_hWnd, IDC_CUSTOM1
	mov m_hWndHE, eax

	invoke GetDlgItem, m_hWnd, IDC_SPLITBTN
	mov m_hWndSplit, eax

	invoke Create@CSplittButton, m_hWndSplit, m_hWndTV, m_hWndHE

	invoke GetClientRect, m_hWnd, addr rect
	invoke MulDiv, rect.right, 3, 4
	mov dwWidth[0*sizeof DWORD], eax
	mov dwWidth[1*sizeof DWORD], -1
	StatusBar_SetParts m_hWndSB, 2, addr dwWidth

	invoke IsStream
	.if (eax)
		invoke ShowWindow, m_hWndSplit, SW_HIDE
		invoke RecalcSize, 0
	.endif

	invoke PostMessage, m_hWnd, WM_COMMAND, IDM_REFRESH, 0

	ret
	align 4

OnInitDialog endp


;*** dialog proc to displays infos about control/container

ViewStorageDlgProc proc uses __this this_:ptr CViewStorageDlg, uMsg:DWORD, wParam:WPARAM, lParam:LPARAM

	mov __this, this_

	mov eax,uMsg
	.if (eax == WM_INITDIALOG)

		invoke OnInitDialog
		invoke ShowWindow, m_hWnd, SW_SHOWNORMAL
		mov eax, 1

	.elseif (eax == WM_COMMAND)

		invoke OnCommand, wParam, lParam

	.elseif (eax == WM_NOTIFY)

		invoke OnNotify, lParam

	.elseif (eax == WM_CLOSE)

		invoke CheckSave, NULL, TRUE
		.if (eax)
			invoke DestroyWindow, m_hWnd
		.endif
		mov eax,1

	.elseif (eax == WM_DESTROY)

		invoke Destroy@CViewStorageDlg, __this

if ?MODELESS
	.elseif (eax == WM_ACTIVATE)

		movzx eax,word ptr wParam
		.if (eax == WA_INACTIVE)
			mov g_hWndDlg, NULL
		.else
			mov eax, m_hWnd
			mov g_hWndDlg, eax
		.endif
endif
	.elseif (eax == WM_ENTERMENULOOP)

		StatusBar_SetSimpleMode m_hWndSB, TRUE

	.elseif (eax == WM_EXITMENULOOP)

		StatusBar_SetSimpleMode m_hWndSB, FALSE

	.elseif (eax == WM_MENUSELECT)

		movzx ecx, word ptr wParam+0
		invoke DisplayStatusBarString, m_hWndSB, ecx
if ?HTMLHELP
	.elseif (eax == WM_HELP)

		invoke DoHtmlHelp, HH_DISPLAY_TOPIC, CStr("viewstoragedialog.htm")
endif

	.else

		xor eax,eax

	.endif
	ret
	align 4

ViewStorageDlgProc endp

;*** WM_COMMAND/IDM_VIEWSTORAGE

Create@CViewStorageDlg proc public uses __this pStorage:LPSTORAGE, pszFile:LPSTR, pObjectItem:ptr CObjectItem

ifdef _DEBUG
local	this_:ptr CViewStorageDlg
endif

	invoke malloc, sizeof CViewStorageDlg
	.if (eax == NULL)
		ret
	.endif
	mov __this,eax
ifdef _DEBUG
	mov this_, eax
endif

	mov eax, ViewStorageDlgProc
	mov m_pDlgProc, eax
	mov eax, pStorage
	mov m_pStorage, eax
	.if (eax)
		invoke vf(eax, IStorage, AddRef)
	.endif
	.if (pszFile)
		invoke lstrlen, pszFile
		inc eax
		invoke malloc, eax
		mov m_pszFile, eax
		invoke lstrcpy, eax, pszFile
	.endif
if ?HANDSOFF
	mov eax, pObjectItem
	mov m_pObjectItem, eax
	.if (eax)
		invoke vf(eax, IObjectItem, AddRef)
	.endif
endif
	invoke Init@CHexEdit, g_hInstance

	return __this
	align 4

Create@CViewStorageDlg endp

Show@CViewStorageDlg proc public uses __this thisarg, hWnd:HWND

	invoke CreateDialogParam, g_hInstance, IDD_VIEWSTORAGEDLG, hWnd, classdialogproc, this@
	ret
	align 4

Show@CViewStorageDlg endp


Destroy@CViewStorageDlg proc public uses __this thisarg

	mov __this, this@
	.if (m_pszFile)
		invoke free, m_pszFile
	.endif
if ?HANDSOFF
	.if (m_pObjectItem)
		invoke vf(m_pObjectItem, IObjectItem, Release)
	.endif
endif
	.if (m_pStorage)
		invoke FindStorage@CObjectItem, m_pStorage
		.if (eax)
			invoke vf(eax, IObjectItem, SetViewStorageDlg), NULL
		.endif
		invoke vf(m_pStorage, IUnknown, Release)
		DebugOut "CViewStorageDlg::Destroy Release(pStorage) returned %u", eax
	.endif
    invoke free, __this
	ret
	align 4

Destroy@CViewStorageDlg endp

	end
