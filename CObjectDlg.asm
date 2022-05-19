
;*** definition of CObjectDlg methods
;*** the object dialog pops up when an object is created
;*** showing "all" interfaces the object supports.

	.386
	.model flat,stdcall
	option casemap :none   ; case sensitive
	option proc:private

	include COMView.inc
	include commctrl.inc
	include servprov.inc
	include statusbar.inc
INSIDE_COBJECTDLG equ 1
	include classes.inc
	include rsrc.inc
	include CEditDlg.inc
	include debugout.inc

externdef IID_DataSourceListener:IID


MAX_OUTGOING		equ 64	;limit scan for outgoing interfaces (winword bug)
EDITMODELESS		equ 1
?MODELESS			equ 1
?MAXCP				equ 12	;maximum of connection points to enumerate
?BOLDCONN			equ 1
?SHOWICON			equ 1	;set icon from DefaultIcon entry as dialog icon
?CHECKREF			equ 1	;check if MultiQI is correct
?CREATEPROXY		equ 0	;

ViewOLEView proto protoViewProc

BEGIN_CLASS CObjectDlg, CDlg
hWndLV				HWND ?
hWndLVOut			HWND ?
hWndSB				HWND ?
iSortCol			DWORD ?
iSortDir			DWORD ?
iSortColOut			DWORD ?
iSortDirOut			DWORD ?
guid				GUID <>
pUnknown			LPUNKNOWN ?
pObjectItem			LPOBJECTITEM ?
pItem				pCInterfaceItem ?		;currently selected item
pItemOut			pCInterfaceItem ?		;currently selected item
dwListId			DWORD ?					;ID of listview having focus
ViewProc			LPVIEWPROC ?			;interface viewer proc
pt					POINT <>				;dialog window pos
if ?SHOWICON
hIcon				HICON ?
endif
bRefresh			BOOLEAN ?				;list refresh currently running
END_CLASS

__this	textequ <edi>
_this	textequ <[__this].CObjectDlg>
thisarg	textequ <this@:ptr CObjectDlg>


	MEMBER hWnd, pDlgProc
	MEMBER hWndLV, hWndLVOut, hWndSB, iSortCol, iSortDir, guid, pUnknown
	MEMBER pObjectItem, pItem, dwListId, iSortColOut, iSortDirOut
	MEMBER ViewProc, bRefresh, pItemOut
	MEMBER pt
if ?SHOWICON
	MEMBER hIcon
endif

SafeRelease	proto :LPUNKNOWN

;--- private methods

IsViewerAvailable	proto
IsDispatch			proto :ptr CInterfaceItem
GetMenuEnableVector	proto
IsConnected			proto :REFIID
GetMenuDefaultCmd	proto :DWORD
OnConnect			proto
OnDisconnect		proto


ShowContextMenu	proto :BOOL
InsertLine		proto :HWND, pIID:REFIID, pUnknown:LPUNKNOWN, iItems:dword, dwIndex:DWORD
RefreshList		proto
OnEdit			proto
OnView			proto pItem:ptr pCInterfaceItem
OnNotify		proto :ptr NMLISTVIEW
OnCommand		proto wParam:WPARAM, lParam:LPARAM
OnInitDialog	proto
ObjectDialog	proto thisarg, message:dword,wParam:WPARAM,lParam:LPARAM

	.const

pColumns label CColHdr
	CColHdr <CStr("IID"),			38>
	CColHdr <CStr("Name"),			30>
	CColHdr <CStr("Addr Interface"),16>
	CColHdr <CStr("Addr VTable"),	16>
NUMCOLS textequ %($ - pColumns) / sizeof CColHdr

GUIDCOL_IN_OBJECT equ 0

pColumnsOut label CColHdr
	CColHdr <CStr("IID"),			42>
	CColHdr <CStr("Name"),			42>
	CColHdr <CStr("Connections"),	16>
NUMCOLSOUT textequ %($ - pColumnsOut) / sizeof CColHdr

pColumnsVtbl label CColHdr
	CColHdr <CStr("#/Name"),		50>
	CColHdr <CStr("Offset"),		20>
	CColHdr <CStr("Value"),			30>
NUMCOLSVTBL textequ %($ - pColumnsVtbl) / sizeof CColHdr

g_szNoTypeInfo	db "no typeinfo available",0
g_szQueryBlanket db "IClientSecurity:QueryBlanket",0

	.data

if ?SHOWICON
g_hIconObj	HICON NULL
endif

	.code

DisplayError proc pszFormatString:LPSTR, hr:DWORD

local	szText[260]:BYTE

		invoke wsprintf, addr szText, pszFormatString, hr
		StatusBar_SetText m_hWndSB, 0, addr szText
		invoke printf@CLogWindow, CStr("%s",10), addr szText
		invoke MessageBeep, MB_OK
		ret
		align 4

DisplayError endp


IsRunning proc

local pOleObject:LPOLEOBJECT
local bRC:BOOL

		mov bRC, FALSE
		invoke vf(m_pUnknown, IUnknown, QueryInterface), addr IID_IOleObject, addr pOleObject
		.if (eax == S_OK)
			invoke OleIsRunning, pOleObject
			mov bRC, eax
			.if (eax == TRUE)
				mov eax, CStr("running")
			.else
				mov eax, CStr("not running")
			.endif
			StatusBar_SetText m_hWndSB, 1, eax
			invoke vf(pOleObject, IUnknown, Release)
			invoke GetDlgItem, m_hWnd, IDC_RUN
			mov ecx, bRC
			xor ecx, 1
			invoke EnableWindow, eax, ecx
		.endif

		return bRC
		align 4

IsRunning endp

EnableLockBtn proc
		invoke vf(m_pObjectItem, IObjectItem, IsLocked)
		push eax
		invoke GetDlgItem, m_hWnd, IDC_LOCK
		pop ecx
		.if (ecx)
			mov ecx, FALSE
		.else
			mov ecx, TRUE
		.endif
		invoke EnableWindow, eax, ecx
		ret
		align 4

EnableLockBtn endp

IsStorageActive proc

local pStorage:LPSTORAGE
local bEnable:BOOL

		mov bEnable, FALSE
		invoke vf(m_pObjectItem, IObjectItem, GetStorage)
		.if (eax)
			mov bEnable, TRUE
		.endif
		invoke GetDlgItem, m_hWnd, IDM_VIEWSTORAGEDLG
		invoke EnableWindow, eax, bEnable
		ret
		align 4

IsStorageActive endp

if 1
IsDispatch proc uses ebx pItem:ptr CInterfaceItem

local	pTypeLib:LPTYPELIB
local	pTypeInfo:LPTYPEINFO
local	pTypeAttr:ptr TYPEATTR
local	dwRC:DWORD

		mov dwRC, FALSE
		mov ebx, pItem
		assume ebx:ptr CInterfaceItem
		invoke LoadRegTypeLib, addr [ebx].TypelibGUID,
			[ebx].dwVerMajor, [ebx].dwVerMinor,	g_LCID, addr pTypeLib
		.if (eax == S_OK)
			invoke vf(pTypeLib, ITypeLib, GetTypeInfoOfGuid), addr [ebx].iid, addr pTypeInfo
			.if (eax == S_OK)
				invoke vf(pTypeInfo, ITypeInfo, GetTypeAttr), addr pTypeAttr
				.if (eax == S_OK)
					mov ecx, pTypeAttr
					.if ([ecx].TYPEATTR.typekind == TKIND_DISPATCH)
						mov dwRC, TRUE
					.endif
					invoke vf(pTypeInfo, ITypeInfo, ReleaseTypeAttr), pTypeAttr
				.endif
				invoke vf(pTypeInfo, ITypeInfo, Release)
			.endif
			invoke vf(pTypeLib, ITypeLib, Release)
		.endif
		return dwRC
		assume ebx:nothing
		align 4

IsDispatch endp
endif

IsViewerAvailable proc uses esi ebx

local	hKey:HANDLE
local	iMax:DWORD
local	wszIID[40]:word
;;local	szIID[40]:byte
local	szKey[128]:byte
local	szValue[128]:byte

		mov esi, m_pItem
		invoke StringFromGUID2, addr [esi].CInterfaceItem.iid, addr wszIID, 40
;;		invoke WideCharToMultiByte,CP_ACP,0,addr wszIID,-1,addr szIID,sizeof szIID,0,0 
		invoke wsprintf, addr szKey, CStr("%s\%S\OLEViewerIViewerCLSID"), addr g_szInterface, addr wszIID

		invoke RegOpenKeyEx, HKEY_CLASSES_ROOT,addr szKey, 0, KEY_READ, addr hKey
		.if (eax == ERROR_SUCCESS)
			mov iMax, sizeof szValue
			invoke RegQueryValueEx, hKey, addr g_szNull, NULL, NULL, addr szValue, addr iMax
			push eax
			invoke RegCloseKey, hKey
			pop eax
			.if (eax == ERROR_SUCCESS)
				mov m_ViewProc, offset ViewOLEView
				return TRUE
			.endif
		.endif

		mov esi, offset InterfaceViewerTab
		.while (dword ptr [esi])
			mov edx, m_pItem
			invoke IsEqualGUID, addr [edx].CInterfaceItem.iid, dword ptr [esi]
			.if (eax)
				mov eax, [esi+4]
				mov m_ViewProc, eax
				return TRUE
			.endif
			add esi, 2 * sizeof DWORD
		.endw
done:
		return FALSE
		align 4

IsViewerAvailable endp


;--- do some menu stuff

MENUF_VIEWOBJECT	equ 1
MENUF_TYPEINFO		equ 2
MENUF_PROPERTY		equ 4
MENUF_CONNECT		equ 8
MENUF_VIEWINTERFACE	equ 10h

ViewObjectIIDs label dword
		DWORD offset IID_IViewObject
		DWORD offset IID_IViewObject2
		DWORD offset IID_IOleObject
		DWORD offset IID_IObjectWithSite
EndViewObjectIIDs label dword

GetMenuEnableVector proc  uses esi ebx

local	pDispatch:LPDISPATCH

local	dwFlags:DWORD

		mov dwFlags, 0
		.if (m_dwListId == IDC_LIST1)
			mov ebx, m_pItem
		.else
			mov ebx, m_pItemOut
		.endif
		.if (!ebx)
			jmp done
		.endif
		assume ebx:ptr CInterfaceItem

		mov esi, offset ViewObjectIIDs
		.while (1)
			.break .if (esi >= offset EndViewObjectIIDs)
			lodsd
			invoke IsEqualGUID,addr [ebx].iid, eax
			.if (eax)
				or dwFlags, MENUF_VIEWOBJECT
				.break
			.endif
		.endw

		.while (ebx)
			.if ([ebx].dwVerMajor != -1)
				or dwFlags, MENUF_TYPEINFO
				.break
			.endif
			invoke IsEqualGUID,addr [ebx].iid, addr IID_IProvideClassInfo
			.if (eax)
				or dwFlags, MENUF_TYPEINFO
				.break
			.endif
			invoke IsEqualGUID,addr [ebx].iid, addr IID_IDispatch
			.if (eax)
				or dwFlags, MENUF_TYPEINFO
			.endif
			.break
		.endw
;------------------------ a properties dialog requires typeinfo
		.if (dwFlags & MENUF_TYPEINFO)
;------------------------ NEW: allow properties dialog if typeinfo available
if 1
			or dwFlags, MENUF_PROPERTY
else
			.if ([ebx].dwVerMajor == -1)
				or dwFlags, MENUF_PROPERTY
			.else
;------------------------ this query is NOT sufficient
if 0
				invoke vf(m_pUnknown, IUnknown, QueryInterface), addr [ebx].iid, addr pDispatch
				.if (eax == S_OK)
					or dwFlags, MENUF_PROPERTY
					invoke vf(pDispatch, IUnknown, Release)
				.endif
else
				invoke IsDispatch, ebx
				.if (eax)
					or dwFlags, MENUF_PROPERTY
				.endif
endif
			.endif
endif
		.endif

;------------------------ alloc "connect" to dispatchable interfaces
;------------------------ IPropertyNotifySink and DataSourceListener only

		.if (m_dwListId == IDC_LIST2)
			.repeat
				invoke IsEqualGUID, addr [ebx].iid, addr IID_DataSourceListener
				.break .if (eax)
				invoke IsEqualGUID, addr [ebx].iid, addr IID_IPropertyNotifySink
				.break .if (eax)
				.if ([ebx].dwVerMajor != -1)
					invoke IsDispatch, ebx
				.endif
			.until (1)
			.if (eax)
				or dwFlags, MENUF_CONNECT
			.endif
		.else
			invoke IsViewerAvailable
			.if (eax)
				or dwFlags, MENUF_VIEWINTERFACE
			.endif
		.endif
done:
		return dwFlags
		assume ebx:nothing
		align 4

GetMenuEnableVector endp

IsConnected proc uses esi riid:REFIID

local	pList:ptr CList
local	pConnection:ptr CConnection

		invoke vf(m_pObjectItem, IObjectItem, GetConnectionList), FALSE
		mov pList,eax
		.if (eax)
			xor esi, esi
			.while (1)
				invoke GetItem@CList, pList, esi
				.break .if (!eax)
				mov pConnection, eax
				invoke IsEqualGUID@CConnection, eax, riid
				.break .if (eax)
				inc esi
			.endw
		.endif
		ret
		align 4

IsConnected endp



GetMenuDefaultCmd proc menuflags:DWORD

		mov eax, menuflags
		.if (eax & MENUF_CONNECT)
			mov ecx, m_pItemOut
			invoke IsConnected, addr [ecx].CInterfaceItem.iid
			.if (eax)
				mov eax, IDM_DISCONNECT
			.else
				mov eax, IDM_CONNECT
			.endif
		.elseif (eax & MENUF_VIEWINTERFACE)
			mov eax, IDM_VIEWINTERFACE
		.elseif (eax & MENUF_PROPERTY)
			mov eax, IDM_PROPERTIESDLG
		.elseif (eax & MENUF_TYPEINFO)
			mov eax, IDM_TYPEINFODLG
		.elseif (eax & MENUF_VIEWOBJECT)
			mov eax, IDM_VIEWOBJECT
		.else
			mov eax, IDM_EDIT
		.endif
		ret
		align 4

GetMenuDefaultCmd endp



OnDisconnect proc uses esi

local	pList:ptr CList
local	dwItem:DWORD
local	pConnection:ptr CConnection
local	szText[32]:byte

		invoke vf(m_pObjectItem, IObjectItem, GetConnectionList), FALSE
		.if (!eax)
			jmp done
		.endif
		mov pList, eax

		xor esi, esi
		.while (1)
			invoke GetItem@CList, pList, esi
			.break .if (!eax)
			mov pConnection, eax
			mov ecx, m_pItemOut
			invoke IsEqualGUID@CConnection, pConnection, addr [ecx].CInterfaceItem.iid
			.if (eax)
				invoke DeleteItem@CList, pList, esi
				invoke Disconnect@CConnection, pConnection, m_hWnd
				.if (eax == S_OK)
					invoke vf(m_pObjectItem, IObjectItem, SetRunLock), FALSE
				.endif
				invoke Destroy@CConnection, pConnection
				invoke RefreshObjectView
if ?BOLDCONN
				invoke ListView_GetNextItem( m_hWndLVOut, -1, LVNI_SELECTED)
				mov dwItem, eax
				invoke FindAllItemData@CList, pList, dwItem
				mov szText, 0
				.if (eax)
					invoke wsprintf, addr szText, CStr("%u"), eax
				.endif
				lea ecx, szText
				ListView_SetItemText m_hWndLVOut, dwItem, 2, ecx
				invoke InvalidateRect, m_hWndLVOut, 0, 1
endif
				invoke IsRunning
				.break
			.endif
			inc esi
		.endw
done:
		ret
		align 4

OnDisconnect endp


OnConnect proc uses esi

local	pConnection:pCConnection
local	pConnectionPointContainer:LPCONNECTIONPOINTCONTAINER
local	pszError:LPSTR
local	dwIndex:DWORD
local	dwItem:DWORD
local	szText[128]:byte

		invoke vf(m_pObjectItem, IObjectItem, GetConnectionList), TRUE
		mov esi, eax
		mov ecx, m_pItemOut
		invoke Create@CConnection, m_pObjectItem, addr [ecx].CInterfaceItem.iid
		.if (eax)
			mov pConnection, eax
			invoke Connect@CConnection, pConnection, m_pUnknown, m_hWnd, addr pszError
			.if (eax != S_OK)
				invoke DisplayError, pszError, eax
				invoke Destroy@CConnection, pConnection
			.else
				invoke AddItem@CList, esi, pConnection
				mov dwIndex, eax
				invoke RefreshObjectView
				invoke vf(m_pObjectItem, IObjectItem, SetRunLock), TRUE
if ?BOLDCONN
				invoke ListView_GetNextItem( m_hWndLVOut, -1, LVNI_SELECTED)
				mov dwItem, eax
				invoke SetItemData@CList, esi, dwIndex, eax
				invoke FindAllItemData@CList, esi, dwItem
				invoke wsprintf, addr szText, CStr("%u"), eax
				lea ecx, szText
				ListView_SetItemText m_hWndLVOut, dwItem, 2, ecx
				invoke InvalidateRect, m_hWndLVOut, 0, 1
endif
			.endif
		.endif

		ret
		align 4

OnConnect endp


;--- WM_INITDIALOG for vtbl dialog

OnInitDialogVtbl proc uses ebx hWnd:HWND

local hKey:HANDLE
local dwNumMethods:DWORD
local dwSize:DWORD
local hWndLV:HWND
local dwVtbl:DWORD
local pTypeLib:LPTYPELIB
local pTypeInfo:LPTYPEINFO
local pTypeAttr:ptr TYPEATTR
local pFuncDesc:ptr FUNCDESC
local dwNames:DWORD
local bstrName:BSTR
local bUseTypeInfo:BOOL
local dwOffset:DWORD
local lvi:LVITEM
local wszIID[40]:word
local szText[128]:byte
local szText2[128]:byte

		invoke GetDlgItem, hWnd, IDC_LIST1
		mov hWndLV, eax
		invoke ListView_SetExtendedListViewStyle( hWndLV, LVS_EX_FULLROWSELECT)
		invoke SetLVColumns, hWndLV, NUMCOLSVTBL, offset pColumnsVtbl

		mov dwNumMethods, 3

		mov ecx, m_pItem
		invoke StringFromGUID2, addr [ecx].CInterfaceItem.iid, addr wszIID, 40

;------------------------------------------ set window title
		invoke wsprintf, addr szText, CStr("%s\%S"), addr g_szInterface, addr wszIID
		invoke RegOpenKeyEx, HKEY_CLASSES_ROOT, addr szText, 0, KEY_READ, addr hKey
		.if (eax == ERROR_SUCCESS)
			mov dwSize, sizeof szText
			invoke RegQueryValueEx, hKey, addr g_szNull, NULL, NULL, addr szText, addr dwSize
			.if (eax == ERROR_SUCCESS)
				invoke GetWindowText, hWnd, addr szText2, sizeof szText2
				sub esp, 256
				mov edx, esp
				invoke wsprintf, edx, CStr("%s %s"), addr szText2, addr szText
				mov edx, esp
				invoke SetWindowText, hWnd, edx
				add esp, 256
			.endif
			invoke RegCloseKey, hKey
		.endif

		invoke wsprintf, addr szText, CStr("%s\%S\NumMethods"), addr g_szInterface, addr wszIID

		invoke RegOpenKeyEx, HKEY_CLASSES_ROOT, addr szText, 0, KEY_READ, addr hKey
		.if (eax == ERROR_SUCCESS)
			mov dwSize, sizeof szText
			invoke RegQueryValueEx, hKey, addr g_szNull, NULL, NULL, addr szText, addr dwSize
			.if (eax == ERROR_SUCCESS)
				invoke String2DWord, addr szText, addr dwNumMethods
			.endif
			invoke RegCloseKey, hKey
		.endif

		mov pTypeInfo, NULL
		mov pTypeAttr, NULL
		mov bUseTypeInfo, FALSE
		mov ebx, m_pItem
		assume ebx:ptr CInterfaceItem
		.if ([ebx].dwVerMajor != -1)
			invoke LoadRegTypeLib, addr [ebx].TypelibGUID,
					[ebx].dwVerMajor, [ebx].dwVerMinor, g_LCID, addr pTypeLib
			.if (eax == S_OK)
				invoke vf(pTypeLib, ITypeLib, GetTypeInfoOfGuid), addr [ebx].iid, addr pTypeInfo
				.if (eax == S_OK)
					invoke vf(pTypeInfo, ITypeInfo, GetTypeAttr), addr pTypeAttr
					.if (eax == S_OK)
						mov ecx, pTypeAttr
;;						.if ([ecx].TYPEATTR.wTypeFlags & TYPEFLAG_FDUAL)
							mov bUseTypeInfo, TRUE
							movzx eax, [ecx].TYPEATTR.cFuncs
							mov dwNumMethods, eax
;;						.else
;;							movzx eax, [ecx].TYPEATTR.cbSizeVft
;;							shr eax, 2
;;							mov dwNumMethods, eax
;;						.endif
					.endif
				.endif
				invoke vf(pTypeLib, ITypeLib, Release)
			.endif
		.endif
		assume ebx:nothing

		invoke ListView_GetNextItem( m_hWndLV, -1, LVNI_SELECTED)
		mov lvi.iItem, eax
		mov lvi.iSubItem, 3
		lea eax, szText
		mov lvi.pszText, eax
		mov lvi.cchTextMax, sizeof szText
		mov lvi.mask_, LVIF_TEXT
		invoke ListView_GetItem( m_hWndLV, addr lvi)
		invoke String2Number, addr szText, addr dwVtbl, 16

		mov lvi.iItem, 0
		.while (dwNumMethods)
			mov lvi.iSubItem, 0
			mov eax, lvi.iItem
			shl eax, 2
			mov dwOffset, eax
			invoke wsprintf, addr szText, CStr("%u"), lvi.iItem
			.if (bUseTypeInfo)
				invoke vf(pTypeInfo, ITypeInfo, GetFuncDesc), lvi.iItem, addr pFuncDesc
				.if (eax == S_OK)
					mov ecx, pFuncDesc
					movzx eax, [ecx].FUNCDESC.oVft
					mov dwOffset, eax
					invoke vf(pTypeInfo, ITypeInfo, GetNames), [ecx].FUNCDESC.memid, addr bstrName, 1, addr dwNames
					.if (eax == S_OK)
						mov ecx, pFuncDesc
						.if ([ecx].FUNCDESC.invkind == INVOKE_PROPERTYGET)
							invoke lstrcpy, addr szText, CStr("get_")
							lea ecx, szText+4
						.elseif ([ecx].FUNCDESC.invkind == INVOKE_PROPERTYPUT)
							invoke lstrcpy, addr szText, CStr("put_")
							lea ecx, szText+4
						.else
							lea ecx, szText
						.endif
						invoke WideCharToMultiByte, CP_ACP, 0, bstrName, -1, ecx, sizeof szText, 0, 0 
						invoke SysFreeString, bstrName
					.endif
					invoke vf(pTypeInfo, ITypeInfo, ReleaseFuncDesc), pFuncDesc
				.endif
			.endif
			lea eax, szText
			mov lvi.pszText, eax
			mov lvi.mask_, LVIF_TEXT
			invoke ListView_InsertItem( hWndLV, addr lvi)

			inc lvi.iSubItem
			invoke wsprintf, addr szText, CStr("%u"), dwOffset
			invoke ListView_SetItem( hWndLV, addr lvi)

			inc lvi.iSubItem
			mov eax, dwOffset
			add eax, dwVtbl
			mov eax, [eax]
			invoke wsprintf, addr szText, CStr("%X"), eax
			invoke ListView_SetItem( hWndLV, addr lvi)

			inc lvi.iItem
			dec dwNumMethods
		.endw
		.if (pTypeInfo)
			.if (pTypeAttr)
				invoke vf(pTypeInfo, ITypeInfo, ReleaseTypeAttr), pTypeAttr
			.endif
			invoke vf(pTypeInfo, ITypeInfo, Release)
		.endif
		ret
		align 4

OnInitDialogVtbl endp

vtbldlgproc proc uses __this hWnd:HWND, message:DWORD, wParam:WPARAM, lParam:LPARAM

		mov eax, message
		.if (eax == WM_INITDIALOG)
			mov __this, lParam
			invoke SetWindowLong, hWnd, DWL_USER, __this
			invoke OnInitDialogVtbl, hWnd
			mov eax, 1
		.elseif (eax == WM_CLOSE)
			invoke EndDialog, hWnd, 0
		.elseif (eax == WM_COMMAND)
			movzx eax, word ptr wParam+0
			.if (eax == IDCANCEL)
				invoke EndDialog, hWnd, 0
			.endif
		.else
			xor eax, eax
		.endif
		ret
		align 4

vtbldlgproc endp


IID_IOleViewViewer IID {0fc37e5bah, 4a8eh, 11ceh, {87h,0bh,08h,00h,36h,8dh,23h,02h}}

BEGIN_INTERFACE IInterfaceViewer, IUnknown
	STDMETHOD View, hwndParent:HWND, riid:REFIID, punk:LPUNKNOWN
END_INTERFACE

LPINTERFACEVIEWER typedef ptr IInterfaceViewer

;--- params not used

ViewOLEView proc uses esi hWnd:HWND, pUnknown:LPUNKNOWN, pCLSID:ptr CLSID

local	hKey:HANDLE
local	iMax:DWORD
local	clsid:CLSID
local	pViewer:LPINTERFACEVIEWER
local	wszIID[40]:word
local	szKey[128]:byte
local	szValue[128]:byte

	mov esi, m_pItem
	invoke StringFromGUID2, addr [esi].CInterfaceItem.iid, addr wszIID, 40
	invoke wsprintf, addr szKey, CStr("%s\%S\OLEViewerIViewerCLSID"), addr g_szInterface, addr wszIID

	invoke RegOpenKeyEx, HKEY_CLASSES_ROOT,addr szKey, 0, KEY_READ, addr hKey
	.if (eax == ERROR_SUCCESS)
		mov iMax, sizeof szValue
		invoke RegQueryValueEx, hKey, addr g_szNull, NULL, NULL, addr szValue, addr iMax
		push eax
		invoke RegCloseKey, hKey
		pop eax
		.if (eax == ERROR_SUCCESS)
			invoke MultiByteToWideChar,CP_ACP,MB_PRECOMPOSED,
					addr szValue, -1, addr wszIID, 40 
			invoke CLSIDFromString, addr wszIID, addr clsid
			invoke CoCreateInstance, addr clsid, NULL, CLSCTX_INPROC_SERVER,
					addr IID_IOleViewViewer, addr pViewer
			.if (eax == S_OK)
				mov ecx, g_pMainDlg
				invoke EnableWindow, [ecx].CDlg.hWnd, FALSE
				invoke vf(pViewer, IInterfaceViewer, View), m_hWnd, addr [esi].CInterfaceItem.iid, m_pUnknown
				invoke vf(pViewer, IUnknown, Release)
				mov ecx, g_pMainDlg
				invoke EnableWindow, [ecx].CDlg.hWnd, TRUE
			.endif
		.endif
	.endif
	ret
	align 4

ViewOLEView endp



;*** user pressed right mouse button, show context menu ***


ShowContextMenu proc uses ebx esi bMouse:BOOL

local	dwFlags1:DWORD
local	dwFlags2:DWORD
local	pt:POINT
local	dwDefault:DWORD
local	hWndLV:HWND
local	pItem:ptr CInterfaceItem
local	pClientSecurity:ptr IClientSecurity

		.if (m_dwListId == IDC_LIST1)
			mov ecx, m_hWndLV
			mov eax, m_pItem
		.else
			mov ecx, m_hWndLVOut
			mov eax, m_pItemOut
		.endif
		mov hWndLV, ecx
		mov pItem, eax
		invoke ListView_GetSelectedCount( ecx)
		.if (!eax)
			ret
		.endif
		.if (m_dwListId == IDC_LIST1)
			invoke GetSubMenu, g_hMenu, ID_SUBMENU_OBJECTDLG
		.else
			invoke GetSubMenu, g_hMenu, ID_SUBMENU_OUT_OBJECTDLG
		.endif
		.if (eax != 0)
			mov ebx, eax

			invoke GetMenuEnableVector
			mov esi, eax

			.if (esi & MENUF_TYPEINFO)
				mov ecx, MF_ENABLED
			.else
				mov ecx, MF_GRAYED
			.endif
			push ecx
			invoke EnableMenuItem, ebx, IDM_TYPEINFODLG, ecx
			pop ecx
			invoke EnableMenuItem, ebx, IDM_TYPELIBDLG, ecx

			.if (m_dwListId == IDC_LIST1)
				.if (esi & MENUF_VIEWINTERFACE)
					mov ecx, MF_ENABLED
				.else
					mov ecx, MF_GRAYED
				.endif
				invoke EnableMenuItem, ebx, IDM_VIEWINTERFACE, ecx

				.if (esi & MENUF_VIEWOBJECT)
					mov ecx, MF_ENABLED
				.else
					mov ecx, MF_GRAYED
				.endif
				invoke EnableMenuItem, ebx, IDM_VIEWOBJECT, ecx

				.if (esi & MENUF_PROPERTY)
					mov ecx, MF_ENABLED
				.else
					mov ecx, MF_GRAYED
				.endif
				invoke EnableMenuItem, ebx, IDM_PROPERTIESDLG, ecx

			.else
				mov dwFlags1, MF_GRAYED or MF_DISABLED
				mov dwFlags2, MF_GRAYED or MF_DISABLED
				.if (esi & MENUF_CONNECT)
					mov dwFlags1, MF_ENABLED
					mov ecx, pItem
					invoke IsConnected, addr [ecx].CInterfaceItem.iid
					.if (eax)
						mov dwFlags2, MF_ENABLED
					.endif
				.endif
				invoke EnableMenuItem, ebx, IDM_CONNECT, dwFlags1
				invoke EnableMenuItem, ebx, IDM_DISCONNECT, dwFlags2
			.endif

			mov dwFlags1, MF_GRAYED or MF_DISABLED
			invoke vf(m_pUnknown, IUnknown, QueryInterface), addr IID_IClientSecurity, addr pClientSecurity
			.if (eax == S_OK)
				mov dwFlags1, MF_ENABLED
				invoke vf(pClientSecurity, IUnknown, Release)
			.endif
			invoke EnableMenuItem, ebx, IDM_SECURITY, dwFlags1

			invoke GetMenuDefaultCmd, esi
			invoke SetMenuDefaultItem, ebx, eax, FALSE

			invoke GetItemPosition, hWndLV, bMouse, addr pt
			invoke TrackPopupMenu, ebx, TPM_LEFTALIGN or TPM_LEFTBUTTON,
					pt.x, pt.y, 0, m_hWnd, NULL

		.endif
		ret
		align 4

ShowContextMenu endp


;*** insert 1 line in listview


InsertLine proc hWndLV:HWND, pIID:REFIID, pUnknown:LPUNKNOWN, iItems:dword, dwIndex:DWORD

local	hSubKey:HANDLE
local	iType:dword
local	lvi:LVITEM
local	pII:ptr CInterfaceItem
local	iMax:dword
local	dwTypeIndex:DWORD
local	pTypeLib:LPTYPELIB
local	pTLibAttr:ptr TLIBATTR
local	pTypeInfo:LPTYPEINFO
local	pTypeInfo2:LPTYPEINFO
local	bstr:BSTR
local	wszIID[40]:word
local	szIID[80]:byte
local	szText[128]:byte

		invoke Create@CInterfaceItem, pIID
		mov pII, eax

		invoke StringFromGUID2, pIID, addr wszIID, LENGTHOF wszIID
		invoke WideCharToMultiByte,CP_ACP,0,addr wszIID,-1,addr szIID,sizeof szIID,0,0

		mov lvi.mask_,LVIF_TEXT or LVIF_PARAM
		mov eax,iItems
		mov lvi.iItem,eax
		mov lvi.iSubItem,0
		lea eax,szIID
		mov lvi.pszText,eax
		mov eax, pII
		mov lvi.lParam, eax
		invoke ListView_InsertItem( hWndLV,addr lvi)
		inc lvi.iSubItem

		mov ecx, pII
		.if (![ecx].CInterfaceItem.pszName)
			invoke GetName@CInterfaceList, g_pcdi, dwIndex
			.if (eax)
				invoke SetName@CInterfaceItem, pII, eax
			.else
				invoke vf(m_pObjectItem, IObjectItem, GetCoClassTypeInfo), addr pTypeInfo
				.if (pTypeInfo)
					invoke vf(pTypeInfo, ITypeInfo, GetContainingTypeLib), addr pTypeLib, addr dwTypeIndex
					.if (eax == S_OK)
						invoke vf(pTypeLib, ITypeLib, GetTypeInfoOfGuid), pIID, addr pTypeInfo2
						.if (eax == S_OK)
							invoke vf(pTypeLib, ITypeLib, GetLibAttr), addr pTLibAttr
							.if (eax == S_OK)
								invoke SetTypeLibAttr@CInterfaceItem, pII, pTLibAttr
								invoke vf(pTypeLib, ITypeLib, ReleaseTLibAttr), pTLibAttr
							.endif
							invoke vf(pTypeInfo2, ITypeInfo, GetDocumentation), MEMBERID_NIL, addr bstr, NULL, NULL, NULL
							.if (eax == S_OK)
								invoke WideCharToMultiByte, CP_ACP, 0, bstr, -1, addr szText, sizeof szText, 0, 0
								invoke SetName@CInterfaceItem, pII, addr szText
								invoke SysFreeString, bstr
							.endif
							invoke vf(pTypeInfo2, IUnknown, Release)
						.endif
						invoke vf(pTypeLib, IUnknown, Release)
					.endif
					invoke vf(pTypeInfo, IUnknown, Release)
				.endif
			.endif
		.endif

		mov lvi.mask_,LVIF_TEXT
		mov ecx, pII
		mov eax, [ecx].CInterfaceItem.pszName
		.if (!eax)
			mov eax, CStr("<?>")
		.endif
		mov lvi.pszText,eax
		invoke ListView_SetItem( hWndLV,addr lvi)
		inc lvi.iSubItem

		.if (pUnknown)
			lea ecx, szText
			mov lvi.pszText, ecx
			invoke wsprintf, ecx, CStr("%08X"), pUnknown
			invoke ListView_SetItem( hWndLV,addr lvi)
			inc lvi.iSubItem

			mov ecx, pUnknown
			invoke wsprintf, addr szText,CStr("%08X"),dword ptr [ecx]
			invoke ListView_SetItem( hWndLV,addr lvi)
		.endif

		ret
		align 4
InsertLine endp


;--- find default interface which will be selected then


FindDefaultInterface proc uses esi dwItem:DWORD, riid:REFIID

local	lvi:LVITEM

		.if (dwItem == -1)
			mov esi, -1
			mov lvi.iSubItem, 0
			mov lvi.mask_, LVIF_PARAM
			.while (1)
				invoke ListView_GetNextItem( m_hWndLV, esi, LVNI_ALL)
				mov esi, eax
				.break .if (eax == -1)
				mov lvi.iItem, esi
				mov lvi.lParam, 0
				invoke ListView_GetItem( m_hWndLV, addr lvi)
				mov ecx, lvi.lParam
				.if (ecx)
					invoke IsEqualGUID, addr [ecx].CInterfaceItem.iid, riid
					.if (eax)
						mov dwItem, esi
						.break
					.endif
					mov ecx, lvi.lParam
					invoke IsEqualGUID, addr [ecx].CInterfaceItem.iid, addr IID_IDispatch
					.if (eax)
						mov dwItem, esi
					.endif
				.endif
			.endw
		.endif
		.if (dwItem != -1)
			invoke ListView_EnsureVisible( m_hWndLV, dwItem, FALSE)
			ListView_SetItemState m_hWndLV, dwItem, LVIS_SELECTED or LVIS_FOCUSED, LVIS_SELECTED or LVIS_FOCUSED
		.endif
		ret
		align 4

FindDefaultInterface endp

IsObject proc

local pOleObject:LPOLEOBJECT

	invoke vf(m_pUnknown, IUnknown, QueryInterface), addr IID_IOleObject,addr pOleObject
	.if (eax == S_OK)
		invoke vf(pOleObject, IUnknown, Release)
		return TRUE
	.endif
	return FALSE
	align 4

IsObject endp

;*** check for all interfaces an object supports

ifdef @StackBase
	option stackbase:ebp
endif
	option prologue:@sehprologue
	option epilogue:@sehepilogue

RefreshList proc uses ebx esi edi

local	this@:ptr CObjectDlg
local	hKey:HANDLE
local	pList:ptr CList
local	hr:DWORD
local	pMQI:ptr MULTI_QI
local	pMultiQI:LPMULTIQI
local	numInterfaces:DWORD
local	pECP:LPENUMCONNECTIONPOINTS
local	pConnectionPoint:LPCONNECTIONPOINT
local	pConnectionPointContainer:LPCONNECTIONPOINTCONTAINER
local	pProvideClassInfo2:LPPROVIDECLASSINFO2
local	pUnknown2:LPUNKNOWN
local	pOleObject:LPOLEOBJECT
local	pTypeInfo:LPTYPEINFO
local	pTypeAttr:ptr TYPEATTR
local   pszKey:LPSTR
local	dwCount:dword
local	dwItem:DWORD
local	dwIndex:DWORD
local	dwFetched:dword
local	bSourceEnumerated:BOOLEAN
local	hCsrOld:HCURSOR
local	iid:IID
local	szText[128]:byte

;---------------------- create interface list if not created yet 

		mov m_bRefresh, TRUE
		invoke SetCursor, g_hCsrWait
		mov hCsrOld, eax
		invoke SetWindowRedraw( m_hWndLV, FALSE)
		invoke SetBusyState@CMainDlg, TRUE

		.if (g_pcdi == NULL)
			invoke Create@CInterfaceList, m_hWnd
			mov g_pcdi,eax
			.if (!eax)
				ret
			.endif
			invoke GetInterfaces@COptions, eax
		.endif

		.if (g_bQueryMI)
			invoke vf(m_pUnknown, IUnknown, QueryInterface), addr IID_IMultiQI, addr pMultiQI
			.if (eax == S_OK)
				mov eax,g_pcdi
				mov eax,[eax].CInterfaceList.cntIID
				mov ecx,sizeof MULTI_QI
				mul ecx
				invoke malloc, eax
				.if (!eax)
					invoke vf(pMultiQI, IUnknown, Release)
					return 0
				.endif
				mov ebx,eax
				mov pMQI,eax
				mov eax,g_pcdi
				mov esi,[eax].CInterfaceList.pRegIIDs
				mov ecx,[eax].CInterfaceList.cntIID
				mov numInterfaces,ecx
				.while (ecx)
					mov [ebx].MULTI_QI.pIID, esi
					mov [ebx].MULTI_QI.pItf, NULL
					mov [ebx].MULTI_QI.hr, 0
					add esi,sizeof IID
					add ebx,sizeof MULTI_QI
					dec ecx
				.endw
if 0
				invoke vf(m_pUnknown, IUnknown, AddRef)
				dec eax
				mov ebx, eax
endif
				invoke vf(pMultiQI, IMultiQI, QueryMultipleInterfaces),
					numInterfaces, pMQI
				.if (FAILED(eax))
if ?CHECKREF
					mov hr, eax
					mov ebx,pMQI
					mov ecx,numInterfaces
					xor esi, esi
					.while (ecx)
						.if ([ebx].MULTI_QI.pItf)
							push ecx
							invoke vf([ebx].MULTI_QI.pItf, IUnknown, Release)
							pop ecx
							inc esi
						.endif
						add ebx,sizeof MULTI_QI
						dec ecx
					.endw
					.if (esi)
		    			invoke DisplayError, CStr("IMultiQI::QueryMultipleInterfaces returned %X, but didn't release all interfaces!"), hr
					.endif
endif
					invoke free, pMQI
					mov pMQI, NULL
				.endif
if 0
				invoke vf(m_pUnknown, IUnknown, Release)
				.if ((pMQI == NULL) && (eax != ebx))
					StatusBar_SetText m_hWndSB, 0, CStr("IMultiQI::QueryMultipleInterfaces implemented wrong!")
				.endif
endif
				invoke vf(m_pObjectItem, IObjectItem, SetMQI), numInterfaces, pMQI
				invoke vf(pMultiQI, IUnknown, Release)
			.endif
		.endif

		invoke ListView_DeleteAllItems( m_hWndLV)
		invoke ListView_DeleteAllItems( m_hWndLVOut)

;---------------------- open \Interface key to  read interface names from

		mov pszKey, offset g_szInterface
		mov hKey,0
		invoke RegOpenKeyEx, HKEY_CLASSES_ROOT, pszKey, 0, KEY_READ, addr hKey

		mov ebx,g_pcdi
		mov esi,[ebx].CInterfaceList.pRegIIDs
		mov ecx,[ebx].CInterfaceList.cntIID
		mov dwCount,ecx
		invoke vf(m_pObjectItem, IObjectItem, GetMQI)
		mov edx, eax
		xor ebx, ebx

		mov dwIndex, 0


		mov this@, __this
		.try

		.while (dwCount)
;---------------------- if MULTI_QI pointer is available, use it
			.if (edx)
				push edx
				.if ([edx].MULTI_QI.hr == S_OK && ([edx].MULTI_QI.pItf))
					invoke InsertLine, m_hWndLV, esi, [edx].MULTI_QI.pItf, ebx, dwIndex
					inc ebx
			    .endif
				pop edx
				add edx,sizeof MULTI_QI
			.else
;---------------------- query local object
				invoke vf(m_pUnknown, IUnknown, QueryInterface),esi,addr pUnknown2
				.if (eax == S_OK)
					invoke InsertLine, m_hWndLV, esi, pUnknown2, ebx, dwIndex
					inc ebx
					invoke vf(pUnknown2, IUnknown, Release)
			    .endif
				xor edx,edx
			.endif
			add esi,sizeof IID
			dec dwCount
			inc dwIndex
		.endw

		.exceptfilter
			mov __this,this@
			mov eax,_exception_info()
			invoke DisplayExceptionInfo, m_hWnd, eax, CStr("QueryInterface"), EXCEPTION_EXECUTE_HANDLER
		.except
			mov __this,this@
		.endtry

		invoke vf(m_pObjectItem, IObjectItem, GetDefaultInterface), FALSE, addr iid
		.if (eax == S_OK)
			mov dwItem, -1
			invoke Find@CInterfaceList, g_pcdi, addr iid
;------------------------------ if default interface is NOT in interface list
;------------------------------ insert it now
			.if (!eax)
				invoke vf(m_pUnknown, IUnknown, QueryInterface), addr iid, addr pUnknown2
				.if (eax == S_OK)
					invoke InsertLine, m_hWndLV, addr iid, pUnknown2, ebx, dwIndex
					mov dwItem, ebx
					invoke vf(pUnknown2, IUnknown, Release)
				.endif
			.endif
			invoke FindDefaultInterface, dwItem, addr iid
		.endif

;--------------------------------------------------------------------

		.try

		mov pConnectionPointContainer, NULL
		mov pECP, NULL
		mov pConnectionPoint, NULL

		mov bSourceEnumerated, FALSE
		invoke vf(m_pUnknown,IUnknown,QueryInterface),
				addr IID_IConnectionPointContainer,addr pConnectionPointContainer
		.if (eax == S_OK)
			invoke vf(m_pObjectItem, IObjectItem, GetConnectionList), FALSE
			mov pList, eax
			xor ebx, ebx
			.if (g_bUseEnumCPs)
				invoke vf(pConnectionPointContainer, IConnectionPointContainer, EnumConnectionPoints), addr pECP
				.if (eax == S_OK)
					mov bSourceEnumerated, TRUE
					mov dwCount, ?MAXCP
					.while (dwCount)
						invoke vf(pECP, IEnumConnectionPoints, Next),1, addr pConnectionPoint, addr dwFetched
						.break .if (eax != S_OK) 
						.break .if (dwFetched == 0) 
						invoke vf(pConnectionPoint, IConnectionPoint, GetConnectionInterface), addr iid
						.if (eax == S_OK)
							invoke InsertLine, m_hWndLVOut, addr iid, NULL, ebx, 0

							.if (pList)
								invoke FindAllItemData@CList, pList, ebx
								mov szText, 0
								.if (eax)
									invoke wsprintf, addr szText, CStr("%u"), eax
								.endif
								lea ecx, szText
								ListView_SetItemText m_hWndLVOut, ebx, 2, ecx
							.endif

							inc ebx
						.endif
						invoke vf(pConnectionPoint, IConnectionPoint, Release)
						mov pConnectionPoint, NULL
						dec dwCount
					.endw
					.if (!dwCount)
						invoke DisplayError, CStr("scan for outgoing interfaces canceled after %u lines"), ?MAXCP
					.endif
				.else
					invoke DisplayError, CStr("IConnectionPointContainer::EnumConnectionPoints failed [%X]"), eax
				.endif
			.endif
			.if (bSourceEnumerated == FALSE)
;---------------------------- connectionpoints not enumerated
;---------------------------- get default source interface from IObjectItem 
				invoke vf(m_pObjectItem, IObjectItem, GetDefaultInterface), TRUE, addr iid
				.if (eax == S_OK)
					invoke vf(pConnectionPointContainer, IConnectionPointContainer, FindConnectionPoint),
						addr iid, addr pConnectionPoint
				.else
if 1	;--- no default source interface, check if IConnectionPoint is exposed directly
					invoke vf(m_pUnknown, IUnknown, QueryInterface),
						addr IID_IConnectionPoint, addr pConnectionPoint
					.if (eax == S_OK)
						invoke vf(pConnectionPoint, IConnectionPoint, GetConnectionInterface), addr iid
					.endif
				.endif
endif
				.if (eax == S_OK)
					invoke InsertLine, m_hWndLVOut, addr iid, NULL, ebx, 0
					.if (pList)
						invoke FindAllItemData@CList, pList, ebx
						mov szText, 0
						.if (eax)
							invoke wsprintf, addr szText, CStr("%u"), eax
						.endif
						lea ecx, szText
						ListView_SetItemText m_hWndLVOut, ebx, 2, ecx
					.endif
				.endif
			.endif
		.endif

		.exceptfilter
			mov __this,this@
			mov eax,_exception_info()
			invoke DisplayExceptionInfo, m_hWnd, eax, CStr("Enum ConnectionPoints"), EXCEPTION_EXECUTE_HANDLER
		.except
			mov __this,this@
		.endtry

		invoke SafeRelease, pConnectionPoint
		invoke SafeRelease, pECP
		invoke SafeRelease, pConnectionPointContainer

;--------------------------------------------------------------------

	    invoke RegCloseKey,hKey

		invoke SetWindowRedraw( m_hWndLV, TRUE)

		invoke GetDlgItem, m_hWnd, IDC_VIEW
		mov ebx, eax
		invoke IsObject
		invoke EnableWindow, ebx, eax

		invoke IsRunning

		invoke SetBusyState@CMainDlg, FALSE
		invoke SetCursor, hCsrOld
		mov m_bRefresh, FALSE
		ret
		align 4

RefreshList endp

	option prologue: prologuedef
	option epilogue: epiloguedef
ifdef @StackBase
	option stackbase:esp
endif

;*** start registry editor dialog

OnEdit proc uses ebx

local	lvi:LVITEM
local	szKey1[64]:byte
local	szKey2[64]:byte
local	szKey3[64]:byte
local	szKey4[64]:byte
local	szKey[260]:byte
local	hKey:HANDLE
local	iType:dword
local	dwSize:dword
local	hInstance:HINSTANCE
local	kp[4]:KEYPAIR
local	pEditDlg:ptr CEditDlg
local	hWnd:HWND

		invoke GetFocus
		.if ((eax != m_hWndLV) && (eax != m_hWndLVOut))
			ret
		.endif
		mov hWnd, eax
		invoke ListView_GetNextItem( hWnd, -1, LVNI_SELECTED)
		.if (eax == -1)
			ret
		.endif
		mov lvi.iItem,eax
		mov lvi.iSubItem, GUIDCOL_IN_OBJECT
		mov lvi.mask_,LVIF_TEXT
		lea eax,szKey1
		mov lvi.pszText,eax
		mov lvi.cchTextMax,sizeof szKey1
		invoke ListView_GetItem( hWnd, addr lvi)

		movzx eax,g_bConfirmDelete
		invoke Create@CEditDlg, m_hWnd, EDITMODELESS, eax
		.if (eax == 0)
			ret
		.endif
		mov pEditDlg,eax

		invoke ZeroMemory, addr kp, sizeof kp
		
		mov kp[0*sizeof KEYPAIR].pszRoot, offset g_szInterface
		lea eax,szKey1
		mov kp[0*sizeof KEYPAIR].pszKey,eax
		mov kp[0*sizeof KEYPAIR].bExpand,TRUE

		invoke wsprintf, addr szKey, CStr("%s\%s\ProxyStubClsid32"), addr g_szInterface, addr szKey1
		invoke RegOpenKeyEx,HKEY_CLASSES_ROOT,addr szKey,0,KEY_READ,addr hKey
		.if (eax == ERROR_SUCCESS)
			mov dwSize,sizeof szKey2
			invoke RegQueryValueEx,hKey, addr g_szNull, NULL,addr iType,addr szKey2,addr dwSize
			.if (szKey2 != 0)
				lea eax,szKey2
				mov kp[1*sizeof KEYPAIR].pszKey,eax
				mov kp[1*sizeof KEYPAIR].pszRoot, offset g_szCLSID
			.endif
			invoke RegCloseKey,hKey
		.endif

		invoke wsprintf, addr szKey, CStr("%s\%s\TypeLib"), addr g_szInterface, addr szKey1
		invoke RegOpenKeyEx,HKEY_CLASSES_ROOT,addr szKey,0,KEY_READ,addr hKey
		.if (eax == ERROR_SUCCESS)
			mov dwSize,sizeof szKey3
			invoke RegQueryValueEx,hKey, addr g_szNull, NULL, addr iType,addr szKey3,addr dwSize
			.if (szKey3 != 0)
				lea eax,szKey3
				mov kp[2*sizeof KEYPAIR].pszKey,eax
				mov kp[2*sizeof KEYPAIR].pszRoot,offset g_szTypeLib
			.endif
		.endif

		invoke SetKeys@CEditDlg, pEditDlg,4, addr kp
		invoke Show@CEditDlg, pEditDlg
if EDITMODELESS eq 0
		invoke Destroy@CEditDlg,pEditDlg
endif
		ret
        assume ebx:nothing
		align 4

OnEdit endp

;*** try to show an ole object (IOleObject/IViewObject)

OnView proc pItem:ptr pCInterfaceItem

	.if (!pItem)
		invoke IsObject
		.if (!eax)
			jmp done
		.endif
	.endif
	mov ecx, g_pMainDlg
	invoke vf(m_pObjectItem, IObjectItem, ShowViewObjectDlg), [ecx].CDlg.hWnd, pItem
	invoke IsRunning
	mov eax, 1
done:
	ret
	align 4

OnView endp


;*** we have 2 listviews in this dialog


SetCurrentItem proc uses ebx pNMHDR:ptr NMHDR, iItem:DWORD

local	dwFlags:DWORD
local	pTypeInfo:LPTYPEINFO
local	iid:IID
local	lvi:LVITEM
local	szStr[64]:byte

	mov ebx,pNMHDR
	assume ebx:ptr NMHDR

	mov eax, iItem
	.if (eax != -1)
		mov lvi.iItem, eax
		mov lvi.iSubItem, 0
		mov lvi.mask_, LVIF_PARAM
		invoke ListView_GetItem( [ebx].hwndFrom, addr lvi)
		mov eax, lvi.lParam
		.if ([ebx].idFrom == IDC_LIST1)
			mov m_pItem, eax
		.else
			mov m_pItemOut, eax
		.endif
	.endif

	mov eax, [ebx].idFrom
	mov m_dwListId, eax

	.if (eax == IDC_LIST1)
		mov dwFlags, FALSE
		invoke vf(m_pObjectItem, IObjectItem, GetCoClassTypeInfo), addr pTypeInfo
		.if (pTypeInfo)
			mov dwFlags, TRUE
			invoke vf(pTypeInfo, ITypeInfo, Release)
		.else
			invoke GetMenuEnableVector
			.if (eax & MENUF_TYPEINFO)
				mov dwFlags, TRUE
			.endif
		.endif
		invoke GetDlgItem, m_hWnd, IDM_PROPERTIESDLG
		invoke EnableWindow, eax, dwFlags
	.endif

	ret
	align 4

SetCurrentItem endp


;--- display security settings


OnSecurity proc

local	pClientSecurity:ptr IClientSecurity
local	pUnknown:LPUNKNOWN
local	dwAuthnSvc:DWORD
local	dwAuthzSvc:DWORD
local	pwszServerPrincName:ptr WORD
local	dwAuthnLevel:DWORD
local	dwImpLevel:DWORD
local	dwCapabilities:DWORD
local	szPrincipalName[128]:byte
local	szText[256]:byte

		invoke vf(m_pUnknown, IUnknown, QueryInterface), addr IID_IClientSecurity, addr pClientSecurity
		.if (eax != S_OK)
			jmp done
		.endif
		mov ecx,m_pItem
		invoke vf(m_pUnknown, IUnknown, QueryInterface), addr [ecx].CInterfaceItem.iid, addr pUnknown
		.if (eax == S_OK)
			invoke vf(pClientSecurity, IClientSecurity, QueryBlanket), pUnknown,
				addr dwAuthnSvc, addr dwAuthzSvc, addr pwszServerPrincName,
				addr dwAuthnLevel, addr dwImpLevel, NULL, addr dwCapabilities
			.if (eax == S_OK)
				mov szPrincipalName, 0
				.if (pwszServerPrincName)
					invoke WideCharToMultiByte,CP_ACP,0, pwszServerPrincName, -1, addr szPrincipalName, sizeof szPrincipalName, 0, 0 
					invoke CoTaskMemFree, pwszServerPrincName
				.endif
				invoke wsprintf, addr szText,
					CStr("Authentication Service: %u",10, "Authorization Service: %u",10, "Principal Name: %s",10, "Authentication Level: %u", 10, "Impersonation Level: %u", 10, "Capabilities: %X",10),
					dwAuthnSvc, dwAuthzSvc, addr szPrincipalName, dwAuthnLevel, dwImpLevel, dwCapabilities
				invoke MessageBox, m_hWnd, addr szText, addr g_szQueryBlanket, MB_OK
			.else
				invoke OutputMessage, m_hWnd, eax, addr g_szQueryBlanket, 0
			.endif
			invoke vf(pUnknown, IUnknown, Release)
		.endif
		invoke vf(pClientSecurity, IUnknown, Release)
done:
		ret
		align 4

OnSecurity endp

if ?CREATEPROXY
OnCreateProxy proc

local	clsid:CLSID
local	psfb:LPPSFACTORYBUFFER
local	pProxy:LPUNKNOWN
local	pClientProxy:LPUNKNOWN
local	pUnknown:LPUNKNOWN
local	pObjectItem:LPOBJECTITEM
local	rect:RECT

		mov ecx,m_pItem
		invoke vf(m_pUnknown, IUnknown, QueryInterface), addr [ecx].CInterfaceItem.iid, addr pUnknown
		.if (eax != S_OK)
			jmp done
		.endif
		mov ecx,m_pItem
		invoke CoGetPSClsid,  addr [ecx].CInterfaceItem.iid, addr clsid
		.if (eax == S_OK)
			invoke CoGetClassObject, addr clsid, CLSCTX_INPROC, 0,
				addr IID_IPSFactoryBuffer, addr psfb
			.if (eax == S_OK)
				mov ecx,m_pItem
				invoke vf(psfb, IPSFactoryBuffer, CreateProxy),
					pUnknown, addr [ecx].CInterfaceItem.iid, addr pProxy, addr pClientProxy
				.if (eax == S_OK)
					.if (pProxy)
						invoke Create@CObjectItem, pProxy, NULL
						.if (eax)
							mov pObjectItem, eax
							invoke vf(pObjectItem, IObjectItem, ShowObjectDlg), m_hWnd
							invoke GetWindowRect, m_hWnd, addr rect
							add rect.left, 16
							add rect.top, 16
							invoke vf(pObjectItem, IObjectItem, GetObjectDlg)
							lea ecx, rect
							invoke SetPosition@CObjectDlg, eax, ecx
							invoke vf(pObjectItem, IObjectItem, Release)
						.endif
						invoke vf(pProxy, IUnknown, Release)
					.endif
					.if (pClientProxy)
						invoke vf(pClientProxy, IUnknown, Release)
					.endif
				.else
					invoke OutputMessage, m_hWnd, eax, CStr("IPSFactoryBuffer::CreateProxy"), 0
				.endif
				invoke vf(psfb, IUnknown, Release)
			.else
				invoke OutputMessage, m_hWnd, eax, CStr("CoGetClassObject"), 0
			.endif
		.endif
		invoke vf(pUnknown, IUnknown, Release)
done:
		ret
		align 4

OnCreateProxy endp
endif

;*** process WM_NOTIFY


CDRF_NEWFONT	equ 000000002h


OnNotify proc uses ebx pNMHDR:ptr NMLISTVIEW

local	hr:BOOL
local	iid:IID
local	szStr[64]:byte

		mov hr, FALSE
		mov ebx,pNMHDR
		assume ebx:ptr NMLISTVIEW

		xor ecx, ecx
		.if (([ebx].hdr.idFrom == IDC_LIST1) || ([ebx].hdr.idFrom == IDC_LIST2))
			inc ecx
		.endif

		.if (ecx && ([ebx].hdr.code == NM_DBLCLK))

			invoke GetMenuEnableVector
			invoke GetMenuDefaultCmd, eax
			invoke PostMessage, m_hWnd, WM_COMMAND, eax, 0

		.elseif (ecx && ([ebx].hdr.code == NM_RCLICK))

			invoke ShowContextMenu, TRUE

		.elseif (ecx && ([ebx].hdr.code == NM_SETFOCUS))

			invoke ListView_GetNextItem( [ebx].hdr.hwndFrom, -1, LVNI_SELECTED)
			invoke SetCurrentItem, ebx, eax

if ?BOLDCONN
		.elseif (ecx && [ebx].hdr.code == NM_CUSTOMDRAW)

			assume ebx:ptr NMLVCUSTOMDRAW

			.if (g_bGrayClrforIF && ([ebx].nmcd.dwDrawStage == CDDS_PREPAINT))
				invoke SetWindowLong, m_hWnd, DWL_MSGRESULT, CDRF_NOTIFYITEMDRAW
				mov hr, TRUE
			.elseif ([ebx].nmcd.dwDrawStage == CDDS_ITEMPREPAINT)

				mov ecx, [ebx].nmcd.lItemlParam
				.if (ecx && [ecx].CInterfaceItem.dwVerMajor == -1)
					mov [ebx].clrText, 0A7A7A7h
					invoke SetWindowLong, m_hWnd, DWL_MSGRESULT, CDRF_NEWFONT
					mov hr, TRUE
				.endif
				.if ([ebx].nmcd.hdr.idFrom == IDC_LIST2)
					invoke vf(m_pObjectItem, IObjectItem, GetConnectionList), FALSE
					.if (eax)
						invoke FindItemData@CList, eax, [ebx].NMLVCUSTOMDRAW.nmcd.dwItemSpec
						.if (eax != -1)
							mov [ebx].clrText, 0C000h
							invoke SetWindowLong, m_hWnd, DWL_MSGRESULT, CDRF_NEWFONT
							mov hr, TRUE
						.endif
					.endif
				.endif
			.endif

			assume ebx:ptr NMLISTVIEW
endif

		.elseif ([ebx].hdr.code == LVN_ITEMCHANGED)

			.if ([ebx].uNewState & LVIS_SELECTED)
				.if ([ebx].uChanged & LVIF_STATE)
					.if (!m_bRefresh)
						invoke StringFromCLSID@CObjectItem, m_pObjectItem, addr szStr
						StatusBar_SetText m_hWndSB, 0, addr szStr
					.endif
					invoke SetCurrentItem, ebx, [ebx].iItem
				.endif
			.endif

		.elseif ([ebx].hdr.code == LVN_COLUMNCLICK)

			.if ([ebx].hdr.idFrom == IDC_LIST1)
				mov eax,[ebx].iSubItem
				.if (eax == m_iSortCol)
					xor m_iSortDir,1
				.else
					mov m_iSortCol,eax
					mov m_iSortDir,0
				.endif	
				invoke LVSort, m_hWndLV, m_iSortCol, m_iSortDir, 0
			.else
				mov eax,[ebx].iSubItem
				.if (eax == m_iSortColOut)
					xor m_iSortDirOut,1
				.else
					mov m_iSortColOut,eax
					mov m_iSortDirOut,0
				.endif	
				invoke LVSort, m_hWndLVOut, m_iSortColOut, m_iSortDirOut, 0
			.endif

		.elseif ([ebx].hdr.code == LVN_KEYDOWN)

			assume ebx:ptr NMLVKEYDOWN
			.if ([ebx].wVKey == VK_APPS)
				invoke ShowContextMenu, FALSE
			.endif
			assume ebx:ptr NMLISTVIEW

		.elseif ([ebx].hdr.code == LVN_DELETEITEM)

			.if ([ebx].NMLISTVIEW.lParam)
				invoke Destroy@CInterfaceItem, [ebx].NMLISTVIEW.lParam
			.endif

		.elseif ([ebx].hdr.code == LVN_DELETEALLITEMS)

			invoke SetWindowLong, m_hWnd, DWL_MSGRESULT, 0
			mov hr, TRUE

		.endif 
		assume ebx:nothing
		return hr
		align 4

OnNotify endp

ifdef @StackBase
	option stackbase:ebp
endif
	option prologue:@sehprologue
	option epilogue:@sehepilogue

CallViewer proc uses esi ebx __this thisarg

	nop
	.try
		invoke m_ViewProc, m_hWnd, m_pUnknown, addr m_guid
	.exceptfilter
		mov __this, this@
		mov eax, _exception_info()
		invoke DisplayExceptionInfo, m_hWnd, eax, CStr("Internal Interface Viewer"), EXCEPTION_EXECUTE_HANDLER
	.except
	.endtry
	ret
	align 4

CallViewer endp

	option prologue: prologuedef
	option epilogue: epiloguedef
ifdef @StackBase
	option stackbase:esp
endif

;*** process WM_COMMAND

OnCommand proc uses ebx wParam:WPARAM, lParam:LPARAM

local	lvi:LVITEM
local	variant:VARIANT
local	pDispatch:LPDISPATCH
local	pTypeInfo:LPTYPEINFO
local	dwIndex:DWORD
local	pTID:ptr CTypeInfoDlg
local	pObjectDlg:ptr CObjectDlg
local	rect:RECT
local	szTypeLib[40]:byte
local	wszGUID[40]:word
local	szText[80]:byte

	movzx eax,word ptr wParam
	.if (eax == IDCANCEL)

		invoke PostMessage,m_hWnd,WM_CLOSE,0,0

	.elseif (eax == IDC_LOCK)

		invoke vf(m_pObjectItem, IObjectItem, Lock_)
		invoke EnableLockBtn

	.elseif (eax == IDOK)					;currently no ok button

		invoke GetMenuEnableVector
		invoke GetMenuDefaultCmd, eax
		invoke PostMessage, m_hWnd, WM_COMMAND, eax, 0

	.elseif (eax == IDM_EDIT)

		invoke OnEdit

	.elseif (eax == IDM_VIEWSTORAGEDLG)

		invoke vf(m_pObjectItem, IObjectItem, ShowViewStorageDlg), NULL

	.elseif (eax == IDM_COPYOBJECT)

		invoke VariantInit, addr variant
		mov ecx, m_pItem
		invoke vf(m_pUnknown, IUnknown, QueryInterface), addr [ecx].CInterfaceItem.iid, addr variant.punkVal
		.if (eax == S_OK)
			mov variant.vt, VT_UNKNOWN
			invoke Create@CDataObject, addr variant, sizeof VARIANT, g_dwMyCBFormat
			.if (eax)
				mov g_pDataObject, eax
				invoke OleSetClipboard, eax
				invoke vf(g_pDataObject, IUnknown, Release)
			.endif
			invoke VariantClear, addr variant
		.endif

	.elseif (eax == IDM_VIEWOBJECT)

		invoke OnView, m_pItem

	.elseif (eax == IDM_TYPELIBDLG)

		.if (m_dwListId == IDC_LIST1)
			mov ebx,m_pItem
		.else
			mov ebx,m_pItemOut
		.endif
		assume ebx:ptr CInterfaceItem
		.repeat
			.if ([ebx].CInterfaceItem.dwVerMajor != -1)
				mov ecx,-1
				invoke Create@CTypeLibDlg, addr [ebx].TypelibGUID, [ebx].dwVerMajor, [ebx].dwVerMinor, ecx, addr [ebx].iid
				.break
			.endif
			invoke IsEqualGUID, addr [ebx].iid, addr IID_IProvideClassInfo
			.if (eax)
				invoke GetTypeInfoFromIProvideClassInfo, m_pUnknown, FALSE
				invoke Create4@CTypeLibDlg, eax
				.break
			.endif
			invoke IsEqualGUID, addr [ebx].iid, addr IID_IDispatch
			.if (eax)
				invoke GetTypeInfoFromIDispatch, m_pUnknown
				invoke Create4@CTypeLibDlg, eax
				.break
			.endif
		.until (1)
		.if (eax)
			invoke Show@CTypeLibDlg, eax, m_hWnd, FALSE
		.else
			invoke MessageBox, m_hWnd, addr g_szNoTypeInfo, 0, MB_OK
		.endif
		assume ebx:nothing

	.elseif (eax == IDM_TYPEINFODLG)

		.repeat
			.if (m_dwListId == IDC_LIST1)
				mov ebx,m_pItem
			.else
				mov ebx,m_pItemOut
			.endif
			.if ([ebx].CInterfaceItem.dwVerMajor != -1)
				invoke Create3@CTypeInfoDlg, addr [ebx].CInterfaceItem.iid,
						addr [ebx].CInterfaceItem.TypelibGUID,
						[ebx].CInterfaceItem.dwVerMajor,
						[ebx].CInterfaceItem.dwVerMinor
				.if (eax)
					mov pTID, eax
					invoke Show@CTypeInfoDlg, pTID, m_hWnd, TYPEINFODLG_FACTIVATE
				.else
					invoke MessageBox, m_hWnd, addr g_szNoTypeInfo, 0, MB_OK
				.endif
				.break
			.endif
;------------------------------------- is it a IProvideClassInfo interface?
			invoke IsEqualGUID, addr [ebx].CInterfaceItem.iid, addr IID_IProvideClassInfo
			.if (eax)
				invoke TypeInfoDlgFromIProvideClassInfo, m_hWnd, m_pUnknown
				.break
			.endif
;---------------------------------- new 12/2002: create TypeInfo dialog from IDispatch
			invoke IsEqualGUID, addr [ebx].CInterfaceItem.iid, addr IID_IDispatch
			.if (eax)
				invoke TypeInfoDlgFromIDispatch, m_hWnd, m_pUnknown
				.break
			.endif
		.until (1)

	.elseif (eax == IDM_PROPERTIESDLG)

		invoke vf(m_pObjectItem, IObjectItem, GetPropDlg)
		mov ecx, eax
		mov eax, m_pItem
		.if (ecx)
			invoke RestoreAndActivateWindow, [ecx].CDlg.hWnd
		.elseif (eax == NULL)
		.else
;----------------------------------- is current interface IProvideClassInfo?
			invoke IsEqualGUID, addr [eax].CInterfaceItem.iid, addr IID_IProvideClassInfo
			.if (eax)
				invoke GetTypeInfoFromIProvideClassInfo, m_pUnknown, FALSE
				.if (eax)
					push eax
					invoke Create2@CPropertiesDlg, m_pUnknown, eax
					pop ecx
					push eax
					invoke vf(ecx, IUnknown, Release)
					pop eax
				.endif
			.else
				mov eax, m_pItem
				.if ([eax].CInterfaceItem.dwVerMajor == -1)
					xor ebx, ebx
					invoke vf(m_pObjectItem, IObjectItem, GetCoClassTypeInfo), addr pTypeInfo
					.if (pTypeInfo)
						invoke GetDefaultInterfaceFromCoClass, pTypeInfo, FALSE
						mov ebx, eax
						invoke vf(pTypeInfo, ITypeInfo, Release)
					.endif
					invoke Create2@CPropertiesDlg, m_pUnknown, ebx
					.if (ebx)
						push eax
						invoke vf(ebx, IUnknown, Release)
						pop eax
					.endif
				.else
					invoke Create@CPropertiesDlg, m_pObjectItem, m_pItem
				.endif
			.endif
			.if (eax)
				.if (g_bPropDlgAsTopLevelWnd)
					mov ecx,NULL
				.else
					mov ecx, m_hWnd
				.endif
				invoke Show@CPropertiesDlg, eax, ecx
			.else
				invoke MessageBox, m_hWnd, addr g_szNoTypeInfo, 0, MB_OK
			.endif
		.endif

	.elseif (eax == IDC_RUN)

		invoke IsRunning
		.if (!eax)
			invoke vf(m_pObjectItem, IObjectItem, SetRunLock), TRUE
			invoke RefreshList
		.endif

	.elseif (eax == IDC_VIEW)

		invoke OnView, NULL

	.elseif (eax == IDM_REFRESH)

		invoke RefreshList

	.elseif (eax == IDM_DISCONNECT)

		invoke OnDisconnect

	.elseif (eax == IDM_CONNECT)

		invoke OnConnect

	.elseif (eax == IDM_VIEWINTERFACE)

		invoke CallViewer, __this

	.elseif (eax == IDM_COPYGUID)

		mov ecx, m_pItem
		invoke StringFromGUID2, addr [ecx].CInterfaceItem.iid, addr wszGUID, 40
		invoke WideCharToMultiByte,CP_ACP,0,addr wszGUID,-1,addr szText, sizeof szText,0,0 
		invoke CopyStringToClipboard, m_hWnd, addr szText

	.elseif (eax == IDM_VIEWVTBL)

		invoke DialogBoxParam, g_hInstance, IDD_VTBLDLG, m_hWnd, vtbldlgproc, __this

	.elseif (eax == IDM_SECURITY)

		invoke OnSecurity
if ?CREATEPROXY
	.elseif (eax == IDM_CREATEPROXY)

		invoke OnCreateProxy
endif
	.else
		xor eax,eax
	.endif

	ret
	align 4

OnCommand endp


;*** process WM_INITDIALOG


OnInitDialog proc uses ebx

local	dwWidth[2]:DWORD
local	rect:RECT

if ?SHOWICON
	invoke SetWindowIcon, m_hWnd, addr m_guid
	mov m_hIcon, eax
	.if (!eax)
		mov eax, g_hIconObj
		.if (!eax)
			invoke LoadIcon,g_hInstance,IDI_OBJECT
			mov g_hIconObj, eax
		.endif
		.if (eax)
			invoke SendMessage, m_hWnd, WM_SETICON, ICON_SMALL, g_hIconObj
			invoke SendMessage, m_hWnd, WM_SETICON, ICON_BIG, g_hIconObj
		.endif
	.endif
endif
	invoke GetDlgItem,m_hWnd,IDC_LIST1
	mov m_hWndLV,eax
	invoke GetDlgItem,m_hWnd,IDC_LIST2
	mov m_hWndLVOut,eax
	invoke GetDlgItem,m_hWnd,IDC_STATUSBAR
	mov m_hWndSB,eax

	invoke GetClientRect, m_hWnd, addr rect		
	invoke MulDiv, rect.right, 4, 5
	mov dwWidth[0*sizeof DWORD], eax
	mov dwWidth[1*sizeof DWORD], -1
	StatusBar_SetParts m_hWndSB, 2, addr dwWidth

	invoke ListView_SetExtendedListViewStyle( m_hWndLV,LVS_EX_FULLROWSELECT or LVS_EX_HEADERDRAGDROP)
	invoke ListView_SetExtendedListViewStyle( m_hWndLVOut,LVS_EX_FULLROWSELECT)

;-------------------------------------- set dialog title

	invoke vf(m_pObjectItem, IObjectItem, SetWindowText_), m_hWnd

	invoke SetLVColumns, m_hWndLV, NUMCOLS, offset pColumns
	invoke SetLVColumns, m_hWndLVOut, NUMCOLSOUT, offset pColumnsOut

	invoke EnableLockBtn

	.if (m_pt.x)
		invoke SetWindowPos, m_hWnd, NULL, m_pt.x, m_pt.y, 0, 0, SWP_NOSIZE or SWP_NOZORDER or SWP_NOACTIVATE
	.endif

	ret
	align 4

OnInitDialog endp


;*** Dialog Proc for "object" dialog


ObjectDialog proc uses __this thisarg, message:dword, wParam:WPARAM, lParam:LPARAM

	mov __this,this@

	mov eax,message
	.if (eax == WM_INITDIALOG)

		invoke CenterWindow, m_hWnd
		invoke OnInitDialog
		invoke vf(m_pObjectItem, IObjectItem, GetFlags)
		.if (eax & OBJITEMF_IGNOREOV)
			and eax, not OBJITEMF_IGNOREOV
			invoke vf(m_pObjectItem, IObjectItem, SetFlags), eax
		.elseif (eax & OBJITEMF_OPENVIEW)
			invoke OnView, NULL
			.if (eax)
				invoke PostMessage, m_hWnd, WM_CLOSE, 0, 0
				mov eax, 1
				jmp exit
			.endif
		.endif
		invoke RefreshList
		invoke ShowWindow, m_hWnd, SW_SHOWNORMAL
		mov eax,1

	.elseif (eax == WM_CLOSE)

if ?MODELESS
		invoke DestroyWindow, m_hWnd
else
		invoke EndDialog, m_hWnd, 0
endif
		mov eax,1

	.elseif (eax == WM_DESTROY)

		invoke ListView_DeleteAllItems( m_hWndLV)
		invoke ListView_DeleteAllItems( m_hWndLVOut)

		invoke Destroy@CObjectDlg, __this

if ?MODELESS
	.elseif (eax == WM_ACTIVATE)

		movzx eax,word ptr wParam
		.if (eax == WA_INACTIVE)
			mov g_hWndDlg, NULL
		.else
			mov eax, m_hWnd
			mov g_hWndDlg, eax
			invoke IsRunning
			invoke IsStorageActive
			invoke EnableLockBtn
		.endif
endif
	.elseif (eax == WM_COMMAND)
		invoke OnCommand, wParam, lParam

	.elseif (eax == WM_NOTIFY)
		invoke OnNotify, lParam

	.elseif (eax == WM_ENTERMENULOOP)

		StatusBar_SetSimpleMode m_hWndSB, TRUE
;		invoke OnEnterMenuLoop, wParam

	.elseif (eax == WM_EXITMENULOOP)

		StatusBar_SetSimpleMode m_hWndSB, FALSE
;		invoke OnExitMenuLoop

	.elseif (eax == WM_MENUSELECT)

		movzx ecx, word ptr wParam+0
		invoke DisplayStatusBarString, m_hWndSB, ecx
if ?HTMLHELP
	.elseif (eax == WM_HELP)

		invoke DoHtmlHelp, HH_DISPLAY_TOPIC, CStr("ObjectDialog.htm")
endif
	.else
		xor eax,eax ;indicates "no processing"
	.endif
exit:
	ret
	align 4

ObjectDialog endp



Create@CObjectDlg proc public uses __this, pObjectItem:ptr CObjectItem

local pPersist:LPPERSIST

	invoke malloc, sizeof CObjectDlg
	.if (!eax)
		ret
	.endif
	mov __this, eax

	mov m_iSortCol,-1
	mov m_iSortColOut,-1
	mov m_pDlgProc,ObjectDialog
	invoke GetUnknown@CObjectItem, pObjectItem
	mov m_pUnknown,eax
;;	invoke vf(m_pUnknown, IUnknown, AddRef)
	mov m_pItem, NULL
	mov m_pItemOut, NULL
	mov eax, pObjectItem
	mov m_pObjectItem, eax
	invoke vf(pObjectItem, IObjectItem, AddRef)

	invoke GetGUID@CObjectItem, pObjectItem, addr m_guid
	return __this
	align 4

Create@CObjectDlg endp


Show@CObjectDlg proc public thisarg, hWnd:HWND

	.if (g_bObjDlgsAsTopLevelWnd)
		mov ecx, NULL
	.else
		mov ecx, hWnd
	.endif
if ?MODELESS
	invoke CreateDialogParam, g_hInstance, IDD_OBJECTDLG, ecx, classdialogproc, this@
else
	invoke DialogBoxParam, g_hInstance, IDD_OBJECTDLG, ecx, classdialogproc, this@
endif
	ret
	align 4

Show@CObjectDlg endp

SetPosition@CObjectDlg proc public thisarg, pPt:ptr POINT
	mov eax, pPt
	mov ecx, this@
	mov edx, [eax].POINT.x
	mov [ecx].CObjectDlg.pt.x, edx
	mov edx, [eax].POINT.y
	mov [ecx].CObjectDlg.pt.y, edx
	ret
	align 4

SetPosition@CObjectDlg endp

Destroy@CObjectDlg proc public uses __this thisarg

local	pOleObject:LPOLEOBJECT

	mov __this,this@

;;	.if (m_pItem)
;;		invoke Destroy@CInterfaceItem, m_pItem
;;	.endif
;;	.if (m_pItemOut)
;;		invoke Destroy@CInterfaceItem, m_pItemOut
;;	.endif

;;	.if (m_pUnknown)
;;		invoke vf(m_pUnknown, IUnknown, Release)
;;		DebugOut "ObjectDlg, IUnknown::Release returned %X", eax
;;	.endif

	invoke vf(m_pObjectItem, IObjectItem, SetObjectDlg), NULL
	invoke vf(m_pObjectItem, IObjectItem, Release)

if ?SHOWICON
	.if (m_hIcon)
		invoke DestroyIcon, m_hIcon
	.endif
endif
	invoke free, __this
	ret
	align 4

Destroy@CObjectDlg endp

	end
