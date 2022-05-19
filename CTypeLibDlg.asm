
;*** definition of CTypeLibDlg methods

	.386
	.model flat,stdcall
	option casemap :none   ; case sensitive
	option proc:private

	include COMView.inc
	include statusbar.inc
INSIDE_CTYPELIBDLG equ 1
	include classes.inc
	include debugout.inc
	include rsrc.inc
	include CEditDlg.inc

?MODELESS		equ 1
?EDITMODELESS	equ 1
?DRAWITEMSB		equ 1
?TYPELIBICON	equ 1		;own icon for typelib dialogs
?CUSTOMDRAW		equ 1

if ?DRAWITEMSB
?SBOWNERDRAW	equ		SBT_OWNERDRAW
	.data?
g_szErrorText	db MAX_PATH+32 dup (?)	;ownerdrawn text in statusbar must be global
else
?SBOWNERDRAW	equ		0
endif

protoTypeLibCB typedef proto :DWORD, :LPTYPELIB, :DWORD
LPTYPELIBCB typedef ptr protoTypeLibCB

BEGIN_CLASS CTypeLibDlg, CDlg
hWndLV		HWND ?
hWndSB		HWND ?
hWndParent	HWND ?
pszTypeLib	LPSTR ?
guid		GUID <>
dwVerMajor	dword ?
dwVerMinor	dword ?
lcid		LCID ?
pTypeLib	LPTYPELIB ?
pIID		dword ?
dwIndex		dword ?
iSortCol	dword ?
iSortDir	dword ?
fnCallBack	LPTYPELIBCB ?
dwCookie	dword ?
dwUserCol	dword ?
bTile		BOOLEAN ?
END_CLASS

	.const

	.data

;----------- column header text

ColumnTable label CColHdr
	CColHdr <CStr("Name"),			20>
	CColHdr <CStr("GUID"),			15>
	CColHdr <CStr("TypeKind"),		15>
	CColHdr <CStr("Flags"),			12>
	CColHdr <CStr("Functions"),		10,	FCOLHDR_RDX10>
	CColHdr <CStr("Variables"),		10,	FCOLHDR_RDX10>
	CColHdr <CStr("Interfaces"),	9,	FCOLHDR_RDX10>
NUMCOLIDX equ ($ - ColumnTable) / sizeof CColHdr
	CColHdr <CStr("Idx"),			9,	FCOLHDR_RDX10>
NUMCOLS textequ %($ - ColumnTable) / sizeof CColHdr

GUIDCOL_IN_TYPELIB equ 1

g_rect			RECT <0,0,0,0>
if ?TYPELIBICON
g_hIconTypeLib	HICON NULL
endif

	.code

__this	textequ <edi>
_this	textequ <[__this].CTypeLibDlg>
thisarg	textequ <this@:ptr CTypeLibDlg>

;*** private members prototypes
;*** thisarg could be omitted cause "this" ptr (=edi) is always set

Destroy@CTypeLibDlg	proto thisarg
GetTypeInfoIndex	proto :DWORD
ShowContextMenu		proto :BOOL
OnCreate			proto :DWORD
OnNotify			proto pNMHDR:ptr NMHDR
OnInitDialog		proto
;OnRegister			proto
;OnUnregister		proto
RefreshList			proto

	MEMBER hWnd, pDlgProc
	MEMBER hWndLV, hWndSB, hWndParent, pszTypeLib
	MEMBER guid, dwVerMajor, dwVerMinor, lcid, pTypeLib, pIID, dwIndex
	MEMBER iSortCol, iSortDir, fnCallBack, dwCookie
	MEMBER dwUserCol, bTile

GetTypeInfoIndex proc iItem:DWORD

local	szStr[32]:byte
local	dwTmp:dword

		mov eax, iItem
		.if (eax == -1)
			invoke ListView_GetNextItem( m_hWndLV, eax, LVNI_SELECTED)
		.endif
		.if (eax != -1)
			lea ecx,szStr
			ListView_GetItemText m_hWndLV, eax, NUMCOLIDX, ecx, sizeof szStr
			invoke String2Number,addr szStr,addr dwTmp,10
			mov eax,dwTmp
		.endif
		ret
		align 4
GetTypeInfoIndex endp

;--- does a certain CLSID/IID exist in Classes/Interface

IsRegistryEntryExisting proc pTypeAttr:ptr TYPEATTR

local bReturn:BOOL
local hKey:HANDLE
local szGUID[64]:byte
local szKey[128]:byte

		mov bReturn, FALSE
		invoke ListView_GetNextItem( m_hWndLV,-1,LVNI_SELECTED)
		.if (eax != -1)
			lea ecx, szGUID
			ListView_GetItemText m_hWndLV, eax, GUIDCOL_IN_TYPELIB, ecx, sizeof szGUID
			mov ecx, pTypeAttr
			.if ([ecx].TYPEATTR.typekind == TKIND_COCLASS)
				mov ecx, offset  g_szRootCLSID
			.else
				mov ecx, offset  g_szRootInterface
			.endif
			invoke wsprintf, addr szKey, CStr("%s\%s"), ecx, addr szGUID
			invoke RegOpenKeyEx, HKEY_CLASSES_ROOT, addr szKey, 0, KEY_READ, addr hKey
			.if (eax == ERROR_SUCCESS)
				invoke RegCloseKey, hKey
				mov bReturn, TRUE
			.endif
		.endif
		mov eax, bReturn
		ret
		align 4

IsRegistryEntryExisting endp

;--- does a certain CLSID/IID exist in Classes/Interface

IsRegistryEntryExisting2 proc pTypeAttr:ptr TYPEATTR

local bReturn:BOOL
local hKey:HANDLE
local szGUID[40]:word
local szKey[128]:byte

		mov bReturn, FALSE
		mov ecx, pTypeAttr
		invoke StringFromGUID2, addr [ecx].TYPEATTR.guid, addr szGUID, 40

		mov ecx, pTypeAttr
		.if ([ecx].TYPEATTR.typekind == TKIND_COCLASS)
			mov ecx, offset  g_szRootCLSID
		.else
			mov ecx, offset  g_szRootInterface
		.endif
		invoke wsprintf, addr szKey, CStr("%s\%S"), ecx, addr szGUID
		invoke RegOpenKeyEx, HKEY_CLASSES_ROOT, addr szKey, 0, KEY_READ, addr hKey
		.if (eax == ERROR_SUCCESS)
			invoke RegCloseKey, hKey
			mov bReturn, TRUE
		.endif
		mov eax, bReturn
		ret
		align 4

IsRegistryEntryExisting2 endp

;*** user pressed right mouse button, show context menu ***

ShowContextMenu proc uses esi bMouse:BOOL

local	pt:POINT
local	hPopupMenu:HMENU
local	dwEditFlags:DWORD
local	dwCopyFlags:DWORD
local	dwHelpFlags:DWORD
local	dwGetClassFlags:DWORD
local	dwIndex:DWORD
local	bstr:BSTR
local	dwCount:DWORD
local	dwContext:DWORD
local	pTypeInfo:LPTYPEINFO
local	pTypeAttr:ptr TYPEATTR

		invoke ListView_GetSelectedCount( m_hWndLV)
		.if (!eax)
			ret
		.endif
		mov dwCount, eax
		invoke GetSubMenu, g_hMenu, ID_SUBMENU_TYPELIBDLG
		.if (!eax)
			jmp done
		.endif
		mov hPopupMenu, eax

;------------ check if selected item describes a coclass

		mov esi, MF_DISABLED or MF_GRAYED or MF_BYCOMMAND
		mov dwGetClassFlags, MF_DISABLED or MF_GRAYED or MF_BYCOMMAND
		mov dwEditFlags, MF_DISABLED or MF_GRAYED or MF_BYCOMMAND
		mov dwCopyFlags, MF_DISABLED or MF_GRAYED or MF_BYCOMMAND
		mov dwHelpFlags, MF_DISABLED or MF_GRAYED or MF_BYCOMMAND

		invoke GetTypeInfoIndex, -1
		mov dwIndex,eax
		invoke vf(m_pTypeLib,ITypeLib,GetTypeInfo),dwIndex,addr pTypeInfo
		.if (eax == S_OK)
			invoke vf(pTypeInfo,ITypeInfo,GetTypeAttr),addr pTypeAttr
			.if (eax == S_OK)
;-------------------------------- CreateInstance enabled only if 1 item selected
				.if (dwCount == 1)
					mov eax, pTypeAttr
					movzx eax,[eax].TYPEATTR.wTypeFlags
					.if (eax & TYPEFLAG_FCANCREATE)
						mov esi, MF_ENABLED or MF_BYCOMMAND
						.if (m_pszTypeLib)
							mov dwGetClassFlags, MF_ENABLED or MF_BYCOMMAND
						.endif
					.endif
				.endif
				mov eax, pTypeAttr
				invoke IsEqualGUID, addr [eax].TYPEATTR.guid, addr IID_NULL
				.if (!eax)
;-------------------------------- Copy GUID enabled only if 1 item selected
					.if (dwCount == 1)
						mov dwCopyFlags, MF_ENABLED or MF_BYCOMMAND
					.endif
					invoke IsRegistryEntryExisting, pTypeAttr
					.if (eax)
						mov dwEditFlags, MF_ENABLED or MF_BYCOMMAND
					.endif
				.endif
				invoke vf(pTypeInfo, ITypeInfo, ReleaseTypeAttr), pTypeAttr
			.endif
;-------------------------------- Help enabled only if 1 item selected
			.if (dwCount == 1)
				invoke vf(pTypeInfo, ITypeInfo, GetDocumentation), MEMBERID_NIL, NULL, NULL, addr dwContext, addr bstr
				.if ((eax == S_OK) && bstr)
					invoke SysFreeString, bstr
					.if (dwContext)
						mov dwHelpFlags, MF_ENABLED or MF_BYCOMMAND
					.endif
				.endif
			.endif
			invoke vf(pTypeInfo, ITypeInfo, Release)
		.endif
		invoke EnableMenuItem, hPopupMenu, IDM_CREATE, esi
		invoke EnableMenuItem, hPopupMenu, IDM_GETCLASS, dwGetClassFlags
		invoke EnableMenuItem, hPopupMenu, IDM_EDIT, dwEditFlags
		invoke EnableMenuItem, hPopupMenu, IDM_COPYGUID, dwCopyFlags
		invoke EnableMenuItem, hPopupMenu, IDM_CONTEXTHELP, dwHelpFlags

		invoke SetMenuDefaultItem, hPopupMenu, IDM_VIEW, FALSE
		invoke GetItemPosition, m_hWndLV, bMouse, addr pt
		invoke TrackPopupMenu, hPopupMenu,TPM_LEFTALIGN or TPM_LEFTBUTTON,
				pt.x, pt.y, 0, m_hWnd, NULL
done:
		ret
		align 4

ShowContextMenu endp


;*** create an object with ITypeInfo::CreateInstance


OnCreate proc dwIndex:DWORD

local	pTypeInfo:LPTYPEINFO
local	pTypeAttr:ptr TYPEATTR
local	pUnknown:LPUNKNOWN
local	pDispatch:LPDISPATCH
local	pObjectItem:LPOBJECTITEM
local	hr:DWORD

		invoke vf(m_pTypeLib, ITypeLib, GetTypeInfo), dwIndex, addr pTypeInfo
		.if (eax == S_OK)
			invoke SetBusyState@CMainDlg, TRUE
			invoke vf(pTypeInfo, ITypeInfo, CreateInstance), NULL, addr IID_IUnknown, addr pUnknown
			mov hr, eax
			invoke SetBusyState@CMainDlg, FALSE
			.if (hr == S_OK)
				invoke vf(pTypeInfo, ITypeInfo, GetTypeAttr),addr pTypeAttr
				.if (eax == S_OK)
					mov eax, pTypeAttr
					invoke Create@CObjectItem, pUnknown, addr [eax].TYPEATTR.guid
					.if (eax)
						mov pObjectItem, eax
						invoke vf(eax, IObjectItem, SetCoClassTypeInfo), pTypeInfo
						invoke vf(pObjectItem, IObjectItem, ShowObjectDlg), m_hWnd
						invoke vf(pObjectItem, IObjectItem, Release)
					.endif
					invoke vf(pTypeInfo, ITypeInfo, ReleaseTypeAttr),pTypeAttr
				.endif
				invoke vf(pUnknown, IUnknown, Release)
			.else
				invoke OutputMessage, m_hWnd, hr, CStr("ITypeInfo::CreateInstance"),0
			.endif
			invoke vf(pTypeInfo, ITypeInfo, Release)
		.endif
		ret
		align 4

OnCreate endp


;*** load library by CoLoadLibrary, then call DllGetClassObject

protoGetClassObject typedef proto :REFGUID, :REFIID, :ptr LPUNKNOWN
;;LPFNGETCLASSOBJECT typedef ptr protoGetClassObject

OnGetClass proc dwIndex:DWORD

local	hModule:HINSTANCE
local	pfnGetClassObject:LPFNGETCLASSOBJECT
local	pTypeInfo:LPTYPEINFO
local	pTypeAttr:ptr TYPEATTR
local	pUnknown:LPUNKNOWN
local	pClassFactory:LPCLASSFACTORY
local	szText[128]:byte
local	wszTypeLib[MAX_PATH]:word

		invoke MultiByteToWideChar,CP_ACP,MB_PRECOMPOSED,\
				m_pszTypeLib, -1, addr wszTypeLib, LENGTHOF wszTypeLib
		invoke CoLoadLibrary, addr wszTypeLib, TRUE
		.if (eax)
			mov hModule, eax
			invoke GetProcAddress, eax, CStr("DllGetClassObject")
			.if (eax)
				mov pfnGetClassObject, eax
				invoke vf(m_pTypeLib, ITypeLib, GetTypeInfo), dwIndex, addr pTypeInfo
				.if (eax == S_OK)
					invoke vf(pTypeInfo, ITypeInfo, GetTypeAttr), addr pTypeAttr
					.if (eax == S_OK)
						mov ecx, pTypeAttr
						invoke pfnGetClassObject, addr [ecx].TYPEATTR.guid, addr IID_IUnknown, addr pUnknown
						.if (eax == S_OK)
							invoke vf(pUnknown, IUnknown, QueryInterface), addr IID_IClassFactory, addr pClassFactory
							.if (eax == S_OK)
								invoke vf(pUnknown, IUnknown, Release)
								invoke vf(pClassFactory, IClassFactory, CreateInstance), NULL, addr IID_IUnknown, addr pUnknown
								.if (eax == S_OK)
									invoke Create@CObjectItem, pUnknown, NULL
									.if (eax)
										push eax
										invoke vf(eax, IObjectItem, ShowObjectDlg), m_hWnd
										pop eax
										invoke vf(eax, IObjectItem, Release)
									.endif
									invoke vf(pUnknown, IUnknown, Release)
								.else
									invoke OutputMessage, m_hWnd, eax, CStr("IClassFactory::CreateInstance"),0
								.endif
							.else
								invoke OutputMessage, m_hWnd, eax, CStr("QueryInterface(IClassFactory)"),0
							.endif
						.else
							invoke OutputMessage, m_hWnd, eax, CStr("DllGetClassObject"),0
						.endif
						invoke vf(pTypeInfo, ITypeInfo, ReleaseTypeAttr), pTypeAttr
					.endif
					invoke vf(pTypeInfo, ITypeInfo, Release)
				.endif
			.else
				invoke MessageBox, m_hWnd, CStr("Entry DllGetClassObject not found"), 0, MB_OK
			.endif
		.else
			invoke MessageBox, m_hWnd, m_pszTypeLib, CStr("CoLoadLibrary failed"), MB_OK
		.endif
		ret
		align 4

OnGetClass endp


;*** WM_NOTIFY/NM_RCLICK for header control ***


OnHeaderRClick proc uses ebx pNMHDR:ptr NMHDR

local	pt:POINT
local	mii:MENUITEMINFO

		invoke CreatePopupMenu
		mov ebx, eax
		invoke MakeUDColumnList, ebx, MODE_TYPEINFO, IDS_GETMOPS
		invoke GetCursorPos, addr pt
		.if (m_dwUserCol)
			invoke CheckMenuItem, ebx, m_dwUserCol, MF_CHECKED 
		.endif
		invoke TrackPopupMenu, ebx, TPM_LEFTALIGN or TPM_LEFTBUTTON or TPM_RETURNCMD,\
				pt.x, pt.y, 0, m_hWnd, NULL
		.if (eax)
			.if (eax == m_dwUserCol)
				mov m_dwUserCol, 0
			.else
				mov m_dwUserCol, eax
			.endif
			invoke RefreshList
		.endif
		invoke DestroyMenu, ebx
		ret
		align 4

OnHeaderRClick endp

OnNotify proc uses ebx pNMHDR:ptr NMHDR

local	lvc:LVCOLUMN

		mov ebx,pNMHDR

		.if ([ebx].NMHDR.idFrom == IDC_LIST1)

			assume ebx:ptr NMLISTVIEW

			.if ([ebx].hdr.code == NM_DBLCLK)

				invoke PostMessage, m_hWnd, WM_COMMAND, IDM_VIEW, 0

			.elseif ([ebx].hdr.code == NM_RCLICK)

				invoke ShowContextMenu, TRUE

if ?CUSTOMDRAW
			.elseif ([ebx].hdr.code == NM_CUSTOMDRAW)

				assume ebx:ptr NMLVCUSTOMDRAW

				xor eax, eax
				.if ([ebx].nmcd.dwDrawStage == CDDS_PREPAINT)
					invoke SetWindowLong, m_hWnd, DWL_MSGRESULT, CDRF_NOTIFYITEMDRAW
					mov eax, TRUE
				.elseif ([ebx].nmcd.dwDrawStage == CDDS_ITEMPREPAINT)
					mov ecx, [ebx].nmcd.lItemlParam
					.if (!ecx)
						mov [ebx].clrText, 0A7A7A7h
						invoke SetWindowLong, m_hWnd, DWL_MSGRESULT, CDRF_NEWFONT
						mov eax, TRUE
					.endif
				.endif

				assume ebx:ptr NMLISTVIEW
endif

			.elseif ([ebx].hdr.code == LVN_KEYDOWN)

				assume ebx:ptr NMLVKEYDOWN

				invoke GetKeyState,VK_CONTROL
				and al,80h
				.if (!ZERO?)				;Ctrl pressed?
					.if ([ebx].wVKey == 'C')
						invoke PostMessage, m_hWnd, WM_COMMAND, IDM_COPY, 0
					.endif
				.elseif ([ebx].wVKey == VK_APPS)
					invoke ShowContextMenu, FALSE
				.elseif ([ebx].wVKey == VK_F6)
					invoke Create@CObjectItem, m_pTypeLib, NULL
					.if (eax)
						push eax
						invoke vf(eax, IObjectItem, ShowObjectDlg), m_hWnd
						pop eax
						invoke vf(eax, IObjectItem, Release)
					.endif
				.endif
				assume ebx:ptr NMLISTVIEW

			.elseif ([ebx].hdr.code == LVN_COLUMNCLICK)

				mov eax,[ebx].iSubItem
				.if (eax == m_iSortCol)
					xor m_iSortDir,1
				.else
					mov m_iSortCol,eax
					@mov m_iSortDir,0
				.endif

				mov eax,m_iSortCol
if 1
				mov lvc.mask_, LVCF_FMT
				lea ecx, lvc
				invoke ListView_GetColumn( m_hWndLV, eax, ecx)
				.if (lvc.fmt & LVCFMT_RIGHT)
else
				.if ([eax * sizeof CColHdr+ColumnTable].CColHdr.wFlags & FCOLHDR_RDXMASK)
endif
					@mov ecx, 1
				.else
					@mov ecx, 0
				.endif
				invoke LVSort, m_hWndLV, m_iSortCol, m_iSortDir, ecx

			.elseif ([ebx].hdr.code == LVN_ITEMCHANGED)

				.if (m_pszTypeLib)
					StatusBar_SetText m_hWndSB, 0, m_pszTypeLib
					StatusBar_SetTipText m_hWndSB, 0, m_pszTypeLib
				.else
					StatusBar_SetText m_hWndSB, 0, addr g_szNull
				.endif

			.endif

		.else
			assume ebx:ptr NMHDR
			.if ([ebx].code == NM_RCLICK)
				invoke ListView_GetHeader( m_hWndLV)
				.if (eax == [ebx].hwndFrom)
					invoke OnHeaderRClick, ebx
				.endif
			.endif
		.endif
		ret
		align 4
		assume ebx:nothing

OnNotify endp

RefreshList proc uses ebx esi

local	hr:DWORD
local	hCsrOld:HCURSOR
local	lvi:LVITEM
local	lvc:LVCOLUMN
local	lcid:LCID
local	rect:RECT
;;local	guid:GUID
local	dwCount:dword
local	pTypeInfo:LPTYPEINFO
local	pTypeAttr:ptr TYPEATTR
local	pTLibAttr:ptr TLIBATTR
local	dwSize:dword
local	iType:dword
local	hKey:HANDLE
local	bstr:BSTR
local	bstr2:BSTR
local	szGUID[40]:byte
local	wszGUID[40]:word
local	szName[80]:byte
local	szDoc[128]:byte
local	szStr[MAX_PATH]:byte
local	wszTypeLib[MAX_PATH]:word

		invoke SetCursor, g_hCsrWait
		mov hCsrOld,eax
		invoke SetWindowRedraw( m_hWndLV, FALSE)

		invoke ListView_DeleteAllItems( m_hWndLV)

		.if (m_dwUserCol)
externdef UserColumnTypeInfo:ptr LPSTR
externdef UserFormatTypeInfo:ptr LPSTR
			mov lvc.mask_,LVCF_TEXT or LVCF_FMT
			mov edx, m_dwUserCol
			sub edx, IDS_GETMOPS
			mov ecx, [edx*4 + offset UserColumnTypeInfo]
			mov eax, [edx*4 + offset UserFormatTypeInfo]
			mov lvc.pszText, ecx
			mov lvc.fmt, eax
			invoke ListView_SetColumn( m_hWndLV, NUMCOLS, addr lvc)
			.if (!eax)
				invoke GetClientRect, m_hWndLV, addr rect
				mov eax, rect.right
				shr eax, 3
				mov lvc.cx_,eax
				mov lvc.mask_,LVCF_TEXT or LVCF_WIDTH or LVCF_FMT
				invoke ListView_InsertColumn( m_hWndLV, NUMCOLS, addr lvc)
			.endif
		.else
			invoke ListView_DeleteColumn( m_hWndLV, NUMCOLS)
		.endif

		.if (m_pTypeLib)
			mov eax, S_OK
		.elseif (m_pszTypeLib == NULL)
			mov eax,m_lcid
			.if (eax == -1)
;;				invoke GetUserDefaultLCID
				mov eax, g_LCID
			.endif
			mov lcid,eax
			invoke LoadRegTypeLib, addr m_guid, m_dwVerMajor, m_dwVerMinor, lcid, addr m_pTypeLib
			.if ((eax != S_OK) && (m_lcid == -1))
				mov lcid,0
				invoke LoadRegTypeLib, addr m_guid, m_dwVerMajor, m_dwVerMinor, lcid, addr m_pTypeLib
			.endif
		.else
			invoke MultiByteToWideChar,CP_ACP,MB_PRECOMPOSED,\
				m_pszTypeLib, -1, addr wszTypeLib, MAX_PATH
			invoke LoadTypeLibEx, addr wszTypeLib, REGKIND_NONE, addr m_pTypeLib
		.endif

		.if (eax != S_OK)
			push eax
			.if (m_pszTypeLib == NULL)
				invoke StringFromGUID2, addr m_guid, addr wszGUID, 40
				invoke wsprintf,addr szStr,CStr("LoadRegTypeLib(%S,%u.%u,0x%X) failed",0ah),
					addr wszGUID, m_dwVerMajor, m_dwVerMinor, lcid
			.else
				invoke wsprintf,addr szStr,CStr("LoadTypeLibEx(%s) failed",0ah),
					m_pszTypeLib
			.endif
			pop ecx
			invoke OutputMessage, m_hWnd,ecx,CStr("COMView"),addr szStr
			mov hr, 0
		.else
;---------------------------------------------------- get name of TypeLib
			mov szName,0
			mov szDoc,0
			invoke vf(m_pTypeLib,ITypeLib,GetDocumentation), MEMBERID_NIL, addr bstr, addr bstr2, NULL, NULL
			.if (eax == S_OK)
				.if (bstr)
					invoke WideCharToMultiByte,CP_ACP,0, bstr, -1, addr szName, sizeof szName,0,0 
					invoke SysFreeString,bstr
				.endif
				.if (bstr2)
					invoke WideCharToMultiByte,CP_ACP,0, bstr2, -1, addr szDoc, sizeof szDoc,0,0 
					invoke SysFreeString,bstr2
				.endif
			.endif

			invoke vf(m_pTypeLib, ITypeLib, GetTypeInfoCount)
			mov dwCount,eax
			invoke vf(m_pTypeLib, ITypeLib, GetLibAttr), addr pTLibAttr
			.if (eax == S_OK)
				mov ebx,pTLibAttr
				movzx eax,[ebx].TLIBATTR.wMajorVerNum
				mov m_dwVerMajor, eax
				movzx ecx,[ebx].TLIBATTR.wMinorVerNum
				mov m_dwVerMinor, ecx
				mov edx, [ebx].TLIBATTR.lcid
				mov m_lcid, edx
				pushad
				lea esi, [ebx].TLIBATTR.guid
				lea edi, m_guid
				movsd
				movsd
				movsd
				movsd
				popad
				invoke vf(m_pTypeLib,ITypeLib,ReleaseTLibAttr),pTLibAttr
				invoke StringFromGUID2, addr m_guid, addr wszGUID, 40
			.endif

			.if (szName)
				invoke wsprintf, addr szStr, CStr("TypeLib %s [%s] "), addr szName, addr szDoc
			.else
				invoke wsprintf, addr szStr, CStr("TypeLib %S"), addr wszGUID
			.endif
			invoke SetWindowText, m_hWnd, addr szStr

			.if (m_pszTypeLib)
				StatusBar_SetText m_hWndSB, 0, m_pszTypeLib
			.endif

			invoke wsprintf, addr szStr, CStr("TypeLib\%S\%u.%u"), addr wszGUID, m_dwVerMajor, m_dwVerMinor
			invoke RegOpenKeyEx, HKEY_CLASSES_ROOT, addr szStr, 0, KEY_READ, addr hKey
			.if (eax != ERROR_SUCCESS)
				invoke GetDlgItem, m_hWnd, IDC_REGISTER
				invoke EnableWindow, eax, TRUE
				invoke GetDlgItem, m_hWnd, IDC_UNREGISTER
				invoke EnableWindow, eax, FALSE
			.else
				invoke RegCloseKey, hKey
			.endif

			mov ebx,0
			mov lvi.iItem,ebx
			.while (ebx < dwCount)
				invoke vf(m_pTypeLib, ITypeLib, GetTypeInfo), ebx, addr pTypeInfo
				.if (eax == S_OK)
					invoke vf(pTypeInfo, ITypeInfo, GetTypeAttr),addr pTypeAttr
					.if (eax == S_OK)
						mov esi,pTypeAttr
						assume esi:ptr TYPEATTR

						mov lvi.mask_,LVIF_TEXT or LVIF_PARAM
						lea eax,szStr
						mov lvi.pszText,eax

;--- flag the COCLASS + CanCreate entries if a CLSID entry exists or not

if ?CUSTOMDRAW
						mov eax, 1
						.if (([esi].typekind == TKIND_COCLASS) && ([esi].wTypeFlags & TYPEFLAG_FCANCREATE))
			 				invoke IsRegistryEntryExisting2, esi
						.endif
endif
						mov lvi.lParam, eax

						mov bstr,NULL
						invoke vf(pTypeInfo, ITypeInfo, GetDocumentation),MEMBERID_NIL,addr bstr,NULL,NULL,NULL
						.if (eax == S_OK)
							invoke WideCharToMultiByte, CP_ACP, 0, bstr, -1, addr szStr,\
								sizeof szStr, 0, 0 
							invoke SysFreeString,bstr
						.else
							mov szStr,0
						.endif
						mov lvi.iSubItem,0
						invoke ListView_InsertItem( m_hWndLV,addr lvi)

						mov lvi.mask_,LVIF_TEXT

						invoke StringFromGUID2,addr [esi].guid,addr wszGUID,40
						invoke WideCharToMultiByte,CP_ACP,0,addr wszGUID,40,addr szStr, sizeof szStr,0,0 
						inc lvi.iSubItem
						invoke ListView_SetItem( m_hWndLV,addr lvi)

						invoke GetTypekindStr,[esi].typekind
						invoke wsprintf,addr szStr,CStr("%s (%u)"),eax,[esi].typekind
						inc lvi.iSubItem
						invoke ListView_SetItem( m_hWndLV,addr lvi)

						movzx ecx,[esi].wTypeFlags
						.if (g_bTypeFlagsAsNumber)
							invoke wsprintf, addr szStr,CStr("%X"), ecx
						.else
							invoke GetTypeFlags, ecx, addr szStr
						.endif
						inc lvi.iSubItem
						invoke ListView_SetItem( m_hWndLV,addr lvi)

						movzx eax,[esi].cFuncs
						invoke wsprintf,addr szStr,CStr("%u"),eax
						inc lvi.iSubItem
						invoke ListView_SetItem( m_hWndLV,addr lvi)

						movzx eax,[esi].cVars
						invoke wsprintf,addr szStr,CStr("%u"),eax
						inc lvi.iSubItem
						invoke ListView_SetItem( m_hWndLV,addr lvi)

						movzx eax,[esi].cImplTypes
						invoke wsprintf,addr szStr,CStr("%u"),eax
						inc lvi.iSubItem
						invoke ListView_SetItem( m_hWndLV,addr lvi)

						invoke wsprintf,addr szStr,CStr("%u"),ebx
						inc lvi.iSubItem
						invoke ListView_SetItem( m_hWndLV,addr lvi)

						.if (m_dwUserCol)
							mov szStr, 0
							.if (m_dwUserCol == IDS_GETMOPS)
								invoke vf(pTypeInfo, ITypeInfo, GetMops), MEMBERID_NIL, addr bstr
								.if (eax == S_OK && bstr)
									invoke WideCharToMultiByte,CP_ACP,0, bstr,-1,addr szStr, sizeof szStr,0,0 
									invoke SysFreeString, bstr
								.endif
							.elseif (m_dwUserCol == IDS_GETHELPCONTEXT)
								invoke vf(pTypeInfo, ITypeInfo, GetDocumentation), MEMBERID_NIL, NULL, NULL, addr dwSize, NULL
								.if (eax == S_OK)
									invoke wsprintf,addr szStr,CStr("%u"),dwSize
								.endif
							.elseif (m_dwUserCol == IDS_GETHELPFILE)
								invoke vf(pTypeInfo, ITypeInfo, GetDocumentation), MEMBERID_NIL, NULL, NULL, NULL, addr bstr
								.if (eax == S_OK && bstr)
									invoke WideCharToMultiByte,CP_ACP,0, bstr,-1,addr szStr, sizeof szStr,0,0 
									invoke SysFreeString, bstr
								.endif
							.elseif (m_dwUserCol == IDS_GETSIZEINST)
								invoke wsprintf, addr szStr, CStr("%u"), [esi].cbSizeInstance
							.elseif (m_dwUserCol == IDS_GETSIZEVFT)
								movzx eax, [esi].cbSizeVft
								invoke wsprintf, addr szStr, CStr("%u"), eax
							.elseif (m_dwUserCol == IDS_GETALIGNMENT)
								movzx eax, [esi].cbAlignment
								invoke wsprintf, addr szStr, CStr("%u"), eax
							.elseif (m_dwUserCol == IDS_GETLCID)
								invoke wsprintf, addr szStr, CStr("%X"), [esi].lcid
							.elseif (m_dwUserCol == IDS_CONSTRUCTOR)
								invoke wsprintf, addr szStr, CStr("%d"), [esi].memidConstructor
							.elseif (m_dwUserCol == IDS_DESTRUCTOR)
								invoke wsprintf, addr szStr, CStr("%d"), [esi].memidDestructor
							.elseif (m_dwUserCol == IDS_IDLFLAGS)
								movzx eax, [esi].idldescType.wIDLFlags
								invoke wsprintf, addr szStr, CStr("%X"), eax
							.endif
							inc lvi.iSubItem
							invoke ListView_SetItem( m_hWndLV,addr lvi)
						.endif

						.if (m_pIID)
							invoke IsEqualGUID,m_pIID,addr [esi].guid
							.if (eax)
								mov eax,lvi.iItem
								mov m_dwIndex,eax
							.endif
						.endif

						inc lvi.iItem
						invoke vf(pTypeInfo, ITypeInfo, ReleaseTypeAttr), pTypeAttr
					.endif
					invoke vf(pTypeInfo, ITypeInfo, Release)
				.endif
				inc ebx
			.endw
			mov hr,1
		.endif

		mov eax,m_iSortCol
		.if (eax != -1)
if 1
			mov lvc.mask_, LVCF_FMT
			lea ecx, lvc
			invoke ListView_GetColumn( m_hWndLV, eax, ecx)
			.if (lvc.fmt & LVCFMT_RIGHT)
else
			.if ([eax * sizeof CColHdr+ColumnTable].CColHdr.wFlags & FCOLHDR_RDXMASK)
endif
				@mov ecx, 1
			.else
				@mov ecx, 0
			.endif
			invoke LVSort, m_hWndLV, m_iSortCol, m_iSortDir, ecx
		.else
			invoke ResetHeaderBitmap, m_hWndLV
		.endif

		invoke SetCursor, hCsrOld
		invoke SetWindowRedraw( m_hWndLV, TRUE)
		return hr
		assume esi:nothing
		align 4

RefreshList endp

OnInitDialog proc


		DebugOut "CTypeLibDlg::OnInitDialog enter"
if ?TYPELIBICON
		mov eax, g_hIconTypeLib
		.if (!eax)
			invoke LoadIcon,g_hInstance,IDI_TYPELIB
			mov g_hIconTypeLib, eax
		.endif
		.if (eax)
			invoke SendMessage, m_hWnd, WM_SETICON, ICON_SMALL, g_hIconTypeLib
			invoke SendMessage, m_hWnd, WM_SETICON, ICON_BIG, g_hIconTypeLib
		.endif
endif
		invoke GetDlgItem, m_hWnd, IDC_LIST1
		mov m_hWndLV,eax
		invoke GetDlgItem, m_hWnd, IDC_STATUSBAR
		mov m_hWndSB,eax

		invoke ListView_SetExtendedListViewStyle( m_hWndLV,	LVS_EX_FULLROWSELECT or LVS_EX_HEADERDRAGDROP or LVS_EX_INFOTIP)

		.if (g_rect.right)
			invoke SetWindowPos, m_hWnd, NULL, g_rect.left, g_rect.top, g_rect.right, g_rect.bottom, SWP_NOZORDER or SWP_NOACTIVATE
		.endif

		invoke SetLVColumns, m_hWndLV, NUMCOLS, addr ColumnTable

		invoke RefreshList

		.if (eax && (m_dwIndex != -1))
			ListView_SetItemState m_hWndLV, m_dwIndex,\
				LVIS_SELECTED or LVIS_FOCUSED, LVIS_SELECTED or LVIS_FOCUSED
			invoke ListView_EnsureVisible( m_hWndLV, m_dwIndex, TRUE)
			mov eax, 1
		.endif
		DebugOut "CTypeLibDlg::OnInitDialog()=%X", eax
		ret
		align 4

OnInitDialog endp

;*** Register a Type library

OnRegister proc

local	szText[MAX_PATH]:byte
local	wszLib[MAX_PATH]:word

		.if (m_pszTypeLib)
			invoke MultiByteToWideChar,CP_ACP,0, m_pszTypeLib, -1, addr wszLib, MAX_PATH
			invoke RegisterTypeLib, m_pTypeLib, addr wszLib, NULL
			.if (eax == S_OK)
				invoke MessageBox, m_hWnd, CStr("Typelib successfully registered"), addr g_szHint, MB_OK
;---------------------------------------- enable "Unregister" button
				invoke GetDlgItem, m_hWnd, IDC_UNREGISTER
				invoke EnableWindow, eax, TRUE
			.else
				invoke wsprintf, addr szText,CStr("RegisterTypeLib() returned %X"), eax
				invoke MessageBox, m_hWnd, addr szText,0, MB_OK
			.endif
		.endif
		ret
		align 4

OnRegister endp


;*** Unregister a Type library


OnUnregister proc 

local	guid:GUID
local	wszGUID[40]:WORD
local	szText[MAX_PATH]:byte

		invoke UnRegisterTypeLib, addr m_guid, m_dwVerMajor, m_dwVerMinor, m_lcid, SYS_WIN32
		.if (eax == S_OK)
			invoke MessageBox, m_hWnd, CStr("Typelib successfully unregistered"), addr g_szHint, MB_OK
			invoke GetDlgItem, m_hWnd, IDCANCEL
			invoke SetFocus, eax
;---------------------------------------- disable "Unregister" button
			invoke GetDlgItem, m_hWnd, IDC_UNREGISTER
			invoke EnableWindow, eax, FALSE
			.if (m_pszTypeLib)
				invoke GetDlgItem, m_hWnd, IDC_REGISTER
				invoke EnableWindow, eax, TRUE
			.endif
		.else
			invoke wsprintf, addr szText,CStr("UnRegisterTypeLib() returned %X"), eax
			invoke MessageBox, m_hWnd, addr szText,0, MB_OK
		.endif
		ret
		align 4

OnUnregister endp

OnEdit proc uses ebx

local	lvi:LVITEM
local	iid:IID
local	hKey:HANDLE
local	pTypeInfo:LPTYPEINFO
local	pTypeAttr:ptr TYPEATTR
local	iType:dword
local	dwSize:dword
local	dwIndex:DWORD
local	typekind:TYPEKIND
local	hInstance:HINSTANCE
local	kp[4]:KEYPAIR
local	pEditDlg:ptr CEditDlg
local	szKey1[64]:byte
local	szKey2[64]:byte
local	szKey3[64]:byte
local	szKey4[64]:byte
local	szKey[260]:byte

		movzx ecx, g_bConfirmDelete
		invoke Create@CEditDlg, m_hWnd, ?EDITMODELESS, ecx
		.if (eax == 0)
			jmp done
		.endif
		mov pEditDlg,eax

		mov ebx, -1
		.while (1)
			invoke ListView_GetNextItem( m_hWndLV, ebx, LVNI_SELECTED)
			.break .if (eax == -1)
			mov ebx, eax
			mov lvi.iItem,eax
			mov lvi.iSubItem, GUIDCOL_IN_TYPELIB
			mov lvi.mask_,LVIF_TEXT
			lea eax,szKey1
			mov lvi.pszText,eax
			mov lvi.cchTextMax,sizeof szKey1
			invoke ListView_GetItem( m_hWndLV, addr lvi)

			invoke GetTypeInfoIndex, ebx
			mov dwIndex,eax
			invoke vf(m_pTypeLib, ITypeLib, GetTypeInfo), dwIndex, addr pTypeInfo
			.if (eax == S_OK)
				invoke vf(pTypeInfo, ITypeInfo, GetTypeAttr), addr pTypeAttr
				.if (eax == S_OK)
					mov eax, pTypeAttr
					mov eax,[eax].TYPEATTR.typekind
					mov typekind, eax
					invoke vf(pTypeInfo, ITypeInfo, ReleaseTypeAttr), pTypeAttr
				.endif
				invoke vf(pTypeInfo, ITypeInfo, Release)
			.endif


			invoke ZeroMemory, addr kp, sizeof kp

			.if (typekind == TKIND_COCLASS)
				mov kp[0*sizeof KEYPAIR].pszRoot, offset  g_szRootCLSID
			.else
				mov kp[0*sizeof KEYPAIR].pszRoot, offset  g_szRootInterface
			.endif
			lea eax,szKey1
			mov kp[0*sizeof KEYPAIR].pszKey,eax
			mov kp[0*sizeof KEYPAIR].bExpand,TRUE

			.if (typekind != TKIND_COCLASS)
				invoke wsprintf, addr szKey, CStr("%s\%s\ProxyStubClsid32"), offset g_szRootInterface, addr szKey1
				invoke RegOpenKeyEx,HKEY_CLASSES_ROOT,addr szKey,0,KEY_READ,addr hKey
				.if (eax == ERROR_SUCCESS)
					mov dwSize,sizeof szKey2
					invoke RegQueryValueEx,hKey,addr g_szNull,NULL,addr iType,addr szKey2,addr dwSize
					.if (szKey2 != 0)
						lea eax,szKey2
						mov kp[1*sizeof KEYPAIR].pszKey,eax
						mov kp[1*sizeof KEYPAIR].pszRoot, offset g_szRootCLSID
					.endif
					invoke RegCloseKey,hKey
				.endif

				invoke wsprintf, addr szKey, CStr("%s\%s\TypeLib"), offset g_szRootInterface, addr szKey1
				invoke RegOpenKeyEx,HKEY_CLASSES_ROOT,addr szKey,0,KEY_READ,addr hKey
				.if (eax == ERROR_SUCCESS)
					mov dwSize,sizeof szKey3
					invoke RegQueryValueEx,hKey,addr g_szNull,NULL,addr iType,addr szKey3,addr dwSize
					.if (szKey3 != 0)
						lea eax,szKey3
						mov kp[2*sizeof KEYPAIR].pszKey,eax
						mov kp[2*sizeof KEYPAIR].pszRoot, offset g_szRootTypeLib
					.endif
					invoke RegCloseKey,hKey
				.endif
			.endif

			invoke SetKeys@CEditDlg, pEditDlg, 4, addr kp
		.endw

		invoke Show@CEditDlg, pEditDlg
if ?EDITMODELESS eq 0
		invoke Destroy@CEditDlg, pEditDlg
endif
done:
		ret
		align 4

OnEdit endp


;--- WM_COMMAND handler


OnCommand proc wParam:WPARAM, lParam:LPARAM

local	pCreateInclude:ptr CCreateInclude
local	dwMode:DWORD
local	pTID:ptr CTypeInfoDlg
local	pTypeInfo:LPTYPEINFO
local	dwContext:DWORD
local	bstr:BSTR
local	szText[MAX_PATH]:byte
local	szGUID[40]:byte


		movzx eax,word ptr wParam
		.if (eax == IDCANCEL)

			invoke PostMessage, m_hWnd,WM_CLOSE,0,0

		.elseif (eax == IDOK)

			invoke PostMessage, m_hWnd, WM_COMMAND,IDM_VIEW,0

		.elseif ((eax == IDC_CREATEINC) || (eax == IDC_CREATESTUB))

if 0
			.if (m_pszTypeLib)
				invoke Create2@CCreateInclude, m_pszTypeLib
			.else
				invoke Create@CCreateInclude, m_pszGUID, m_lcid, m_dwVerMajor, m_dwVerMinor
			.endif
else
			invoke Create3@CCreateInclude, m_pTypeLib
endif
			.if (eax)
				mov pCreateInclude, eax
				movzx eax,word ptr wParam
				.if (eax == IDC_CREATEINC)
					mov dwMode, INCMODE_BASIC
				.else
					mov dwMode, INCMODE_DISPHLP
				.endif
				invoke Run@CCreateInclude, pCreateInclude, m_hWnd, dwMode
				invoke Destroy@CCreateInclude, pCreateInclude
			.endif

		.elseif (eax == IDC_REGISTER)

			invoke OnRegister

		.elseif (eax == IDC_UNREGISTER)

			invoke MessageBox, m_hWnd, CStr("Are you sure?"),\
						CStr("Unregister type library"),\
						MB_YESNO or MB_DEFBUTTON2 or MB_ICONQUESTION
			.if (eax == IDYES)
				invoke OnUnregister
			.endif

		.elseif (eax == IDM_VIEW)

			push ebx
			mov ebx, -1
			.while (1)
				invoke ListView_GetNextItem( m_hWndLV, ebx, LVNI_SELECTED)
				.break .if (eax == -1)
				mov ebx, eax
				invoke GetTypeInfoIndex, ebx
				.if (eax != -1)
					.if (m_fnCallBack)
						invoke [m_fnCallBack], m_dwCookie, m_pTypeLib, eax
						.break
					.endif
					invoke Create@CTypeInfoDlg, m_pTypeLib, eax
					.if (eax)
						mov pTID, eax
						invoke SetTab@CTypeInfoDlg, pTID, -1
						invoke Show@CTypeInfoDlg, pTID, m_hWnd, TYPEINFODLG_FACTIVATE or TYPEINFODLG_FTILE
					.endif
				.endif
			.endw
			pop ebx

		.elseif (eax == IDM_CREATE)

			invoke GetTypeInfoIndex, -1
			.if (eax != -1)	
				invoke OnCreate, eax
			.endif

		.elseif (eax == IDM_EDIT)

			invoke OnEdit

		.elseif (eax == IDM_COPYGUID)

			invoke ListView_GetNextItem( m_hWndLV,-1,LVNI_SELECTED)
			.if (eax != -1)
				lea ecx, szGUID
				ListView_GetItemText m_hWndLV, eax, GUIDCOL_IN_TYPELIB, ecx, sizeof szGUID
				invoke CopyStringToClipboard, m_hWnd, addr szGUID
			.endif

		.elseif (eax == IDM_CONTEXTHELP)

			invoke GetTypeInfoIndex, -1
			mov ecx, eax
			invoke vf(m_pTypeLib, ITypeLib, GetTypeInfo), ecx, addr pTypeInfo
			.if (eax == S_OK)
				invoke vf(pTypeInfo, ITypeInfo, GetDocumentation), MEMBERID_NIL, NULL, NULL, addr dwContext, addr bstr
				.if ((eax == S_OK) && bstr)
					invoke WideCharToMultiByte,CP_ACP,0, bstr, -1, addr szText, sizeof szText,0,0
					invoke SysFreeString, bstr
					.if (dwContext)
						invoke ShowHtmlHelp, addr szText, HH_HELP_CONTEXT, dwContext
						.if (!eax)
							push esi
if ?DRAWITEMSB
							mov esi, offset g_szErrorText
else
							sub esp, MAX_PATH+32
							mov esi, esp
endif
							invoke wsprintf, esi, CStr("HtmlHelp('%s', %u) failed"), addr szText, dwContext
							StatusBar_SetText m_hWndSB, 0 or ?SBOWNERDRAW, esi
							StatusBar_SetTipText m_hWndSB, 0, esi
							invoke MessageBeep, MB_OK
if ?DRAWITEMSB eq 0
							add esp, MAX_PATH+32
endif
							pop esi
						.endif
					.endif
				.endif
				invoke vf(pTypeInfo, IUnknown, Release)
			.else
				invoke MessageBeep, MB_OK
			.endif

		.elseif (eax == IDM_GETCLASS)

			invoke GetTypeInfoIndex, -1
			.if (eax != -1)	
				invoke OnGetClass, eax
			.endif

		.elseif (eax == IDM_SELECTALL)

			ListView_SetItemState m_hWndLV, -1, LVIS_SELECTED, LVIS_SELECTED

		.elseif (eax == IDM_COPY)

			invoke Create@CProgressDlg, m_hWndLV, NULL, SAVE_CLIPBOARD, -1
			invoke DialogBoxParam, g_hInstance, IDD_PROGRESSDLG, m_hWnd, classdialogproc, eax

		.elseif (eax == IDM_REFRESH)

			mov m_iSortCol, -1
			mov m_iSortDir, 0
			invoke RefreshList

		.endif
		ret
		align 4

OnCommand endp


;--- WM_SIZE message 

	.const
BtnTab dd IDC_REGISTER, IDC_UNREGISTER, IDC_CREATEINC, IDC_CREATESTUB, IDCANCEL
NUMBUTTONS textequ %($ - BtnTab) / sizeof DWORD
	.code

SetStatusParts proc

local	wszGUID[40]:word
local	szText[MAX_PATH]:byte

	.const
dwSBParts dd 50
	.code

	invoke SetSBParts, m_hWndSB, offset dwSBParts, LENGTHOF dwSBParts + 1

	invoke StringFromGUID2, addr m_guid, addr wszGUID, 40
	invoke wsprintf, addr szText, CStr("%S, Version %u.%u, LCID %X"), addr wszGUID, m_dwVerMajor, m_dwVerMinor, m_lcid
	StatusBar_SetText m_hWndSB, 1, addr szText
	StatusBar_SetTipText m_hWndSB, 1, addr szText
	ret
	align 4

SetStatusParts endp

OnSize proc uses ebx esi dwType:dword, dwWidth:dword, dwHeight:dword

local dwRim:DWORD
local dwHeightBtn:DWORD
local dwWidthBtn:DWORD
local dwXPos:DWORD
local dwYPos:DWORD
local dwAddX:DWORD
local dwHeightSB:DWORD
local dwHeightLV:DWORD
local rect:RECT

	invoke GetWindowRect, m_hWndSB, addr rect
	mov eax, rect.bottom
	sub eax, rect.top
	mov dwHeightSB, eax

	invoke GetWindowRect, m_hWndLV, addr rect
	invoke ScreenToClient, m_hWnd, addr rect
	mov eax, rect.left
	mov dwRim, eax

	shl eax, 1
	sub dwWidth, eax

	mov dwWidthBtn, 0
	mov esi, offset BtnTab
	mov ecx, NUMBUTTONS
	.while (ecx)
		push ecx
		lodsd
		invoke GetDlgItem, m_hWnd, eax
		lea ecx, rect
		invoke GetWindowRect, eax, ecx
		mov eax, rect.right
		sub eax, rect.left
		add dwWidthBtn, eax
		pop ecx
		.if (ecx == NUMBUTTONS)
			mov eax, rect.bottom
			sub eax, rect.top
			mov dwHeightBtn, eax
		.endif
		dec ecx
	.endw

	invoke BeginDeferWindowPos, 2 + NUMBUTTONS
	mov ebx, eax

	mov eax, dwHeight
	sub eax, dwHeightSB
	sub eax, dwRim
	sub eax, dwHeightBtn
	sub eax, dwRim
	sub eax, dwRim
	mov dwHeightLV, eax
	test eax, eax
	.if (SIGN?)
		mov dwHeightLV, 0
	.endif
	invoke DeferWindowPos, ebx, m_hWndLV, NULL, 0, 0, dwWidth, dwHeightLV, SWP_NOMOVE or SWP_NOZORDER or SWP_NOACTIVATE

	mov eax, dwRim
	add eax, dwHeightLV
	add eax, dwRim
	mov dwYPos, eax

	mov eax, dwWidth
	sub eax, dwWidthBtn
	.if (!CARRY?)
		xor edx, edx
		mov ecx, NUMBUTTONS - 1
		div ecx
		mov dwAddX, eax
	.else
		mov dwAddX, 1
	.endif
	mov eax, dwRim
	mov dwXPos, eax

	mov esi, offset BtnTab
	mov ecx, NUMBUTTONS
	.while (ecx)
		push ecx
		lodsd
		invoke GetDlgItem, m_hWnd, eax
		push eax
		lea ecx, rect
		invoke GetWindowRect, eax, ecx
		pop eax
		invoke DeferWindowPos, ebx, eax, NULL, dwXPos, dwYPos, 0, 0, SWP_NOSIZE or SWP_NOZORDER or SWP_NOACTIVATE
		mov eax, rect.right
		sub eax, rect.left
		add eax, dwAddX
		add dwXPos, eax
		pop ecx
		dec ecx
	.endw

	invoke DeferWindowPos, ebx, m_hWndSB, NULL, 0, 0, 0, 0, SWP_NOZORDER or SWP_NOACTIVATE

	invoke EndDeferWindowPos, ebx

	invoke SetStatusParts

	ret
	align 4

OnSize endp

OnSizing proc pRect:ptr RECT

local dwMinHeight:DWORD
local dwMinWidth:DWORD
local rect:RECT

		invoke GetWindowRect, m_hWndLV, addr rect
		invoke ScreenToClient, m_hWnd, addr rect
;------------------------------- calc minimal height (rim * 5)
		mov eax, rect.top
		shl eax, 2
		add eax, rect.top
		mov dwMinHeight, eax
		mov dwMinWidth, eax
		invoke GetWindowRect, m_hWndSB, addr rect
		mov eax, rect.bottom
		sub eax, rect.top
;------------------------------- add statusbar height
		add dwMinHeight, eax
		invoke GetDlgItem, m_hWnd, IDC_REGISTER
		mov ecx, eax
		invoke GetWindowRect, ecx, addr rect
		mov eax, rect.bottom
		sub eax, rect.top
;------------------------------- add button height
		add dwMinHeight, eax
;------------------------------- add height of header control (= button)
		add dwMinHeight, eax
		mov eax, rect.right
		sub eax, rect.left
		mov ecx, eax
		shl eax, 2
		add eax, ecx
		add dwMinWidth, eax


		invoke SetRect, addr rect, 0, 0, dwMinWidth, dwMinHeight
		invoke AdjustWindowRect, addr rect, WS_OVERLAPPEDWINDOW, FALSE
		mov edx, rect.bottom
		sub edx, rect.top
		mov ecx, pRect
		mov eax, [ecx].RECT.bottom
		sub eax, [ecx].RECT.top
		sub eax, edx
		.if (CARRY?)
			neg eax
			add [ecx].RECT.bottom, eax
		.endif
		mov edx, rect.right
		sub edx, rect.left
		mov eax, [ecx].RECT.right
		sub eax, [ecx].RECT.left
		sub eax, edx
		.if (CARRY?)
			neg eax
			add [ecx].RECT.right, eax
		.endif
		mov eax, 1
		ret
		align 4

OnSizing endp


;*** Dialog Proc for "typelib" dialog


TypeLibDialog proc uses __this thisarg, message:dword,wParam:dword,lParam:dword

		mov __this,this@

		mov eax,message
		.if (eax == WM_INITDIALOG)

			invoke OnInitDialog
			.if (eax)
;;				invoke CenterWindow, m_hWnd
if ?MODELESS
				invoke ShowWindow, m_hWnd, SW_NORMAL
endif
				.if (m_bTile)
					invoke GetWindowRect, m_hWnd, addr g_rect
					mov eax, g_rect.left
					sub g_rect.right, eax
					mov eax, g_rect.top
					sub g_rect.bottom, eax
					invoke GetSystemMetrics, SM_CXFULLSCREEN
					mov edx, 16
					mov ecx, g_rect.left
					add ecx, g_rect.right
					add ecx, edx
					.if (eax < ecx)
						mov g_rect.left, edx
					.else
						add g_rect.left, edx
					.endif
					invoke GetSystemMetrics, SM_CYFULLSCREEN
					mov edx, 16
					mov ecx, g_rect.top
					add ecx, g_rect.bottom
					add ecx, edx
					.if (eax < ecx)
						mov g_rect.top, edx
					.else
						add g_rect.top, edx
					.endif
				.endif
			.else
				DebugOut "TypeLibDialog: OnInitDialog failed, closing"
				invoke PostMessage, m_hWnd, WM_CLOSE, 0, 0
			.endif
			mov eax,1

		.elseif (eax == WM_CLOSE)

;---------------------------------------- get normal window pos & size

			invoke SaveNormalWindowPos, m_hWnd, addr g_rect
if ?MODELESS
			.if ( m_hWndParent )
				invoke SetActiveWindow, m_hWndParent
			.endif
			invoke DestroyWindow, m_hWnd
else
			invoke EndDialog, m_hWnd, 0
endif
			mov eax,1

		.elseif (eax == WM_DESTROY)

			invoke Destroy@CTypeLibDlg, __this
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
		.elseif (eax == WM_SIZE)

			.if (wParam != SIZE_MINIMIZED)
				movzx eax, word ptr lParam+0
				movzx ecx, word ptr lParam+2
				invoke OnSize, wParam, eax, ecx
			.endif

		.elseif (eax == WM_SIZING)

			invoke OnSizing, lParam

		.elseif (eax == WM_COMMAND)

			invoke OnCommand, wParam, lParam

		.elseif (eax == WM_NOTIFY)

			invoke OnNotify, lParam

		.elseif (eax == WM_ENTERMENULOOP)

			StatusBar_SetSimpleMode m_hWndSB, TRUE

		.elseif (eax == WM_EXITMENULOOP)

			StatusBar_SetSimpleMode m_hWndSB, FALSE

		.elseif (eax == WM_MENUSELECT)

			movzx ecx, word ptr wParam+0
			invoke DisplayStatusBarString, m_hWndSB, ecx
if ?DRAWITEMSB
		.elseif (eax == WM_DRAWITEM)

			.if (wParam == IDC_STATUSBAR)
				push esi
				mov esi, lParam
				invoke SetTextColor, [esi].DRAWITEMSTRUCT.hDC, 000000C0h
				invoke SetBkMode, [esi].DRAWITEMSTRUCT.hDC, TRANSPARENT
				add [esi].DRAWITEMSTRUCT.rcItem.left, 4
				invoke DrawTextEx, [esi].DRAWITEMSTRUCT.hDC,
					[esi].DRAWITEMSTRUCT.itemData, -1, addr [esi].DRAWITEMSTRUCT.rcItem,
					DT_LEFT or DT_SINGLELINE or DT_VCENTER, NULL
				pop esi
			.endif
			mov eax, 1
endif
if ?HTMLHELP
		.elseif (eax == WM_HELP)

			invoke DoHtmlHelp, HH_DISPLAY_TOPIC, CStr("typelibdialog.htm")
endif
		.else
			xor eax,eax ;indicates "no processing"
		.endif
		ret
		align 4

TypeLibDialog endp


;*** constructor


Create@CTypeLibDlg proc public uses __this pGuid:ptr GUID, dwVerMajor:dword,dwVerMinor:dword,lcid:LCID,pIID:ptr IID

		invoke malloc, sizeof CTypeLibDlg
		.if (!eax)
			ret
		.endif
		mov __this, eax
		pushad
		mov esi, pGuid
		lea edi, m_guid
		movsd
		movsd
		movsd
		movsd
		popad
		mov eax,dwVerMajor
		mov m_dwVerMajor,eax
		mov eax,dwVerMinor
		mov m_dwVerMinor,eax
		mov eax,lcid
		mov m_lcid,eax
		mov eax,pIID
		mov m_pIID,eax
		mov m_pDlgProc,TypeLibDialog
		mov m_iSortCol,-1
		mov m_pszTypeLib,NULL
		mov m_dwIndex, -1
		return __this
		align 4

Create@CTypeLibDlg endp


Create2@CTypeLibDlg proc public uses __this pszTypeLib:LPSTR, pIID:ptr IID, bSupressErr:BOOL

		invoke malloc, sizeof CTypeLibDlg
		.if (!eax)
			ret
		.endif
		mov __this, eax
		mov eax,1
		mov m_dwVerMajor,eax
		mov eax,0
		mov m_dwVerMinor,eax
		mov eax,0
		mov m_lcid,eax
		mov eax,pIID
		mov m_pIID,eax
		mov m_pDlgProc,TypeLibDialog
		mov m_iSortCol,-1
		invoke lstrlen, pszTypeLib
		inc eax
		invoke malloc, eax
		.if (!eax)
			ret
		.endif
		mov m_pszTypeLib,eax
		invoke lstrcpy, eax, pszTypeLib
		mov m_dwIndex, -1
		.if (bSupressErr)
			sub esp, MAX_PATH*2
			mov edx, esp
			invoke MultiByteToWideChar,CP_ACP,MB_PRECOMPOSED,
				m_pszTypeLib, -1, edx, MAX_PATH
			mov edx, esp
			invoke LoadTypeLibEx, edx, REGKIND_NONE, addr m_pTypeLib
			add esp, MAX_PATH*2
			.if (!m_pTypeLib)
				invoke Destroy@CTypeLibDlg, __this
				return 0
			.endif
		.endif
		return __this
		align 4

Create2@CTypeLibDlg endp

Create3@CTypeLibDlg proc public uses __this pTypeLib:LPTYPELIB

		invoke malloc, sizeof CTypeLibDlg
		.if (!eax)
			ret
		.endif
		mov __this, eax
		mov m_pDlgProc,TypeLibDialog
		mov m_iSortCol,-1
		mov eax, pTypeLib
		mov m_pTypeLib, eax
		invoke vf(pTypeLib, IUnknown, AddRef)
		mov m_dwIndex, -1
		return __this

Create3@CTypeLibDlg endp

Create4@CTypeLibDlg proc public uses __this pTypeInfo:LPTYPEINFO

		.if (!pTypeInfo)
			return 0
		.endif
		invoke malloc, sizeof CTypeLibDlg
		.if (!eax)
			ret
		.endif
		mov __this, eax
		mov m_pDlgProc,TypeLibDialog
		mov m_iSortCol,-1
		invoke vf(pTypeInfo, ITypeInfo, GetContainingTypeLib), addr m_pTypeLib, addr m_dwIndex
		.if (eax != S_OK)
			invoke Destroy@CTypeLibDlg, __this
			return 0
		.endif
		return __this
		align 4

Create4@CTypeLibDlg endp

Show@CTypeLibDlg proc public uses __this thisarg, hWnd:HWND, bTile:BOOL

		mov __this, this@
		mov eax, bTile
		mov m_bTile, al
		mov eax, hWnd
		mov m_hWndParent, eax
if ?MODELESS
		.if (g_bTLibDlgAsTopLevelWnd)
			mov ecx, NULL
		.else
			mov ecx, hWnd
		.endif
		invoke CreateDialogParam, g_hInstance, IDD_TYPELIBDLG,
			ecx, classdialogproc, __this
else
		invoke DialogBoxParam,g_hInstance,IDD_TYPELIBDLG,
			hWnd, classdialogproc, __this
endif
		DebugOut "Show@CTypeLibDlg=%X", eax
		ret
		align 4

Show@CTypeLibDlg endp


SetCallBack@CTypeLibDlg proc public thisarg, fnCallBack:LPVOID, dwCookie:DWORD

		mov ecx, fnCallBack
		mov edx, dwCookie
		mov eax, this@
		mov [eax].CTypeLibDlg.fnCallBack, ecx
		mov [eax].CTypeLibDlg.dwCookie, edx
		ret

SetCallBack@CTypeLibDlg endp



Destroy@CTypeLibDlg proc public uses __this thisarg

		mov __this,this@

		.if (m_hWnd)
			invoke BroadCastMessage, WM_WNDDESTROYED, 0, m_hWnd
		.endif
		.if (m_pTypeLib)
			invoke vf(m_pTypeLib,ITypeLib,Release)
			mov m_pTypeLib, NULL
		.endif
		.if (m_pszTypeLib)
			invoke free, m_pszTypeLib
			mov m_pszTypeLib, NULL
		.endif
		invoke free, __this
		ret
		align 4

Destroy@CTypeLibDlg endp

;*** end of CTypeLibDlg methods ***

	end
