
;*** definition of CTypeInfoPageDlg and CTypeInfoDlg methods

;--- the typeinfo dialog consists of 3 pages:
;---

	.386
	.model flat,stdcall
	option casemap :none   ; case sensitive
	option proc:private

_LDT_ENTRY_DEFINED equ <>	;makes Masm v8 accept winnt.inc
PLDT_ENTRY typedef ptr

INSIDE_CTYPEINFODLG equ 1
INSIDE_CTYPEINFOPAGEDLG equ 1
	include COMView.inc
	include statusbar.inc
	include classes.inc
	include rsrc.inc
	include debugout.inc

?MODELESS		equ 1
?DRAWITEMSB		equ 1
?DESTROYDLG		equ 0	;all tabs have same dialog res, so avoid creating dialogs

if ?DRAWITEMSB
?ERRORPART	equ		0 or SBT_OWNERDRAW
	.data?
g_szErrorText	db MAX_PATH+32 dup (?)	;ownerdrawn text in statusbar must be global
else
?ERRORPART	equ		0
endif

TAB_FUNCTIONS	equ 0
TAB_VARIABLES	equ 1
TAB_INTERFACES	equ 2

;--- for functions and variables

LPARAMSTRUCT struct
memid	DWORD ?		;memberid
iIndex	DWORD ?		;typeinfo index 
LPARAMSTRUCT ends

	.const

TabDlgPages label CTabDlgPage
	CTabDlgPage {CStr("Functions"),IDD_TYPEINFOPAGE,FunctionsDialog}
	CTabDlgPage {CStr("Variables"),IDD_TYPEINFOPAGE,VariablesDialog}
	CTabDlgPage {CStr("Interfaces"),IDD_TYPEINFOPAGE,InterfacesDialog}
NUMDLGS textequ %($ - TabDlgPages) / sizeof CTabDlgPage

	.data

ColTabFunctions label CColHdr
	CColHdr <CStr("Name"),						20>
	CColHdr <CStr("memid"),						8,		FCOLHDR_RDX16>
	CColHdr <CStr("FuncKind,InvKind,CallConv"),	24>
	CColHdr <CStr("rcType"),					8>
	CColHdr <CStr("Params"),					25>
	CColHdr <CStr("Flags"),						5,		FCOLHDR_RDX10>
	CColHdr <CStr("ofsVft/Entry"),				10,		FCOLHDR_RDX10>
NUMFUNCCOLS textequ %($ - ColTabFunctions) / sizeof CColHdr


ColTabVariables label CColHdr
	CColHdr <CStr("Name"),						20>
	CColHdr <CStr("memid"),						15,		FCOLHDR_RDX16>
	CColHdr <CStr("varkind"),					15>
	CColHdr <CStr("Type"),						20>
	CColHdr <CStr("Value/Offset"),				15>
	CColHdr <CStr("Flags"),						15,		FCOLHDR_RDX10>
NUMVARSCOLS textequ %($ - ColTabVariables) / sizeof CColHdr


ColTabInterfaces label CColHdr
	CColHdr <CStr("Name"),						30>
	CColHdr <CStr("IID"),						30>
	CColHdr <CStr("Flags"),						30>
	CColHdr <CStr("hRefType"),					10>
NUMINTERFACECOLS textequ %($ - ColTabInterfaces) / sizeof CColHdr


BEGIN_CLASS CTypeInfoPageDlg, CDlg
hWndTab			HWND ?
hWndLV			HWND ?
hWndSB			HWND ?
pLParam			LPVOID ?
pTypeInfo		LPTYPEINFO ?
iItem			DWORD ?				;index of current selected item
iSortCol		DWORD ?				;index of current sort column
iSortDir		DWORD ?
iType			DWORD ?				;0=Functions, 1=Variables, 2=Interfaces
wFlags			WORD ?				;activate window when first showing?
bInitialized	BOOLEAN ?
bFromTypeLib	BOOLEAN ?			;created from typelib dialog?
END_CLASS

;*** private methods, "this" parameter is edi

SetChildDlgPos			proto :HWND, :HWND, :HWND
GetDocumentationText	proto dwID:DWORD, pszText:LPSTR, iMax:DWORD
FunctionsOnNotify		proto pNMLV:ptr NMLISTVIEW
FunctionsRefresh		proto
VariablesOnNotify		proto pNMLV:ptr NMLISTVIEW
VariablesOnInitDialog	proto
StartNewDialog			proto pITypeInfo:LPTYPEINFO, iItem:DWORD
ShowContextMenu			proto :BOOL
InterfacesOnNotify		proto pNMLV:ptr NMLISTVIEW
InterfacesOnInitDialog	proto

g_rect	RECT <>

	.code

__this	textequ <edi>
_this	textequ <[__this].CTypeInfoPageDlg>
thisarg	textequ <this@:ptr CTypeInfoPageDlg>


	MEMBER hWnd, pDlgProc
	MEMBER hWndTab, hWndLV, hWndSB, iType, iItem, pLParam
	MEMBER pTypeInfo, iSortCol, iSortDir
	MEMBER bInitialized, bFromTypeLib


;--- adjust page dialog + listview to fit in client area of tab control

SetChildDlgPos proc uses ebx hWndTab:HWND, hWndDlg:HWND, hWndLV:HWND

local	rect:RECT
local	point:POINT

		invoke GetChildPos, hWndTab
		movzx ecx,ax
		mov point.x,ecx
		shr eax,16
		mov point.y,eax
		invoke GetClientRect, hWndTab,addr rect
		invoke TabCtrl_AdjustRect( hWndTab, FALSE, addr rect)

		mov edx,rect.right
		sub edx,rect.left
		mov ecx,rect.bottom
		sub ecx,rect.top
		mov eax,point.x
		add rect.left,eax
		mov eax,point.y
		add rect.top,eax
		invoke SetWindowPos, hWndDlg, HWND_TOP, rect.left, rect.top,\
			edx, ecx, SWP_NOACTIVATE or SWP_SHOWWINDOW
		invoke GetClientRect, hWndDlg, addr rect
		invoke SetWindowPos, hWndLV, 0, rect.left, rect.top,\
			rect.right, rect.bottom,\
			SWP_NOZORDER or SWP_NOACTIVATE or SWP_SHOWWINDOW
		ret
		align 4

SetChildDlgPos endp


;*** get documentation text of a member id


GetDocumentationText proc dwID:DWORD, pszText:LPSTR, iMax:DWORD

local	bstr:BSTR

		mov eax,pszText
		mov byte ptr [eax],0
		invoke vf(m_pTypeInfo,ITypeInfo,GetDocumentation),
			dwID, NULL, addr bstr, NULL, NULL
		.if ((eax == S_OK) && bstr)
			invoke WideCharToMultiByte,CP_ACP,0,bstr,-1,pszText,iMax,0,0 
			invoke SysFreeString, bstr
		.endif
		ret
		align 4

GetDocumentationText endp

;--- set status bar text on WM_INITDIALOG

SetInitStatusText proc uses ebx pTypeAttr:ptr TYPEATTR

local pszTypeKind:LPSTR
local szText[256]:byte
local wszGUID[40]:word

	mov ebx, pTypeAttr
	assume ebx:ptr TYPEATTR

	invoke StringFromGUID2, addr [ebx].guid, addr wszGUID, 40

	invoke GetTypekindStr,[ebx].typekind
	mov pszTypeKind, eax

	movzx ecx, [ebx].wTypeFlags
	movzx edx, [ebx].idldescType.wIDLFlags
	invoke wsprintf, addr szText, CStr("%S, LCID=0x%X, %s, TypeFlags=%X, IDLFlags=%X"),
		addr wszGUID, [ebx].lcid, pszTypeKind, ecx, edx
	StatusBar_SetText m_hWndSB, 0, addr szText

	ret
	assume ebx:nothing
	align 4

SetInitStatusText endp

;--- since listview may be sorted, we cannot use listview index
;--- as function index in typeinfo

_GetFuncDesc proc iItem:DWORD, ppFuncDesc:ptr ptr FUNCDESC

local	lvi:LVITEM

	mov eax, iItem
	mov lvi.iItem, eax
	mov lvi.iSubItem, 0
	mov lvi.mask_, LVIF_PARAM
	invoke ListView_GetItem( m_hWndLV, addr lvi)
	.if (eax)
		mov ecx, lvi.lParam
		invoke vf(m_pTypeInfo, ITypeInfo, GetFuncDesc), [ecx].LPARAMSTRUCT.iIndex, ppFuncDesc
	.else
		mov eax, E_FAIL
	.endif
	ret
	align 4
_GetFuncDesc endp


_GetVarDesc proc iItem:DWORD, ppVarDesc:ptr ptr VARDESC

local	lvi:LVITEM

	mov eax, iItem
	mov lvi.iItem, eax
	mov lvi.iSubItem, 0
	mov lvi.mask_, LVIF_PARAM
	invoke ListView_GetItem( m_hWndLV, addr lvi)
	.if (eax)
		mov ecx, lvi.lParam
		invoke vf(m_pTypeInfo, ITypeInfo, GetVarDesc), [ecx].LPARAMSTRUCT.iIndex, ppVarDesc
	.else
		mov eax, E_FAIL
	.endif
	ret
	align 4
_GetVarDesc endp


;--- get referenced type info from current item


GetRefTypeInfo proc

local pFuncDesc:ptr FUNCDESC
local pVarDesc:ptr VARDESC
local pTypeInfoRef:LPTYPEINFO

		mov pTypeInfoRef, NULL
		.if (m_iType == TAB_FUNCTIONS)
			invoke _GetFuncDesc, m_iItem, addr pFuncDesc
			mov ecx, pFuncDesc
			lea ecx, [ecx].FUNCDESC.elemdescFunc.tdesc
		.elseif (m_iType == TAB_VARIABLES)
			invoke _GetVarDesc, m_iItem, addr pVarDesc
			mov ecx, pVarDesc
			lea ecx, [ecx].VARDESC.elemdescVar.tdesc
		.else
			mov eax, E_FAIL
		.endif
		.if (eax == S_OK)
			mov eax, ecx
			.if ([eax].TYPEDESC.vt == VT_PTR)
				mov eax, [eax].TYPEDESC.lptdesc
			.endif
			.if ([eax].TYPEDESC.vt == VT_USERDEFINED)
				mov ecx, [eax].TYPEDESC.hreftype
				invoke vf(m_pTypeInfo, ITypeInfo, GetRefTypeInfo), ecx, addr pTypeInfoRef
			.endif
			.if (m_iType == TAB_FUNCTIONS)
				invoke vf(m_pTypeInfo,ITypeInfo,ReleaseFuncDesc),pFuncDesc
			.else
				invoke vf(m_pTypeInfo,ITypeInfo,ReleaseVarDesc),pVarDesc
			.endif
		.endif
		return pTypeInfoRef
		align 4

GetRefTypeInfo endp


;*** show context menu for all tabs ***


ShowContextMenu proc uses esi bMouse:BOOL

local	dwFlags:DWORD
local	dwSelItems:DWORD
local	pt:POINT
local	bstr:BSTR
local	dwContext:DWORD
local	lvi:LVITEM

		invoke ListView_GetSelectedCount( m_hWndLV)
		mov dwSelItems, eax
		.if (eax)
			invoke GetSubMenu,g_hMenu, ID_SUBMENU_TYPEINFODLG
			.if (eax != 0)
				mov esi, eax

				mov ecx, MF_GRAYED
				.if (dwSelItems == 1)
					.if (m_iType == TAB_VARIABLES)
						invoke SetMenuDefaultItem, esi, -1, FALSE
						mov ecx, MF_GRAYED
					.else
						invoke SetMenuDefaultItem, esi, IDM_VIEW, FALSE
						mov ecx, MF_ENABLED
					.endif
				.endif
				invoke EnableMenuItem, esi, IDM_VIEW, ecx

				invoke GetRefTypeInfo
				.if (eax && (dwSelItems == 1))
					invoke vf(eax, IUnknown, Release)
					mov ecx, MF_ENABLED
				.else
					mov ecx, MF_GRAYED
				.endif
				invoke EnableMenuItem, esi, IDM_TYPEINFOREF, ecx

				mov dwFlags, MF_GRAYED
				.if ((m_iType != TAB_INTERFACES) && (dwSelItems == 1))
					invoke ListView_GetNextItem( m_hWndLV, -1, LVNI_SELECTED)
					.if (eax != -1)
						mov lvi.iItem, eax
						mov lvi.iSubItem, 0
						mov lvi.mask_, LVIF_PARAM
						invoke ListView_GetItem( m_hWndLV, addr lvi)
						.if (eax)
							mov ecx, lvi.lParam
							invoke vf(m_pTypeInfo, ITypeInfo, GetDocumentation), [ecx].LPARAMSTRUCT.memid, NULL, NULL, addr dwContext, addr bstr
							.if ((eax == S_OK) && bstr)
								invoke SysFreeString, bstr
								.if (dwContext)
									mov dwFlags, MF_ENABLED
								.endif
							.endif
						.endif
					.endif
				.endif
				invoke EnableMenuItem, esi, IDM_CONTEXTHELP, dwFlags

				invoke GetItemPosition, m_hWndLV, bMouse, addr pt
				invoke TrackPopupMenu, esi, TPM_LEFTALIGN or TPM_LEFTBUTTON,
						pt.x,pt.y,0,m_hWnd,NULL
			.endif
		.endif

		ret
		align 4

ShowContextMenu endp


;--- handle WM_NOTIFY for all tabs


OnNotifyGeneral proc uses ebx pNMLV:ptr NMLISTVIEW

		mov ebx,pNMLV

		.if ([ebx].NMLISTVIEW.hdr.code == LVN_COLUMNCLICK)

			mov eax,[ebx].NMLISTVIEW.iSubItem
			.if (eax == m_iSortCol)
				xor m_iSortDir,1
			.else
				mov m_iSortCol,eax
				@mov m_iSortDir,0
			.endif	
			mov eax, m_iSortCol
			mov ecx, m_iType
			.if (ecx == TAB_FUNCTIONS)
				mov ecx, offset ColTabFunctions
			.elseif (ecx == TAB_VARIABLES)
				mov ecx, offset ColTabVariables
			.else
				mov ecx, offset ColTabInterfaces
			.endif

			.if ([eax * sizeof CColHdr + ecx].CColHdr.wFlags & FCOLHDR_RDXMASK)
				@mov ecx, 1
			.else
				@mov ecx, 0
			.endif
			invoke LVSort, m_hWndLV, m_iSortCol, m_iSortDir, ecx

;-------------------------------------- refresh index of selected item

			invoke ListView_GetNextItem( m_hWndLV, -1, LVIS_SELECTED)
			.if (eax != -1)
				mov m_iItem, eax
			.endif


		.elseif ([ebx].NMLISTVIEW.hdr.code == LVN_KEYDOWN)

			invoke GetKeyState, VK_CONTROL
			and 	al,80h
			.if (!ZERO?)				;Ctrl pressed?
				.if ([ebx].NMLVKEYDOWN.wVKey == 'C')
					invoke PostMessage, m_hWnd, WM_COMMAND, IDM_COPY, 0
				.endif
			.elseif ([ebx].NMLVKEYDOWN.wVKey == VK_APPS)
				invoke ShowContextMenu, FALSE
			.elseif ([ebx].NMLVKEYDOWN.wVKey == VK_F4)
				.if (!m_bFromTypeLib)
					invoke Create4@CTypeLibDlg, m_pTypeInfo
					.if (eax)
						invoke Show@CTypeLibDlg, eax, m_hWnd, FALSE
					.endif
				.endif
			.elseif ([ebx].NMLVKEYDOWN.wVKey == VK_F6)
				invoke Create@CObjectItem, m_pTypeInfo, NULL
				.if (eax)
					push eax
					invoke vf(eax, IObjectItem, ShowObjectDlg), m_hWnd
					pop eax
					invoke vf(eax, IObjectItem, Release)
				.endif
			.endif

		.elseif ([ebx].NMHDR.idFrom == IDC_LIST1)

			.if ([ebx].NMLISTVIEW.hdr.code == NM_RCLICK)

				invoke ShowContextMenu, TRUE

			.endif

		.endif
		ret
		align 4

OnNotifyGeneral endp

;*** WM_NOTIFY of "functions" dialog page

FunctionsOnNotify proc uses ebx pNMLV:ptr NMLISTVIEW

local	pFuncDesc:ptr FUNCDESC
local	szText[160]:byte

		mov ebx,pNMLV

		.if ([ebx].NMLISTVIEW.hdr.code == LVN_ITEMCHANGED)

			.if ([ebx].NMLISTVIEW.uNewState & LVIS_SELECTED)
				mov eax, [ebx].NMLISTVIEW.iItem
				mov m_iItem, eax

				mov ecx, [ebx].NMLISTVIEW.lParam
				invoke GetDocumentationText, [ecx].LPARAMSTRUCT.memid, addr szText, sizeof szText
				StatusBar_SetText m_hWndSB, 0, addr szText
				invoke _GetFuncDesc, m_iItem, addr pFuncDesc
				.if (eax == S_OK)
					mov eax, pFuncDesc
					movzx ecx, [eax].FUNCDESC.wFuncFlags
					invoke GetFuncFlags, ecx, addr szText
					invoke vf(m_pTypeInfo, ITypeInfo, ReleaseFuncDesc),pFuncDesc
					StatusBar_SetText m_hWndSB, 1, addr szText
				.endif
			.endif

		.elseif ([ebx].NMHDR.idFrom == IDC_LIST1)

			.if ([ebx].NMLISTVIEW.hdr.code == NM_DBLCLK)

				invoke PostMessage, m_hWnd, WM_COMMAND, IDM_VIEW, 0

			.elseif ([ebx].NMLISTVIEW.hdr.code == NM_RETURN)

				invoke PostMessage, m_hWnd, WM_COMMAND, IDM_VIEW, 0

			.else

				invoke OnNotifyGeneral, ebx

			.endif

		.endif

		ret
		align 4

FunctionsOnNotify endp


;*** refresh "functions" display


FunctionsRefresh proc uses ebx esi

local	lvi:LVITEM
local	pTypeAttr:ptr TYPEATTR
local	pFuncDesc:ptr FUNCDESC
local	pTypeInfo:LPTYPEINFO
local	rect:RECT
local	dwReturn:dword
local	pbstr:ptr BSTR
local	bstrDll:BSTR
local	bstrName:BSTR
local	pArray:ptr
local   dwNumNames:dword
local   wOrdinal:word
local	szName[128]:byte
local	szDll[MAX_PATH]:byte
local	szStr[1024]:byte

		mov m_iSortCol,-1

		.if (m_pLParam)
			invoke free, m_pLParam
			mov m_pLParam, NULL
		.endif

		invoke SetWindowRedraw( m_hWndLV, FALSE)

		invoke ListView_DeleteAllItems( m_hWndLV)

		mov eax,m_pTypeInfo
		mov pTypeInfo,eax				;we need this later
		invoke vf(m_pTypeInfo,ITypeInfo,GetTypeAttr),addr pTypeAttr
		.if (eax == S_OK)
			mov ebx,0
			mov esi,pTypeAttr
			invoke SetInitStatusText, esi
			mov lvi.iItem,0

			movzx eax, [esi].TYPEATTR.cFuncs
			.if (eax)
				shl eax, 3
				invoke malloc, eax
				mov m_pLParam, eax
			.endif

			.while (bx < [esi].TYPEATTR.cFuncs)
				push esi
				mov lvi.iSubItem,0
				lea eax,szStr
				mov lvi.pszText,eax
				mov lvi.mask_,LVIF_TEXT or LVIF_PARAM
				invoke vf(m_pTypeInfo,ITypeInfo,GetFuncDesc), ebx, addr pFuncDesc
				.if (eax == S_OK)

					mov esi,pFuncDesc

					mov ecx, m_pLParam
					lea edx, [ecx+ebx*sizeof LPARAMSTRUCT]
					mov eax, [esi].FUNCDESC.memid
					mov [edx].LPARAMSTRUCT.memid, eax
					mov [edx].LPARAMSTRUCT.iIndex, ebx
					mov lvi.lParam, edx

					movzx ecx,[esi].FUNCDESC.cParams
					inc ecx
					mov dwNumNames, ecx
					mov eax,sizeof BSTR
					mul ecx
					invoke malloc, eax
					mov pbstr, eax

					invoke vf(m_pTypeInfo,ITypeInfo,GetNames),[esi].FUNCDESC.memid, pbstr, dwNumNames, addr dwReturn

					mov edx,pbstr
					.if (eax == S_OK && dword ptr [edx])
						invoke WideCharToMultiByte,CP_ACP,0,[edx],-1,addr szStr,sizeof szStr,0,0 
						mov edx,pbstr
						invoke SysFreeString,[edx]
					.else
						invoke lstrcpy,addr szStr,CStr("?")
					.endif

;;					mov eax,[esi].FUNCDESC.memid
;;					mov lvi.lParam,eax
					invoke ListView_InsertItem( m_hWndLV,addr lvi)

					mov lvi.mask_,LVIF_TEXT
					.if (g_bMemIdInDecimal)
						mov edx, CStr("%d")
					.else
						mov edx, CStr("0x%X")
					.endif
					invoke wsprintf, addr szStr, edx, [esi].FUNCDESC.memid
					inc lvi.iSubItem
					invoke ListView_SetItem( m_hWndLV,addr lvi)

					invoke GetFuncKind,[esi].FUNCDESC.funckind
					push eax
					invoke GetCallConv,[esi].FUNCDESC.callconv
					push eax
					invoke GetInvokeKind,[esi].FUNCDESC.invkind
					pop edx
					pop ecx
					invoke wsprintf,addr szStr,CStr("%s, %s, %s"),ecx,eax,edx
					inc lvi.iSubItem
					invoke ListView_SetItem( m_hWndLV,addr lvi)

					invoke GetParameterType, m_pTypeInfo,\
						addr [esi].FUNCDESC.elemdescFunc.tdesc, addr szStr, sizeof szStr
					inc lvi.iSubItem
					invoke ListView_SetItem( m_hWndLV,addr lvi)

					movzx eax, [esi].FUNCDESC.elemdescFunc.paramdesc.wParamFlags
;;					DebugOut "Return: %s Flags=%X", addr szStr, eax

					pushad
					mov edi,pbstr
					add edi, sizeof BSTR
					movzx ecx,[esi].FUNCDESC.cParams		;load into registers to
					mov esi,[esi].FUNCDESC.lprgelemdescParam
					mov szStr,0
					.while (ecx)
						push ecx
						mov szName,0
						.if (dword ptr [edi])
							invoke WideCharToMultiByte,CP_ACP,0,[edi],-1,addr szName,sizeof szName,0,0 
							invoke SysFreeString,[edi]
						.endif
						invoke lstrcat, addr szStr, addr szName
						invoke lstrcat, addr szStr, CStr(":")
						invoke lstrlen, addr szStr
						lea ecx, szStr
						mov edx, sizeof szStr
						sub edx,eax
						add eax,ecx
						invoke GetParameterType, pTypeInfo, esi, eax, edx

						movzx eax, [esi].ELEMDESC.paramdesc.wParamFlags
;;						DebugOut "%.192s Flags=%X", addr szStr, eax

						pop ecx
						push ecx
						.if (ecx > 1)
							invoke lstrcat,addr szStr, CStr(", ")
						.endif
						pop ecx
						dec ecx
						add esi,sizeof ELEMDESC
						add edi, sizeof BSTR
					.endw
					invoke free, pbstr
					popad

					inc lvi.iSubItem
					invoke ListView_SetItem( m_hWndLV,addr lvi)

					movzx edx,[esi].FUNCDESC.wFuncFlags
					invoke wsprintf,addr szStr,CStr("%X"),edx
					inc lvi.iSubItem
					invoke ListView_SetItem( m_hWndLV,addr lvi)

					.if ([esi].FUNCDESC.funckind == FUNC_STATIC)
						mov ecx, [esi].FUNCDESC.invkind
						invoke vf(m_pTypeInfo, ITypeInfo, GetDllEntry),\
							[esi].FUNCDESC.memid, ecx, addr bstrDll, addr bstrName, addr wOrdinal
						.if (eax == S_OK)
							invoke WideCharToMultiByte,CP_ACP,0,bstrDll,-1,addr szDll,sizeof szDll,0,0 
							invoke SysFreeString, bstrDll
							.if (bstrName)
								invoke WideCharToMultiByte,CP_ACP,0,bstrName,-1,addr szName,sizeof szName,0,0 
								invoke wsprintf,addr szStr,CStr("%s:%s"),addr szDll, addr szName
								invoke SysFreeString, bstrName
							.else
								movzx ecx, wOrdinal
								invoke wsprintf,addr szStr,CStr("%s:%u"),addr szDll, ecx
							.endif
						.else
							invoke wsprintf,addr szStr,CStr("GetDllEntry failed [%X]"),eax
						.endif
					.else
						movzx edx,[esi].FUNCDESC.oVft			;avoid "invoke" bug
						invoke wsprintf,addr szStr,CStr("%u"),edx
					.endif
					inc lvi.iSubItem
					invoke ListView_SetItem( m_hWndLV,addr lvi)

					invoke vf(m_pTypeInfo,ITypeInfo,ReleaseFuncDesc),pFuncDesc
					inc lvi.iItem
				.endif
				inc ebx
				pop esi
			.endw
			invoke vf(m_pTypeInfo,ITypeInfo,ReleaseTypeAttr),pTypeAttr
		.endif
		invoke SetWindowRedraw( m_hWndLV, TRUE)
		ret
		align 4

FunctionsRefresh endp


;--- show the function in a bit more detail


viewdetailproc proc uses ebx esi __this hWnd:HWND, message:DWORD, wParam:WPARAM, lParam:LPARAM

local	hWndLV:HWND
local	pFuncDesc:ptr FUNCDESC
local	lvi:LVITEM
local	var:VARIANT
local	orgvt:WORD
local	szDefault[64]:byte
local	szType[128]:byte
local	szText[2048]:byte
local	szText2[1024]:byte

		mov eax, message
		.if (eax == WM_INITDIALOG)
			mov __this, lParam
			invoke ListView_GetNextItem( m_hWndLV, -1, LVNI_SELECTED)
			.if (eax != -1)
				mov lvi.iItem, eax
				invoke _GetFuncDesc, lvi.iItem, addr pFuncDesc
				.if (eax != S_OK)
					invoke MessageBeep, MB_OK
					mov eax, 1
					jmp done
				.endif

				mov esi, pFuncDesc
				assume esi:ptr FUNCDESC

				mov lvi.mask_,LVIF_TEXT
				lea eax, szText2
				mov lvi.pszText, eax
				mov lvi.cchTextMax, sizeof szText2
;------------------------------------------------- save this so we can use edi
				mov eax, m_hWndLV
				mov hWndLV, eax

;------------------------------------------------- get function flags
				push edi

				lea edi, szText
				mov al,'['
				stosb
				movzx ecx, [esi].FUNCDESC.wFuncFlags
				invoke GetFuncFlags, ecx, edi
				invoke lstrlen, edi
				add edi, eax
				mov ax," ]"
				stosw

;------------------------------------------------- get returncode

				mov lvi.iSubItem,3
				invoke ListView_GetItem( hWndLV, addr lvi)

				push esi
				lea esi, szText2
				.while (1)
					lodsb
					.break .if (al == 0)
					stosb
				.endw
				pop  esi
				mov ax, "  "
				stosw

;------------------------------------------------- get function name
				mov lvi.iSubItem,0
				invoke ListView_GetItem( hWndLV, addr lvi)

				push esi
				lea esi, szText2
				.while (1)
					lodsb
					.break .if (al == 0)
					stosb
				.endw
				pop  esi

				mov ax, "( "
				stosw
				mov al, " "
				stosb

;------------------------------------------------- get function params
				mov lvi.iSubItem,4
				invoke ListView_GetItem( hWndLV, addr lvi)

				movzx ecx,[esi].FUNCDESC.cParams		;load into registers to
				mov esi,[esi].FUNCDESC.lprgelemdescParam
				assume esi:ptr ELEMDESC
				lea ebx, szText2
				.while (ecx)
					push ecx

					movzx ecx, [esi].paramdesc.wParamFlags
					invoke GetParamFlags, ecx, addr szType
					movzx ecx, [esi].paramdesc.wParamFlags
					mov szDefault, 0
					.if (ecx & PARAMFLAG_FHASDEFAULT)
						invoke VariantInit, addr var
						mov ecx, [esi].paramdesc.pparamdescex
						mov edx, [ecx].PARAMDESCEX.cBytes
						sub edx, sizeof VARIANT
						add edx, ecx
						mov ax, [edx].VARIANT.vt
						mov orgvt, ax
						invoke VariantChangeType, addr var, edx, 0, VT_BSTR
						.if (eax == S_OK)
							invoke WideCharToMultiByte,CP_ACP,0,var.bstrVal,-1,addr szDefault,sizeof szDefault,0,0 
						.endif
						invoke VariantClear, addr var
					.endif
					mov ax, 0A0Dh
					stosw
					mov ax,"[ "
					stosw
					push esi
					lea esi, szType
					.while (1)
						lodsb
						.break .if (al == 0)
						stosb
					.endw
					mov ax," ]"
					stosw
					mov esi, ebx
					.while (1)
						lodsb
						.break .if ((al == 0) || (al == ','))
						stosb
					.endw
					push eax
					mov ebx, esi
					.if (szDefault)
						mov ax, "= "
						stosw
						stosb
						.if (orgvt == VT_BSTR)
							mov al, '"'
							stosb
						.endif
						lea esi, szDefault
						lodsb
						.while (al)
							stosb
							lodsb
						.endw
						.if (orgvt == VT_BSTR)
							mov al, '"'
							stosb
						.endif
					.endif
					pop eax
					.if (al)
						stosb
					.endif
					pop esi

					pop ecx
					dec ecx
					add esi,sizeof ELEMDESC
				.endw
				mov ax, ") "
				stosw

if 0
				mov esi, pFuncDesc
				assume esi:ptr FUNCDESC

				mov ax, 0A0Dh
				stosw
				movzx edx, [esi].cParams
				movzx ecx, [esi].cParamsOpt
				invoke wsprintf, edi, CStr("cParams=%u, cParamsOpt=%u",13,10), edx, ecx
				add edi, eax
				movzx edx, [esi].oVft
				movzx ecx, [esi].cScodes
				invoke wsprintf, edi, CStr("oVft=%u, cScodes=%u",13,10), edx, ecx
				add edi, eax
				movzx edx, [esi].elemdescFunc.paramdesc.wParamFlags
				movzx ecx, [esi].elemdescFunc.tdesc.vt
				invoke wsprintf, edi, CStr("elemdescFunc.paramdesc.wParamFlags=%X elemdescFunc.tdesc.vt=%X",13,10), edx, ecx
				add edi, eax
endif

				mov al,0
				stosb

				pop edi

				invoke vf(m_pTypeInfo,ITypeInfo,ReleaseFuncDesc),pFuncDesc

				invoke SetDlgItemText, hWnd, IDC_EDIT1, addr szText
			.endif
			mov eax, 1
		.elseif (eax == WM_CLOSE)
			invoke EndDialog, hWnd, 0
		.elseif (eax == WM_COMMAND)
			movzx eax, word ptr wParam+0
			.if ((eax == IDCANCEL) || (eax == IDOK))
				invoke EndDialog, hWnd, 0
			.endif
		.else
			xor eax, eax
		.endif
done:
		ret
		align 4
		assume esi:nothing

viewdetailproc endp

;--- start htmlhelp to display context help

DisplayContextHelp proc

local	pVarDesc:ptr VARDESC
local	pFuncDesc:ptr FUNCDESC
local	dwContext:DWORD
local	bstr:BSTR
local	szText[MAX_PATH]:byte

		invoke ListView_GetNextItem( m_hWndLV, -1, LVNI_SELECTED)
		.if (eax != -1)
			mov ecx, eax
			.if (m_iType == TAB_VARIABLES)
				invoke _GetVarDesc, ecx, addr pVarDesc
			.else
				invoke _GetFuncDesc, ecx, addr pFuncDesc
			.endif
			.if (eax == S_OK)
				.if (m_iType == TAB_VARIABLES)
					mov ecx, pVarDesc
					mov ecx, [ecx].VARDESC.memid
				.else
					mov ecx, pFuncDesc
					mov ecx, [ecx].FUNCDESC.memid
				.endif
				invoke vf(m_pTypeInfo, ITypeInfo, GetDocumentation), ecx, NULL, NULL, addr dwContext, addr bstr
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
							StatusBar_SetText m_hWndSB, ?ERRORPART, esi
							StatusBar_SetTipText m_hWndSB, 0, esi
							invoke MessageBeep, MB_OK
if ?DRAWITEMSB eq 0
							add esp, MAX_PATH+32
endif
							pop esi
						.endif
					.endif
				.endif
				.if (m_iType == TAB_VARIABLES)
					invoke vf(m_pTypeInfo, ITypeInfo, ReleaseVarDesc), pVarDesc
				.else
					invoke vf(m_pTypeInfo, ITypeInfo, ReleaseFuncDesc), pFuncDesc
				.endif
			.endif
		.endif
		ret
		align 4
DisplayContextHelp endp


FunctionsDialog proc uses __this thisarg, message:dword,wParam:WPARAM,lParam:LPARAM

local	dwContext:DWORD
local	bstr:BSTR
local	pFuncDesc:ptr FUNCDESC
local	szText[MAX_PATH]:byte

		mov __this,this@

		mov eax,message
		.if (eax == WM_INITDIALOG)

			invoke GetDlgItem, m_hWnd, IDC_LIST1
			mov m_hWndLV,eax
			.if (!m_bInitialized)
				invoke SetChildDlgPos, m_hWndTab, m_hWnd, m_hWndLV
				invoke SetLVColumns, m_hWndLV, NUMFUNCCOLS, addr ColTabFunctions
				mov m_bInitialized, TRUE
			.endif
			invoke FunctionsRefresh
			mov eax,1

		.elseif (eax == WM_COMMAND)

			movzx eax,word ptr wParam
			.if (eax == IDM_VIEW)
				.if (m_iItem != -1)	
					invoke DialogBoxParam, g_hInstance, IDD_VIEWDETAIL, m_hWnd, viewdetailproc, __this
				.else
					invoke MessageBeep, MB_OK
				.endif
			.elseif (eax == IDM_TYPEINFOREF)
				.if (m_iItem == -1)
					invoke MessageBeep, MB_OK
				.else
					invoke GetRefTypeInfo
					.if (eax)
						push eax
						invoke Create2@CTypeInfoDlg, eax
						.if (eax)
							invoke Show@CTypeInfoDlg, eax, m_hWnd, TYPEINFODLG_FACTIVATE or TYPEINFODLG_FTILE
						.endif
						pop eax
						invoke vf(eax, ITypeInfo, Release)
					.endif
				.endif

			.elseif (eax == IDM_CONTEXTHELP)

				invoke DisplayContextHelp

			.else
				invoke GetParent, m_hWnd
				invoke SendMessage, eax, message, wParam, lParam
			.endif
			xor eax,eax

		.elseif (eax == WM_NOTIFY)
			invoke FunctionsOnNotify, lParam
		.else
			xor eax,eax ;indicates "no processing"
		.endif
		ret
		align 4
FunctionsDialog endp

;--- this proc is called by CCreateInclude and CPropertiesDlg!

GetVariant proc public uses esi pVariant:ptr VARIANT, pStr:LPSTR,dwMax:dword, pszHexPrefix:LPSTR

local	variant:VARIANT

		invoke VariantInit,addr variant
		mov esi, pszHexPrefix
		.if (esi)
			mov ecx, pVariant
			mov ax, [ecx].VARIANT.vt
			mov edx, [ecx].VARIANT.lVal
			.if ((ax == VT_I4) || (ax == VT_UI4))
				;
			.elseif ((ax == VT_I2) || (ax == VT_UI2))
				movzx edx, dx
			.elseif ((ax == VT_I1) || (ax == VT_UI1))
				movzx edx, dl
			.elseif ((ax == VT_BSTR) && edx && (byte ptr [esi] == '0'))
				mov ecx, edx
				.while (1)
					mov ax, [ecx]
					.break .if (ax < ' ')
					inc ecx
					inc ecx
				.endw
				.if (!ax)
					jmp step1
				.else
					push edi
					push esi
					mov edi, pStr
					mov esi, edx
					sub dwMax, 6
					.while (1)
						lodsw
						.break .if (!ax)
						movzx eax, ax
						invoke wsprintf, edi, CStr("%s%X,"), pszHexPrefix, eax
						add edi, eax
						sub dwMax, eax
						.if (CARRY?)
							mov eax, "..."
							stosd
							.break
						.endif
					.endw
					.if (!ax)
						dec edi
						stosb
					.endif
					pop esi
					pop edi
					jmp done
				.endif
			.else
				jmp step1
			.endif
			invoke wsprintf, pStr, CStr("%s%X"), esi, edx
		.else
step1:
			invoke VariantChangeType,addr variant,pVariant,0,VT_BSTR
			.if (eax == S_OK)
				invoke SysStringByteLen,variant.bstrVal
				mov ecx,eax
				invoke WideCharToMultiByte,CP_ACP,0,variant.bstrVal,ecx,pStr,dwMax,0,0 
			.else
				invoke lstrcpy,pStr,CStr("???")
			.endif
			invoke VariantClear,addr variant
		.endif
done:
		ret
		align 4
GetVariant endp

;*** WM_NOTIFY of "functions" dialog page

VariablesOnNotify proc uses ebx pNMLV:ptr NMLISTVIEW

local	pVarDesc:ptr VARDESC
local	szText[128]:byte

		mov ebx,pNMLV

		.if ([ebx].NMLISTVIEW.hdr.code == LVN_ITEMCHANGED)

			.if ([ebx].NMLISTVIEW.uNewState & LVIS_SELECTED)
				mov eax, [ebx].NMLISTVIEW.iItem
				mov m_iItem, eax
				mov ecx, [ebx].NMLISTVIEW.lParam
				invoke GetDocumentationText, [ecx].LPARAMSTRUCT.memid, addr szText, sizeof szText
				StatusBar_SetText m_hWndSB, 0, addr szText

				invoke _GetVarDesc, m_iItem, addr pVarDesc
				.if (eax == S_OK)
					mov eax, pVarDesc
					movzx ecx, [eax].VARDESC.wVarFlags
					invoke GetVarFlags, ecx, addr szText
					invoke vf(m_pTypeInfo, ITypeInfo, ReleaseVarDesc),pVarDesc
					StatusBar_SetText m_hWndSB, 1, addr szText
				.endif

			.endif

		.else

			invoke OnNotifyGeneral, ebx

		.endif
		ret
		align 4

VariablesOnNotify endp


VariablesOnInitDialog proc uses ebx esi

local	lvc:LVCOLUMN
local	lvi:LVITEM
local	pTypeAttr:ptr TYPEATTR
local	pVarDesc:ptr VARDESC
local	szStr[260]:byte
local	szVariant[260]:byte
local	rect:RECT
local	dwReturn:dword
local	bstr:BSTR
local	pArray:ptr
local	pVariant:ptr VARIANT

		invoke SetLVColumns, m_hWndLV, NUMVARSCOLS, addr ColTabVariables

		.if (m_pLParam)
			invoke free, m_pLParam
			mov m_pLParam, NULL
		.endif

		mov m_iSortCol,-1

		invoke ListView_DeleteAllItems( m_hWndLV)

		invoke vf(m_pTypeInfo,ITypeInfo,GetTypeAttr),addr pTypeAttr
		.if (eax == S_OK)
			mov ebx,0
			mov esi,pTypeAttr
			invoke SetInitStatusText, esi
			mov lvi.iItem,0

			movzx eax, [esi].TYPEATTR.cVars
			.if (eax)
				shl eax, 3
				invoke malloc, eax
				mov m_pLParam, eax
			.endif

			.while (bx < [esi].TYPEATTR.cVars)
				mov lvi.mask_,LVIF_TEXT or LVIF_PARAM
				mov lvi.iSubItem,0
				lea eax,szStr
				mov lvi.pszText,eax
				invoke vf(m_pTypeInfo,ITypeInfo,GetVarDesc), ebx, addr pVarDesc
				push ebx
				.if (eax == S_OK)

					mov ecx, m_pLParam
					lea edx, [ecx+ebx*sizeof LPARAMSTRUCT]
					mov [edx].LPARAMSTRUCT.iIndex, ebx

					mov ebx,pVarDesc

					mov eax, [ebx].VARDESC.memid
					mov [edx].LPARAMSTRUCT.memid, eax
					mov lvi.lParam, edx


					mov bstr, NULL
;;					mov eax,[ebx].VARDESC.memid
;;					mov lvi.lParam,eax
					invoke vf(m_pTypeInfo,ITypeInfo,GetNames), [ebx].VARDESC.memid, addr bstr,1, addr dwReturn

					.if (bstr)
						invoke WideCharToMultiByte, CP_ACP, 0, bstr, -1, addr szStr, sizeof szStr,0,0 
						invoke SysFreeString,bstr
					.else
						invoke lstrcpy,addr szStr,CStr("?")
					.endif
					invoke ListView_InsertItem( m_hWndLV, addr lvi)
					inc lvi.iSubItem

					mov lvi.mask_,LVIF_TEXT
					.if (g_bMemIdInDecimal)
						mov edx, CStr("%d")
					.else
						mov edx, CStr("0x%X")
					.endif
					invoke wsprintf, addr szStr, edx, [ebx].VARDESC.memid
					invoke ListView_SetItem( m_hWndLV,addr lvi)
					inc lvi.iSubItem

					invoke GetVarKind, [ebx].VARDESC.varkind
					invoke wsprintf,addr szStr,CStr("%s"), eax
					invoke ListView_SetItem( m_hWndLV,addr lvi)
					inc lvi.iSubItem

					mov szStr,0
					invoke GetParameterType, m_pTypeInfo,\
						addr [ebx].VARDESC.elemdescVar.tdesc, addr szStr, sizeof szStr
					invoke ListView_SetItem( m_hWndLV,addr lvi)
					inc lvi.iSubItem

					mov szVariant, 0
					.if ([ebx].VARDESC.varkind == VAR_CONST)
						.if (g_bValueInDecimal)
							xor ecx, ecx
						.else
							mov ecx, CStr("0x")
						.endif
						invoke GetVariant, [ebx].VARDESC.lpvarValue,
							addr szVariant, sizeof szVariant, ecx
					.elseif ([ebx].VARDESC.varkind == VAR_PERINSTANCE)
						invoke wsprintf,addr szVariant,CStr("0x%X"),[ebx].VARDESC.oInst
					.else
						mov szVariant,0
					.endif
					invoke wsprintf,addr szStr,CStr("%s"),addr szVariant
					invoke ListView_SetItem( m_hWndLV,addr lvi)
					inc lvi.iSubItem

					movzx edx,[ebx].VARDESC.wVarFlags
					invoke wsprintf,addr szStr,CStr("%u"), edx
					invoke ListView_SetItem( m_hWndLV,addr lvi)

					invoke vf(m_pTypeInfo,ITypeInfo,ReleaseVarDesc), ebx
					inc lvi.iItem
				.endif
				pop ebx
				inc ebx
			.endw
			invoke vf(m_pTypeInfo,ITypeInfo,ReleaseTypeAttr),pTypeAttr
		.endif
		ret
		align 4
VariablesOnInitDialog endp



VariablesDialog proc uses __this thisarg, message:dword,wParam:WPARAM,lParam:LPARAM

local	rect:RECT

		mov __this,this@

		mov eax,message
		.if (eax == WM_INITDIALOG)

			invoke GetDlgItem, m_hWnd, IDC_LIST1
			mov m_hWndLV,eax
			invoke SetChildDlgPos, m_hWndTab, m_hWnd, m_hWndLV
			invoke VariablesOnInitDialog

			mov eax,1
		.elseif (eax == WM_COMMAND)

			movzx eax,word ptr wParam
			.if (eax == IDM_TYPEINFOREF)
				.if (m_iItem == -1)
					invoke MessageBeep, MB_OK
				.else
					invoke GetRefTypeInfo
					.if (eax)
						push eax
						invoke Create2@CTypeInfoDlg, eax
						.if (eax)
							invoke Show@CTypeInfoDlg, eax, m_hWnd, TYPEINFODLG_FACTIVATE or TYPEINFODLG_FTILE
						.endif
						pop eax
						invoke vf(eax, ITypeInfo, Release)
					.endif
				.endif

			.elseif (eax == IDM_CONTEXTHELP)

				invoke DisplayContextHelp

			.else
				invoke GetParent, m_hWnd
				invoke SendMessage, eax, message, wParam, lParam
			.endif
			xor eax,eax

		.elseif (eax == WM_NOTIFY)

			invoke VariablesOnNotify, lParam

		.else
			xor eax,eax			;indicates "no processing"
		.endif
		ret
		align 4

VariablesDialog endp

;*** start a new dialog box

StartNewDialog proc pTypeInfo:LPTYPEINFO, iItem:DWORD

local	reftype:HREFTYPE
local	pTypeInfoRef:LPTYPEINFO
local	pTID:ptr CTypeInfoDlg
local	lvi:LVITEM

		mov eax, iItem
		mov lvi.iItem, eax
		@mov lvi.iSubItem, 0
		mov lvi.mask_, LVIF_PARAM
		invoke ListView_GetItem( m_hWndLV, addr lvi)

		invoke vf(pTypeInfo,ITypeInfo,GetRefTypeOfImplType), lvi.lParam, addr reftype
		.if (eax == S_OK)
			invoke vf(pTypeInfo,ITypeInfo,GetRefTypeInfo), reftype, addr pTypeInfoRef
			.if (eax == S_OK)
				invoke Create2@CTypeInfoDlg, pTypeInfoRef
				.if (eax)
					mov pTID, eax
					invoke Show@CTypeInfoDlg, eax, m_hWnd, TYPEINFODLG_FACTIVATE or TYPEINFODLG_FTILE
				.endif
				invoke vf(pTypeInfoRef,ITypeInfo,Release)
			.endif
		.endif
		ret
		align 4

StartNewDialog endp



;*** WM_NOTIFY of "interfaces" dialog page


InterfacesOnNotify proc pNMLV:ptr NMLISTVIEW

local	pt:POINT

		mov eax,pNMLV
		.if ([eax].NMHDR.idFrom == IDC_LIST1)

			.if ([eax].NMLISTVIEW.hdr.code == NM_DBLCLK)

				invoke StartNewDialog, m_pTypeInfo, [eax].NMLISTVIEW.iItem

			.elseif ([eax].NMLISTVIEW.hdr.code == NM_RETURN)

				invoke PostMessage, m_hWnd, WM_COMMAND, IDM_VIEW, 0

			.else

				invoke OnNotifyGeneral, eax

			.endif

		.endif
		ret
		align 4

InterfacesOnNotify endp

InterfacesOnInitDialog proc uses ebx esi

local	lvc:LVCOLUMN
local	lvi:LVITEM
local	pTypeAttr:ptr TYPEATTR
local	pTypeAttr2:ptr TYPEATTR
local	implTypeFlags:dword
local	szText1[260]:byte
local	szText2[40]:byte
local	szText3[128]:byte
local	szText4[32]:byte
;local	szIID[40]:byte
local	wszIID[40]:word
local	rect:RECT
local	dwReturn:dword
local	pArray:ptr
local	pVariant:ptr VARIANT
local	pTypeInfoRef:LPTYPEINFO
local	reftype:HREFTYPE
local	pszUndef:LPSTR
local	bstr:BSTR
local	dwCnt:DWORD

		mov pszUndef,CStr("???")

		invoke SetLVColumns, m_hWndLV, NUMINTERFACECOLS, addr ColTabInterfaces

		invoke ListView_DeleteAllItems( m_hWndLV)

		invoke vf(m_pTypeInfo,ITypeInfo,GetTypeAttr),addr pTypeAttr
		.if (eax == S_OK)
			mov esi,pTypeAttr
			invoke SetInitStatusText, esi

			mov ebx,0
;------------------------------------- if type=DISPATCH+DUAL, get ref. TKIND_INTERFACE
			.if (([esi].TYPEATTR.typekind == TKIND_DISPATCH) &&	([esi].TYPEATTR.wTypeFlags & TYPEFLAG_FDUAL))
				mov ebx,-1
			.endif

			mov lvi.iItem,0

			movzx eax, [esi].TYPEATTR.cImplTypes
			mov dwCnt, eax

			.while (dwCnt)

				invoke lstrcpy, addr szText1, pszUndef
				invoke lstrcpy, addr szText2, pszUndef
				invoke lstrcpy, addr szText4, pszUndef

				.if (ebx != -1)
					invoke lstrcpy, addr szText3, pszUndef
					invoke vf(m_pTypeInfo,ITypeInfo,GetImplTypeFlags), ebx, addr implTypeFlags
					.if (eax == S_OK)
						invoke GetImplTypeFlags_, implTypeFlags, addr szText3, sizeof szText3
					.endif
				.else
					mov szText3, 0
				.endif
				invoke vf(m_pTypeInfo,ITypeInfo,GetRefTypeOfImplType), ebx, addr reftype
				.if (eax == S_OK)
					invoke wsprintf,addr szText4,CStr("0x%X"),reftype
					invoke vf(m_pTypeInfo,ITypeInfo,GetRefTypeInfo), reftype, addr pTypeInfoRef
					.if (eax == S_OK)
						mov bstr,0
						invoke vf(pTypeInfoRef,ITypeInfo,GetDocumentation),\
							MEMBERID_NIL,addr bstr,NULL,NULL,NULL
						.if (bstr)
				 			invoke WideCharToMultiByte,CP_ACP,0,bstr,-1,addr szText1,sizeof szText1,0,0 
							invoke SysFreeString,bstr
						.endif
						invoke vf(pTypeInfoRef,ITypeInfo,GetTypeAttr), addr pTypeAttr2
						.if (eax == S_OK)
							mov ecx,pTypeAttr2
							invoke StringFromGUID2, addr [ecx].TYPEATTR.guid, addr wszIID, 40
							invoke wsprintf,addr szText2,CStr("%S"),addr wszIID
							invoke vf(pTypeInfoRef,ITypeInfo,ReleaseTypeAttr), pTypeAttr2
						.endif
						invoke vf(pTypeInfoRef,ITypeInfo,Release)
					.endif
				.endif

				mov lvi.mask_,LVIF_TEXT or LVIF_PARAM
				mov lvi.lParam, ebx
				mov lvi.iSubItem,0

				lea eax,szText1
				mov lvi.pszText,eax
				invoke ListView_InsertItem( m_hWndLV,addr lvi)
				inc lvi.iSubItem

				mov lvi.mask_,LVIF_TEXT
				lea eax,szText2
				mov lvi.pszText,eax
				invoke ListView_SetItem( m_hWndLV,addr lvi)
				inc lvi.iSubItem

				lea eax,szText3
				mov lvi.pszText,eax
				invoke ListView_SetItem( m_hWndLV,addr lvi)
				inc lvi.iSubItem

				lea eax,szText4
				mov lvi.pszText,eax
				invoke ListView_SetItem( m_hWndLV,addr lvi)

				inc lvi.iItem
				inc ebx
				dec dwCnt

			.endw
			invoke vf(m_pTypeInfo,ITypeInfo,ReleaseTypeAttr),pTypeAttr
		.endif

		ret
		align 4

InterfacesOnInitDialog endp


InterfacesDialog proc uses __this thisarg, message:dword, wParam:WPARAM, lParam:LPARAM

		mov __this,this@

		mov eax,message
		.if (eax == WM_INITDIALOG)

			invoke GetDlgItem, m_hWnd, IDC_LIST1
			mov m_hWndLV,eax
			invoke SetChildDlgPos, m_hWndTab, m_hWnd, m_hWndLV
			invoke InterfacesOnInitDialog

			mov eax,1

		.elseif (eax == WM_NOTIFY)

			invoke InterfacesOnNotify, lParam

		.elseif (eax == WM_COMMAND)

			movzx eax,word ptr wParam
			.if (eax == IDM_VIEW)
				invoke ListView_GetNextItem( m_hWndLV, -1, LVNI_SELECTED)
				.if (eax != -1)	
					invoke StartNewDialog, m_pTypeInfo, eax
				.else
					invoke MessageBeep, MB_OK
				.endif
			.else
				invoke GetParent, m_hWnd
				invoke SendMessage, eax, message, wParam, lParam
			.endif

		.else
			xor eax,eax ;indicates "no processing"
		.endif
		ret
		align 4

InterfacesDialog endp

;-------------------------------------- new class -------------------------------

BEGIN_CLASS CTypeInfoDlg
Frame		CDlg <>
dwRim		dword ?				;height of lower rim
iTabIndex	dword ?				;should be last member before CTypeInfoPageDlg
			CTypeInfoPageDlg <>
END_CLASS

__this	textequ <edi>
_this	textequ <[__this].CTypeInfoDlg>
thisarg	textequ <this@:ptr CTypeInfoDlg>

Destroy@CTypeInfoDlg	proto thisarg
OnInitDialog			proto
OnNotify				proto pNMHDR:ptr NMHDR
SelectTabDialog			proto iIndex:dword

	MEMBER Frame, iTabIndex, dwRim
	MEMBER hWnd, pDlgProc
	MEMBER hWndLV, hWndSB, hWndTab, pLParam
	MEMBER pTypeInfo, pTypeLib, iIndex, iType, iSortCol, iSortDir
	MEMBER wFlags, bInitialized

SelectTabDialog proc iIndex:dword


		mov eax,iIndex
		mov ecx,sizeof CTabDlgPage
		mul ecx
		add eax,offset TabDlgPages

		mov ecx,[eax].CTabDlgPage.pDlgProc
		mov m_pDlgProc,ecx
		mov ecx,iIndex
		mov m_iType, ecx
		mov ecx,[eax].CTabDlgPage.dwResID
		.if (m_hWnd == NULL)
;------------------------------- calc this_ for page dialog object
			lea eax, m_iTabIndex + sizeof CTypeInfoDlg.iTabIndex
			invoke CreateDialogParam, g_hInstance, ecx, m_Frame.hWnd, classdialogproc, eax
			invoke ListView_SetExtendedListViewStyle( m_hWndLV,	LVS_EX_FULLROWSELECT or LVS_EX_HEADERDRAGDROP or LVS_EX_INFOTIP)
		.else
			invoke SetWindowRedraw( m_hWndLV, FALSE)
;------------------------------- calc this_ for page dialog object
			lea eax, m_iTabIndex + sizeof CTypeInfoDlg.iTabIndex
			invoke SendMessage, m_hWnd, WM_INITDIALOG, 0, eax
			invoke SetWindowRedraw( m_hWndLV, TRUE)
		.endif
		mov eax, iIndex
		mov m_iTabIndex, eax

		ret
		align 4

SelectTabDialog endp

	.const
BtnTab dd IDCANCEL
NUMBUTTONS textequ %($ - BtnTab) / sizeof DWORD
	.code

OnSize proc uses ebx esi hWnd:HWND, dwType:dword, dwWidth:dword, dwHeight:dword

local dwRim:DWORD
local dwHeightBtn:DWORD
local dwWidthBtn:DWORD
local dwXPos:DWORD
local dwYPos:DWORD
local dwAddX:DWORD
local dwHeightSB:DWORD
local dwHeightLV:DWORD
local rect:RECT
local hWndBtn[NUMBUTTONS]:HWND

	invoke GetWindowRect, m_hWndSB, addr rect
	mov eax, rect.bottom
	sub eax, rect.top
	mov dwHeightSB, eax

	invoke GetWindowRect, m_hWndTab, addr rect
	invoke ScreenToClient, hWnd, addr rect
	mov eax, rect.left
	mov dwRim, eax

	shl eax, 1
	sub dwWidth, eax

	@mov dwWidthBtn, 0
	mov esi, offset BtnTab
	lea ebx, hWndBtn
	mov ecx, NUMBUTTONS
	.while (ecx)
		push ecx
		lodsd
		invoke GetDlgItem, hWnd, eax
		mov [ebx], eax
		add ebx, sizeof HWND
		lea ecx, rect
		invoke GetWindowRect, eax, ecx
		mov eax, rect.right
		sub eax, rect.left
		add dwWidthBtn, eax
		pop ecx
		mov eax, rect.bottom
		sub eax, rect.top
		mov dwHeightBtn, eax
		dec ecx
	.endw

	invoke BeginDeferWindowPos, 2 + NUMBUTTONS
	mov ebx, eax

	mov eax, dwHeight
	sub eax, dwRim
	sub eax, m_dwRim
	mov dwHeightLV, eax
	test eax, eax
	.if (SIGN?)
		@mov dwHeightLV, 0
	.endif
	invoke DeferWindowPos, ebx, m_hWndTab, NULL, 0, 0, dwWidth, dwHeightLV, SWP_NOMOVE or SWP_NOZORDER or SWP_NOACTIVATE

	mov eax, m_dwRim
	sub eax, dwHeightSB
	sub eax, dwHeightBtn
	shr eax, 1
	add eax, dwHeightLV
	add eax, dwRim
	mov dwYPos, eax

	mov eax, dwWidth
	sub eax, dwWidthBtn
	shr eax, 1
	mov dwXPos, eax

	lea esi, hWndBtn
	mov ecx, NUMBUTTONS
	.while (ecx)
		push ecx
		lodsd
		push eax
		lea ecx, rect
		invoke GetWindowRect, eax, ecx
		pop eax
		invoke DeferWindowPos, ebx, eax, NULL, dwXPos, dwYPos, 0, 0, SWP_NOSIZE or SWP_NOZORDER or SWP_NOACTIVATE
		mov eax, rect.right
		sub eax, rect.left
		pop ecx
		dec ecx
	.endw

	invoke DeferWindowPos, ebx, m_hWndSB, NULL, 0, 0, 0, 0, SWP_NOZORDER or SWP_NOACTIVATE

	invoke EndDeferWindowPos, ebx

	invoke SetChildDlgPos, m_hWndTab, m_hWnd, m_hWndLV
	ret

OnSize endp


OnNotify proc uses ebx pNMHDR:ptr NMHDR

local	point:POINT
local	hImage:HANDLE

		mov ebx,pNMHDR
		assume ebx:ptr NMHDR

		mov eax,[ebx].code
		.if (eax == TCN_SELCHANGE)
			invoke TabCtrl_GetCurSel( m_hWndTab)
			invoke SelectTabDialog, eax
		.elseif (eax == TCN_SELCHANGING)
			.if (m_hWnd != NULL)
;;				invoke SendMessage, m_hWnd, WM_COMMAND, IDOK, 0
				StatusBar_SetText m_hWndSB, 0, addr g_szNull
				StatusBar_SetText m_hWndSB, 1, addr g_szNull
if ?DESTROYDLG
				invoke DestroyWindow,m_hWnd
				mov m_hWnd,NULL
else
				invoke SetWindowRedraw( m_hWndLV,FALSE)
;--------------------------------------- always delete items before colums
				invoke ListView_DeleteAllItems( m_hWndLV)
				.repeat
					invoke ListView_DeleteColumn( m_hWndLV,0)
				.until (eax == FALSE)
				invoke SetWindowRedraw( m_hWndLV,TRUE)
endif
				mov m_bInitialized, FALSE
			.endif	
		.endif

		assume ebx:nothing

		ret
		align 4

OnNotify endp

SetWindowCaption proc

local	bstr:BSTR
local	bstr2:BSTR
local	szText[256]:byte
local	szName[80]:byte
local	szDoc[128]:byte

		mov bstr, NULL
		mov bstr2, NULL
		invoke lstrcpy, addr szName, CStr("<noname>")
		mov szDoc, 0
		invoke vf(m_pTypeInfo,ITypeInfo,GetDocumentation),MEMBERID_NIL, addr bstr, addr bstr2, NULL, NULL
		.if (eax == S_OK)
			.if (bstr)
				invoke WideCharToMultiByte, CP_ACP, 0, bstr, -1, addr szName, sizeof szName, 0, 0
				invoke SysFreeString,bstr
			.endif
			.if (bstr2)
				invoke WideCharToMultiByte, CP_ACP, 0, bstr2, -1, addr szDoc, sizeof szDoc, 0, 0
				invoke SysFreeString,bstr2
			.endif
		.endif
		invoke wsprintf,addr szText, CStr("TypeInfo %s [%s]"), addr szName, addr szDoc
		invoke SetWindowText, m_Frame.hWnd, addr szText
		ret
		align 4
SetWindowCaption endp

;--- WM_COMMAND

OnCommand proc wParam:WPARAM, lParam:LPARAM

		movzx eax, word ptr wParam
		.if (eax == IDCANCEL)

			invoke PostMessage,m_Frame.hWnd,WM_CLOSE,0,0

		.elseif (eax == IDOK)							;no OK button available
if 0
			invoke MessageBeep,MB_OK					;so beep (user pressed RETURN)
else
			invoke GetFocus								;or send message to dlg control
			.if (eax == m_hWndLV)
				sub esp,sizeof NMHDR
				mov [esp].NMHDR.hwndFrom,eax
				mov [esp].NMHDR.idFrom,IDC_LIST1
				mov [esp].NMHDR.code,NM_RETURN
				mov eax,esp
				invoke SendMessage, m_hWnd, WM_NOTIFY, IDC_LIST1, eax
				add esp,sizeof NMHDR
			.endif
endif
		.elseif (eax == IDM_SELECTALL)

			ListView_SetItemState m_hWndLV, -1, LVIS_SELECTED, LVIS_SELECTED

		.elseif (eax == IDM_COPY)

			invoke Create@CProgressDlg, m_hWndLV, NULL, SAVE_CLIPBOARD, -1
			invoke DialogBoxParam, g_hInstance, IDD_PROGRESSDLG, m_hWnd, classdialogproc, eax

		.endif
		ret
		align 4
OnCommand endp


;--- WM_INITDIALOG 


OnInitDialog proc uses ebx

local	pTypeAttr:ptr TYPEATTR
local	tci:TC_ITEM
local	dwWidth[2]:dword
local	rect:RECT
local	szText[64]:byte
ifdef _DEBUG
local	this_:ptr CTypeInfoDlg
		mov this_, __this
endif

		invoke GetDlgItem, m_Frame.hWnd, IDC_TAB1
		mov m_hWndTab,eax
		invoke GetWindowRect, m_hWndTab, addr rect
		invoke ScreenToClient, m_Frame.hWnd, addr rect.right
		push rect.bottom
		invoke GetClientRect, m_Frame.hWnd, addr rect
		pop ecx
		mov eax, rect.bottom
		sub eax, ecx
		mov m_dwRim, eax

		invoke GetDlgItem,m_Frame.hWnd,IDC_STATUSBAR
		mov m_hWndSB,eax

		invoke GetClientRect, m_hWndSB, addr rect
		mov eax, rect.right
		shr eax, 2
		mov ecx, eax
		shl eax, 1
		add eax, ecx
		mov dwWidth[0*sizeof DWORD], eax
		mov dwWidth[1*sizeof DWORD], -1
		StatusBar_SetParts m_hWndSB, 2, addr dwWidth

		mov m_hWnd,NULL

		mov tci.mask_,TCIF_TEXT or TCIF_PARAM
		mov ebx,offset TabDlgPages
		mov ecx,0
		.while (ecx < NUMDLGS)
			push ecx
			mov tci.lParam,ebx
			mov eax,[ebx].CTabDlgPage.pTabName
			mov tci.pszText,eax
			invoke TabCtrl_InsertItem( m_hWndTab,ecx,addr tci)
			add ebx,sizeof CTabDlgPage
			pop ecx
			inc ecx
		.endw

		invoke vf(m_pTypeInfo,ITypeInfo,GetTypeAttr),addr pTypeAttr
		.if (eax == S_OK)
			mov ebx,pTypeAttr
			assume ebx:ptr TYPEATTR

			.if ([ebx].typekind == TKIND_ALIAS)
				invoke lstrcpy, addr szText, CStr("Alias for ")
				invoke GetParameterType, m_pTypeInfo, addr [ebx].tdescAlias, addr szText+10, sizeof szText-10
				StatusBar_SetText m_hWndSB, 1, addr szText
			.endif

			.if (m_iTabIndex == -1)
				.if ([ebx].typekind == TKIND_COCLASS)
					mov m_iTabIndex, TAB_INTERFACES
				.elseif (([ebx].typekind == TKIND_ENUM) || \
						([ebx].typekind == TKIND_RECORD) || \
						([ebx].typekind == TKIND_UNION))
					mov m_iTabIndex, TAB_VARIABLES
				.elseif ([ebx].typekind == TKIND_MODULE)
					.if ([ebx].cFuncs)
						mov m_iTabIndex, TAB_FUNCTIONS
					.else
						mov m_iTabIndex, TAB_VARIABLES
					.endif
				.else
					mov m_iTabIndex, TAB_FUNCTIONS
				.endif
			.endif

			invoke vf(m_pTypeInfo,ITypeInfo,ReleaseTypeAttr), ebx
		.endif

		invoke SetWindowCaption
		invoke TabCtrl_SetCurSel( m_hWndTab, m_iTabIndex)
		invoke SelectTabDialog, m_iTabIndex
		ret
		align 4

OnInitDialog endp


;*** Dialog Proc for "typeinfo" dialog


TypeInfoDialog proc uses ebx __this thisarg, message:dword,wParam:WPARAM,lParam:LPARAM

local pTypeLib:LPTYPELIB
local dwIndex:DWORD
local rect:RECT

		mov __this,this@

		mov eax,message
		.if (eax == WM_INITDIALOG)

			invoke OnInitDialog
			.if (m_wFlags & TYPEINFODLG_FTILE)
				invoke GetWindow, m_Frame.hWnd, GW_OWNER
				mov ecx, eax
				invoke GetWindowRect, ecx, addr rect
				add rect.left, 20
				add rect.top, 20
				invoke SetWindowPos, m_Frame.hWnd, NULL, rect.left, rect.top, 0, 0,\
					SWP_NOZORDER or SWP_NOACTIVATE or SWP_NOSIZE
			.else
				invoke CenterWindow, m_Frame.hWnd
			.endif
if ?MODELESS
			.if (m_wFlags & TYPEINFODLG_FACTIVATE)
				invoke ShowWindow, m_Frame.hWnd, SW_SHOWNORMAL
			.endif
endif
			mov eax,1

		.elseif (eax == WM_CLOSE)

if ?MODELESS
			invoke DestroyWindow, m_Frame.hWnd
else
			invoke EndDialog, m_Frame.hWnd, 0
endif
			mov eax,1

		.elseif (eax == WM_DESTROY)

			invoke Destroy@CTypeInfoDlg, __this
if ?MODELESS
		.elseif (eax == WM_ACTIVATE)

			movzx eax,word ptr wParam
			.if (eax == WA_INACTIVE)
				mov g_hWndDlg, NULL
			.else
				mov eax, m_Frame.hWnd
				mov g_hWndDlg, eax
			.endif
endif
		.elseif (eax == WM_SIZE)

			.if (wParam != SIZE_MINIMIZED)
				movzx eax, word ptr lParam+0
				movzx ecx, word ptr lParam+2
				invoke OnSize, m_Frame.hWnd, wParam, eax, ecx
			.endif

		.elseif (eax == WM_COMMAND)

			invoke OnCommand, wParam, lParam

		.elseif (eax == WM_NOTIFY)

			invoke OnNotify, lParam
if ?DRAWITEMSB
		.elseif (eax == WM_DRAWITEM)

			.if (wParam == IDC_STATUSBAR)
				push esi
				mov esi, lParam
				invoke SetTextColor, [esi].DRAWITEMSTRUCT.hDC, 000000C0h
				invoke SetBkMode, [esi].DRAWITEMSTRUCT.hDC, TRANSPARENT
				add [esi].DRAWITEMSTRUCT.rcItem.left, 4
				invoke DrawTextEx, [esi].DRAWITEMSTRUCT.hDC,\
					[esi].DRAWITEMSTRUCT.itemData, -1, addr [esi].DRAWITEMSTRUCT.rcItem,\
					DT_LEFT or DT_SINGLELINE or DT_VCENTER, NULL
				pop esi
			.endif
			mov eax, 1
endif
if ?HTMLHELP
		.elseif (eax == WM_HELP)

			invoke DoHtmlHelp, HH_DISPLAY_TOPIC, CStr("typeinfodialog.htm")
endif
		.else
			xor eax,eax ;indicates "no processing"
		.endif
		ret
		align 4

TypeInfoDialog endp


;*** constructor

Create@CTypeInfoDlg proc public uses __this pTypeLib:LPTYPELIB, iIndex:dword

		invoke malloc, sizeof CTypeInfoDlg
		.if (!eax)
			ret
		.endif
		mov __this,eax
		mov m_Frame.pDlgProc,TypeInfoDialog
		mov m_iTabIndex,TAB_FUNCTIONS

		mov m_iItem, -1
		mov m_bFromTypeLib, TRUE

		invoke vf(pTypeLib, ITypeLib, GetTypeInfo), iIndex, addr m_pTypeInfo
		.if (!m_pTypeInfo)
			invoke Destroy@CTypeInfoDlg, __this
			return NULL
		.endif

		return __this
		align 4

Create@CTypeInfoDlg endp

Create2@CTypeInfoDlg proc public uses __this pTypeInfo:LPTYPEINFO

local pTypeAttr:ptr TYPEATTR

		invoke malloc, sizeof CTypeInfoDlg
		.if (!eax)
			ret
		.endif
		mov __this,eax
		mov m_Frame.pDlgProc,TypeInfoDialog

		mov ecx, pTypeInfo
		mov m_pTypeInfo,ecx
		invoke vf(m_pTypeInfo, ITypeInfo, AddRef)

		invoke vf(m_pTypeInfo, ITypeInfo, GetTypeAttr), addr pTypeAttr
		.if (eax == S_OK)
			mov eax, pTypeAttr
			.if ([eax].TYPEATTR.typekind == TKIND_ENUM)
				mov m_iTabIndex, TAB_VARIABLES
			.elseif ([eax].TYPEATTR.typekind == TKIND_COCLASS)
				mov m_iTabIndex, TAB_INTERFACES
			.endif
			invoke vf(m_pTypeInfo, ITypeInfo, ReleaseTypeAttr), pTypeAttr
		.endif

		mov m_iItem, -1

		return __this
		align 4

Create2@CTypeInfoDlg endp


Create3@CTypeInfoDlg proc public uses __this refIID:REFIID, refTLBID:REFGUID, dwVerMajor:DWORD, dwVerMinor:DWORD

local	pTypeLib:LPTYPELIB

		invoke malloc, sizeof CTypeInfoDlg
		.if (!eax)
			ret
		.endif
		mov __this,eax
		mov m_Frame.pDlgProc,TypeInfoDialog

		invoke LoadRegTypeLib, refTLBID, dwVerMajor, dwVerMinor, g_LCID,addr pTypeLib
		.if (eax == S_OK)
			invoke vf(pTypeLib, ITypeLib, GetTypeInfoOfGuid), refIID, addr m_pTypeInfo
			invoke vf(pTypeLib, IUnknown, Release)
		.endif

		mov m_iItem, -1

		.if (!m_pTypeInfo)
			invoke Destroy@CTypeInfoDlg, __this
			return NULL
		.endif

		return __this
		align 4

Create3@CTypeInfoDlg endp

Destroy@CTypeInfoDlg proc public uses __this thisarg

		mov __this, this@

		.if (m_Frame.hWnd)
			invoke BroadCastMessage, WM_WNDDESTROYED, 0, m_Frame.hWnd
		.endif
		.if (m_pLParam)
			invoke free, m_pLParam
			mov m_pLParam, NULL
		.endif
		.if (m_pTypeInfo)
			invoke vf(m_pTypeInfo, ITypeInfo, Release)
		.endif
;------------------------ do this manually
		.if (m_hWnd)
			invoke SetWindowLong, m_hWnd, DWL_USER, NULL
		.endif
		invoke free, __this
		ret
		align 4

Destroy@CTypeInfoDlg endp


Show@CTypeInfoDlg proc public thisarg, hWnd:HWND, dwFlags:DWORD

		mov ecx, this@
		mov eax, dwFlags
		mov [ecx].CTypeInfoDlg.wFlags, ax
if ?MODELESS
		invoke CreateDialogParam, g_hInstance, IDD_TYPEINFODLG,\
			hWnd, classdialogproc, ecx
else
		invoke DialogBoxParam,g_hInstance,IDD_TYPEINFODLG,\
			hWnd, classdialogproc, ecx
endif
		ret
		align 4

Show@CTypeInfoDlg endp



SetTab@CTypeInfoDlg proc public thisarg, dwTabIndex:DWORD

		mov eax, this@
		mov ecx, dwTabIndex
		mov [eax].CTypeInfoDlg.iTabIndex,ecx
		ret
		align 4

SetTab@CTypeInfoDlg endp

if 0
SetWindowPos@CTypeInfoDlg proc public thisarg, ptPos:ptr POINT
		mov ecx, this@
		mov edx, ptPos
		mov eax, [edx].POINT.x
		mov [ecx].CTypeInfoDlg.ptPos.x,eax
		mov eax, [edx].POINT.y
		mov [ecx].CTypeInfoDlg.ptPos.y,eax
		ret
		align 4
SetWindowPos@CTypeInfoDlg endp
endif

FindFunc@CTypeInfoDlg proc public uses __this thisarg, dwID:MEMBERID, invkind:INVOKEKIND

local	iItem:DWORD
local	pFuncDesc:ptr FUNCDESC
local	lvfi:LVFINDINFO
local	nmhdr:NMHDR

		mov __this, this@
		.if (m_iTabIndex != TAB_FUNCTIONS)
			mov nmhdr.code, TCN_SELCHANGING
			invoke OnNotify, addr nmhdr
			invoke TabCtrl_SetCurSel( m_hWndTab, TAB_FUNCTIONS)
			invoke SelectTabDialog, TAB_FUNCTIONS
		.endif
		.if (!m_pLParam)
			jmp exit
		.endif
		invoke ListView_GetItemCount( m_hWndLV)
		mov ecx, eax
		mov edx, m_pLParam
		mov eax, dwID
		@mov iItem, 0
		.while (ecx)
			.if (eax == [edx].LPARAMSTRUCT.memid)
				pushad
				mov ecx, edx
				invoke vf(m_pTypeInfo, ITypeInfo, GetFuncDesc), [ecx].LPARAMSTRUCT.iIndex, addr pFuncDesc
				clc
				.if (eax == S_OK)
					mov ecx, pFuncDesc
					push [ecx].FUNCDESC.invkind
					invoke vf(m_pTypeInfo, ITypeInfo, ReleaseFuncDesc), pFuncDesc
					pop ecx
					clc
					.if (ecx == invkind)
						stc
					.endif
				.endif
				popad
				.break .if CARRY?
			.endif
			inc iItem
			add edx, sizeof LPARAMSTRUCT
			dec ecx
		.endw
		.if (!ecx)
			jmp exit
		.endif
		mov iItem, -1
		mov lvfi.flags, LVFI_PARAM 
		mov lvfi.lParam, edx
		invoke ListView_FindItem( m_hWndLV, -1, addr lvfi)
		.if (eax != -1)
			mov iItem, eax
			invoke ListView_EnsureVisible( m_hWndLV, iItem, FALSE)
			mov eax, -1
			.while (1)
				invoke ListView_GetNextItem( m_hWndLV, eax, LVNI_SELECTED)
				.break .if (eax == -1)
				push eax
				ListView_SetItemState m_hWndLV, eax, 0, LVIS_SELECTED
				pop eax
			.endw
			ListView_SetItemState m_hWndLV, iItem, LVIS_SELECTED, LVIS_SELECTED
		.endif
exit:
		ret
		align 4

FindFunc@CTypeInfoDlg endp

FindVar@CTypeInfoDlg proc public uses __this thisarg, dwID:MEMBERID

local	iItem:DWORD
local	pVarDesc:ptr VARDESC
local	lvfi:LVFINDINFO
local	nmhdr:NMHDR

		mov __this, this@
		.if (m_iTabIndex != TAB_VARIABLES)
			mov nmhdr.code, TCN_SELCHANGING
			invoke OnNotify, addr nmhdr
			invoke TabCtrl_SetCurSel( m_hWndTab, TAB_VARIABLES)
			invoke SelectTabDialog, TAB_VARIABLES
		.endif
		.if (!m_pLParam)
			jmp exit
		.endif
		invoke ListView_GetItemCount( m_hWndLV)
		mov ecx, eax
		mov edx, m_pLParam
		mov eax, dwID
		@mov iItem, 0
		.while (ecx)
			.if (eax == [edx].LPARAMSTRUCT.memid)
				pushad
				mov ecx, edx
				invoke vf(m_pTypeInfo, ITypeInfo, GetVarDesc), [ecx].LPARAMSTRUCT.iIndex, addr pVarDesc
				clc
				.if (eax == S_OK)
					invoke vf(m_pTypeInfo, ITypeInfo, ReleaseVarDesc), pVarDesc
					stc
				.endif
				popad
				.break .if CARRY?
			.endif
			inc iItem
			add edx, sizeof LPARAMSTRUCT
			dec ecx
		.endw
		.if (!ecx)
			jmp exit
		.endif
		mov iItem, -1
		mov lvfi.flags, LVFI_PARAM 
		mov lvfi.lParam, edx
		invoke ListView_FindItem( m_hWndLV, -1, addr lvfi)
		.if (eax != -1)
			mov iItem, eax
			invoke ListView_EnsureVisible( m_hWndLV, iItem, FALSE)
			mov eax, -1
			.while (1)
				invoke ListView_GetNextItem( m_hWndLV, eax, LVNI_SELECTED)
				.break .if (eax == -1)
				push eax
				ListView_SetItemState m_hWndLV, eax, 0, LVIS_SELECTED
				pop eax
			.endw
			ListView_SetItemState m_hWndLV, iItem, LVIS_SELECTED, LVIS_SELECTED
		.endif
exit:
		ret
		align 4

FindVar@CTypeInfoDlg endp

SetTypeInfo@CTypeInfoDlg proc public uses __this thisarg, pTypeInfo:LPTYPEINFO

local	nmhdr:NMHDR

		mov __this, this@
		mov eax, m_pTypeInfo
		.if (eax == pTypeInfo)
			return 0
		.endif
		.if (eax)
;------------------------------- simulate a WM_NOTIFY to clear listview
			.if (m_iTabIndex != TAB_FUNCTIONS)
				mov nmhdr.code, TCN_SELCHANGING
				invoke OnNotify, addr nmhdr
			.endif
			invoke vf(m_pTypeInfo, ITypeInfo, Release)
		.endif
		mov eax, pTypeInfo
		mov m_pTypeInfo, eax
		invoke vf(m_pTypeInfo, ITypeInfo, AddRef)
		invoke SetWindowCaption
		invoke TabCtrl_SetCurSel( m_hWndTab, TAB_FUNCTIONS)
		invoke SelectTabDialog, TAB_FUNCTIONS
		return 1
		align 4

SetTypeInfo@CTypeInfoDlg endp

	end
