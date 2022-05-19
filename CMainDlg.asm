
;*** definition of CMainDlg (main dialog) methods 
;*** this class handles all WM_COMMAND/IDM_xxx messages

;*** this dialog display a tab control with a listview in it,
;*** and a statusbar is placed below
;*** the listview is a virtual listview
;*** data is stored in class CDocument, which is rebuild on tab changes

	.386
	.model flat,stdcall
	option casemap:none   ; case sensitive
	option proc:private

INSIDE_CMAINDLG equ 1

	include COMView.inc
	include shellapi.inc
	include shlobj.inc
	include dde.inc
	include statusbar.inc


	include classes.inc
	include rsrc.inc
	include CEditDlg.inc
	include debugout.inc

?MULTIDOC		equ 1
?DRAWITEMSB		equ 1
?SAVECOLORDER	equ 1
?STATEIMAGE		equ 0	;???
;;?SUPPLICENSEKEY	equ 1
?DROPSOURCE		equ 1	;listview may be a drop source
?DDESUPPORT		equ 1	;comview supports DDE


CMainView	struct
pDoc		pCDocument ?
iSortCol	DWORD ?
iSortDir	DWORD ?
iTopIndex	DWORD ?
if ?SAVECOLORDER
dwColumns	DWORD ?
pdwColOrder	LPVOID ?
endif
CMainView	ends

BEGIN_CLASS CMainDlg, CDlg
hWndTab		HWND ?
hWndLV		HWND ?
hWndHdr		HWND ?
hMenu		HMENU ?
iMode		dword ?			;view mode
pMode		pCMode ?
iNumCols	dword ?
pszRoot		LPSTR ?
    		CMainView <>
pDropTarget LPDROPTARGET ?
if ?MULTIDOC
savedView	CMainView MODE_MAX+1 dup (<>)
endif
END_CLASS

STATUSCLASSNAME equ <CStr("msctls_statusbar32")>
EDITMODELESS	equ 1	;1 = edit registry dialog is modeless dialog
MAXFINDLEN		equ 80
?LVSTYLE		equ LVS_EX_FULLROWSELECT or LVS_EX_HEADERDRAGDROP or LVS_EX_INFOTIP

CO_S_NOTALLINTERFACES	equ 00080012h	;value not defined in windows.inc

	.const

g_szCLSID	db "CLSID", 0
g_szTypeLib	db "TypeLib", 0
g_szInterface db "Interface", 0

g_dwSBParts dd 50, 15, 20

	.data

g_hAccel	HACCEL 0
g_hAccel2	HACCEL 0
g_hWndSB    HWND 0
g_hWndEditRoot HWND 0
g_hLibRE	HINSTANCE 0
g_uMsgFind	DWORD 0
g_hWndFind	HWND 0
g_lpstrFind	LPSTR 0
g_pszFilename LPSTR NULL
g_dwTimer	DWORD NULL
g_hConOut	DWORD -1
if ?DDESUPPORT
g_aApplication	DWORD 0
g_aSystem		DWORD 0
endif
g_szMachine	BYTE MAXINPUTTEXT dup (0)
g_bAcceptDrop BOOLEAN TRUE
g_bMenuLoop BOOLEAN FALSE
g_bColumnsChanged BOOLEAN FALSE
g_bFirstOut	BOOLEAN FALSE

;*** define header colums of all listview modes

ColumnsCLSID label CColHdr
	CColHdr <g_szCLSID			, 15>
	CColHdr <CStr("Text")		, 25>
	CColHdr <CStr("Type")		, 10>
	CColHdr <CStr("Type Value")	, 20>
	CColHdr <CStr("ProgID")		, 15>
	CColHdr <g_szTypeLib		, 15>
NUMCOLS_CLSID equ ($ - ColumnsCLSID) / sizeof CColHdr
CtrlColCLSID CColHdr <0					, 15>			; place for user defined column

CLSIDCOL_IN_CLSID	equ 0
TYPECOL_IN_CLSID	equ 2
PATHCOL_IN_CLSID	equ 3
PROGIDCOL_IN_CLSID	equ 4
TYPELIBCOL_IN_CLSID	equ 5

ColumnsTypeLib label CColHdr
	CColHdr <g_szTypeLib		,25>
	CColHdr <CStr("Text")		,31>
	CColHdr <CStr("Path (win32)"),30>
	CColHdr <CStr("Version")	,7>
	CColHdr <CStr("LCID")		,7,		FCOLHDR_RDX16>	;numeric + hex
NUMCOLS_TYPELIB equ ($ - ColumnsTypeLib) / sizeof CColHdr

PATHCOL_IN_TYPELIB equ 2

ColumnsInterface label CColHdr
	CColHdr <CStr("IID")		,20>
	CColHdr <CStr("Text")		,30>
	CColHdr <CStr("ProxyStubClsid32"),30>
	CColHdr <g_szTypeLib		,20>
NUMCOLS_INTERFACE equ ($ - ColumnsInterface) / sizeof CColHdr
CtrlColInterface CColHdr <0				, 15>			; place for user defined column

CLSIDCOL_IN_INTERFACE	equ 2
TYPELIBCOL_IN_INTERFACE equ 3

ColumnsAppID label CColHdr
	CColHdr <CStr("AppID")		,20>
	CColHdr <CStr("Text")		,20>
	CColHdr <CStr("[AppID]")	,20>
	CColHdr <CStr("[AuthLvl]")	,7>
	CColHdr <CStr("[RunAs]")	,13>
	CColHdr <CStr("[DllSurrogate]")	,10>
	CColHdr <CStr("[LocalService]")	,10>
NUMCOLS_APPID equ ($ - ColumnsAppID) / sizeof CColHdr

APPIDCOL_IN_APPID	equ 2

ColumnsCompCat label CColHdr
	CColHdr <CStr("GUID")		,50>
	CColHdr <CStr("Text")		,50>
NUMCOLS_COMPCAT equ ($ - ColumnsCompCat) / sizeof CColHdr

ColumnsHKCR label CColHdr
	CColHdr <CStr("Key")				,25>
	CColHdr <CStr("Value")				,25>
	CColHdr <g_szCLSID					,25>
	CColHdr <CStr("Shell\Open\Command")	,25>
NUMCOLS_HKCR equ ($ - ColumnsHKCR) / sizeof CColHdr
CtrlColHKCR CColHdr <0					, 15>			; place for user defined column

CLSIDCOL_IN_HKCR equ 2
PATHCOL_IN_HKCR equ 3

ColumnsObject label CColHdr
	CColHdr <g_szCLSID			,30>
	CColHdr <CStr("Text")		,35>
	CColHdr <CStr("IMoniker Display Name"),30>
	CColHdr <CStr("Connections"),5>
NUMCOLS_OBJECT equ ($ - ColumnsObject) / sizeof CColHdr

ColumnsROT label CColHdr
		CColHdr <g_szCLSID			,25>
		CColHdr <CStr("DisplayName"),25>
		CColHdr <CStr("Type")		,15>
		CColHdr <CStr("Running")	,10>
		CColHdr <CStr("Text")		,25>
NUMCOLS_ROT equ ($ - ColumnsROT) / sizeof CColHdr


;*** this table describes the modes

ModeDesc label CMode
	CMode {MODE_CLSID,		ColumnsCLSID,	NUMCOLS_CLSID,	IDM_CLSID,	g_szRootCLSID,	g_szCLSID}
	CMode {MODE_TYPELIB,	ColumnsTypeLib,	NUMCOLS_TYPELIB,IDM_TYPELIB,g_szRootTypeLib,g_szTypeLib}
	CMode {MODE_INTERFACE,	ColumnsInterface,NUMCOLS_INTERFACE,IDM_INTERFACE,g_szRootInterface, g_szInterface}
	CMode {MODE_APPID,		ColumnsAppID,	NUMCOLS_APPID,	IDM_APPID,	g_szRootAppID,	CStr("AppID")}
	CMode {MODE_COMPCAT,	ColumnsCompCat,	NUMCOLS_COMPCAT,IDM_COMPCAT,g_szRootCompCat,CStr("Component Category")}
	CMode {MODE_HKCR,		ColumnsHKCR,	NUMCOLS_HKCR,	IDM_HKCR,	NULL,			CStr("HKCR")}
	CMode {MODE_OBJECT,		ColumnsObject,	NUMCOLS_OBJECT,	IDM_OBJECT,	g_szRootCLSID,	CStr("Created Objects")}
	CMode {MODE_ROT,		ColumnsROT,		NUMCOLS_ROT,	IDM_ROT,	g_szRootCLSID,	CStr("ROT")}
NUMMODES equ ($ - ModeDesc) / sizeof CMode

	.code

;*** CMainDlg methods ***

__this	textequ <edi>
_this	textequ <[__this].CMainDlg>
thisarg equ <this@:ptr CMainDlg>

;*** private members prototypes
;*** thisarg could be omitted cause "this" ptr (=edi) is always set

ResizeClients			proto dwNewSize:dword
SetMyLVColumns			proto iNewMode:DWORD
OnCheckFile				proto
OnCheckTypelib			proto
OnCheckCLSID			proto
OnCheckProgID			proto
OnCheckAppID			proto
IsFileLink				proto iItem:DWORD
RefreshView				proto iNewMode:dword
RefreshLine				proto pszItem:ptr EDITITEMDESC
OnCreateInstance		proto
OnCreateInstanceOn		proto
SetStatusPane1			proto
SetStatusPane2			proto
ShowContextMenu			proto pNMLV:ptr NMLISTVIEW, bMouse:BOOL
OnEdit					proto
OnOption				proto
OnAmbientProperties		proto
OnTypeLibDlg			proto
OnObjectDlg				proto
OnLoadTypeLib			proto
OnLoadFile				proto	;create an embedded object from a file
OnCreateLink			proto	;create a linked object from a file
OnOpenStorage			proto
OnOpenStream			proto
OnSaveAs				proto
OnCopy					proto
OnPaste					proto
OnSelectAll				proto
OnInvertSelection		proto
OnCopyGUID				proto
OnFind					proto
OnFindNextCross			proto
OnExplore				proto
if ?REMOVEITEM
OnRemoveItem			proto
endif
OnProperties			proto
OnRegister				proto
OnUnregister			proto
OnOleReg				proto
SearchString			proto pszFind:LPSTR, dwFlags:DWORD, dwItem:DWORD, iColumn:DWORD
OnFindMsg				proto :ptr FINDREPLACE
SortListView			proto
GetDispInfo				proto :ptr NMLVDISPINFO
OnNotifyLV				proto pNMLV:ptr NMLISTVIEW
OnNotify				proto pNMHDR:ptr NMHDR
OnCommand				proto wParam:WPARAM, lParam:LPARAM
OnInitDialog			proto
MainDialog				proto thisarg, :dword, wParam:WPARAM, lParam:LPARAM
BindToObject			proto pMoniker:LPMONIKER, bDisplayError:BOOL


	MEMBER hWnd, pDlgProc
	MEMBER hWndTab, hWndLV, hWndHdr, hMenu
	MEMBER pDoc, iMode, pMode, iNumCols, iSortCol, iSortDir, pszRoot
	MEMBER pDropTarget
if ?MULTIDOC
	MEMBER savedView
endif


if ?DROPSOURCE

ifdef @StackBase
	option stackbase:ebp
endif

Create@CHDrop proc uses esi ebx pdwSize:ptr DWORD

local	dwSize:DWORD
local	dwIndex:DWORD
local	dwItems:DWORD
local	dwESP:DWORD
local	pszFile:LPSTR
local	lvi:LVITEM
LOCAL	szPath[MAX_PATH]:BYTE

		mov dwSize,SIZEOF DROPFILES + 1
		mov dwIndex, -1
		mov dwItems, 0
		mov dwESP, esp
		.while (1)
			invoke ListView_GetNextItem( m_hWndLV, dwIndex, LVNI_SELECTED)
			.break .if (eax == -1)
			mov dwIndex,eax
			mov lvi.iItem,eax
			.if (m_iMode == MODE_CLSID)
				invoke IsFileLink, eax
				.continue .if (!eax)
				mov lvi.iSubItem, PATHCOL_IN_CLSID
			.elseif (m_iMode == MODE_TYPELIB)
				mov lvi.iSubItem, PATHCOL_IN_TYPELIB
			.endif
			lea eax, szPath
			mov lvi.pszText, eax
			mov lvi.cchTextMax,SIZEOF szPath
			mov lvi.mask_,LVIF_TEXT
			invoke ListView_GetItem( m_hWndLV, addr lvi)

			invoke lstrlen, addr szPath
			.if (eax)
				lea edx, szPath
				.if ((eax > 1) && (byte ptr [edx] == '"') && (byte ptr [edx+eax-1] == '"'))
					mov byte ptr [edx+eax-1],0
					inc edx
					dec eax
					dec eax
				.endif
				mov pszFile, edx
				.while (eax)
					mov ecx, [edx+eax-1]
					or ecx, 20202000h
					.if ((ecx == "exe.") && (byte ptr [edx+eax+3] == ' '))
						mov byte ptr [edx+eax+3],0
						.break
					.endif
					dec eax
				.endw
				sub esp, MAX_PATH
				invoke strchr, pszFile, '\'
				.if (!eax)
					mov eax, esp
					invoke GetSystemDirectory, eax, MAX_PATH
					add eax, esp
					mov byte ptr [eax],'\'
					inc eax
				.else
					mov eax, esp
				.endif
				invoke ExpandEnvironmentStrings, pszFile, eax, MAX_PATH
;--------------------------- WinNT bug, so dont rely on return code (MSDN Q234874)
				invoke lstrlen, esp
				inc eax
				add dwSize,eax
				inc dwItems
				.break .if (dwItems > 999)
			.endif
		.endw

		.if (dwItems)
			invoke GlobalAlloc, GPTR, dwSize
			mov ebx,eax
			.if (!eax)
				jmp exit
			.endif
			mov [ebx].DROPFILES.pFiles,SIZEOF DROPFILES
			mov [ebx].DROPFILES.fNC,0
			mov [ebx].DROPFILES.fWide,0
			push edi
			lea edi, [ebx+sizeof DROPFILES]
			mov esi, dwESP
			.while (dwItems)
				sub esi, MAX_PATH
				push esi
				DebugOut "Create@CHDrop: %s", esi
				.repeat
					lodsb
					stosb
				.until (al == 0)
				pop esi
				dec dwItems
			.endw
			mov byte ptr [edi], 0
			pop edi
			mov esp, dwESP
		.else
			xor ebx, ebx
		.endif
exit:
		mov	eax, dwSize
		mov ecx, pdwSize
		mov [ecx], eax
		return ebx
		align 4

Create@CHDrop endp

ifdef @StackBase
	option stackbase:esp
endif

OnBeginDrag proc

LOCAL	dwSize:DWORD
local	hDrop:HANDLE
LOCAL	pDropSource:LPDROPSOURCE
LOCAL	pDataObject:LPDATAOBJECT
LOCAL	deffect:DWORD

		invoke Create@CHDrop, addr dwSize
		mov hDrop, eax
		.if (eax)
			invoke Create@CDropSource
			mov pDropSource, eax
			invoke Create@CDataObject, hDrop, dwSize, CF_HDROP
			mov pDataObject, eax

			mov g_bAcceptDrop, FALSE

;;			mov deffect, DROPEFFECT_COPY or DROPEFFECT_LINK
			mov deffect, 0

			invoke DoDragDrop, pDataObject, pDropSource, DROPEFFECT_COPY or DROPEFFECT_LINK, ADDR deffect

			mov g_bAcceptDrop, TRUE

			invoke vf(pDropSource, IDropSource, Release)
			invoke vf(pDataObject, IDataObject, Release)
		.endif

		ret
		align 4

OnBeginDrag endp
endif

;*** set text "busy"/"ready" in status line part 0

SetBusyState@CMainDlg proc public uses eax dwStatus:BOOL

if ?DRAWITEMSB
?BUSYPART	equ 3 or SBT_OWNERDRAW
else
?BUSYPART	equ 3
endif
		.if (dwStatus)
			StatusBar_SetText g_hWndSB, ?BUSYPART, CStr("busy")
		.else
			StatusBar_SetText g_hWndSB, ?BUSYPART, CStr("ready")
		.endif
		ret
		align 4
SetBusyState@CMainDlg endp



SetErrorText@CMainDlg proc public uses __this thisarg, pszError:LPSTR, hr:DWORD, bBeep:BOOL

local szText[256]:byte
		mov __this, this@
		invoke wsprintf, addr szText, pszError, hr
		StatusBar_SetText g_hWndSB, 0, addr szText
		.if (bBeep)
			invoke MessageBeep, MB_OK
		.endif
		ret
		align 4
SetErrorText@CMainDlg endp

GetDefaultCommand proc
		.if (m_iMode == MODE_TYPELIB)
			mov eax, IDM_TYPELIBDLG
		.elseif ((m_iMode == MODE_OBJECT) || (m_iMode == MODE_ROT))
			mov eax, IDM_OBJECTDLG
		.else
			mov eax, IDM_EDIT
		.endif
		ret
		align 4
GetDefaultCommand endp

;*** adjust client window sizes to changed main window size

ResizeClients proc uses ebx dwNewSize:dword

local	dwCX:dword
local	dwCY:dword
local	rect:RECT
local	rect2:RECT

		invoke BeginDeferWindowPos,3
		mov ebx,eax

if 0
		mov eax,dwNewSize
		movzx ecx,ax
		mov dwCX,ecx
		shr eax,16
		mov dwCY,eax

		invoke GetChildPos, m_hWndLV
		movzx ecx,ax
		shl ecx,1
		sub dwCX,ecx

		shr eax,16
		shl eax,1
		sub dwCY,eax
		invoke DeferWindowPos,ebx, m_hWndLV,NULL,0,0,dwCX,dwCY,
				SWP_NOZORDER or SWP_NOMOVE or SWP_NOACTIVATE 
else
		invoke GetClientRect, m_hWnd, addr rect

		invoke GetClientRect, g_hWndSB, addr rect2
		mov eax,rect2.bottom
		sub eax,rect2.top
		sub rect.bottom,eax
		invoke DeferWindowPos,ebx, g_hWndSB, NULL,
				0, rect.bottom, rect.right, rect2.bottom,
				SWP_NOZORDER or SWP_NOACTIVATE 

		invoke DeferWindowPos,ebx, m_hWndTab, NULL,
				0, 0, rect.right, rect.bottom,
				SWP_NOZORDER or SWP_NOACTIVATE

;		invoke GetClientRect, m_hWndTab, addr rect2
		invoke TabCtrl_AdjustRect( m_hWndTab, FALSE, addr rect)

		mov eax,rect.left
		sub rect.right,eax
		mov eax,rect.top
		sub rect.bottom,eax

		invoke DeferWindowPos, ebx, m_hWndLV, HWND_TOP,
				rect.left,rect.top,rect.right,rect.bottom,
				SWP_NOACTIVATE 
endif
		invoke EndDeferWindowPos,ebx

		invoke SetSBParts, g_hWndSB, offset g_dwSBParts, LENGTHOF g_dwSBParts + 1

		ret
		align 4

ResizeClients endp

;*** if object dialog requires no special handling, use this wrapper

ShowObject proc pUnknown:LPUNKNOWN

	invoke Create@CObjectItem, pUnknown, NULL
	.if (eax)
		push eax
		invoke vf(eax, IObjectItem, ShowObjectDlg), m_hWnd
		pop eax
		invoke vf(eax, IObjectItem, Release)
	.endif
	ret
	align 4

ShowObject endp


;*** listview mode changes, set new listview header


SetMyLVColumns proc uses ebx esi iNewMode:DWORD

local	rect:RECT
local	dwCX:dword
local	iWidthTotal:SDWORD
local	bRefresh:BOOL
local	lvc:LVCOLUMN

		@SetLocalThis CMainDlg
;----------------------------------------- transfer user def. columns
		mov bRefresh, FALSE
		xor eax, eax
		.if (m_iMode == MODE_CLSID)
			mov ebx, NUMCOLS_CLSID
			mov esi, offset CtrlColCLSID
			.if (g_szUserColCLSID)
				mov eax, offset g_szUserColCLSID
			.endif
			.if (eax != CtrlColCLSID.pColText)
				mov CtrlColCLSID.pColText, eax
				mov bRefresh, TRUE
			.endif
		.elseif (m_iMode == MODE_HKCR)
			mov ebx, NUMCOLS_HKCR
			mov esi, offset CtrlColHKCR
			.if (g_szUserColHKCR)
				mov eax,offset g_szUserColHKCR
			.endif
			.if (eax != CtrlColHKCR.pColText)
				mov CtrlColHKCR.pColText, eax
				mov bRefresh, TRUE
			.endif
		.elseif (m_iMode == MODE_INTERFACE)
			mov ebx, NUMCOLS_INTERFACE
			mov esi, offset CtrlColInterface
			.if (g_szUserColInterface)
				mov eax,offset g_szUserColInterface
			.endif
			.if (eax != CtrlColInterface.pColText)
				mov CtrlColInterface.pColText, eax
				mov bRefresh, TRUE
			.endif
		.endif

;----------------------------------------- no new mode, but header changed (usercol)

		.if ((iNewMode == -1) && (bRefresh == FALSE))
			mov eax, [esi].CColHdr.pColText
			.if (eax)
				mov lvc.pszText,eax
				mov lvc.mask_,LVCF_TEXT
				invoke ListView_SetColumn( m_hWndLV, ebx, addr lvc)
			.endif
			jmp done
		.endif
;----------------------------------------- delete all columns from listview
		.repeat
			invoke ListView_DeleteColumn( m_hWndLV,0)
		.until (eax == 0)

;----------------------------------------- calc width of listview
		invoke GetClientRect, m_hWndLV, addr rect

		invoke GetWindowLong, m_hWndLV, GWL_STYLE
		.if !(eax & WS_VSCROLL)
			invoke GetSystemMetrics,SM_CXVSCROLL
			mov dwCX,eax
			sub rect.right,eax
		.endif
		dec rect.right

;*** copy m_iNumCols, pszRoot from CMODE to CMainDlg

		mov edx, m_pMode
		mov eax,[edx].CMode.iNumCols
		.if (CtrlColCLSID.pColText && (m_iMode == MODE_CLSID))
			inc eax
		.elseif (CtrlColHKCR.pColText && (m_iMode == MODE_HKCR))
			inc eax
		.elseif (CtrlColInterface.pColText && (m_iMode == MODE_INTERFACE))
			inc eax
		.endif
		mov m_iNumCols,eax

		mov eax,[edx].CMode.pszRoot
		mov m_pszRoot,eax

;--------------------------- get sum of iWidths values
		xor ebx, ebx
		mov esi, m_pMode
		mov esi,(CMode ptr [esi]).pColDesc
		mov iWidthTotal,ebx
		.while (ebx < m_iNumCols)
			movzx eax, [esi].CColHdr.wWidth
			.if ([esi].CColHdr.wFlags & FCOLHDR_ABSOLUTE)
				movzx eax, ax
				invoke MulDiv, eax, 100, rect.right
			.endif
			movzx eax,ax
			add iWidthTotal,eax
			inc ebx
			add esi,sizeof CColHdr
		.endw

		mov esi, m_pMode
		mov esi,(CMode ptr [esi]).pColDesc
		xor ebx, ebx
		.while (ebx < m_iNumCols)
			mov eax, [esi].CColHdr.pColText
			.if (eax)
				mov lvc.pszText,eax
				mov eax,rect.right
				movzx ecx, [esi].CColHdr.wWidth
				mov lvc.mask_,LVCF_TEXT or LVCF_WIDTH
				.if ([esi].CColHdr.wFlags & FCOLHDR_RDXMASK)
					mov lvc.mask_,LVCF_TEXT or LVCF_WIDTH or LVCF_FMT
					mov lvc.fmt, LVCFMT_RIGHT
				.endif
				.if ([esi].CColHdr.wFlags & FCOLHDR_ABSOLUTE)
					mov eax, ecx
				.else
					mul ecx
					mov ecx, iWidthTotal
					div ecx
				.endif
				mov lvc.cx_,eax
				invoke ListView_InsertColumn( m_hWndLV, ebx, addr lvc)
			.endif
			inc ebx
			add esi,sizeof CColHdr
		.endw
done:
		ret
		align 4
SetMyLVColumns endp

SaveLVColumns proc uses ebx esi

local	lvc:LVCOLUMN

		mov esi, m_pMode
		mov esi, [esi].CMode.pColDesc
		xor ebx, ebx
		.while (1)
			mov lvc.mask_, LVCF_WIDTH
			invoke ListView_GetColumn( m_hWndLV, ebx, addr lvc)
			.break .if (!eax)
			mov eax, lvc.cx_
			mov [esi].CColHdr.wWidth, ax 
			or [esi].CColHdr.wFlags, FCOLHDR_ABSOLUTE
			add esi, sizeof CColHdr
			inc ebx
		.endw
		ret
		align 4

SaveLVColumns endp

;--- check if type of item in CLSID view is a file link

IsFileLink proc uses ebx iItem:DWORD

		invoke GetItemData@CDocument, m_pDoc, iItem, TYPECOL_IN_CLSID
		mov ebx, eax
		.if (!eax)
			jmp done
		.endif
;--------------------------------------- dont check TreatAs, AutoConvertTo,...
		invoke lstrcmp, ebx, CStr("TreatAs")
		.if (eax)
			invoke lstrcmp, ebx, CStr("AutoConvertTo")
			.if (eax)
				invoke lstrcmp, ebx, CStr("Ole1Class")
				.if (eax)
					invoke lstrcmp, ebx, CStr("RTFClassName")
				.endif
			.endif
		.endif
done:
		ret
		align 4

IsFileLink endp

;---

SetImageListLV proc uses esi bSet:BOOL

if ?STATEIMAGE

		.if (bSet)
			mov esi, g_himlLV
			mov ecx, ?LVSTYLE or LVS_EX_CHECKBOXES
		.else
			xor esi, esi
			mov ecx, ?LVSTYLE
		.endif
		invoke ListView_SetExtendedListViewStyle( m_hWndLV,	ecx)
		invoke ListView_SetImageList( m_hWndLV, esi, LVSIL_STATE)
else
		.if (bSet)
			invoke ListView_SetImageList( m_hWndLV, g_himlLV, LVSIL_SMALL)
		.endif
endif
		invoke SetImageListHdr, m_hWndLV
		ret
		align 4

SetImageListLV endp


;*** WM_COMMAND/IDM_CHECKFILE: read all listview entries (depending on mode),
;*** check if path to module exists (this check is not foolproofed)


OnCheckFile proc uses ebx 

local	hOldCursor:HCURSOR
local	pszName:LPSTR
local	dwSize:DWORD
local	iColumn:dword
local	bState:BOOL
local	dwCount:DWORD
local	dwItems:DWORD
local	bPathIncluded:BOOL
local	szKey[MAX_PATH]:byte
local	szSysDir[MAX_PATH]:byte
local	szWinDir[MAX_PATH]:byte
local	szStrEx[260]:byte

		.if (m_iMode == MODE_CLSID)
			mov iColumn,PATHCOL_IN_CLSID
		.elseif (m_iMode == MODE_TYPELIB)
			mov iColumn,PATHCOL_IN_TYPELIB
		.elseif (m_iMode == MODE_HKCR)
			mov iColumn,PATHCOL_IN_HKCR
		.else
			ret 				;not in modes "Interface" or "AppID"
		.endif

		invoke SetCursor,g_hCsrWait
		mov hOldCursor,eax
		invoke SetBusyState@CMainDlg, TRUE

		invoke GetSystemDirectory, addr szSysDir, MAX_PATH
		invoke GetWindowsDirectory, addr szWinDir, MAX_PATH
		
		invoke SetImageListLV, TRUE

		ListView_SetItemState m_hWndLV, -1, 0, LVIS_SELECTED

		invoke GetItemCount@CDocument, m_pDoc
		mov dwItems,eax
		mov dwCount,0
		mov ebx,0
		.while (ebx < dwItems)
			mov bState,0
			invoke GetItemData@CDocument, m_pDoc, ebx, iColumn
			.if (eax != 0)
				mov pszName, eax

				.if (m_iMode == MODE_CLSID)
					invoke IsFileLink, ebx
					.if (!eax)
						jmp nextline
					.endif
				.endif

				invoke ExpandEnvironmentStrings, pszName, addr szKey, sizeof szKey
;--------------------------- WinNT bug, so dont rely on return code (MSDN Q234874)
				invoke lstrlen, addr szKey
				inc eax
;---------------------------- throw away double quotes
				lea edx,szKey
				.if ((eax > 2) && (byte ptr [edx] == '"'))
					.if (m_iMode == MODE_HKCR)
						inc edx
						dec eax
						xor ecx, ecx
						.while (ecx < eax)
							.if (byte ptr [edx+ecx] == '"')
								mov byte ptr [edx+ecx], 0
								mov eax, ecx
								.break
							.endif
							inc ecx
						.endw
					.else
						.if (byte ptr [edx+eax-2] == '"')
							mov byte ptr [edx+eax-2],0
							inc edx
							dec eax
							dec eax
						.endif
					.endif
				.endif
				mov pszName, edx
				mov dwSize, eax

				.if (m_iMode == MODE_HKCR)
					.if (word ptr [edx] == "1%")
						jmp nextline
					.endif
					mov eax, [edx+0]
					mov ecx, [edx+4]
					or eax, 20202020h
					or ecx, 20202020h
					.if ((eax == "dnur") && (ecx == "23ll"))
						.if (byte ptr [edx+8] == ' ')
							jmp nextline
						.endif
					.endif
				.endif

;---------------------------- throw away parameters
				xor eax, eax
				mov edx, pszName
				.while (eax < dwSize)
					mov ecx, [edx+eax]
					or ecx,20202000h
					.if ((ecx == "exe.") && (byte ptr [edx+eax+4] == ' '))
						mov byte ptr [edx+eax+4],0
						.break
					.endif
					inc eax
				.endw


				mov cl,'\'
				invoke strchr, pszName, ecx			;check if a path is included
				.if (eax == 0)
					mov bPathIncluded, FALSE
					invoke wsprintf, addr szStrEx, CStr("%s\%s"), addr szSysDir, pszName
				.else
					mov bPathIncluded, TRUE
					invoke lstrcpy, addr szStrEx, pszName
				.endif

				.if (m_iMode == MODE_TYPELIB)
;---------------------------- throw away any "\n" resource id suffix
;---------------------------- will work for IDs 1-9 only
					invoke lstrlen, addr szStrEx
					lea edx, szStrEx
					.if ((eax > 2) && (byte ptr [edx+eax-2] == '\'))
						mov cl,[edx+eax-1]
						.if ((cl >= '1') && (cl <= '9'))
							mov byte ptr [edx+eax-2],0
						.endif
					.endif
				.endif

if 0
				invoke CreateFile, addr szStrEx,GENERIC_READ,
								FILE_SHARE_READ or FILE_SHARE_WRITE,
								NULL,OPEN_EXISTING,0,0
				.if (eax != -1)
					invoke CloseHandle, eax
else
				invoke GetFileAttributes, addr szStrEx
				.if ((eax != -1) && (!(eax & FILE_ATTRIBUTE_DIRECTORY)))
endif
				.else
					.if ((m_iMode == MODE_HKCR) && (bPathIncluded == FALSE))
						invoke wsprintf, addr szStrEx, CStr("%s\%s"), addr szWinDir, pszName
if 0
						invoke CreateFile, addr szStrEx,GENERIC_READ,
									FILE_SHARE_READ or FILE_SHARE_WRITE,
									NULL,OPEN_EXISTING,0,0
						.if (eax != -1)
							invoke CloseHandle, eax
else
						invoke GetFileAttributes, addr szStrEx
						.if ((eax != -1) && (!(eax & FILE_ATTRIBUTE_DIRECTORY)))
endif
						.else
							mov bState,FLAG_IMAGE
							inc dwCount
						.endif
					.else
						mov bState,FLAG_IMAGE
						inc dwCount
					.endif
				.endif
			.endif
nextline:
			invoke SetItemFlag@CDocument, m_pDoc, ebx, bState, FLAG_IMAGE
			
			.if (bState)
				ListView_SetItemState m_hWndLV, ebx, LVIS_SELECTED, LVIS_SELECTED
			.endif

			inc ebx

		.endw

		invoke SetErrorText@CMainDlg, __this, CStr("%u links failed"), dwCount, FALSE

		invoke SetCursor,hOldCursor		;reset cursor
		invoke SetBusyState@CMainDlg, FALSE

		ret
		align 4

OnCheckFile endp

;--- used by OnCheckTypeLib, OnCheckCLSID, OnCheckProgID

CheckRegistryReference proc uses ebx pszFormat:LPSTR, iColumn:DWORD

local	hOldCursor:HCURSOR
local	hKey:HANDLE
local	bState:BOOL
local	dwCount:DWORD
local	dwItems:DWORD
local	szKey[260]:byte

		invoke SetCursor, g_hCsrWait
		mov hOldCursor,eax

		invoke SetImageListLV, TRUE

		ListView_SetItemState m_hWndLV, -1, 0, LVIS_SELECTED

		invoke GetItemCount@CDocument, m_pDoc
		mov dwItems,eax
		mov dwCount,0
		mov ebx,0
		.while (ebx < dwItems)
			mov bState,0
;------------------------------ in HKCR view ProgId check only if item starts with '.'
			.if ((m_iMode == MODE_HKCR) && (iColumn == 1))
				invoke GetItemData@CDocument, m_pDoc, ebx, 0
				.if ((eax) && (byte ptr [eax] != '.'))
					jmp nextline
				.endif
			.endif
			invoke GetItemData@CDocument, m_pDoc, ebx, iColumn
			.if (eax != 0)
				invoke wsprintf,addr szKey,pszFormat, eax
				invoke RegOpenKey, HKEY_CLASSES_ROOT, addr szKey, addr hKey
				.if (eax != ERROR_SUCCESS)
;------------------------------ handle special case in HKCR view
					.if ((m_iMode == MODE_HKCR) && (iColumn == 1))
						invoke GetItemData@CDocument, m_pDoc, ebx, 0
						push eax
						invoke GetItemData@CDocument, m_pDoc, ebx, iColumn
						pop ecx
						invoke wsprintf,addr szKey, CStr("%s\%s"), ecx, eax
						invoke RegOpenKey, HKEY_CLASSES_ROOT, addr szKey, addr hKey
						.if (eax == ERROR_SUCCESS)
							invoke RegCloseKey,hKey
							jmp nextline
						.endif
					.endif
					mov bState,FLAG_IMAGE
					inc dwCount
				.else
					invoke RegCloseKey,hKey
				.endif
			.endif
nextline:
			invoke SetItemFlag@CDocument, m_pDoc, ebx, bState, FLAG_IMAGE
			.if (bState)
				ListView_SetItemState m_hWndLV, ebx, LVIS_SELECTED, LVIS_SELECTED
			.endif
			inc ebx
		.endw

		invoke SetErrorText@CMainDlg, __this, CStr("%u links failed"), dwCount, FALSE

		invoke SetCursor,hOldCursor		;reset cursor
		ret
		align 4

CheckRegistryReference endp


;*** WM_COMMAND/IDM_CHECKTYPELIB: check existance of HKCR\TypeLib\???


OnCheckTypelib proc uses ebx

local	iColumn:dword

		.if (m_iMode == MODE_CLSID)
			mov iColumn,TYPELIBCOL_IN_CLSID
		.elseif (m_iMode == MODE_INTERFACE)
			mov iColumn,TYPELIBCOL_IN_INTERFACE
		.else
			ret
		.endif

		invoke CheckRegistryReference, CStr("TypeLib\%s"), iColumn

		ret
		align 4

OnCheckTypelib endp


;*** WM_COMMAND/IDM_CHECKCLSID: check existance of HKCR\CLSID\???


OnCheckCLSID proc uses ebx

local	iColumn:dword

		.if (m_iMode == MODE_INTERFACE)
			mov iColumn,CLSIDCOL_IN_INTERFACE
		.elseif (m_iMode == MODE_HKCR)
			mov iColumn,CLSIDCOL_IN_HKCR
		.else
			ret
		.endif

		invoke CheckRegistryReference, CStr("CLSID\%s"), iColumn
		ret
		align 4

OnCheckCLSID endp


;*** WM_COMMAND/IDM_CHECKPROGID: check existance of HKCR\???


OnCheckProgID proc uses ebx

local	iColumn:dword

		.if (m_iMode == MODE_CLSID)
			mov iColumn,PROGIDCOL_IN_CLSID
		.elseif (m_iMode == MODE_HKCR)
			mov iColumn, 1
		.else
			ret
		.endif

		invoke CheckRegistryReference, CStr("%s"), iColumn
		ret
		align 4

OnCheckProgID endp


;*** WM_COMMAND/IDM_CHECKAPPID: check existance of HKCR\{clsid}[AppID]


OnCheckAppID proc uses ebx

local	iColumn:dword
local	hOldCursor:HCURSOR
local	hKey:HANDLE
local	hKey2:HANDLE
local	bState:BOOL
local	dwCount:DWORD
local	dwItems:DWORD
local	dwSize:DWORD
local	dwType:DWORD
local	szKey[260]:byte
local	szValue[260]:byte

		.if (m_iMode == MODE_CLSID)
			mov iColumn,0
		.elseif (m_iMode == MODE_APPID)
			mov iColumn,APPIDCOL_IN_APPID
		.else
			ret
		.endif

		invoke SetCursor, g_hCsrWait
		mov hOldCursor,eax

		invoke SetImageListLV, TRUE

		invoke GetItemCount@CDocument, m_pDoc
		mov dwItems,eax
		mov dwCount,0
		mov ebx,0
		.while (ebx < dwItems)
			mov bState,0
			invoke GetItemData@CDocument, m_pDoc, ebx, iColumn
			.if (eax != 0)
				.if (m_iMode == MODE_CLSID)
;------------------------------------- mode CLSID
					invoke wsprintf,addr szKey,CStr("CLSID\%s"), eax
					invoke RegOpenKey,HKEY_CLASSES_ROOT,addr szKey,addr hKey
					.if (eax == ERROR_SUCCESS)
						mov dwSize,sizeof szValue
						invoke RegQueryValueEx,hKey,CStr("AppID"),NULL,addr dwType,addr szValue,addr dwSize
						.if (eax == ERROR_SUCCESS)
							invoke wsprintf, addr szKey, CStr("AppID\%s"), addr szValue
							invoke RegOpenKey, HKEY_CLASSES_ROOT, addr szKey, addr hKey2
							.if (eax != ERROR_SUCCESS)
								mov bState,FLAG_IMAGE
								inc dwCount
							.else
								invoke RegCloseKey,hKey2
							.endif
						.endif
						invoke RegCloseKey,hKey
					.endif
				.else
;------------------------------------- mode APPID
					invoke wsprintf, addr szKey, CStr("AppID\%s"), eax
					invoke RegOpenKey, HKEY_CLASSES_ROOT, addr szKey, addr hKey2
					.if (eax != ERROR_SUCCESS)
						mov bState,FLAG_IMAGE
						inc dwCount
					.else
						invoke RegCloseKey,hKey2
					.endif
				.endif
			.endif

			invoke SetItemFlag@CDocument, m_pDoc, ebx, bState, FLAG_IMAGE
			.if (bState)
				ListView_SetItemState m_hWndLV, ebx, LVIS_SELECTED, LVIS_SELECTED
			.endif
			
			inc ebx
		.endw

		invoke SetErrorText@CMainDlg, __this, CStr("%u links failed"), dwCount, FALSE

		invoke SetCursor,hOldCursor		;reset cursor
		ret
		align 4

OnCheckAppID endp


;*** update state of menu items depending on view mode ***


UpdateMenu proc uses ebx esi

local	iSelCount:dword
local	iItem:DWORD
local	hKey:HANDLE
local	bIsFileLink:BOOL
local	szStr[128]:byte

;--------------------------- menu item state may depend on # of items 
		invoke ListView_GetSelectedCount( m_hWndLV)
		mov iSelCount,eax
		.if (eax)
			invoke ListView_GetNextItem( m_hWndLV, -1, LVNI_SELECTED)
			mov iItem, eax
		.endif
		mov esi, m_iMode
;--------------------------- menu item IDM_CHECKFILE

		mov ebx,MF_GRAYED or MF_BYCOMMAND
		.if ((esi == MODE_CLSID) || (esi == MODE_TYPELIB) || (esi == MODE_HKCR))
			mov ebx,MF_ENABLED or MF_BYCOMMAND
		.endif
		invoke EnableMenuItem, m_hMenu, IDM_CHECKFILE, ebx

;--------------------------- menu item IDM_CHECKLIB

		mov ebx,MF_GRAYED or MF_BYCOMMAND
		.if ((esi == MODE_CLSID) || (esi == MODE_INTERFACE))
			mov ebx,MF_ENABLED or MF_BYCOMMAND
		.endif
		invoke EnableMenuItem, m_hMenu, IDM_CHECKTYPELIB, ebx

;--------------------------- menu item IDM_CHECKCLSID

		mov ebx,MF_GRAYED or MF_BYCOMMAND
		.if ((esi == MODE_INTERFACE) || (esi == MODE_HKCR))
			mov ebx,MF_ENABLED or MF_BYCOMMAND
		.endif
		invoke EnableMenuItem, m_hMenu, IDM_CHECKCLSID, ebx

;--------------------------- menu item IDM_CHECKPROGID

		mov ebx,MF_GRAYED or MF_BYCOMMAND
		.if ((esi == MODE_CLSID) || (esi == MODE_HKCR))
			mov ebx,MF_ENABLED or MF_BYCOMMAND
		.endif
		invoke EnableMenuItem, m_hMenu, IDM_CHECKPROGID, ebx

;--------------------------- menu item IDM_CHECKAPPID

		mov ebx,MF_GRAYED or MF_BYCOMMAND
		.if ((esi == MODE_CLSID) || (esi == MODE_APPID))
			mov ebx,MF_ENABLED or MF_BYCOMMAND
		.endif
		invoke EnableMenuItem, m_hMenu, IDM_CHECKAPPID, ebx

;--------------------------- menu item IDM_CREATEINSTANCE/IDM_CREATEINSTON

		mov ebx, MF_GRAYED or MF_BYCOMMAND

		.if (iSelCount == 1)							;allow only 1 selected item for CoCreateInst
			.if (esi == MODE_CLSID)
				mov ebx,MF_ENABLED or MF_BYCOMMAND
			.elseif (esi == MODE_HKCR)
				invoke GetItemData@CDocument, m_pDoc, iItem, CLSIDCOL_IN_HKCR
				.if (eax)
					mov ebx,MF_ENABLED or MF_BYCOMMAND
				.endif
			.endif
		.endif
		invoke EnableMenuItem, m_hMenu, IDM_CREATEINSTANCE, ebx
		invoke EnableMenuItem, m_hMenu, IDM_CREATEINSTON, ebx
		invoke EnableMenuItem, m_hMenu, IDM_GETCLASSFACT, ebx

;--------------------------- menu item IDM_EDIT/IDM_COPY

		mov ebx, MF_GRAYED or MF_BYCOMMAND
		.if (iSelCount > 0)
			mov ebx, MF_ENABLED or MF_BYCOMMAND
		.endif
		invoke EnableMenuItem, m_hMenu, IDM_EDIT, ebx
		invoke EnableMenuItem, m_hMenu, IDM_COPY, ebx
		invoke EnableMenuItem, m_hMenu, IDM_REMOVEITEM, ebx

;--------------------------- menu item IDM_COPYGUID

		mov ebx,MF_GRAYED or MF_BYCOMMAND
		.if (iSelCount == 1)
			.if (esi == MODE_HKCR)
				invoke GetItemData@CDocument, m_pDoc, iItem, CLSIDCOL_IN_HKCR
				.if (eax)
					mov ebx,MF_ENABLED or MF_BYCOMMAND
				.endif
			.else
				mov ebx,MF_ENABLED or MF_BYCOMMAND
			.endif
		.endif
		invoke EnableMenuItem, m_hMenu, IDM_COPYGUID, ebx

;--------------------------- menu item IDM_TYPELIBDLG

		mov ebx,MF_GRAYED or MF_BYCOMMAND
		.if (iSelCount > 0)
			.if (esi == MODE_TYPELIB)
				mov ebx,MF_ENABLED or MF_BYCOMMAND
			.elseif (esi == MODE_CLSID)
				invoke GetItemData@CDocument, m_pDoc, iItem, TYPELIBCOL_IN_CLSID
				.if (eax)
					mov ebx,MF_ENABLED or MF_BYCOMMAND
				.endif
			.elseif (esi == MODE_INTERFACE)
				invoke GetItemData@CDocument, m_pDoc, iItem, TYPELIBCOL_IN_INTERFACE
				.if (eax)
					mov ebx,MF_ENABLED or MF_BYCOMMAND
				.endif
			.elseif (esi == MODE_HKCR)
				invoke GetItemData@CDocument, m_pDoc, iItem, CLSIDCOL_IN_HKCR
				.if (eax)
					invoke wsprintf, addr szStr, CStr("CLSID\%s\TypeLib"), eax
					invoke RegOpenKeyEx,HKEY_CLASSES_ROOT,addr szStr,0,KEY_READ,addr hKey
					.if (eax == ERROR_SUCCESS)
						invoke RegCloseKey, hKey
						mov ebx,MF_ENABLED or MF_BYCOMMAND
					.endif
				.endif
			.endif
		.endif
		invoke EnableMenuItem, m_hMenu, IDM_TYPELIBDLG, ebx

		.if ((iSelCount) && (esi == MODE_CLSID))
			invoke IsFileLink, iItem
			mov bIsFileLink, eax
		.else
			mov bIsFileLink, TRUE
		.endif

;--------------------------- menu item IDM_EXPLORE

		mov ebx,MF_GRAYED or MF_BYCOMMAND
		.if (iSelCount == 1)
			.if (((esi == MODE_CLSID) && (bIsFileLink)) || (esi == MODE_TYPELIB))
				mov ebx,MF_ENABLED or MF_BYCOMMAND
			.elseif (esi == MODE_HKCR)
				invoke GetItemData@CDocument, m_pDoc, iItem, PATHCOL_IN_HKCR
				.if (eax)
					mov ebx,MF_ENABLED or MF_BYCOMMAND
				.endif
			.endif
		.endif
		invoke EnableMenuItem, m_hMenu, IDM_EXPLORE, ebx

;--------------------------- menu item IDM_PROPERTIES

		mov ebx,MF_GRAYED or MF_BYCOMMAND
		.if (iSelCount == 1)
			.if (((esi == MODE_CLSID) && (bIsFileLink)) || (esi == MODE_TYPELIB))
				mov ebx,MF_ENABLED or MF_BYCOMMAND
			.endif
		.endif
		invoke EnableMenuItem, m_hMenu, IDM_PROPERTIES, ebx

;--------------------------- menu item IDM_UNREGISTER

		mov ebx,MF_GRAYED or MF_BYCOMMAND
		.if (iSelCount == 1)
			.if ((esi == MODE_CLSID) && (bIsFileLink))
				mov ebx,MF_ENABLED or MF_BYCOMMAND
			.endif
		.endif
		invoke EnableMenuItem, m_hMenu, IDM_UNREGISTER, ebx

;--------------------------- menu item IDM_OBJECTDLG

		.if ((esi == MODE_OBJECT) || (esi == MODE_ROT))
			.if (iSelCount == 0)
				mov ebx, MF_GRAYED or MF_BYCOMMAND
			.else
				mov ebx, MF_ENABLED or MF_BYCOMMAND
			.endif
			invoke EnableMenuItem, m_hMenu, IDM_OBJECTDLG, ebx
		.endif

;--------------------------- menu item IDM_UNLOCK

		.if (esi == MODE_OBJECT)
			mov ebx, MF_GRAYED or MF_BYCOMMAND
			.if (iSelCount > 0)
				invoke GetItemData@CDocument, m_pDoc, iItem, DATACOL_IN_OBJECT
				.if (eax)
					invoke vf(eax, IObjectItem, IsLocked)
					.if (eax)
						mov ebx, MF_ENABLED or MF_BYCOMMAND
					.endif
				.endif
			.endif
			invoke EnableMenuItem, m_hMenu, IDM_UNLOCK, ebx

;--------------------------- menu item IDM_VIEWMONIKER

			mov ebx, MF_GRAYED or MF_BYCOMMAND
			.if (iSelCount == 1)
				invoke GetItemData@CDocument, m_pDoc, iItem, DATACOL_IN_OBJECT
				.if (eax)
					invoke vf(eax, IObjectItem, GetMoniker)
					.if (eax)
						mov ebx, MF_ENABLED or MF_BYCOMMAND
					.endif
				.endif
			.endif
			invoke EnableMenuItem, m_hMenu, IDM_VIEWMONIKER, ebx

		.endif

		ret
		align 4

UpdateMenu endp


OnEnterMenuLoop proc uses ebx bIsTrackPopupMenu:BOOL

local	pDataObject:LPDATAOBJECT

		invoke GetSubMenu, m_hMenu, 1

		.if (bIsTrackPopupMenu)
			push eax
			invoke GetDefaultCommand
			pop ecx
			invoke SetMenuDefaultItem, ecx, eax, FALSE
		.else
			invoke SetMenuDefaultItem, eax, -1, FALSE

			mov ebx, MF_BYCOMMAND or MF_GRAYED or MF_DISABLED
			invoke OleGetClipboard, addr pDataObject
			.if (eax == S_OK)
				invoke vf(pDataObject, IDataObject, Release)
				mov ebx, MF_BYCOMMAND or MF_ENABLED
			.endif
			invoke EnableMenuItem, m_hMenu, IDM_PASTE, ebx

			mov ebx, MF_BYCOMMAND or MF_GRAYED or MF_DISABLED
			.if (g_pStorage)
				mov ebx, MF_BYCOMMAND or MF_ENABLED
			.endif
			invoke EnableMenuItem, m_hMenu, IDM_VIEWSTORAGE, ebx

			mov ebx, MF_BYCOMMAND or MF_GRAYED or MF_DISABLED
			.if (g_pStream)
				mov ebx, MF_BYCOMMAND or MF_ENABLED
			.endif
			invoke EnableMenuItem, m_hMenu, IDM_VIEWSTREAM, ebx

;--------------------------- now set check mark of view mode

			mov ebx,offset ModeDesc
			mov ecx,NUMMODES
			.while (ecx > 0)
				push ecx
				mov eax,[ebx].CMode.iMode
				.if (eax == m_iMode)
					mov ecx, MF_BYCOMMAND or MF_CHECKED
				.else
					mov ecx, MF_BYCOMMAND or MF_UNCHECKED
				.endif
				invoke CheckMenuItem, m_hMenu, [ebx].CMode.iCmdID, ecx
				add ebx,sizeof CMode
				pop	ecx
				dec ecx
			.endw
		.endif

		ret
		align 4

OnEnterMenuLoop endp

OnExitMenuLoop proc bIsTrackPopupMenu:BOOL
		ret
		align 4
OnExitMenuLoop endp

ListView_SetTopIndex proc hWndLV:HWND, iTopIndex:DWORD

local	dwItems:DWORD

		invoke ListView_GetItemCount( m_hWndLV)
		mov dwItems, eax
		invoke ListView_GetCountPerPage( m_hWndLV)
		add eax, iTopIndex
		.if (eax > dwItems)
			mov eax, dwItems
		.endif
		dec eax
		invoke ListView_EnsureVisible( hWndLV, eax, FALSE)
		invoke ListView_EnsureVisible( hWndLV, iTopIndex, FALSE)
		ret
		align 4
ListView_SetTopIndex endp

;*** set new mode (if iNewMode != -1),
;*** recreate CDocument


RefreshView proc uses esi iNewMode:dword
	
local	hOldCursor:HCURSOR
local	pSavedView:ptr CMainView
local	szText[128]:byte
ifdef _DEBUG
local	this_:ptr CMainDlg
		mov this_, __this
endif

		invoke SetWindowRedraw( m_hWndLV, FALSE)
		invoke SetCursor, g_hCsrWait
		mov hOldCursor,eax

		invoke SetBusyState@CMainDlg, TRUE

if ?MULTIDOC
		.if (g_bMultiDoc)
			.if (iNewMode == -1)
				invoke Reset@CMainDlg, __this
			.else
				mov esi, m_iMode
				.if (esi != -1)
					mov eax, esi
					mov ecx, sizeof CMainView
					mul ecx
					lea esi, m_savedView[eax]
					invoke ListView_GetTopIndex( m_hWndLV)
					mov [esi].CMainView.iTopIndex, eax
					mov eax, m_iSortCol
					mov [esi].CMainView.iSortCol, eax
					mov eax, m_iSortDir
					mov [esi].CMainView.iSortDir, eax
if ?SAVECOLORDER
					invoke Header_GetItemCount( m_hWndHdr)
					.if (eax)
						mov [esi].CMainView.dwColumns, eax
						push eax
						shl eax, 2
						invoke malloc, eax
						xchg eax, [esi].CMainView.pdwColOrder
						invoke free, eax
						pop ecx
						invoke Header_GetOrderArray( m_hWndHdr, ecx, [esi].CMainView.pdwColOrder)
					.endif
endif
				.endif
			.endif
		.else
endif
			invoke Reset@CMainDlg, __this
if ?MULTIDOC
		.endif
endif

;;		invoke ListView_DeleteAllItems( m_hWndLV)

		invoke SetImageListLV, FALSE


;-------------------------- set listview header if view mode changes

		mov eax,iNewMode
		.if ((eax != -1) || (g_bColumnsChanged == TRUE))
			.if (eax != -1)
if ?MULTIDOC
				.if (!g_bMultiDoc)
endif
					mov m_iSortCol,-1
if ?MULTIDOC
				.endif
endif
				mov edx,offset ModeDesc
				mov ecx,NUMMODES
				.while (ecx > 0)
					.break .if (eax == [edx].CMode.iMode)
					add edx,sizeof CMode
					dec ecx
				.endw
				mov m_pMode,edx
				mov m_iMode,eax
			.endif
			invoke SetMyLVColumns, iNewMode
			invoke TabCtrl_SetCurSel( m_hWndTab, m_iMode)
		.endif

;-------------------------- build CDocument

if ?MULTIDOC
		.if (g_bMultiDoc)
			mov eax, m_iMode
			mov ecx, sizeof CMainView
			mul ecx
			lea esi, m_savedView[eax]
			mov pSavedView, esi
if ?SAVECOLORDER
			.if ([esi].CMainView.pdwColOrder)
;;				Header_GetItemCount m_hWndHdr
;;				DebugOut "Header_GetItemCount=%u, OrderArray=%X", eax, [esi].CMainView.pdwColOrder
				invoke Header_SetOrderArray( m_hWndHdr, [esi].CMainView.dwColumns, [esi].CMainView.pdwColOrder)
			.endif
endif
			mov eax, [esi].CMainView.pDoc
;-------------------------- always refresh ROT
			.if (eax && (m_iMode == MODE_ROT))

				invoke Destroy@CDocument, [esi].CMainView.pDoc
				xor eax, eax
				mov [esi].CMainView.pDoc, eax
				mov [esi].CMainView.iTopIndex, eax
			.endif
			.if (!eax)
				invoke Create@CDocument, m_hWnd, m_iMode, m_iNumCols, m_pszRoot
				mov [esi].CMainView.pDoc, eax
			.endif
			mov ecx, [esi].CMainView.iSortCol
			mov m_iSortCol, ecx
			mov ecx, [esi].CMainView.iSortDir
			mov m_iSortDir, ecx
		.else
endif
			invoke Create@CDocument, m_hWnd, m_iMode, m_iNumCols, m_pszRoot
if ?MULTIDOC
		.endif
endif
		mov m_pDoc,eax

		invoke GetItemCount@CDocument, m_pDoc
		push eax
		invoke ListView_SetItemCount( m_hWndLV, eax)
		invoke ListView_SetCallbackMask( m_hWndLV, LVIS_SELECTED or LVIS_FOCUSED)
		pop ecx
if ?MULTIDOC
		xor esi, esi
		.while (ecx)
			push ecx
			invoke GetItemFlag@CDocument, m_pDoc, esi, LVIS_SELECTED or LVIS_FOCUSED
			ListView_SetItemState m_hWndLV, esi, eax, LVIS_SELECTED or LVIS_FOCUSED
			pop ecx
			inc esi
			dec ecx
		.endw
endif
		invoke SetStatusPane1
		invoke SetStatusPane2

if ?MULTIDOC
		.if (g_bMultiDoc)
			mov ecx, pSavedView
			invoke ListView_SetTopIndex, m_hWndLV, [ecx].CMainView.iTopIndex
			.if (m_iSortCol != -1)
				.if (g_bColumnsChanged)
					invoke SortListView
if ?HDRBMPS
				.else
					invoke SetHeaderBitmap, m_hWndLV, m_iSortCol, m_iSortDir
endif
				.endif
			.endif
		.else
endif
			.if (m_iSortCol != -1)
				invoke SortListView
			.endif
if ?MULTIDOC
		.endif
endif

		invoke DeleteMenu, m_hMenu, IDM_OBJECTDLG, MF_BYCOMMAND
		invoke DeleteMenu, m_hMenu, IDM_VIEWMONIKER, MF_BYCOMMAND
		invoke DeleteMenu, m_hMenu, IDM_UNLOCK, MF_BYCOMMAND
		.if ((m_iMode == MODE_OBJECT) || (m_iMode == MODE_ROT))
			invoke InsertMenu, m_hMenu, IDM_CREATEINSTANCE, MF_GRAYED or MF_BYCOMMAND, IDM_OBJECTDLG, CStr("&View Object",9,"F6")
		.endif
		.if (m_iMode == MODE_OBJECT)
			invoke InsertMenu, m_hMenu, IDM_CREATEINSTANCE, MF_GRAYED or MF_BYCOMMAND, IDM_UNLOCK, CStr("&Unlock Object")
			invoke InsertMenu, m_hMenu, IDM_CREATEINSTANCE, MF_GRAYED or MF_BYCOMMAND, IDM_VIEWMONIKER, CStr("View &Moniker")
		.endif
		invoke UpdateMenu

		mov g_bColumnsChanged, FALSE

		invoke SetWindowRedraw( m_hWndLV, TRUE)

		invoke SetCursor,hOldCursor

		invoke SetBusyState@CMainDlg, FALSE

		StatusBar_SetText g_hWndSB, 0, addr g_szNull
		
		ret
		align 4

RefreshView endp


;*** refresh one line in listview (not implemented yet)


RefreshLine proc uses esi pszItem:ptr EDITITEMDESC

if 0
		mov eax,pszItem
		invoke RefreshItem@CDocument, m_pDoc, [eax].EDITITEMDESC.dwCookie,\
				[eax].EDITITEMDESC.pszKey
endif
		ret
		align 4

RefreshLine endp

if 0
			invoke lstrcpy, addr szCLSID, eax
			.while (1)
				invoke DialogBoxParam, g_hInstance, IDD_ENTERCLSID, m_hWnd, inputdlgproc, addr szCLSID
				.if (!eax)
					jmp done
				.endif
				invoke MultiByteToWideChar,CP_ACP,MB_PRECOMPOSED,
					addr szCLSID, -1, addr wszCLSID, 40
				invoke CLSIDFromString,addr wszCLSID,addr clsid
				.break .if (eax == S_OK)
				invoke OutputMessage, m_hWnd, eax, CStr("CLSIDFromString()"), 0
			.endw
endif

OnGetClassFactory proc

local item:DWORD
local pClassFactory:LPCLASSFACTORY
local hr:dword
local clsid:CLSID
local szCLSID[MAXINPUTTEXT]:BYTE
local wszCLSID[40]:WORD

		invoke ListView_GetNextItem( m_hWndLV, -1, LVNI_SELECTED)
		.if (eax != -1)
			mov item, eax
			.if (m_iMode == MODE_CLSID)
				mov ecx, CLSIDCOL_IN_CLSID
			.else
				mov ecx, CLSIDCOL_IN_HKCR
			.endif
			invoke GetItemData@CDocument, m_pDoc, eax, ecx
			lea ecx, wszCLSID
			invoke MultiByteToWideChar,CP_ACP,MB_PRECOMPOSED,
				eax, -1, ecx, 40
			invoke CLSIDFromString,addr wszCLSID,addr clsid
			.if (eax != S_OK)
				invoke OutputMessage, m_hWnd, eax, CStr("CLSIDFromString()"), 0
				jmp done
			.endif
			invoke GetCoCreateFlags@COptions
			mov ecx, eax
			invoke CoGetClassObject, addr clsid, ecx, NULL,
				addr IID_IUnknown, addr pClassFactory
;				addr IID_IClassFactory, addr pClassFactory
			.if (eax == S_OK)
				invoke Create@CObjectItem, pClassFactory, NULL
				.if (eax)
					push eax
					invoke vf(eax, IObjectItem, ShowObjectDlg), m_hWnd
					pop eax
					invoke vf(eax, IObjectItem, Release)
				.endif
				invoke vf(pClassFactory, IUnknown, Release)
			.else
				invoke OutputMessage, m_hWnd, eax, CStr("CoGetClassObject()"), 0
			.endif
		.endif
done:
		ret
		align 4

OnGetClassFactory endp

ifdef @StackBase
	option stackbase:ebp
endif
	option prologue:@sehprologue
	option epilogue:@sehepilogue

;;CLASS_E_NOTLICENSED	equ 80040112h

CreateLicencedObject proc pClsid:ptr CLSID, ppUnknown:ptr LPUNKNOWN

local pClassFactory2:LPCLASSFACTORY2
local bstrKey:BSTR
local hr:dword
local licinfo:LICINFO

		invoke GetCoCreateFlags@COptions
		mov ecx, eax
		invoke CoGetClassObject, pClsid, ecx, NULL,
				addr IID_IClassFactory2, addr pClassFactory2
		.if (eax == S_OK)
			invoke vf(pClassFactory2, IClassFactory2, GetLicInfo), addr licinfo
			mov bstrKey,NULL
			.if (licinfo.fRuntimeKeyAvail)
				invoke vf(pClassFactory2, IClassFactory2, RequestLicKey), NULL, addr bstrKey
			.endif
			invoke vf(pClassFactory2, IClassFactory2, CreateInstanceLic),
				NULL, NULL, addr IID_IUnknown, bstrKey, ppUnknown
			mov hr, eax
			.if (bstrKey)
				invoke SysFreeString, bstrKey
			.endif
			invoke vf(pClassFactory2, IClassFactory2, Release)
			mov eax, hr
		.endif
		ret
		align 4
CreateLicencedObject endp

;*** WM_COMMAND/IDM_CREATEINSTANCE: do CoCreateInstance() for selected entry 
;*** after creation, start modal dialog, then "release" created object


OnCreateInstance proc uses ebx esi edi

local	this@:ptr CMainDlg
local	clsid:CLSID
local	dwErr:dword
local	pUnknown:LPUNKNOWN
local	dwExc:dword
local	dwExcAddr:dword
local	wszCLSID[40]:word
local	szStr[260]:byte

		mov this@,__this		;save this pointer

		.if ((m_iMode != MODE_CLSID) && (m_iMode != MODE_HKCR))
			jmp exit
		.endif

		invoke ListView_GetNextItem( m_hWndLV, -1, LVNI_SELECTED)
		.if (eax == -1)
			jmp exit
		.endif
		.if (m_iMode == MODE_CLSID)
			mov ecx, CLSIDCOL_IN_CLSID
		.elseif (m_iMode == MODE_HKCR)
			mov ecx, CLSIDCOL_IN_HKCR
		.endif
		invoke GetItemData@CDocument, m_pDoc, eax, ecx
		.if (!eax)
			jmp exit
		.endif

		mov ecx,eax
		invoke MultiByteToWideChar,CP_ACP,MB_PRECOMPOSED,
				ecx,-1,addr wszCLSID, 40 

		invoke CLSIDFromString,addr wszCLSID,addr clsid

		.if (eax != ERROR_SUCCESS)
			invoke MessageBox, m_hWnd, CStr("No valid CLSID"), 0, MB_OK
			jmp exit
		.endif

		.try
			invoke SetBusyState@CMainDlg, TRUE
			mov eax, S_FALSE
			.if (g_bUseClassFactory2)
				invoke CreateLicencedObject, addr clsid, addr pUnknown
			.endif
			.if (eax != S_OK)
				invoke GetCoCreateFlags@COptions
				mov ecx, eax
				invoke CoCreateInstance, addr clsid, NULL,
					ecx, addr IID_IUnknown, addr pUnknown
			.endif
		.exceptfilter
			mov eax,_exception_info()
			mov eax,(EXCEPTION_POINTERS ptr [eax]).ExceptionRecord
			mov ecx,(EXCEPTION_RECORD ptr [eax]).ExceptionCode
			mov dwExc,ecx
			mov ecx,(EXCEPTION_RECORD ptr [eax]).ExceptionAddress
			mov dwExcAddr,ecx

			mov __this,this@	;reload edi!
			invoke wsprintf, addr szStr, CStr("Exception 0x%08X occured at 0x%08X.",0ah,"Do you want to continue?"),dwExc,dwExcAddr
			invoke MessageBox, m_hWnd, addr szStr, CStr("Error executing CoCreateInstance"), MB_YESNO
			.if (eax == IDNO)
				mov eax,EXCEPTION_EXECUTE_HANDLER
			.else
				mov eax,EXCEPTION_CONTINUE_SEARCH
			.endif
		.except
			mov __this,this@	;reload edi!
			mov eax, E_UNEXPECTED
		.endtry
		push eax
		invoke SetBusyState@CMainDlg, FALSE
		pop eax
		.if (eax == S_OK)
			.if (!pUnknown)
				invoke MessageBox, m_hWnd, CStr("IUnknown pointer is NULL"), 0, MB_OK
			.else
				invoke Create@CObjectItem, pUnknown, addr clsid
				.if (eax)
					push eax
					invoke vf(eax, IObjectItem, ShowObjectDlg), m_hWnd
					pop eax
					invoke vf(eax, IObjectItem, Release)
				.endif
				invoke vf(pUnknown, IUnknown, Release)
			.endif
		.else
			invoke OutputMessage, m_hWnd, eax, CStr("CoCreateInstance()"), 0
		.endif
exit:
		ret
		align 4

OnCreateInstance endp


;*** WM_COMMAND/ONCREATEINSTANCEON

;--- get CoCreateInstanceEx dynamically because DCOM has to be installed for that

protoCoCreateInstanceEx typedef proto :ptr GUID, :LPUNKNOWN, :DWORD, :ptr COSERVERINFO, :DWORD, :ptr MULTI_QI
LPFNCOCREATEINSTANCEEX typedef ptr protoCoCreateInstanceEx

OnCreateInstanceOn proc uses ebx esi edi

local	this@:ptr CMainDlg
local	csi:COSERVERINFO
local	mqi:MULTI_QI
local	clsid:CLSID
local	dwErr:dword
local	dwExc:dword
local	dwExcAddr:dword
local	hLibOle32:HINSTANCE
local	pfnCoCreateInstanceEx:LPFNCOCREATEINSTANCEEX
local	wszCLSID[40]:word
local	wszMachine[MAXINPUTTEXT]:word
local	szStr[260]:byte

		mov this@,__this		;save this pointer

		invoke GetModuleHandle, CStr("OLE32.DLL")
		mov hLibOle32, eax
		.if (eax)
			invoke GetProcAddress, hLibOle32, CStr("CoCreateInstanceEx")
			mov pfnCoCreateInstanceEx, eax
		.endif
		.if (!eax)
			invoke MessageBox, m_hWnd, CStr("DCOM not installed"), 0, MB_OK
			jmp exit
		.endif

		invoke DialogBoxParam, g_hInstance, IDD_CREATEREMOTE, m_hWnd, inputdlgproc, addr g_szMachine
		.if (!eax)
			jmp exit
		.endif

		invoke ListView_GetNextItem( m_hWndLV, -1, LVNI_SELECTED)
		.if (eax == -1)
			jmp exit
		.endif

		.if (m_iMode == MODE_CLSID)
			mov ecx, CLSIDCOL_IN_CLSID
		.elseif (m_iMode == MODE_HKCR)
			mov ecx, CLSIDCOL_IN_HKCR
		.endif
		invoke GetItemData@CDocument, m_pDoc, eax, ecx
		.if (!eax)
			jmp exit
		.endif
		mov ecx,eax
		invoke MultiByteToWideChar,CP_ACP,MB_PRECOMPOSED,
				ecx,-1,addr wszCLSID, 40 
		invoke CLSIDFromString,addr wszCLSID,addr clsid
		.if (eax != ERROR_SUCCESS)
			invoke MessageBox, m_hWnd, CStr("No valid CLSID"), 0, MB_OK
			jmp exit
		.endif

		invoke SetBusyState@CMainDlg, TRUE

		.if (g_szMachine)
			invoke MultiByteToWideChar,CP_ACP,MB_PRECOMPOSED,
				addr g_szMachine,-1,addr wszMachine, MAXINPUTTEXT
			xor edx, edx
			lea eax,wszMachine
			mov csi.dwReserved1, edx
			mov csi.pwszName,eax
			mov csi.pAuthInfo, edx
			mov csi.dwReserved2, edx
			lea ecx, csi
			mov edx, CLSCTX_REMOTE_SERVER
		.else
			invoke GetCoCreateFlags@COptions
			mov edx, eax
			or edx, CLSCTX_REMOTE_SERVER
			xor ecx,ecx			;pServerInfo is NULL
		.endif

		mov mqi.pIID,offset IID_IUnknown
		mov mqi.pItf, NULL
		mov mqi.hr, 0

		.try
			invoke pfnCoCreateInstanceEx, addr clsid, NULL,
							edx, ecx, 1, addr mqi
		.exceptfilter
			mov eax,_exception_info()
			mov eax,(EXCEPTION_POINTERS ptr [eax]).ExceptionRecord
			mov ecx,(EXCEPTION_RECORD ptr [eax]).ExceptionCode
			mov dwExc,ecx
			mov ecx,(EXCEPTION_RECORD ptr [eax]).ExceptionAddress
			mov dwExcAddr,ecx

			mov __this,this@
			invoke wsprintf, addr szStr, CStr("Exception 0x%08X occured at 0x%08X.",0ah,"Do you want to continue?"),dwExc,dwExcAddr
			invoke MessageBox, m_hWnd, addr szStr, CStr("Error executing CoCreateInstance"), MB_YESNO
			.if (eax == IDNO)
				mov eax,EXCEPTION_EXECUTE_HANDLER
			.else
				mov eax,EXCEPTION_CONTINUE_SEARCH
			.endif
		.except
			mov __this,this@
			mov eax, E_UNEXPECTED
		.endtry

		push eax
		invoke SetBusyState@CMainDlg, FALSE
		pop eax

		.if (eax == S_OK)
			invoke Create@CObjectItem, mqi.pItf, addr clsid
			.if (eax)
				push eax
				invoke vf(eax, IObjectItem, ShowObjectDlg), m_hWnd
				pop eax
				invoke vf(eax, IObjectItem, Release)
			.endif
			invoke vf(mqi.pItf, IUnknown, Release)
		.else
			invoke OutputMessage, m_hWnd, eax, CStr("CoCreateInstanceEx()"), 0
		.endif
exit:
		ret
		align 4

OnCreateInstanceOn endp

	option prologue: prologuedef
	option epilogue: epiloguedef
ifdef @StackBase
	option stackbase:esp
endif

;*** user pressed right mouse button, show context menu (submenu 1) ***

ShowContextMenu proc pNMLV:ptr NMLISTVIEW, bMouse:BOOL

local	hSubMenu:HMENU
local	pt:POINT

		invoke GetSubMenu, m_hMenu, 1	;get "Edit" submenu
		.if (eax != 0)
			mov hSubMenu,eax
			invoke GetItemPosition, m_hWndLV, bMouse, addr pt
			invoke TrackPopupMenu, hSubMenu, TPM_LEFTALIGN or TPM_LEFTBUTTON,
						pt.x, pt.y, 0, m_hWnd, NULL
		.endif
		ret
		align 4

ShowContextMenu endp


;*** process WM_COMMAND/IDM_EDIT
;*** prepares, creates and shows the "edit" dialog(s)


OnEdit proc uses ebx esi

local	pszCLSID:LPSTR
local	hKey:HANDLE
local	dwSize:dword
local	kp[7]:KEYPAIR
local	pKP:ptr KEYPAIR
local	dwNumPairs:dword
local	dwNumSelItems:dword
local	pEditDlg:ptr CEditDlg
local	dwType:DWORD
;local	szKey1[128]:byte
local	szKey2[128]:byte
local	szKey3[128]:byte
local	szKey4[256]:byte
local	szKey5[40]:byte
local	szKey[260]:byte

		invoke ListView_GetSelectedCount( m_hWndLV)
		.if (eax == 0)
			ret
		.endif
		mov dwNumSelItems,eax

;------------------------------ create a dialog object (not shown yet)

		movzx ecx,g_bConfirmDelete
		invoke Create@CEditDlg, m_hWnd, EDITMODELESS, ecx
		.if (eax == 0)
			ret
		.endif
		mov pEditDlg,eax

;------------------------------ all selected lines are processed here!

		mov esi,-1
		.while (1) 

;------------------------------ get next selected line
			invoke ListView_GetNextItem( m_hWndLV, esi, LVNI_SELECTED)
			.break .if (eax == -1)
			mov esi,eax

;------------------------------ init key struct (maximum is 5 keys)

			invoke ZeroMemory, addr kp, sizeof kp

			invoke GetItemData@CDocument, m_pDoc, esi, 0
			mov ecx,eax
			mov pszCLSID,eax

;------------------------------ set keys for editor to show/edit
;------------------------------ key[0] is equal for all views (clsid)
			mov eax, m_pszRoot
			mov kp[0*sizeof KEYPAIR].pszRoot,eax
			.if (eax)
				mov kp[0*sizeof KEYPAIR].pszKey,ecx
			.else
				mov kp[0*sizeof KEYPAIR].pszRoot,ecx
			.endif
;------------------------------ expand first key, but only if 1 line selected
			.if (dwNumSelItems < 2)
				mov kp[0*sizeof KEYPAIR].bExpand,TRUE
			.endif
			mov dwNumPairs,1
			lea ebx, [kp + 1*sizeof KEYPAIR]
		
			mov eax,m_iMode

;------------------------------ keys[1]-[4] are mode dependant

			.if (eax == MODE_CLSID)

				.if (g_bAddAutoTreatEntries)
					invoke wsprintf, addr szKey, CStr("CLSID\%s\AutoConvertTo"), pszCLSID
					invoke RegOpenKey, HKEY_CLASSES_ROOT, addr szKey, addr hKey
					.if (eax == ERROR_SUCCESS)
						mov dwSize, sizeof szKey2
						invoke RegQueryValue, hKey, NULL, addr szKey2, addr dwSize
						invoke RegCloseKey, hKey
						lea eax,szKey2
						.if (byte ptr [eax])
							mov [ebx].KEYPAIR.pszKey, eax
							mov [ebx].KEYPAIR.pszRoot,offset g_szRootCLSID
							add ebx, sizeof KEYPAIR
							inc dwNumPairs
						.endif
					.endif

					invoke wsprintf, addr szKey, CStr("CLSID\%s\TreatAs"), pszCLSID
					invoke RegOpenKey, HKEY_CLASSES_ROOT, addr szKey, addr hKey
					.if (eax == ERROR_SUCCESS)
						mov dwSize, sizeof szKey3
						invoke RegQueryValue, hKey, NULL, addr szKey3, addr dwSize
						invoke RegCloseKey, hKey
						lea eax,szKey3
						.if (byte ptr [eax])
							mov [ebx].KEYPAIR.pszKey, eax
							mov [ebx].KEYPAIR.pszRoot,offset g_szRootCLSID
							add ebx, sizeof KEYPAIR
							inc dwNumPairs
						.endif
					.endif
				.endif

				invoke GetItemData@CDocument, m_pDoc, esi, TYPELIBCOL_IN_CLSID
				.if (eax)
					mov [ebx].KEYPAIR.pszKey,eax
					mov [ebx].KEYPAIR.pszRoot,offset g_szRootTypeLib
					add ebx, sizeof KEYPAIR
					inc dwNumPairs
				.endif

				invoke GetItemData@CDocument, m_pDoc, esi, PROGIDCOL_IN_CLSID
				.if (eax)
					mov [ebx].KEYPAIR.pszRoot,eax
					mov [ebx].KEYPAIR.pszKey,NULL
					add ebx, sizeof KEYPAIR
					inc dwNumPairs
				.endif

				invoke wsprintf,addr szKey,CStr("CLSID\%s\VersionIndependentProgID"),pszCLSID
				invoke RegOpenKey, HKEY_CLASSES_ROOT, addr szKey, addr hKey
				.if (eax == ERROR_SUCCESS)
					mov dwSize, sizeof szKey4
					invoke RegQueryValue, hKey, NULL, addr szKey4, addr dwSize
					invoke RegCloseKey, hKey
					lea eax,szKey4
					mov [ebx].KEYPAIR.pszRoot,eax
					mov [ebx].KEYPAIR.pszKey,NULL
					add ebx, sizeof KEYPAIR
					inc dwNumPairs
				.endif

				invoke wsprintf,addr szKey,CStr("CLSID\%s"),pszCLSID
				invoke RegOpenKey, HKEY_CLASSES_ROOT, addr szKey, addr hKey
				.if (eax == ERROR_SUCCESS)
					mov dwSize, sizeof szKey5
					invoke RegQueryValueEx, hKey, CStr("AppId"), NULL, addr dwType, addr szKey5, addr dwSize
					push eax
					invoke RegCloseKey, hKey
					pop eax
					.if (eax == ERROR_SUCCESS)
						lea eax,szKey5
						mov [ebx].KEYPAIR.pszKey, eax
						mov [ebx].KEYPAIR.pszRoot, CStr("AppId")
						add ebx, sizeof KEYPAIR
						inc dwNumPairs
					.endif
				.endif

			.elseif (eax == MODE_INTERFACE)

				.if (g_bExcludeProxy == 0)
					invoke GetItemData@CDocument, m_pDoc, esi,
						CLSIDCOL_IN_INTERFACE
					.if (eax && byte ptr [eax])
						mov [ebx].KEYPAIR.pszKey,eax
						mov [ebx].KEYPAIR.pszRoot,offset g_szRootCLSID
						add ebx, sizeof KEYPAIR
						inc dwNumPairs
					.endif
				.endif

				.if (g_bExcludeTypeLib == 0)
					invoke GetItemData@CDocument, m_pDoc, esi,
						TYPELIBCOL_IN_INTERFACE
					.if (eax && byte ptr [eax])
						mov [ebx].KEYPAIR.pszKey,eax
						mov [ebx].KEYPAIR.pszRoot,offset g_szRootTypeLib
						add ebx, sizeof KEYPAIR
						inc dwNumPairs
					.endif
				.endif

			.elseif (eax == MODE_HKCR)

				invoke GetItemData@CDocument, m_pDoc, esi, CLSIDCOL_IN_HKCR
				.if (eax)
					mov [ebx].KEYPAIR.pszKey,eax
					mov [ebx].KEYPAIR.pszRoot,offset g_szRootCLSID
					add ebx, sizeof KEYPAIR
					inc dwNumPairs
				.endif
;----------------------------------- if key begins with '.', add ProgId
				mov eax, pszCLSID
				.if (byte ptr [eax] == '.')
					invoke GetItemData@CDocument, m_pDoc, esi, 1
					.if (eax && (byte ptr [eax]))
						mov [ebx].KEYPAIR.pszRoot, eax
						add ebx, sizeof KEYPAIR
						inc dwNumPairs
					.endif
;----------------------------------- if subkey PersistentHandler exists, add CLSID
					invoke wsprintf,addr szKey,CStr("%s\PersistentHandler"), pszCLSID
					invoke RegOpenKey, HKEY_CLASSES_ROOT, addr szKey, addr hKey
					.if (eax == ERROR_SUCCESS)
						mov dwSize, sizeof szKey4
						invoke RegQueryValueEx, hKey, addr g_szNull, NULL, addr dwType, addr szKey4, addr dwSize
						push eax
						invoke RegCloseKey, hKey
						pop eax
						.if (eax == ERROR_SUCCESS)
							lea eax,szKey4
							mov [ebx].KEYPAIR.pszKey, eax
							mov [ebx].KEYPAIR.pszRoot, offset g_szRootCLSID
							add ebx, sizeof KEYPAIR
							inc dwNumPairs
						.endif
					.endif
				.endif

			.elseif (eax == MODE_APPID)

				invoke GetItemData@CDocument, m_pDoc, esi,
					APPIDCOL_IN_APPID
				.if (eax)
					mov kp[1*sizeof KEYPAIR].pszKey,eax
					mov kp[1*sizeof KEYPAIR].pszRoot,offset g_szRootAppID
					mov kp[1*sizeof KEYPAIR].bExpand,TRUE
					mov dwNumPairs,2
				.endif

			.endif

;------------------------------ transfer keys to edit dialog

			invoke SetKeys@CEditDlg, pEditDlg, dwNumPairs, addr kp

		.endw

;------------------------------ set a cookie if only 1 item to edit
;------------------------------ so we will be informed of changes
		.if (dwNumSelItems == 1)
			invoke SetCookie@CEditDlg, pEditDlg, esi
		.endif
;------------------------------ now show edit dialog

		invoke Show@CEditDlg, pEditDlg
if EDITMODELESS eq 0
		invoke Destroy@CEditDlg, pEditDlg
endif
		ret
		assume ebx:nothing
		align 4

OnEdit endp

;*** process WM_COMMAND/IDM_OLEREG: edit general OLE settings

OnOleReg proc uses ebx esi

local	hWnd:HWND
local	hKey:HANDLE
local	dwNumKeys:DWORD
local	kp[3]:KEYPAIR
local	szText[128]:byte

		movzx ecx,g_bConfirmDelete
		invoke Create@CEditDlg, m_hWnd, EDITMODELESS, ecx
		.if (eax)
			lea ebx, kp
			assume ebx:ptr KEYPAIR
			mov esi, eax
			invoke SetRoot@CEditDlg, esi, HKEY_LOCAL_MACHINE
			mov [ebx].pszRoot, CStr("Software\Microsoft")
			mov [ebx].pszKey, CStr("OLE")
			mov [ebx].bExpand, TRUE
			mov dwNumKeys, 1
			invoke RegOpenKeyEx, HKEY_LOCAL_MACHINE, CStr("Software\Microsoft\COM3"), NULL,\
				KEY_READ, addr hKey
			.if (eax == ERROR_SUCCESS)
				add ebx, sizeof KEYPAIR
				mov [ebx].pszRoot, CStr("Software\Microsoft")
				mov [ebx].pszKey, CStr("COM3")
				mov [ebx].bExpand, FALSE
				inc dwNumKeys
				invoke RegCloseKey, hKey
			.endif
			invoke RegOpenKeyEx, HKEY_LOCAL_MACHINE, CStr("Software\Microsoft\RPC"), NULL,\
				KEY_READ, addr hKey
			.if (eax == ERROR_SUCCESS)
				add ebx, sizeof KEYPAIR
				mov [ebx].pszRoot, CStr("Software\Microsoft")
				mov [ebx].pszKey, CStr("RPC")
				mov [ebx].bExpand, FALSE
				inc dwNumKeys
				invoke RegCloseKey, hKey
			.endif
			invoke SetKeys@CEditDlg, esi, dwNumKeys, addr kp
			invoke Show@CEditDlg, esi
			.if (eax)
				mov hWnd, eax
				invoke GetWindowText, hWnd, addr szText, sizeof szText
				invoke lstrcat, addr szText, CStr(" HKEY_LOCAL_MACHINE")
				invoke SetWindowText, hWnd, addr szText
			.endif
if EDITMODELESS eq 0
			invoke Destroy@CEditDlg, esi
endif
		.endif
		ret
		assume ebx:nothing
		align 4

OnOleReg endp

;*** process WM_COMMAND/IDM_AMBIENTPROP

OnAmbientProperties proc

		invoke Create@CAmbientPropDlg
		.if (eax)
			invoke Show@CAmbientPropDlg, eax, m_hWnd
		.endif
		ret
		align 4

OnAmbientProperties endp

;*** process WM_COMMAND/IDM_OPTION

OnOption proc

		invoke Create@COptionDlg
		.if (eax)
if 0
			push eax
			invoke Show@COptionDlg, eax, m_hWnd
			pop eax
			invoke Destroy@COptionDlg, eax
else
			invoke Show@COptionDlg, eax, m_hWnd
endif
		.endif

		ret
		align 4

OnOption endp


;*** process WM_COMMAND/IDM_TYPELIBDLG
;*** show the typelib dialog

OnTypeLibDlg proc

local	guid:GUID
local	hKey:HANDLE
local	dwSize:dword
local	iType:dword
local	iItem:dword
local	dwVerMajor
local	dwVerMinor
local	iid:IID
local	pIID:ptr IID
local	lcid:LCID
local	bTile:BOOL
local   szGUID[64]:byte
local   szGUID2[64]:byte
local   wszGUID[40]:word
local   szVersion[64]:byte
local   szStr[260]:byte

		invoke ListView_GetSelectedCount( m_hWndLV)
		.if (eax > 1)
			mov bTile, TRUE
			.if (eax > 12)
				invoke MessageBox, m_hWnd, CStr("More than 12 typelibs selected.",10,"Do you really want to open so many typelib dialogs?"), addr g_szHint, MB_YESNO
				.if (eax == IDNO)
					ret
				.endif
			.endif
		.else
			mov bTile, FALSE
		.endif
		mov iItem, -1
		.while (1)
			invoke ListView_GetNextItem( m_hWndLV, iItem, LVNI_SELECTED)
			.break .if (eax == -1)
			mov iItem, eax

			mov dwVerMajor,1
			mov dwVerMinor,0
			mov lcid,-1
			mov pIID,NULL

			mov eax,m_iMode
			.if (eax == MODE_CLSID)
				mov ecx,TYPELIBCOL_IN_CLSID
			.elseif (eax == MODE_TYPELIB)
				mov ecx,0
			.elseif (eax == MODE_INTERFACE)
				mov ecx,TYPELIBCOL_IN_INTERFACE
			.elseif (eax == MODE_HKCR)
				invoke GetItemData@CDocument, m_pDoc,
							iItem, CLSIDCOL_IN_HKCR
				.continue .if (!eax)
				invoke lstrcpy, addr szGUID2, eax
				invoke wsprintf, addr szStr, CStr("CLSID\%s\TypeLib"), addr szGUID2
				invoke RegOpenKeyEx, HKEY_CLASSES_ROOT, addr szStr, 0, KEY_READ, addr hKey
				.continue .if (eax != ERROR_SUCCESS)
				mov dwSize, sizeof szGUID
				invoke RegQueryValueEx,hKey, addr g_szNull, NULL, NULL,addr szGUID,addr dwSize
				invoke RegCloseKey, hKey
				mov ecx, -1
			.endif

			.if (ecx != -1)
;--------------- get guid of Typelib
				lea eax,szGUID
				mov byte ptr [eax],0
				ListView_GetItemText m_hWndLV, iItem, ecx, eax, sizeof szGUID
				.continue .if (!szGUID)
;--------------- get guid of coclass/interface
				lea eax,szGUID2
				ListView_GetItemText m_hWndLV, iItem, 0, eax, sizeof szGUID2
			.endif

			invoke wsprintf,addr szStr,CStr("%s\%s"), m_pszRoot, addr szGUID2

			mov eax,m_iMode
			.if (eax == MODE_CLSID || eax == MODE_HKCR)
				.if (eax == MODE_HKCR)
					invoke wsprintf,addr szStr,CStr("CLSID\%s"), addr szGUID2
				.endif
				invoke lstrcat, addr szStr, CStr("\Version")
				mov szVersion, 0
				invoke RegOpenKeyEx, HKEY_CLASSES_ROOT, addr szStr, 0, KEY_READ, addr hKey
				.if (eax == ERROR_SUCCESS)
					mov dwSize,sizeof szVersion
					invoke RegQueryValueEx,hKey,addr g_szNull,NULL,addr iType,addr szVersion,addr dwSize
					invoke RegCloseKey,hKey
				.else
					invoke wsprintf,addr szStr,CStr("TypeLib\%s"), addr szGUID
					invoke RegOpenKeyEx, HKEY_CLASSES_ROOT, addr szStr, 0, KEY_READ, addr hKey
					.if (eax == ERROR_SUCCESS)
						mov dwSize,sizeof szVersion
						invoke RegEnumKeyEx, hKey, 0, addr szVersion, addr dwSize, NULL, NULL, NULL, NULL
						invoke RegCloseKey, hKey
					.endif
				.endif
				.if (szVersion)
					invoke String22DWords,addr szVersion,addr dwVerMajor,addr dwVerMinor
				.endif

				invoke MultiByteToWideChar,CP_ACP,MB_PRECOMPOSED,
						addr szGUID2,40,addr wszGUID, 40 
				invoke IIDFromString,addr wszGUID,addr iid
				lea eax,iid
				mov pIID,eax

			.elseif (eax == MODE_INTERFACE)

				invoke lstrcat,addr szStr,CStr("\TypeLib")
				invoke RegOpenKeyEx,HKEY_CLASSES_ROOT,addr szStr,0,KEY_READ,addr hKey
				.if (eax == ERROR_SUCCESS)
					mov dwSize,sizeof szVersion
					invoke RegQueryValueEx,hKey,CStr("Version"),NULL,addr iType,addr szVersion,addr dwSize
					invoke String22DWords,addr szVersion,addr dwVerMajor,addr dwVerMinor
					invoke RegCloseKey,hKey
				.endif

				invoke MultiByteToWideChar,CP_ACP,MB_PRECOMPOSED,
						addr szGUID2,40,addr wszGUID, 40 
				invoke IIDFromString,addr wszGUID,addr iid
				lea eax,iid
				mov pIID,eax

			.elseif (eax == MODE_TYPELIB)
;-------------------------------- get Version from Column 3
				lea eax,szVersion
				ListView_GetItemText m_hWndLV, iItem, 3, eax, sizeof szVersion
				invoke String22DWords,addr szVersion,addr dwVerMajor,addr dwVerMinor
;-------------------------------- get LCID from Column 4
				lea eax,szVersion
				ListView_GetItemText m_hWndLV, iItem, 4, eax, sizeof szVersion
				invoke String2Number,addr szVersion,addr lcid,16
			.endif
		
;-------------------------------- now create and show the typelib dialog

			invoke MultiByteToWideChar,CP_ACP,MB_PRECOMPOSED, addr szGUID,-1, addr wszGUID, LENGTHOF wszGUID
			invoke CLSIDFromString, addr wszGUID,addr guid
			invoke Create@CTypeLibDlg, addr guid, dwVerMajor, dwVerMinor, lcid,pIID
			.if (eax)
				invoke Show@CTypeLibDlg, eax, m_hWnd, bTile
			.endif
		.endw
		ret
		align 4
OnTypeLibDlg endp


;--- WM_COMMAND/IDM_VIEWMONIKER


OnViewMoniker proc

local	pObjectItem:ptr CObjectItem
local	pMoniker:LPMONIKER
local	clsid:CLSID

		invoke ListView_GetNextItem( m_hWndLV, -1, LVNI_SELECTED)
		.if (eax == -1)
			ret
		.endif
		invoke GetItemData@CDocument, m_pDoc, eax, DATACOL_IN_OBJECT
		.if (eax)
			mov pObjectItem, eax
			invoke vf(eax, IObjectItem, GetMoniker)
			.if (eax)
				mov pMoniker, eax
				invoke Find@CObjectItem, pMoniker
				.if (eax)
					invoke vf(eax, IObjectItem, ShowObjectDlg), m_hWnd
				.else
					invoke vf(pMoniker, IMoniker, GetClassID), addr clsid
					invoke Create@CObjectItem, pMoniker, addr clsid
					.if (eax)
						push eax
						invoke vf(eax, IObjectItem, ShowObjectDlg), m_hWnd
						pop eax
						invoke vf(eax, IObjectItem, Release)
					.endif
				.endif
			.endif
		.endif
		ret
		align 4

OnViewMoniker endp


;*** process WM_COMMAND/IDM_OBJECTDLG
;*** shows an object

OnObjectDlg proc uses ebx

local	pROT:LPRUNNINGOBJECTTABLE
local	pUnknown:LPUNKNOWN
local	pMoniker:LPMONIKER
local	pszCLSID:LPSTR
local	pObjectItem:ptr CObjectItem
local	clsid:CLSID

;----------------------------------- get first selected item

		invoke ListView_GetNextItem( m_hWndLV, -1, LVNI_SELECTED)
		.if (eax == -1)
			ret
		.endif
		.if (m_iMode == MODE_OBJECT)
			invoke GetItemData@CDocument, m_pDoc, eax, DATACOL_IN_OBJECT

;----------------------------------- create "object" dialog
			.if (eax)
				mov pObjectItem, eax
				invoke vf(eax, IObjectItem, GetViewObjectDlg)
				.if (eax)
					invoke RestoreAndActivateWindow, [eax].CDlg.hWnd
				.else
					invoke vf(pObjectItem, IObjectItem, ShowObjectDlg), m_hWnd
				.endif
			.else
				invoke MessageBox, m_hWnd, CStr("object not found"), 0, MB_OK
			.endif


		.elseif (m_iMode == MODE_ROT)
			push eax
			invoke GetItemData@CDocument, m_pDoc, eax, 0
			mov pszCLSID, eax
			pop eax
			invoke FindROTItem@CDocument, m_pDoc, eax
			.if (eax)
				mov pMoniker, eax
if 1
				invoke GetRunningObjectTable, NULL, addr pROT
				invoke vf(pROT, IRunningObjectTable, GetObject_), pMoniker, addr pUnknown
				push eax
				invoke vf(pROT, IUnknown, Release)
				pop eax
				.if (eax == S_OK)
					mov ecx, pszCLSID
					.if (ecx)
						invoke GUIDFromLPSTR, ecx, addr clsid
						lea ecx, clsid
					.endif
					invoke Create@CObjectItem, pUnknown, ecx
					.if (eax)
						mov pObjectItem, eax
						invoke vf(eax, IObjectItem, SetFlags), OBJITEMF_INIT or OBJITEMF_ROT
						invoke vf(pObjectItem, IObjectItem, SetMoniker), pMoniker
						invoke vf(pObjectItem, IObjectItem, ShowObjectDlg), m_hWnd
						invoke vf(pObjectItem, IObjectItem, Release)
					.endif
					invoke vf(pUnknown, IUnknown, Release)
				.else
					invoke MessageBox, m_hWnd, CStr("IRunningObjectTable::GetObject failed"), 0, MB_OK
				.endif
else
				invoke BindToObject, pMoniker, TRUE
endif
				invoke vf(pMoniker, IUnknown, Release)
			.else
				invoke MessageBox, m_hWnd, CStr("Object not found"), 0, MB_OK
			.endif
		.endif
		ret
		align 4

OnObjectDlg endp

OnUnlock proc

		invoke ListView_GetNextItem( m_hWndLV, -1, LVNI_SELECTED)
		.if (eax == -1)
			ret
		.endif
		invoke GetItemData@CDocument, m_pDoc, eax, DATACOL_IN_OBJECT
		.if (eax)
			invoke vf(eax, IObjectItem, Unlock)
		.endif
		ret
		align 4

OnUnlock endp

if ?UPDATECHK

OnCheckUpdate proc public
	invoke CheckUpdate, m_hWnd
	.if (eax == TRUE)
		invoke MessageBox, m_hWnd, CStr("Newer Version of COMView is available.",10,10,"Download now?"), addr g_szHint, MB_YESNO
		.if (eax == IDYES)
			invoke ShellExecute, m_hWnd, CStr("open"), addr g_szCOMViewHttp,
				NULL, NULL, SW_SHOWNORMAL 
		.endif
	.elseif (!eax)
		invoke MessageBox, m_hWnd, CStr("No newer version of COMView available."), addr g_szHint, MB_OK
	.endif
	ret
	align 4
OnCheckUpdate endp

endif

;*** process WM_COMMAND/IDM_LOADTYPELIB
;*** load a typelib resource from a *.dll,*.tlb or *.exe file

OnLoadTypeLib proc

local	szStr1[MAX_PATH]:byte
local	szStr3[128]:byte

;------------------------------- prepare GetOpenFileName dialog

		mov szStr1,0

		invoke ZeroMemory,addr szStr3,sizeof szStr3
		invoke lstrcpy,addr szStr3,CStr("Type Libraries (*.tlb;*.dll;*.exe;*.ocx;*.olb)")
		invoke lstrlen,addr szStr3
		inc eax
		lea ecx,szStr3
		add ecx,eax
		invoke lstrcpy,ecx,CStr("*.tlb;*.dll;*.exe;*.ocx;*.olb")
	
		invoke MyGetFileName, m_hWnd, addr szStr1, sizeof szStr1, addr szStr3, sizeof szStr3, 2, NULL
		.if (eax)

;--------------------------------- create typelib dialog with filename param

			invoke Create2@CTypeLibDlg, addr szStr1, NULL, FALSE
			invoke Show@CTypeLibDlg, eax, m_hWnd, FALSE
		.endif

		ret
		align 4

OnLoadTypeLib endp


Register@CMainDlg proc public uses __this this_:ptr CMainDlg, pszFile:LPSTR

local	pszCaption:LPSTR
local	hWndDlg:HWND
local	hLib:HANDLE
local	sei:SHELLEXECUTEINFO
local	szText[MAX_PATH+32]:byte

		mov __this, this_
		mov pszCaption, NULL
		invoke lstrlen, pszFile
		mov ecx, pszFile
		mov ecx, [ecx+eax-4]
		or ecx, 20202000h
		.if (ecx == "exe.")
			mov sei.cbSize, sizeof SHELLEXECUTEINFO
			mov sei.fMask, SEE_MASK_INVOKEIDLIST 
			mov eax, m_hWnd
			mov sei.hwnd, eax 
			mov sei.lpVerb, NULL
			@mov sei.lpFile, pszFile
			mov sei.lpParameters, CStr("/RegServer")
			mov sei.lpDirectory, NULL
			mov sei.nShow, SW_SHOWDEFAULT
			mov sei.lpIDList, NULL
			invoke ShellExecuteEx, addr sei
			.if (eax)
				invoke wsprintf, addr szText, CStr("ShellExecute",10,"%s",10,"/RegServer succeeded"), pszFile
				invoke MessageBox, m_hWnd, addr szText, offset g_szHint, MB_OK
			.else
				invoke GetLastError
				invoke wsprintf, addr szText, CStr("ShellExecute",10,"%s",10,"/RegServer failed [%X]"), pszFile, eax
				invoke MessageBox, m_hWnd, addr szText, 0, MB_OK
			.endif
			jmp done
		.endif
		invoke LoadLibrary, pszFile
		mov hLib, eax
		.if (eax >= 32)
			invoke GetProcAddress, hLib, CStr("DllRegisterServer")
			.if (eax)
				call eax
				.if (eax == S_OK)
					invoke wsprintf, addr szText, CStr("Server successfully registered")
					mov pszCaption, offset g_szHint
				.else
					invoke wsprintf, addr szText, CStr("DllRegisterServer failed [%X]"), eax
				.endif
			.else
				invoke wsprintf, addr szText, CStr("Function DllRegisterServer not found in",10,"%s"), pszFile
			.endif
		.else
			invoke wsprintf, addr szText, CStr("LoadLibrary",10,"%s",10,"failed [%X]"), pszFile, eax
		.endif
		.if (szText)
			invoke MessageBox, m_hWnd, addr szText, pszCaption, MB_OK
		.endif
		.if (hLib >= 32)
			invoke FreeLibrary, hLib
		.endif
done:
		ret
		align 4

Register@CMainDlg endp

OnRegister proc

local	szStr1[MAX_PATH]:byte
local	szStr3[128]:byte

		mov szStr1,0

		invoke ZeroMemory,addr szStr3,sizeof szStr3
		invoke lstrcpy,addr szStr3,CStr("Executables (*.dll;*.ocx;*.exe)")
		invoke lstrlen,addr szStr3
		inc eax
		lea ecx,szStr3
		add ecx,eax
		invoke lstrcpy,ecx,CStr("*.dll;*.ocx;*.exe")
	
		invoke MyGetFileName, m_hWnd, addr szStr1, sizeof szStr1, addr szStr3, sizeof szStr3, 0, NULL
		.if (eax)
			invoke Register@CMainDlg, __this, addr szStr1
		.endif
		ret
		align 4

OnRegister endp

;--- bind to an object via a moniker

BindToObject proc pMoniker:LPMONIKER, bDisplayError:BOOL

local pBindCtx:LPBINDCTX
local pUnknown:LPUNKNOWN
local pStorage:LPSTORAGE
local pObjectItem:ptr CObjectItem
local pszError:LPSTR
local hr:DWORD
local clsid:CLSID
local bind_opts:BIND_OPTS

		invoke CreateBindCtx, NULL, addr pBindCtx
;------------------------------- set transacted mode
		mov bind_opts.cbStruct, sizeof BIND_OPTS
		invoke vf(pBindCtx, IBindCtx, GetBindOptions), addr bind_opts
		.if (eax == S_OK)
			or bind_opts.grfFlags, BIND_MAYBOTHERUSER
			mov bind_opts.grfMode, STGM_READWRITE or STGM_SHARE_EXCLUSIVE or STGM_TRANSACTED
			invoke vf(pBindCtx, IBindCtx, SetBindOptions), addr bind_opts
		.endif
		mov pStorage, NULL
		invoke vf(pMoniker, IMoniker, BindToStorage), pBindCtx, NULL, addr IID_IStorage, addr pStorage
		DebugOut "IMoniker::BindToStorage returned %X", eax
		.if (eax == S_OK)
;			mov hr, eax
;			mov pszError, CStr("IMoniker::BindToStorage")
			invoke ReadClassStg, pStorage, addr clsid
			mov hr, eax
			mov pszError, CStr("ReadClassStorage")
			.if (eax == S_OK)
				invoke CoCreateInstance, addr clsid, NULL, CLSCTX_SERVER, addr IID_IUnknown, addr pUnknown
				mov hr, eax
				mov pszError, CStr("CoCreateInstance")
				.if (eax == S_OK)
					invoke Create@CObjectItem, pUnknown, addr clsid
					.if (eax)
						mov pObjectItem, eax
						invoke vf(eax, IObjectItem, SetFlags), OBJITEMF_OPENVIEW
						invoke vf(pObjectItem, IObjectItem, SetStorage), pStorage
						invoke vf(pObjectItem, IObjectItem, SetMoniker), pMoniker
						invoke vf(pObjectItem, IObjectItem, ShowObjectDlg), m_hWnd
						invoke vf(pObjectItem, IObjectItem, Release)
					.endif
					invoke vf(pUnknown, IUnknown, Release)
				.endif
			.endif
			invoke vf(pStorage, IStorage, Release)
		.else
			invoke vf(pMoniker, IMoniker, BindToObject), pBindCtx, NULL, addr IID_IUnknown, addr pUnknown
			DebugOut "IMoniker::BindToObject returned %X", eax
			mov hr, eax
			mov pszError, CStr("IMoniker::BindToObject")
			.if (eax == S_OK)
				invoke Create@CObjectItem, pUnknown, NULL
				.if (eax)
					mov pObjectItem, eax
					invoke vf(pObjectItem, IObjectItem, SetFlags), OBJITEMF_INIT or OBJITEMF_OPENVIEW
					invoke vf(pObjectItem, IObjectItem, SetMoniker), pMoniker
					invoke vf(pObjectItem, IObjectItem, ShowObjectDlg), m_hWnd
					invoke vf(pObjectItem, IObjectItem, Release)
				.endif
				invoke vf(pUnknown, IUnknown, Release)
			.endif
		.endif
		invoke vf(pBindCtx, IUnknown, Release)
		.if (bDisplayError && (hr != S_OK))
			invoke OutputMessage, m_hWnd, hr, pszError, 0
		.endif
		return hr
		align 4

BindToObject endp

;*** process WM_COMMAND/IDM_LOADFILE
;*** create an embedded object from a file

LoadFile@CMainDlg proc public uses __this this_:ptr CMainDlg, pszFile:LPSTR, bDisplayError:BOOL

local hr:DWORD
local pMoniker:LPMONIKER
local wszFile[MAX_PATH]:WORD

	mov __this, this_
	invoke MultiByteToWideChar,CP_ACP,MB_PRECOMPOSED, pszFile, -1, addr wszFile, MAX_PATH
	invoke CreateFileMoniker, addr wszFile, addr pMoniker
	mov hr, eax
	.if (eax == S_OK)
		invoke BindToObject, pMoniker, bDisplayError
		mov hr, eax
		invoke vf(pMoniker, IUnknown, Release)
	.else
		.if (bDisplayError)
			invoke OutputMessage, m_hWnd, hr, CStr("CreateFileMoniker"), 0
		.endif
	.endif
	return hr
	align 4

LoadFile@CMainDlg endp

OnLoadFile proc

local szFile[MAX_PATH]:byte

	mov szFile, 0
	invoke MyGetFileName, m_hWnd, addr szFile, MAX_PATH, NULL, 0, 0, NULL
	.if (eax)
		invoke LoadFile@CMainDlg, __this, addr szFile, TRUE
	.endif
	ret
	align 4
OnLoadFile endp

;--- open object dialog for a storage object

OpenStorage@CMainDlg proc public uses esi __this this_:ptr CMainDlg, pszFile:LPSTR, pStorage:LPSTORAGE

local hr:DWORD
local pObjectDlg:ptr CObjectDlg
local wszFile[MAX_PATH]:WORD

	mov __this, this_
	.if (pszFile)
		invoke MultiByteToWideChar,CP_ACP,MB_PRECOMPOSED, pszFile, -1, addr wszFile, MAX_PATH
		invoke StgOpenStorage, addr wszFile, NULL,
			STGM_READWRITE or STGM_SHARE_EXCLUSIVE or STGM_TRANSACTED,
			NULL, NULL, addr pStorage
		mov hr, eax
		.if (eax != S_OK)
			invoke OutputMessage, m_hWnd, eax, CStr("StgOpenStorage Error"), 0
			jmp done
		.endif
	.endif
	mov esi, offset InterfaceViewerTab
	.while (dword ptr [esi])
		invoke IsEqualGUID, [esi], addr IID_IStorage
		.break .if (eax)
		add esi, 8
		xor eax, eax
	.endw
	.if (eax)
		invoke (LPVIEWPROC ptr [esi+4]), m_hWnd, pStorage, NULL
	.endif
	.if (pszFile)
		invoke vf(pStorage, IUnknown, Release)
	.endif
done:
	return hr
	align 4

OpenStorage@CMainDlg endp

;*** process WM_COMMAND/IDM_OPENSTORAGE
;*** create a storage object from a file

OnOpenStorage proc

local szFile[MAX_PATH]:byte

	mov szFile, 0
	invoke MyGetFileName, m_hWnd, addr szFile, MAX_PATH, NULL, 0, 0, NULL
	.if (eax)
		invoke OpenStorage@CMainDlg, __this, addr szFile, NULL
	.endif
	ret
	align 4
OnOpenStorage endp

OnViewStorage proc

	invoke OpenStorage@CMainDlg, __this, NULL, g_pStorage
	ret
	align 4

OnViewStorage endp

;--- open a stream object

OpenStream@CMainDlg proc public uses esi __this this_:ptr CMainDlg, pszFile:LPSTR

local hr:DWORD
local hFile:DWORD
local dwSize:DWORD
local dwRead:DWORD
local hGlobal:HANDLE
local pStream:LPSTREAM

	mov __this, this_
	invoke CreateFile, pszFile, GENERIC_READ,
			FILE_SHARE_READ or FILE_SHARE_WRITE,
			NULL, OPEN_EXISTING, 0, 0
	.if (eax != INVALID_HANDLE_VALUE)
		mov hFile,eax
		invoke GetFileSize, hFile, NULL
		mov dwSize, eax
		invoke GlobalAlloc, GMEM_MOVEABLE , eax
		.if (eax)
			mov hGlobal, eax
			invoke GlobalLock, hGlobal
			lea ecx, dwRead
			invoke ReadFile, hFile, eax, dwSize, ecx, NULL
			push eax
			invoke GlobalUnlock, hGlobal
			invoke CloseHandle, hFile
			pop eax
			.if (eax)
				invoke CreateStreamOnHGlobal, hGlobal, TRUE, addr pStream
				.if (eax == S_OK)
					invoke Create@CViewStorageDlg, pStream, pszFile, NULL
					.if (eax)
						invoke Show@CViewStorageDlg, eax, NULL
					.endif
				.else
					invoke OutputMessage, m_hWnd, eax, CStr("CreateStreamOnHGlobal"), 0
				.endif
			.else
				invoke GetLastError
				invoke SetErrorText@CMainDlg, __this, CStr("ReadFile failed [%X]"), eax, TRUE
			.endif
		.else
			invoke CloseHandle, hFile
			invoke SetErrorText@CMainDlg, __this, CStr("GlobalAlloc failed"), eax, TRUE
		.endif
	.else
		invoke GetLastError
		invoke SetErrorText@CMainDlg, __this, CStr("CreateFile failed [%X]"), eax, TRUE
	.endif
	return hr
	align 4

OpenStream@CMainDlg endp


;*** process WM_COMMAND/IDM_VIEWSTREAM

OnViewStream proc

local	pObjectDlg:ptr CObjectDlg
local	pUnknown:LPUNKNOWN

	invoke Create@CViewStorageDlg, g_pStream, NULL, NULL
	.if (eax)
		invoke Show@CViewStorageDlg, eax, NULL
	.endif
	ret
	align 4

OnViewStream endp

;*** process WM_COMMAND/IDM_OPENSTREAM
;*** create a storage object from a file

OnOpenStream proc

local szFile[MAX_PATH]:byte

	mov szFile, 0
	invoke MyGetFileName, m_hWnd, addr szFile, MAX_PATH, NULL, 0, 0, NULL
	.if (eax)
		invoke OpenStream@CMainDlg, __this, addr szFile
	.endif
	ret
	align 4
OnOpenStream endp


;*** process WM_COMMAND/IDM_CREATELINK
;*** create a linked object from a file

CreateLink@CMainDlg proc public uses __this this_:ptr CMainDlg, pszFile:LPSTR

local pObjectItem:LPOBJECTITEM
local pOleLink:LPOLELINK
local pMoniker:LPMONIKER
local pStorage:LPSTORAGE
local pOleObject:LPOLEOBJECT
local wszFile[MAX_PATH]:WORD

	mov __this, this_
	invoke MultiByteToWideChar,CP_ACP,MB_PRECOMPOSED, pszFile, -1, addr wszFile, MAX_PATH
	invoke StgCreateDocfile, NULL,
			STGM_READWRITE or STGM_SHARE_EXCLUSIVE or STGM_TRANSACTED or STGM_DELETEONRELEASE,
			NULL,addr pStorage
	.if (eax != S_OK)
		invoke OutputMessage, m_hWnd, eax, CStr("StgCreateDocFile"), 0
		jmp done
	.endif
	invoke OleCreateLinkToFile, addr wszFile, addr IID_IOleObject, OLERENDER_DRAW, NULL,\
		NULL, pStorage, addr pOleObject
	.if (eax == S_OK)
		invoke Create@CObjectItem, pOleObject, NULL
		.if (eax)
			mov pObjectItem, eax
			invoke vf(eax, IObjectItem, SetStorage), pStorage
			invoke vf(pOleObject, IUnknown, QueryInterface), addr IID_IOleLink, addr pOleLink
			.if (eax == S_OK)
				invoke vf(pOleLink, IOleLink, GetSourceMoniker), addr pMoniker
				.if (eax == S_OK)
					invoke vf(pObjectItem, IObjectItem, SetMoniker), pMoniker
					invoke vf(pMoniker, IUnknown, Release)
				.endif
				invoke vf(pOleLink, IUnknown, Release)
			.endif
			invoke vf(pObjectItem, IObjectItem, ShowObjectDlg), m_hWnd
			invoke vf(pObjectItem, IObjectItem, Release)
		.endif
		invoke vf(pOleObject, IUnknown, Release)
	.else
		invoke OutputMessage, m_hWnd, eax, CStr("OleCreateLinkToFile"), 0
	.endif
	invoke vf(pStorage, IStorage, Release)
done:
	ret
	align 4
CreateLink@CMainDlg endp

OnCreateLink proc

local szFile[MAX_PATH]:byte

	mov szFile, 0
	invoke MyGetFileName, m_hWnd, addr szFile, MAX_PATH, NULL, 0, 0, NULL
	.if (eax)
		invoke CreateLink@CMainDlg, __this, addr szFile
	.endif
	ret
	align 4
OnCreateLink endp

OnSHDesktopFolder proc

local pUnknown:LPUNKNOWN

	invoke SHGetDesktopFolder, addr pUnknown
	.if (eax == S_OK)
		invoke ShowObject, pUnknown
		invoke vf(pUnknown, IUnknown, Release)
	.endif
	ret
	align 4

OnSHDesktopFolder endp


OnViewMalloc proc

local pUnknown:LPUNKNOWN

	invoke CoGetMalloc, 1, addr pUnknown
	.if (eax == S_OK)
		invoke ShowObject, pUnknown
		invoke vf(pUnknown, IUnknown, Release)
	.endif
	ret
	align 4

OnViewMalloc endp

;*** process WM_COMMAND/IDM_SAVEAS
;*** writes content of listview into a file


OnSaveAs proc

local	szStr1[MAX_PATH]:byte
local	szStr3[128]:byte

;------------------------------- prepare GetSaveFileName dialog

		mov szStr1,0

		invoke ZeroMemory, addr szStr3, sizeof szStr3
		invoke lstrcpy, addr szStr3,CStr("Text files (*.txt)")
		invoke lstrlen, addr szStr3
		inc eax
		lea ecx,szStr3
		add ecx,eax
		invoke lstrcpy, ecx, CStr("*.txt")

		invoke MyGetFileName, m_hWnd, addr szStr1, sizeof szStr1, addr szStr3, sizeof szStr3, 1, NULL
		.if (eax)

;------------------------------- we use the "progress" dialog
;------------------------------- this will start a second thread to save data

			invoke Create@CProgressDlg, m_hWndLV, addr szStr1, SAVE_DISK, m_iNumCols
			invoke DialogBoxParam, g_hInstance, IDD_PROGRESSDLG, m_hWnd, classdialogproc, eax
		.endif
		ret
		align 4

OnSaveAs endp


;*** process WM_COMMAND/IDM_COPY
;*** copies selected lines in listview to the clipboard


OnCopy proc

;------------------------------- we use the "progress" dialog
;------------------------------- this will start a second thread to write data

		invoke Create@CProgressDlg, m_hWndLV, NULL, SAVE_CLIPBOARD, m_iNumCols
		invoke DialogBoxParam, g_hInstance, IDD_PROGRESSDLG, m_hWnd, classdialogproc, eax

		ret
		align 4

OnCopy endp


OnSelectAll proc

		ListView_SetItemState m_hWndLV, -1 , LVIS_SELECTED, LVIS_SELECTED
		invoke SetStatusPane2
		invoke UpdateMenu
		ret
		align 4

OnSelectAll endp



OnInvertSelection proc uses ebx esi

local	hCsrOld:dword

		invoke SetCursor,g_hCsrWait
		mov hCsrOld, eax
		invoke SetBusyState@CMainDlg, TRUE
		invoke SetWindowRedraw( m_hWndLV, FALSE)
		invoke ListView_GetItemCount( m_hWndLV)
		mov ebx, eax
		mov esi, 0
		.while (esi < ebx)
			invoke ListView_GetItemState( m_hWndLV, esi, LVIS_SELECTED)
			xor eax, LVIS_SELECTED
			ListView_SetItemState m_hWndLV, esi, eax, LVIS_SELECTED
			inc esi
		.endw
		invoke SetWindowRedraw( m_hWndLV, TRUE)
		invoke SetStatusPane1
		invoke SetStatusPane2
		invoke UpdateMenu
		invoke SetCursor, hCsrOld
		invoke SetBusyState@CMainDlg, FALSE
		ret
		align 4

OnInvertSelection endp


OnPaste proc

local	pDataObject:LPDATAOBJECT
local	pOleObject:LPOLEOBJECT
local	pUnknown:LPUNKNOWN
local	pStorage:LPSTORAGE
local	pObjectItem:LPOBJECTITEM
local	szText[128]:byte

		invoke OleGetClipboard, addr pDataObject
		.if (eax == S_OK)
			mov pOleObject, NULL
			mov pStorage, NULL
			invoke OleQueryCreateFromData, pDataObject
			.if (eax == S_OK)
				invoke StgCreateDocfile, NULL,
					STGM_READWRITE or STGM_SHARE_EXCLUSIVE or STGM_TRANSACTED or STGM_DELETEONRELEASE,
					NULL, addr pStorage
				invoke OleCreateFromData, pDataObject, addr IID_IOleObject,
					OLERENDER_DRAW, NULL, NULL, pStorage, addr pOleObject
				.if (eax == S_OK)
					mov ecx, pOleObject
				.else
					push eax
					invoke SetErrorText@CMainDlg, __this, CStr("OleCreateFromData failed [%X]"), eax, TRUE
					pop eax
				.endif
			.elseif (eax == OLE_S_STATIC)
				invoke StgCreateDocfile, NULL,
					STGM_READWRITE or STGM_SHARE_EXCLUSIVE or STGM_TRANSACTED or STGM_DELETEONRELEASE,
					NULL, addr pStorage
				invoke OleCreateStaticFromData, pDataObject, addr IID_IOleObject,
					OLERENDER_DRAW, NULL, NULL, pStorage, addr pOleObject
				.if (eax == S_OK)
					mov ecx, pOleObject
				.else
					push eax
					invoke SetErrorText@CMainDlg, __this, CStr("OleCreateStaticFromData failed [%X]"), eax, TRUE
					pop eax
				.endif
			.else
				mov eax, S_OK
				mov ecx, pDataObject
			.endif
			.if (eax == S_OK)
				mov pUnknown, ecx
				invoke Create@CObjectItem, pUnknown, NULL
				.if (eax)
					mov pObjectItem, eax
					.if (pOleObject)
						invoke vf(eax, IObjectItem, SetFlags), OBJITEMF_OPENVIEW
					.endif
					.if (pStorage)
						invoke vf(pObjectItem, IObjectItem, SetStorage), pStorage
					.endif
					invoke vf(pObjectItem, IObjectItem, ShowObjectDlg), m_hWnd
					invoke vf(pObjectItem, IObjectItem, Release)
				.endif
			.endif
			.if (pOleObject)
				invoke vf(pOleObject, IOleObject, Release)
			.endif
			.if (pStorage)
				invoke vf(pStorage, IStorage, Release)
			.endif
			invoke vf(pDataObject, IDataObject, Release)
		.endif
		ret
		align 4
OnPaste endp

OnRemoveItem proc uses ebx

		invoke SetCursor,g_hCsrWait
		mov ebx,eax
		invoke SetBusyState@CMainDlg, TRUE
		invoke SetWindowRedraw( m_hWndLV, FALSE)
		.while (1)
			invoke ListView_GetNextItem( m_hWndLV, -1, LVNI_SELECTED)
			.break .if (eax == -1)
			push eax
			invoke RemoveItem@CDocument, m_pDoc, eax
			pop eax
			invoke ListView_DeleteItem( m_hWndLV, eax)
		.endw
		invoke SetWindowRedraw( m_hWndLV, TRUE)
		invoke SetStatusPane1
		invoke SetCursor, ebx
		invoke SetBusyState@CMainDlg, FALSE
		ret
		align 4
OnRemoveItem endp

;*** process WM_COMMAND/IDM_COPYGUID
;*** copies GUID to clipboard


OnCopyGUID proc


		invoke ListView_GetNextItem( m_hWndLV, -1, LVNI_SELECTED)
		.if (eax != -1)
			.if (m_iMode == MODE_HKCR)
				mov ecx, CLSIDCOL_IN_HKCR
			.elseif (m_iMode == MODE_APPID)
				push eax
;----------------------------------- does [AppId] entry exist?
				invoke GetItemData@CDocument, m_pDoc, eax, APPIDCOL_IN_APPID
				mov ecx, eax
				pop eax
				.if (ecx)
					mov ecx, APPIDCOL_IN_APPID
				.else
					mov ecx, 0
				.endif
			.else
				mov ecx, 0
			.endif
			invoke GetItemData@CDocument, m_pDoc, eax, ecx
			.if (eax)
				invoke CopyStringToClipboard, m_hWnd, eax
			.endif
		.endif
		ret
		align 4
OnCopyGUID endp


;*** Find/Replace hook: just used to update global var g_hWndDlg


FRHook proc hWnd:HWND, uMsg:DWORD, wParam:WPARAM, lParam:LPARAM

		xor eax,eax
		.if (uMsg == WM_INITDIALOG)
			mov eax,1
		.elseif (uMsg == WM_ACTIVATE)
			movzx ecx,word ptr wParam
			.if (ecx == WA_INACTIVE)
				mov g_hWndDlg,NULL
			.else
				mov ecx,hWnd
				mov g_hWndDlg,ecx
			.endif
		.endif
		ret
		align 4

FRHook endp

;*** process WM_COMMAND/IDM_FIND
;*** Find a string in listview

FINDMSGSTRING textequ  <CStr("commdlg_FindReplace")>

OnFindCommon proc uses ebx dwMode:DWORD

local	pfr:ptr FINDREPLACE

		.if (g_hWndFind)
			invoke ShowWindow, g_hWndFind, SW_RESTORE
			ret
		.endif
		invoke malloc, sizeof FINDREPLACE
		.if (!eax)
			ret
		.endif
		mov ebx,eax
		assume ebx:ptr FINDREPLACE

		mov [ebx].lStructSize, sizeof FINDREPLACE
		mov eax, m_hWnd
		mov [ebx].hwndOwner, eax
		mov [ebx].lpfnHook, FRHook
		mov eax, dwMode
		mov [ebx].lCustData, eax
		mov [ebx].Flags,FR_HIDEUPDOWN or FR_HIDEWHOLEWORD or FR_ENABLEHOOK
		mov eax,g_lpstrFind
		.if (eax == NULL)
	 		invoke malloc,MAXFINDLEN
			.if (!eax)
				ret
			.endif
			mov g_lpstrFind,eax
			mov byte ptr [eax],0
		.endif
		mov [ebx].lpstrFindWhat,eax
		mov [ebx].wFindWhatLen,MAXFINDLEN

		.if (!g_uMsgFind)
			invoke RegisterWindowMessage, FINDMSGSTRING
			mov g_uMsgFind,eax
		.endif

		invoke FindText, ebx
		mov g_hWndFind,eax

		ret
		align 4
		assume ebx:nothing

OnFindCommon endp

OnFind proc 
		invoke OnFindCommon, 0
		ret
		align 4
OnFind endp

;*** process WM_COMMAND/IDM_FINDALL
;*** remove lines which dont match

OnFindAll proc
		invoke OnFindCommon, 1
		ret
		align 4
OnFindAll endp

;*** Do the real "find" work here. But keep it simple.
;*** No need for sophisticated find features here.

SEARCH_MATCHCASE	equ 1
SEARCH_WHOLEITEM	equ 2
SEARCH_WRAP			equ 4
SEARCH_FINDALL		equ 8

SearchString proc uses ebx esi lpszFindWhat:LPSTR, dwFlags:DWORD, iItemStart:DWORD, iColumn:DWORD

local	iMax:DWORD
local	iStrLength:DWORD
local	szText[260]:byte
local	szFindText[MAXFINDLEN]:byte
local	bFound:BOOL
local	iColumnStart:DWORD
local	iColumnEnd:DWORD

	invoke lstrcpy, addr szFindText, lpszFindWhat

	.if (!(dwFlags & SEARCH_MATCHCASE))
		invoke CharUpper, addr szFindText
	.endif

	invoke lstrlen, addr szFindText
	mov iStrLength,eax				;length of string to find

	mov eax, iColumn
	.if (eax == -1)
		mov ecx, m_iNumCols
		xor eax, eax
	.else
		mov ecx, eax
		inc ecx
	.endif
	mov iColumnStart, eax
	mov iColumnEnd, ecx

	mov ebx,iItemStart				;search start
	.if (dwFlags & SEARCH_FINDALL)
		ListView_SetItemState m_hWndLV, -1 , 0, LVIS_SELECTED
		xor ebx, ebx
	.endif

	invoke ListView_GetItemCount( m_hWndLV)
	mov iMax,eax					;number of lines in listview

	mov bFound,FALSE
nextpass:
	.while (ebx < iMax && (bFound == FALSE))
		mov esi, iColumnStart
		.while (esi < iColumnEnd && (bFound == FALSE))
			mov szText,0
			invoke GetItemData@CDocument, m_pDoc, ebx, esi
			.if (eax)
				invoke lstrcpy, addr szText, eax
			.endif
			.if (!(dwFlags & SEARCH_MATCHCASE))
				invoke CharUpper,  addr szText
			.endif
			invoke lstrlen, addr szText
			mov ecx,eax
			.if ((!(dwFlags & SEARCH_WHOLEITEM)) && ecx)
				mov ecx, 1
			.endif
			push edi
			push esi
			lea edi,szText
			lea esi,szFindText
			.while (ecx >= iStrLength)
				mov al,[edi]
				.if (al == [esi])
					pushad
					mov ecx,iStrLength
					repz cmpsb
					popad
					.if (ZERO?)
						mov bFound,TRUE
						.break
					.endif
				.endif
				inc edi			;next character
				dec ecx
			.endw
			pop esi
			pop edi
			inc esi				;next column
		.endw
		.if (dwFlags & SEARCH_FINDALL)
			.if (bFound == FALSE)
				ListView_SetItemState m_hWndLV, ebx , LVIS_SELECTED, LVIS_SELECTED
			.else
				mov bFound, FALSE
			.endif
		.endif
		inc ebx					;next line
	.endw
	.if ((!bFound) && (dwFlags & SEARCH_WRAP))
		and dwFlags, not SEARCH_WRAP
		mov eax, iItemStart
		mov iMax, eax
		xor ebx, ebx
		jmp nextpass
	.endif

	mov eax, -1
	.if (bFound)
		dec ebx
		mov eax, ebx
	.endif
	ret
	align 4

SearchString endp


;*** received FINDMSGSTRING from "FindText" dialog


OnFindMsg proc uses ebx lpfr:ptr FINDREPLACE

		mov ebx,lpfr
		assume ebx:ptr FINDREPLACE

		.if ([ebx].Flags & FR_DIALOGTERM)
			invoke free, ebx
			mov g_hWndFind,NULL
		.elseif ([ebx].Flags & FR_FINDNEXT)
			invoke ListView_GetNextItem( m_hWndLV, -1, LVNI_SELECTED)
			inc eax
			mov ecx, SEARCH_WHOLEITEM
			.if ([ebx].Flags & FR_MATCHCASE)
				or ecx, SEARCH_MATCHCASE
			.endif
			.if ([ebx].lCustData)
				or ecx, SEARCH_FINDALL
			.endif
			invoke SearchString, [ebx].lpstrFindWhat, ecx, eax, -1
			.if ([ebx].lCustData)
				invoke OnRemoveItem
			.elseif (eax != -1)
				mov ebx, eax
;---------------------------- reset all selections first
				ListView_SetItemState m_hWndLV, -1 , 0, LVIS_SELECTED
;---------------------------- now select,focus+make visible item found
				invoke ListView_EnsureVisible( m_hWndLV, ebx, FALSE)
				ListView_SetItemState m_hWndLV, ebx , LVIS_FOCUSED or LVIS_SELECTED, LVIS_FOCUSED or LVIS_SELECTED
			.else
				invoke MessageBox, m_hWnd, CStr("No (more) occurances found"), 0, MB_OK
			.endif
		.endif
		ret
		assume ebx:nothing
		align 4

OnFindMsg endp

OnFindNextCross proc uses ebx

local dwItems:DWORD
local dwStart:DWORD

		invoke ListView_GetNextItem( m_hWndLV, -1, LVNI_SELECTED)
		mov ebx, eax
		mov dwStart, eax
		invoke GetItemCount@CDocument, m_pDoc
		mov dwItems, eax
		inc ebx
		.while (ebx < dwItems)
			invoke GetItemFlag@CDocument, m_pDoc, ebx, FLAG_IMAGE
			.if (eax & FLAG_IMAGE)
				.if (dwStart != -1)
					ListView_SetItemState m_hWndLV, dwStart, 0, LVIS_SELECTED
				.endif
				invoke ListView_EnsureVisible( m_hWndLV, ebx, FALSE)
				ListView_SetItemState m_hWndLV, ebx, LVIS_SELECTED, LVIS_SELECTED
				jmp done
			.endif
			inc ebx
		.endw
		invoke MessageBeep, MB_OK
done:
		ret
		align 4

OnFindNextCross endp


;--- open explorer window at path location

;;LPITEMIDLIST typedef ptr
protoSHParseDisplayName typedef proto :ptr WORD, :DWORD, :ptr LPITEMIDLIST, :DWORD, :ptr DWORD
LPSHPARSEDISPLAYNAME typedef ptr protoSHParseDisplayName
protoSHOpenFolderAndSelectItems typedef proto :LPITEMIDLIST, :DWORD, :ptr LPITEMIDLIST, :DWORD
LPSHOPENFOLDERANDSELECTITEMS typedef ptr protoSHOpenFolderAndSelectItems

OnExplore proc uses ebx esi

local	pfnSHParseDisplayName:LPSHPARSEDISPLAYNAME
local	pfnSHOpenFolderAndSelectItems:LPSHOPENFOLDERANDSELECTITEMS
local	pidl:LPVOID
local	pidl2:LPVOID
local	pMalloc:LPMALLOC
local	dwSFGAO:DWORD
local	hShell32:HANDLE
local	pszFile:LPSTR
local	lvi:LVITEM
local	sei:SHELLEXECUTEINFO
local	szPath[MAX_PATH]:byte
local	szPath2[MAX_PATH]:byte
		

		.if (m_iMode == MODE_CLSID)
			mov lvi.iSubItem, PATHCOL_IN_CLSID
		.elseif (m_iMode == MODE_TYPELIB)
			mov lvi.iSubItem, PATHCOL_IN_TYPELIB
		.elseif (m_iMode == MODE_HKCR)
			mov lvi.iSubItem, PATHCOL_IN_HKCR
		.else
			jmp error
		.endif
		invoke ListView_GetNextItem( m_hWndLV, -1, LVNI_SELECTED)
		.if (eax == -1)
			jmp error
		.endif
		mov lvi.iItem, eax

		.if (m_iMode == MODE_CLSID)
			invoke IsFileLink, eax
			.if (!eax)
				jmp error
			.endif
		.endif

		lea eax, szPath2
		mov lvi.pszText, eax
		mov lvi.cchTextMax,SIZEOF szPath2
		mov lvi.mask_,LVIF_TEXT
		invoke ListView_GetItem( m_hWndLV, addr lvi)
		lea ebx, szPath
		invoke ExpandEnvironmentStrings, addr szPath2, ebx, MAX_PATH
;--------------------------- WinNT bug, so dont rely on return code (MSDN Q234874)
		invoke lstrlen, ebx
;-------------------------------- skip ""
		.if ((eax > 1 ) && (byte ptr [ebx] == '"'))
			.if (m_iMode == MODE_HKCR)
				inc ebx
				dec eax
				xor ecx, ecx
				.while (ecx < eax)
					.if (byte ptr [ebx+ecx] == '"')
						mov byte ptr [ebx+ecx],0
						mov eax, ecx
						.break
					.endif
					inc ecx
				.endw
			.elseif (byte ptr [ebx+eax-1] == '"')
				mov byte ptr [ebx+eax-1],0
				inc ebx
				dec eax
				dec eax
			.endif
		.endif
		.if (!eax)
			jmp error
		.endif

		xor edx, edx
		.while (edx < eax)
			mov ecx, [ebx+edx]
			or ecx,20202000h
			.if ((ecx == "exe.") && (byte ptr [ebx+edx+4] == ' '))
				mov byte ptr [ebx+edx+4],0
				lea eax, [edx+4]
				.break
			.endif
			inc edx
		.endw

;-------------------------- ebx -> start of path
;-------------------------- eax -> length of path

		mov pszFile, NULL
		.while (eax)
			mov cl, [ebx+eax-1]
			.if (cl == '\')
				lea edx, [ebx+eax-1]
				mov byte ptr [edx],0
				inc edx
				invoke lstrcpy, addr szPath2, edx
				lea eax, szPath2
				mov pszFile, eax
				.break
			.endif
			dec eax
		.endw
		.if (!eax)
			mov pszFile, ebx
			lea ebx, szPath2
			invoke GetSystemDirectory, ebx, sizeof szPath2
		.else
			invoke GetFileAttributes, ebx
			.if ((eax != -1) && (!(eax & FILE_ATTRIBUTE_DIRECTORY)))
				invoke lstrlen, ebx
				.while (eax)
					mov cl, [ebx+eax-1]
					.if (cl == '\')
						lea edx, [ebx+eax-1]
						mov byte ptr [edx],0
						inc edx
						invoke lstrcpy, addr szPath2, edx
						.break
					.endif
					dec eax
				.endw
			.endif
		.endif

;------------------------- ebx -> directory
;------------------------- pszFile -> file

		invoke GetModuleHandle, CStr("shell32")
		.if (eax)
			mov hShell32, eax
			invoke GetProcAddress, eax, CStr("SHParseDisplayName")
			.if (eax)
				mov pfnSHParseDisplayName, eax
				invoke GetProcAddress, hShell32, CStr("SHOpenFolderAndSelectItems")
				mov pfnSHOpenFolderAndSelectItems, eax
			.endif
		.endif

		.if (eax)
			sub esp, MAX_PATH
			mov edx, esp
			invoke wsprintf, edx, CStr("%s\%s"), ebx, pszFile
			mov edx, esp
			sub esp, MAX_PATH * 2
			mov esi, esp
			invoke MultiByteToWideChar,CP_ACP,MB_PRECOMPOSED, edx, -1, esi, MAX_PATH 
			invoke pfnSHParseDisplayName, esi, NULL, addr pidl, NULL, addr dwSFGAO
			add esp, MAX_PATH * 3
			.if (eax == S_OK)
				invoke pfnSHOpenFolderAndSelectItems, pidl, 0, NULL, NULL
				.if (eax != S_OK)
					invoke OutputMessage, m_hWnd, eax, CStr("SHOpenFolderAndSelectItems"), 0
				.endif
				invoke SHGetMalloc, addr pMalloc
				.if (eax == S_OK)
					invoke vf(pMalloc, IMalloc, Free), pidl
					invoke vf(pMalloc, IMalloc, Release)
				.endif
				jmp done
			.endif
		.endif

		mov sei.cbSize, sizeof SHELLEXECUTEINFO
		mov sei.fMask, 0
		mov eax, m_hWnd
		mov sei.hwnd, eax 
		mov sei.lpVerb, CStr("open")
		mov	sei.lpFile, ebx
		mov sei.lpParameters, NULL
		mov sei.lpDirectory, NULL
		mov sei.nShow, SW_SHOWDEFAULT
		invoke ShellExecuteEx, addr sei
done:
		ret
error:
		invoke MessageBeep, MB_OK
		ret
		align 4

OnExplore endp


;--- Properties/Unregister


FileOperation@CMainDlg proc public uses ebx __this thisarg, dwOperation:DWORD, pszFile:LPSTR

local	hLib:HANDLE
local	pszCaption:LPSTR
local	pProc:DWORD
local	pszChar:LPSTR
local	cSaved:BYTE
local	bDll:BOOL
local	lvi:LVITEM
local	sei:SHELLEXECUTEINFO
local	szPath[MAX_PATH]:byte
local	szPath2[MAX_PATH]:byte
local	szText[MAX_PATH+32]:byte
		
		mov __this,this@
		.if (!pszFile)
			.if (m_iMode == MODE_CLSID)
				mov lvi.iSubItem, PATHCOL_IN_CLSID
			.elseif (m_iMode == MODE_TYPELIB)
				mov lvi.iSubItem, PATHCOL_IN_TYPELIB
			.else
				jmp error
			.endif
			invoke ListView_GetNextItem( m_hWndLV, -1, LVNI_SELECTED)
			.if (eax == -1)
				jmp error
			.endif
			mov lvi.iItem, eax

			.if (m_iMode == MODE_CLSID)
				invoke IsFileLink, eax
				.if (!eax)
					jmp error
				.endif
			.endif

			lea eax, szPath2
			mov lvi.pszText, eax
			mov lvi.cchTextMax,SIZEOF szPath
			mov lvi.mask_,LVIF_TEXT
			invoke ListView_GetItem( m_hWndLV, addr lvi)
			lea ebx, szPath
			invoke ExpandEnvironmentStrings, addr szPath2, ebx, MAX_PATH
		.else
			mov ebx, pszFile
		.endif

;--------------------------- WinNT bug, so dont rely on return code (MSDN Q234874)
		invoke lstrlen, ebx
;-------------------------------- skip ""
		.if ((eax > 1 ) && (byte ptr [ebx] == '"') && (byte ptr [ebx+eax-1] == '"'))
			mov byte ptr [ebx+eax-1],0
			inc ebx
			dec eax
			dec eax
		.endif
		.if (!eax)
			jmp error
		.endif
		xor edx, edx
		mov bDll, TRUE
		.while (edx < eax)
			mov ecx, [ebx+edx]
			or ecx,20202000h
			.if ((ecx == "exe.") && (byte ptr [ebx+edx+4] <= ' '))
				mov byte ptr [ebx+edx+4],0
				mov bDll, FALSE
				lea eax, [edx+4]
				.break
			.endif
			inc edx
		.endw
		.while (eax)
			mov cl, [ebx+eax-1]
			.if (cl == '\')
				.break
			.endif
			dec eax
		.endw
		.if (!eax)
			invoke GetSystemDirectory, addr szPath2, sizeof szPath2
			lea ecx, szPath2
			mov word ptr [ecx+eax],'\'
			invoke lstrcat, addr szPath2, ebx
			lea ebx, szPath2
		.endif
		invoke GetFileAttributes, ebx
		.if (eax == -1)
			invoke lstrlen, ebx
			mov pszChar, NULL
			.while (eax)
				lea edx, [ebx+eax-1]
				mov cl, [edx]
;------------------------------------- this is for TYPELIB resource
				.if (cl == '\')
					xor al, al
					xchg al, [edx]
					mov cSaved, al
					mov pszChar, edx
					.break
				.endif
				dec eax
			.endw
			invoke GetFileAttributes, ebx
			.if ((eax == -1) || (eax & FILE_ATTRIBUTE_DIRECTORY))
				.if (pszChar)
					mov edx, pszChar
					mov al, cSaved
					mov [edx],al
				.endif
				invoke wsprintf, addr szText, CStr("File",10,"%s",10,"not found"), ebx
				invoke MessageBox, m_hWnd, addr szText, NULL, MB_OK
				jmp done
			.endif
		.endif

		.if (dwOperation == FILEOP_PROPERTIES)
			mov sei.cbSize, sizeof SHELLEXECUTEINFO
			mov sei.fMask, SEE_MASK_INVOKEIDLIST 
			mov eax, m_hWnd
			mov sei.hwnd, eax 
			mov sei.lpVerb, CStr("properties")
			mov sei.lpFile, ebx
			mov sei.lpParameters, NULL
			mov sei.lpDirectory, NULL
			mov sei.nShow, SW_SHOWDEFAULT
			mov sei.lpIDList, NULL
			invoke ShellExecuteEx, addr sei
		.elseif (dwOperation == FILEOP_UNREGISTER)
			mov pszCaption, NULL

;-------------------------------------- set current directory to path of dll
;-------------------------------------- so dependant dll may be found
			invoke lstrcpy, addr szText, ebx
			invoke lstrlen, ebx
			lea ecx, szText
			.while (eax)
				dec eax
				.if (byte ptr [eax+ecx] == '\')
					mov byte ptr [eax+ecx],0
					.break
				.endif
			.endw
			invoke SetCurrentDirectory, addr szText

			.if (bDll)
				invoke LoadLibrary, ebx
				mov hLib, eax
				.if (eax >= 32)
					invoke GetProcAddress, hLib, CStr("DllUnregisterServer")
					.if (eax)
						mov pProc, eax
						invoke MessageBox, m_hWnd, CStr("Are you sure?"), CStr("Unregister Server"),
							MB_YESNO or MB_DEFBUTTON2 or MB_ICONQUESTION
						.if (eax == IDYES)
							call pProc
							.if (eax == S_OK)
								invoke wsprintf, addr szText, CStr("Server successfully unregistered")
								mov pszCaption, offset g_szHint
							.else
								invoke wsprintf, addr szText, CStr("DllUnregisterServer failed [%X]"), eax
							.endif
						.else
							mov szText, 0
						.endif
					.else
						invoke wsprintf, addr szText, CStr("Function DllUnregisterServer not found in",10,"%s"), ebx
					.endif
				.else
					invoke wsprintf, addr szText, CStr("LoadLibrary",10,"%s",10,"failed [%X]"), ebx, eax
				.endif
				.if (szText)
					invoke MessageBox, m_hWnd, addr szText, pszCaption, MB_OK
				.endif
				.if (hLib >= 32)
					invoke FreeLibrary, hLib
				.endif
			.else
				invoke MessageBox, m_hWnd, CStr("Are you sure?"), CStr("Unregister Server"),\
					MB_YESNO or MB_DEFBUTTON2 or MB_ICONQUESTION
				.if (eax == IDYES)
					mov sei.cbSize, sizeof SHELLEXECUTEINFO
					mov sei.fMask, SEE_MASK_INVOKEIDLIST 
					mov eax, m_hWnd
					mov sei.hwnd, eax 
					mov sei.lpVerb, NULL
					mov sei.lpFile, ebx
					mov sei.lpParameters, CStr("/UnregServer")
					mov sei.lpDirectory, NULL
					mov sei.nShow, SW_SHOWDEFAULT
					mov sei.lpIDList, NULL
					invoke ShellExecuteEx, addr sei
					.if (eax)
						invoke wsprintf, addr szText, CStr("ShellExecute",10,"%s",10,"/UnregServer succeeded"), ebx
						invoke MessageBox, m_hWnd, addr szText, offset g_szHint, MB_OK
					.else
						invoke GetLastError
						invoke wsprintf, addr szText, CStr("ShellExecute",10,"%s",10,"/UnregServer failed [%X]"), ebx, eax
						invoke MessageBox, m_hWnd, addr szText, 0, MB_OK
					.endif
				.endif
			.endif
		.endif
		ret
error:
		invoke MessageBeep, MB_OK
done:
		ret
		align 4

FileOperation@CMainDlg endp

;--- open properties window at file location

OnProperties proc
		invoke FileOperation@CMainDlg, __this, FILEOP_PROPERTIES, NULL
		ret
		align 4
OnProperties endp

;--- unregister DLL

OnUnregister proc
		invoke FileOperation@CMainDlg, __this, FILEOP_UNREGISTER, NULL
		ret
		align 4
OnUnregister endp


;*** do listview sort. called from WM_NOTIFY/LVN_COLUMNCLICK
;*** sorting is a job for the document, but we must renew
;*** all selections + focus after the sort


SortListView proc uses ebx

local	dwItems:DWORD
local	hCsrOld:HCURSOR

		invoke SetCursor, g_hCsrWait
		mov hCsrOld,eax

;------------------------ call document to do the sort
		mov eax, m_iSortCol
		mov edx, m_pMode
		mov edx,[edx].CMode.pColDesc
		movzx eax,[edx + eax * sizeof CColHdr].CColHdr.wFlags
		and eax, FCOLHDR_RDXMASK
;------------------------ ax=0 -> string format, else 10/16 (radix)
		invoke Sort@CDocument, m_pDoc, m_iSortCol, m_iSortDir, eax

;------------------------ set item states
		invoke GetItemCount@CDocument, m_pDoc
		mov dwItems, eax
		xor ebx, ebx
		.while (ebx < dwItems)
			invoke GetItemFlag@CDocument, m_pDoc, ebx,
				LVIS_SELECTED or LVIS_FOCUSED
			movzx eax, al
			ListView_SetItemState m_hWndLV, ebx, eax, LVIS_SELECTED or LVIS_FOCUSED
			inc ebx
		.endw

		invoke InvalidateRect, m_hWndLV, 0, 1
if ?HDRBMPS
		invoke ResetHeaderBitmap, m_hWndLV
		invoke SetHeaderBitmap, m_hWndLV, m_iSortCol, m_iSortDir
endif
		invoke SetCursor, hCsrOld
		ret
		align 4

SortListView endp


;*** WM_NOTIFY/LVN_DISPINFO processing for listview ***


GetDispInfo proc uses ebx esi pNMLV:ptr NMLVDISPINFO

		mov ebx,pNMLV
		assume ebx:ptr NMLVDISPINFO

		.if (m_iMode != MODE_OBJECT)	;not for object list, itemdata used
if ?STATEIMAGE
			.if ([ebx].item.imask & LVIF_STATE)
				invoke GetItemFlag@CDocument, m_pDoc, [ebx].item.iItem, FLAG_IMAGE
				.if (al)
					or [ebx].item.state,1000h
					or [ebx].item.stateMask, LVIS_STATEIMAGEMASK
				.endif
			.endif
else
			.if ([ebx].item.mask_ & LVIF_IMAGE)
				invoke GetItemFlag@CDocument, m_pDoc, [ebx].item.iItem, FLAG_IMAGE
				.if (al)
					mov [ebx].item.iImage,0
				.endif
			.endif
endif
		.endif
		.if ([ebx].item.mask_ & LVIF_STATE)
			invoke GetItemFlag@CDocument, m_pDoc, [ebx].item.iItem, LVIS_SELECTED or LVIS_FOCUSED
			mov [ebx].item.state,eax
			or [ebx].item.stateMask,LVIS_SELECTED or LVIS_FOCUSED
		.endif
		.if ([ebx].item.mask_ & LVIF_TEXT)
			invoke GetItemData@CDocument, m_pDoc, [ebx].item.iItem, [ebx].item.iSubItem
			.if (!eax)
				mov eax,offset g_szNull
			.endif
if 0
			invoke lstrcpyn, [ebx].item.pszText, eax, [ebx].item.cchTextMax
else
			mov [ebx].item.pszText, eax
endif
		.endif
		ret
		assume ebx:nothing
		align 4

GetDispInfo endp


SetStatusPane1 proc

local	szText[128]:byte

		invoke ListView_GetItemCount( m_hWndLV)
		invoke wsprintf, addr szText, CStr("%u items"), eax
		StatusBar_SetText g_hWndSB, 1, addr szText
		ret
		align 4
SetStatusPane1 endp

SetStatusPane2 proc

local	szText[128]:byte

		invoke ListView_GetSelectedCount( m_hWndLV)
		.if (eax > 1)
			invoke wsprintf, addr szText, CStr("%u items selected"), eax
		.else
			mov szText,0
		.endif
		StatusBar_SetText g_hWndSB, 2, addr szText
		ret
		align 4

SetStatusPane2 endp


;*** WM_NOTIFY processing for listview ***


OnNotifyLV proc pNMLV:ptr NMLISTVIEW

		mov edx,pNMLV
		assume edx:ptr NMLISTVIEW

		.if ([edx].hdr.code == NM_DBLCLK)

			.if (m_iMode == MODE_TYPELIB)
				mov ecx, IDM_TYPELIBDLG
			.elseif ((m_iMode == MODE_OBJECT) || (m_iMode == MODE_ROT))
				mov ecx, IDM_OBJECTDLG
			.else
				mov ecx, IDM_EDIT
			.endif
			invoke PostMessage, m_hWnd, WM_COMMAND, ecx, 0

		.elseif ([edx].hdr.code == NM_RCLICK)

			invoke ShowContextMenu, pNMLV, TRUE

		.elseif ([edx].hdr.code == LVN_GETDISPINFO)

			invoke GetDispInfo, edx

		.elseif ([edx].hdr.code == LVN_ODSTATECHANGED)

			assume edx:ptr NMLVODSTATECHANGE
;;			DebugOut "LVN_ODSTATECHANGED, %u-%u, Old=%X, New=%X", [edx].iFrom, [edx].iTo, [edx].uOldState, [edx].uNewState
			push esi
			mov esi, [edx].iFrom
			mov eax, [edx].uNewState
			mov ecx, [edx].uOldState
			and eax, LVIS_SELECTED or LVIS_FOCUSED
			and ecx, LVIS_SELECTED or LVIS_FOCUSED
			xor ecx, eax
			mov edx, [edx].iTo
			.while (esi <= edx)
				push eax
				push edx
				push ecx
				invoke SetItemFlag@CDocument, m_pDoc, esi,
					eax, ecx
				pop ecx
				pop edx
				pop eax
				inc esi
			.endw
			pop esi
			assume edx:ptr NMLISTVIEW

		.elseif ([edx].hdr.code == LVN_ITEMCHANGED)

			.if ([edx].uChanged & LVIF_STATE)
				push edx
;;				DebugOut "LVN_ITEMCHANGED, %u, Old=%X, New=%X", [edx].iItem, [edx].uOldState, [edx].uNewState
				mov eax, [edx].uNewState
				mov ecx, [edx].uOldState
				and eax, LVIS_SELECTED or LVIS_FOCUSED
				and ecx, LVIS_SELECTED or LVIS_FOCUSED
				xor ecx, eax
				invoke SetItemFlag@CDocument, m_pDoc, [edx].iItem,
					eax, ecx
				pop edx
			.endif

			.if (!g_dwTimer)
				invoke SetTimer, m_hWnd, 1, 50, NULL
				mov g_dwTimer, eax
			.endif

		.elseif ([edx].hdr.code == LVN_KEYDOWN)

			assume edx:ptr NMLVKEYDOWN
			.if ([edx].wVKey == VK_APPS)
				invoke ShowContextMenu, pNMLV, FALSE
			.endif
			assume edx:ptr NMLISTVIEW

		.elseif ([edx].hdr.code == LVN_COLUMNCLICK)

			mov eax,[edx].iSubItem
			.if (eax == m_iSortCol)
				xor m_iSortDir,1
			.else
				mov m_iSortCol,eax
				mov m_iSortDir,0
			.endif
			invoke SortListView
;			ListView_SetSelectedColumn m_hWndLV, m_iSortCol

		.elseif ([edx].hdr.code == LVN_ODFINDITEM)

			assume edx:ptr NMLVFINDITEM
			mov ecx, m_iSortCol
			.if (ecx == -1)
				inc ecx
			.endif
			xor eax, eax
			.if ([edx].lvfi.flags & LVFI_WRAP)
				mov eax, SEARCH_WRAP
			.endif
			invoke SearchString, [edx].lvfi.psz, eax, [edx].iStart, ecx
			invoke SetWindowLong, m_hWnd, DWL_MSGRESULT, eax
			mov eax, 1
			assume edx:ptr NMLISTVIEW
if ?DROPSOURCE
		.elseif (([edx].hdr.code == LVN_BEGINDRAG) || ([edx].hdr.code == LVN_BEGINRDRAG))

			.if ((m_iMode == MODE_CLSID) || (m_iMode == MODE_TYPELIB))
				invoke OnBeginDrag
			.endif
endif
if 0
		.elseif ([edx].hdr.code == NM_CUSTOMDRAW)

	.data
g_clrTextBk COLORREF 808080h
	.code

			assume edx:ptr NMLVCUSTOMDRAW
			.if ([edx].nmcd.dwDrawStage == CDDS_PREPAINT)
				.if (m_iSortCol != -1)
					invoke SetWindowLong, m_hWnd, DWL_MSGRESULT, CDRF_NOTIFYITEMDRAW
					mov eax, 1
				.endif
			.elseif ([edx].nmcd.dwDrawStage == CDDS_ITEMPREPAINT)
				mov eax, [edx].clrTextBk
				mov g_clrTextBk, eax 
				invoke SetWindowLong, m_hWnd, DWL_MSGRESULT, CDRF_NOTIFYSUBITEMDRAW
				mov eax, 1
			.elseif ([edx].nmcd.dwDrawStage == (CDDS_ITEMPREPAINT or CDDS_SUBITEM))
				mov ecx, m_iSortCol
				.if (ecx == [edx].iSubItem)
					mov eax, 0FFFFFFh
				.else
					mov eax, g_clrTextBk
				.endif
				mov [edx].clrTextBk, eax
			.endif
endif
		.endif
		ret
		align 4
		assume edx:nothing

OnNotifyLV endp

;*** WM_NOTIFY/NM_RCLICK for header control ***

OnHeaderRClick proc uses ebx pNMHDR:ptr NMHDR

local	pt:POINT
local	mii:MENUITEMINFO
local	szText[64]:byte

		.if ((m_iMode == MODE_CLSID) || (m_iMode == MODE_HKCR) || (m_iMode == MODE_INTERFACE))
			invoke CreatePopupMenu
			mov ebx, eax
			invoke MakeUDColumnList, ebx, m_iMode, 1
			invoke GetCursorPos, addr pt
			mov g_pszMenuHelp, \
				CStr("Select one of these items as userdefined column. Select checked item to remove userdefined column.")
			invoke TrackPopupMenu, ebx, TPM_LEFTALIGN or TPM_LEFTBUTTON or TPM_RETURNCMD,
					pt.x, pt.y, 0, m_hWnd, NULL
			mov g_pszMenuHelp, NULL
			.if (eax)
				mov ecx, eax
				mov mii.cbSize, sizeof MENUITEMINFO
				mov mii.fMask, MIIM_TYPE
				lea eax, szText
				mov mii.dwTypeData, eax
				mov mii.cch, sizeof szText
				invoke GetMenuItemInfo, ebx, ecx, FALSE, addr mii
				.if (m_iMode == MODE_CLSID)
					mov edx, offset g_szUserColCLSID
				.elseif (m_iMode == MODE_INTERFACE)
					mov edx, offset g_szUserColInterface
				.else
					mov edx, offset g_szUserColHKCR
				.endif
				push edx
				invoke lstrcmp, edx, addr szText
				pop edx
				.if (eax)
					invoke lstrcpy, edx, addr szText
				.else
					mov byte ptr [edx], 0
				.endif
				mov g_bColumnsChanged, TRUE
				invoke RefreshView@CMainDlg, g_pMainDlg, m_iMode
			.endif
			invoke DestroyMenu, ebx
		.endif
		ret
		align 4

OnHeaderRClick endp


;*** WM_NOTIFY processing ***


OnNotify proc uses ebx pNMHDR:ptr NMHDR

local	tci:TC_ITEM

		xor eax, eax
		mov ebx,pNMHDR
		.if ([ebx].NMHDR.idFrom == IDC_LIST1)

			invoke OnNotifyLV, ebx

		.elseif ([ebx].NMHDR.idFrom == IDC_TAB1)

			mov eax,[ebx].NMHDR.code
			.if (eax == TCN_SELCHANGE)
				invoke TabCtrl_GetCurSel( m_hWndTab)
				mov ecx,eax
				mov tci.mask_,TCIF_PARAM
				invoke TabCtrl_GetItem( m_hWndTab, ecx, addr tci)
				invoke PostMessage, m_hWnd, WM_COMMAND, tci.lParam, 0
;;			.elseif (eax == TCN_SELCHANGING)
			.endif

		.else

			mov ecx, [ebx].NMHDR.hwndFrom 
			.if ((ecx == m_hWndHdr) && ([ebx].NMHDR.code == NM_RCLICK))
				invoke OnHeaderRClick, ebx
			.elseif ((ecx == m_hWndHdr) && ([ebx].NMHDR.code == HDN_ITEMCHANGED))
				mov ecx, [ebx].NMHEADER.pitem
				.if ([ecx].HDITEM.mask_ & HDI_WIDTH)
					invoke SaveLVColumns
				.endif
			.endif

		.endif
		ret
		align 4

OnNotify endp

if ?HTMLHELP

ShowHtmlHelp proc public uses ebx pszHelpfile:LPSTR, uCommand:DWORD, dwData:DWORD

local hKey:HANDLE
local dwSize:DWORD
local pszError:LPSTR
local szLib[MAX_PATH]:byte
local szPath[MAX_PATH+32]:byte
local szKey[64]:byte

if 0
		includelib "HtmlHelp.Lib"

HtmlHelpA proto :HWND, :LPSTR, :DWORD, :DWORD
HtmlHelp equ <HtmlHelpA>

		invoke HtmlHelp, m_hWnd, addr szPath, uCommand, dwData
else
		.data

g_hLibHHCtrl	HANDLE NULL
g_pfnHtmlHelp	LPHTMLHELP NULL
g_dwHelpCookie	DWORD 0

		.code

		.if (!g_hLibHHCtrl)
;----------------------------------------- load HHCtrl.OCX
			mov pszError,NULL
			invoke RegOpenKey, HKEY_CLASSES_ROOT, CStr("CLSID\{ADB880A6-D8FF-11CF-9377-00AA003B7A11}\InProcServer32"), addr hKey
			.if (eax == ERROR_SUCCESS)
				mov dwSize, sizeof szLib
				invoke RegQueryValueEx,hKey,addr g_szNull,NULL,NULL,addr szLib,addr dwSize
				.if (eax == ERROR_SUCCESS)
					invoke LoadLibrary, addr szLib
					.if (eax > 32)
						mov g_hLibHHCtrl, eax
						invoke GetProcAddress, eax, 14
						mov g_pfnHtmlHelp, eax
						.if (eax)
							invoke g_pfnHtmlHelp, NULL, NULL, HH_INITIALIZE, addr g_dwHelpCookie
						.endif
					.else
						invoke wsprintf, addr szPath, CStr("Library",10,"%s",10,"not found"), addr szLib
						lea eax, szPath
						mov pszError, eax
					.endif
				.else
					mov pszError, CStr("Invalid InProcServer32 Entry for HTML Help Control")
				.endif
				invoke RegCloseKey, hKey
			.else
				mov pszError, CStr("HTML Help Control not installed")
			.endif
			.if (pszError)
				invoke MessageBox, m_hWnd, pszError, 0, MB_OK
				jmp done
			.endif
		.endif
		.if (g_pfnHtmlHelp)
			invoke GetDesktopWindow
			invoke g_pfnHtmlHelp, eax, pszHelpfile, uCommand, dwData
		.endif
endif
done:
		ret
		align 4

ShowHtmlHelp endp

DoHtmlHelp proc public uCommand:DWORD, dwData:DWORD

local szPath[MAX_PATH]:byte

		invoke GetModuleFileName, NULL, addr szPath,  MAX_PATH
		lea edx, szPath
		mov dword ptr [edx+eax-3],"mhc"
		invoke ShowHtmlHelp, edx, uCommand, dwData
		.if (!eax)
			sub esp, MAX_PATH+32
			mov edx, esp
			invoke wsprintf, edx, CStr("HtmlHelp('%s') failed"), addr szPath
			mov edx, esp
			invoke MessageBox, m_hWnd, edx, 0, MB_OK
			add esp, MAX_PATH+32
		.endif
		ret
		align 4

DoHtmlHelp endp

endif

;*** WM_COMMAND processing ***

		.const

COMMANDPROC typedef proto
LPCOMMANDPROC typedef ptr COMMANDPROC

CommandItem struct
dwID	dd ?
pProc	LPCOMMANDPROC ?
CommandItem ends

CommandItemTab	label CommandItem
	CommandItem <IDM_LOADTYPELIB,	OnLoadTypeLib>
	CommandItem <IDM_REGISTER,		OnRegister>
	CommandItem <IDM_LOADFILE,		OnLoadFile>
	CommandItem <IDM_PASTE,			OnPaste>
	CommandItem <IDM_CREATELINK,	OnCreateLink>
	CommandItem <IDM_SHDESKTOPFOLDER, OnSHDesktopFolder>
	CommandItem <IDM_OPENSTORAGE,	OnOpenStorage>
	CommandItem <IDM_VIEWSTORAGE,	OnViewStorage>
	CommandItem <IDM_OPENSTREAM,	OnOpenStream>
	CommandItem <IDM_VIEWSTREAM,	OnViewStream>
	CommandItem <IDM_SAVEAS,		OnSaveAs>
	CommandItem <IDM_OLEREG,		OnOleReg>

	CommandItem <IDM_EDIT,			OnEdit>
	CommandItem <IDM_COPY,			OnCopy>
	CommandItem <IDM_COPYGUID,		OnCopyGUID>
if ?REMOVEITEM
	CommandItem <IDM_REMOVEITEM,	OnRemoveItem>
endif
	CommandItem <IDM_INVERT,		OnInvertSelection>

	CommandItem <IDM_CREATEINSTANCE,OnCreateInstance>
	CommandItem <IDM_CREATEINSTON,	OnCreateInstanceOn>
	CommandItem <IDM_TYPELIBDLG,	OnTypeLibDlg>
	CommandItem <IDM_OBJECTDLG,		OnObjectDlg>
	CommandItem <IDM_VIEWMONIKER,	OnViewMoniker>
	CommandItem <IDM_GETCLASSFACT,	OnGetClassFactory>
	CommandItem <IDM_UNLOCK,		OnUnlock>
	CommandItem <IDM_FIND,			OnFind>
	CommandItem <IDM_FINDALL,		OnFindAll>
	CommandItem <IDM_UNREGISTER,	OnUnregister>
	CommandItem <IDM_EXPLORE,		OnExplore>
	CommandItem <IDM_PROPERTIES,	OnProperties>

	CommandItem <IDM_CHECKFILE,		OnCheckFile>
	CommandItem <IDM_CHECKTYPELIB,	OnCheckTypelib>
	CommandItem <IDM_CHECKCLSID,	OnCheckCLSID>
	CommandItem <IDM_CHECKPROGID,	OnCheckProgID>
	CommandItem <IDM_CHECKAPPID,	OnCheckAppID>
	CommandItem <IDM_AMBIENTPROP,	OnAmbientProperties>
	CommandItem <IDM_OPTIONS,		OnOption>
if ?UPDATECHK
	CommandItem <IDM_CHECKUPD,		OnCheckUpdate>
endif
	CommandItem <IDM_SELECTALL,		OnSelectAll>
	CommandItem <IDM_FINDNEXTCROSS,	OnFindNextCross>
	CommandItem <IDM_VIEWMALLOC,	OnViewMalloc>
COMMANDTABSIZE equ ($ - CommandItemTab) / sizeof CommandItem

		.code

OnCommand proc wParam:WPARAM, lParam:LPARAM

		movzx eax,word ptr wParam		;use only LOWORD(wParam)

;-------------------------------- scan table of commands without parameters

		mov	ecx, COMMANDTABSIZE
		mov edx, offset CommandItemTab
		.while (ecx)
			.if (eax == [edx])
				invoke [edx].CommandItem.pProc
				ret
			.endif
			add edx,sizeof CommandItem
			dec ecx
		.endw

;-------------------------------- scan table of mode commands

		mov edx,offset ModeDesc
		mov ecx,NUMMODES
		.while (ecx)
			.if (eax == [edx].CMode.iCmdID)
				invoke RefreshView, [edx].CMode.iMode
				ret
			.endif
			add edx,sizeof CMode
			dec ecx
		.endw

;-------------------------------- scan rest of commands

		.if (eax == IDM_EXIT)

			invoke PostMessage, m_hWnd, WM_CLOSE, 0, 0

		.elseif (eax == IDM_REFRESH)

			invoke RefreshView, -1

		.elseif (eax == IDM_REFRESHVIEW)

			invoke RefreshView@CMainDlg, g_pMainDlg, m_iMode

		.elseif (eax == IDM_REFRESHLINE)

			invoke RefreshLine, lParam

		.elseif (eax == IDM_LOGWINDOW)

			invoke Create@CLogWindow

		.elseif (eax == IDM_ABOUT)

			invoke DialogBoxParam, g_hInstance, IDD_ABOUTDLG, m_hWnd, aboutdialogproc, 0

		.elseif (eax == IDOK)

			invoke ListView_GetSelectedCount( m_hWndLV)
			.if (eax)
				invoke GetDefaultCommand
				invoke PostMessage, m_hWnd, WM_COMMAND, eax, 0
			.endif
			xor eax,eax

		.elseif (eax == IDCANCEL)

			xor eax,eax

		.elseif (eax == IDM_HELP)

			invoke DoHtmlHelp, HH_DISPLAY_TOC, NULL

		.else

			invoke MessageBox, m_hWnd, CStr("unknown command"), 0, MB_OK
			xor eax,eax

		.endif
		ret
		align 4

OnCommand endp


;*** WM_INITDIALOG processing ***


OnInitDialog proc uses ebx 

local	tci:TC_ITEM

		DebugOut "CMainWnd::OnInitDialog"

		invoke GetMenu, m_hWnd
		mov m_hMenu, eax
		invoke GetDlgItem, m_hWnd, IDC_TAB1
		mov m_hWndTab, eax
		invoke GetDlgItem, m_hWnd, IDC_LIST1
		mov m_hWndLV, eax
		invoke ListView_GetHeader( eax)
		mov m_hWndHdr, eax

		invoke CreateWindowEx, 0, STATUSCLASSNAME, NULL,
				WS_CHILD or WS_VISIBLE or SBARS_SIZEGRIP or CCS_BOTTOM,
				0,0,0,0, m_hWnd, IDC_STATUSBAR, g_hInstance, NULL
		.if (eax)
			mov g_hWndSB, eax
			StatusBar_SetText g_hWndSB, 255 or SBT_NOBORDERS, addr g_szNull
;;			invoke SetSBParts, g_hWndSB, offset g_dwSBParts, LENGTHOF g_dwSBParts + 1
		.endif

		mov tci.mask_,TCIF_TEXT or TCIF_PARAM
		mov ebx,0
		.while (ebx < NUMMODES)
			mov eax,sizeof CMode
			mul ebx
			lea ecx,[eax + offset ModeDesc]
			mov eax,[ecx].CMode.pszTabText
			mov tci.pszText,eax
			mov eax,[ecx].CMode.iCmdID
			mov tci.lParam,eax
			invoke TabCtrl_InsertItem( m_hWndTab, ebx, addr tci)
			inc ebx
		.endw

		invoke ListView_SetExtendedListViewStyle( m_hWndLV, ?LVSTYLE)

		invoke Create@CDropTarget, m_hWnd
		mov m_pDropTarget, eax
		invoke RegisterDragDrop, m_hWndLV, eax
if ?UPDATECHK eq 0
		invoke DeleteMenu, g_hMenu, IDM_CHECKUPD, MF_BYCOMMAND 
endif
		ret
		align 4

OnInitDialog endp


;--- WM_HELP


OnHelp proc pHelpInfo:ptr HELPINFO

local mii:MENUITEMINFO
local szText[32]:BYTE
local szText2[128]:BYTE

		mov ecx, pHelpInfo
		.if ([ecx].HELPINFO.iContextType == HELPINFO_MENUITEM)
			mov szText, 0
			mov mii.cbSize, sizeof MENUITEMINFO
			mov mii.fMask, MIIM_TYPE 
			lea eax, szText
			mov mii.dwTypeData, eax
			mov mii.cch, sizeof szText
			invoke GetMenuItemInfo, m_hMenu, [ecx].HELPINFO.hItemHandle, FALSE, addr mii
;			DebugOut "WM_HELP, %s", addr szText
			.if (eax)
				lea eax, szText
				.if (byte ptr [eax] == '&')
					inc eax
				.endif
				invoke wsprintf, addr szText2, CStr("mainwindowmenu.htm#%sSubMenu"), eax
				invoke DoHtmlHelp, HH_DISPLAY_TOPIC, addr szText2
			.endif
		.else
			invoke DoHtmlHelp, HH_DISPLAY_TOPIC, CStr("mainwindowviews.htm")
		.endif
		ret
		align 4

OnHelp endp

if ?DDESUPPORT

OnDDEExecute proc uses esi pszCommand:LPSTR

local szCmd[32]:BYTE
local szFile[MAX_PATH]:BYTE
local dwRC:DWORD

	DebugOut "OnDDEExecute %s", pszCommand
	mov dwRC, FALSE
	mov esi, pszCommand
	.if (byte ptr [esi] == '[')
		inc esi
	.endif
	push edi
	lea edi, szCmd
	mov ecx, sizeof szCmd - 1
	.while (ecx)
		lodsb
		.break .if ((al == '(') || (al == 0))
		stosb
		dec ecx
	.endw
	mov byte ptr [edi], 0
	.if (al == '(')
		lea edi, szFile
		mov ecx, MAX_PATH-1
		mov ah, 00
		.while (ecx)
			lodsb
			.break .if (al == 0)
			.if (al == '"')
				xor ah,1
			.elseif ((al == ')') && (ah == 0))
				.break
			.else
				stosb
				dec ecx
			.endif
		.endw
	.endif
	mov al, 00
	stosb
	pop edi
	invoke lstrcmpi, addr szCmd, CStr("OPEN")
	.if (!eax)
		mov dwRC, TRUE
		.if (g_bBindIsDefault)
			invoke SmartLoad@CMainDlg, __this, addr szFile
		.else
			invoke Create2@CTypeLibDlg, addr szFile, NULL, FALSE
			invoke Show@CTypeLibDlg, eax, m_hWnd, FALSE
		.endif
	.endif
exit:
	return dwRC
	align 4

OnDDEExecute endp

	.data
g_wDDEStatus dw 0
g_hwndDDEServer HWND NULL
g_hDDEMem HANDLE NULL
	.code

endif

;*** main dialog proc message processing ***


MainDialog proc uses __this thisarg, message:dword, wParam:WPARAM, lParam:LPARAM

if ?DDESUPPORT
local lo:DWORD
local hi:DWORD
endif

if 0; def _DEBUG
		.if ((message != WM_ENTERIDLE))
			DebugOut "MainDialog %X, %X, %X", message, wParam, lParam
		.endif
endif
		mov __this,this@

		mov eax,message
		.if (eax == WM_INITDIALOG)

if ?DDESUPPORT
			.if (g_bOneInstance)
				mov g_wDDEStatus, 0
				DebugOut "Sending WM_DDE_INITIATE"
				invoke SendMessage, HWND_BROADCAST, WM_DDE_INITIATE, m_hWnd, g_aApplication
				.if (g_wDDEStatus)
					.if (g_pszFilename)
						sub esp, MAX_PATH+32
						mov edx, esp
						invoke wsprintf, edx, CStr("[open(",22h,"%s",22h,")]"), g_pszFilename
						mov edx, esp
						add eax, 4
						and al, 0FCh
						push eax
						shl eax, 1
						invoke GlobalAlloc, GMEM_MOVEABLE, eax
						mov g_hDDEMem, eax
						invoke GlobalLock, eax
						DebugOut "DDE global memory block=%X", g_hDDEMem
						pop ecx
						mov edx, esp
if 0
						invoke MultiByteToWideChar,CP_ACP,MB_PRECOMPOSED,
							edx, -1, eax, ecx
else
						invoke lstrcpy, eax, edx
endif
						invoke GlobalUnlock, g_hDDEMem
						add esp, MAX_PATH+32
						DebugOut "Posting WM_DDE_Execute"
						invoke PostMessage, g_hwndDDEServer, WM_DDE_EXECUTE, m_hWnd, g_hDDEMem
						mov g_pszFilename, NULL
					.else
						DebugOut "Posting WM_DDE_TERMINATE"
						invoke PostMessage, g_hwndDDEServer, WM_DDE_TERMINATE, m_hWnd, NULL
						invoke RestoreAndActivateWindow, g_hwndDDEServer
					.endif
					mov eax, 1
					jmp done
				.endif
			.endif
endif
			invoke OnInitDialog

			push esi
			invoke GetSystemMetrics, SM_CXFULLSCREEN
			push eax
			invoke GetSystemMetrics, SM_CYFULLSCREEN
			pop ecx
			mov edx, g_rectMain.right
			mov esi, g_rectMain.bottom
			.if (edx && (ecx > edx) && (eax > esi))
				invoke SetWindowPos, m_hWnd, NULL, 0, 0,
					g_rectMain.right, g_rectMain.bottom, SWP_NOMOVE or SWP_NOZORDER or SWP_NOACTIVATE
			.endif
			pop esi

			invoke ResizeClients, 0

			invoke CenterWindow, m_hWnd

;			invoke SendMessage, m_hWnd, WM_SETICON, ICON_SMALL, g_hIconApp
;			invoke SendMessage, m_hWnd, WM_SETICON, ICON_BIG, g_hIconApp
			invoke RefreshView, m_iMode
;;			invoke SendMessage, m_hWnd, WM_COMMAND, IDM_CLSID, 0

			invoke ShowWindow, m_hWnd,SW_NORMAL
			mov eax,1

		.elseif (eax == WM_CLOSE)

			.if (m_pDropTarget)
				invoke RevokeDragDrop, m_hWndLV
				invoke vf(m_pDropTarget, IUnknown, Release)
				mov m_pDropTarget, NULL
			.endif

			sub esp, sizeof WINDOWPLACEMENT
			mov [esp].WINDOWPLACEMENT.length_, sizeof WINDOWPLACEMENT
			invoke GetWindowPlacement, m_hWnd, esp
			invoke CopyRect, addr g_rectMain, addr [esp].WINDOWPLACEMENT.rcNormalPosition
			add esp, sizeof WINDOWPLACEMENT
			mov eax, g_rectMain.right
			sub eax, g_rectMain.left
			mov g_rectMain.right, eax
			mov g_rectMain.left, 0
			mov eax, g_rectMain.bottom
			sub eax, g_rectMain.top
			mov g_rectMain.bottom, eax
			mov g_rectMain.top, 0
;------------------------- since there may exist other top level windows
;------------------------- do this broadcast before object termination
			invoke BroadCastMessage, WM_CLOSE, m_hWnd, 0

			invoke PostQuitMessage,0
			mov eax,0

		.elseif (eax == WM_SIZE)
			.if ((wParam == SIZE_RESTORED) || (wParam == SIZE_MAXIMIZED))
				invoke ResizeClients, lParam
			.endif
			mov eax,1

		.elseif (eax == WM_COMMAND)

			invoke OnCommand, wParam, lParam

		.elseif (eax == WM_NOTIFY)

			invoke OnNotify, lParam

		.elseif (eax == WM_ERASEBKGND)

			mov eax, 1

		.elseif (eax == WM_ENTERMENULOOP)

			StatusBar_SetSimpleMode g_hWndSB, TRUE
			invoke OnEnterMenuLoop, wParam

		.elseif (eax == WM_EXITMENULOOP)

			StatusBar_SetSimpleMode g_hWndSB, FALSE
			invoke OnExitMenuLoop, wParam

		.elseif (eax == WM_MENUSELECT)

;;			DebugOut "WM_MENUSELECT, wParam=%X", wParam
			movzx ecx, word ptr wParam+0
			invoke DisplayStatusBarString, g_hWndSB, ecx

		.elseif (eax == WM_ACTIVATE)

			movzx eax,word ptr wParam
			.if (eax == WA_INACTIVE)
				mov g_hWndDlg,NULL
				mov eax,g_hAccel2
				mov g_hCurAccel,eax
			.else
				mov eax, m_hWnd
				mov g_hWndDlg,eax
				mov eax,g_hAccel
				mov g_hCurAccel,eax
			.endif

		.elseif (eax == g_uMsgFind)

			invoke OnFindMsg, lParam

		.elseif (eax == WM_HELP)

if ?HTMLHELP eq 0
			.if (g_hLibRE == 0)
				invoke LoadLibrary, CStr("RICHED32.DLL")
				mov g_hLibRE,eax
			.endif
			invoke DialogBoxParam, g_hInstance, IDD_HELPDLG, m_hWnd, helpdialogproc, 0
else
			invoke OnHelp, lParam
endif
if ?DDESUPPORT
		.elseif ((eax == WM_DDE_INITIATE) && g_aApplication)

			DebugOut "WM_DDE_INITIATE, wParam=%X, lParam=%X", wParam, lParam
;------------------------------- respond to windows from this thread only
			movzx ecx, word ptr lParam+0
			movzx edx, word ptr lParam+2
			mov eax, wParam
			.if ((ecx == g_aApplication) && (eax != m_hWnd))
				DebugOut "WM_DDE_INITIATE accepted, my hWnd is %X", m_hWnd
				mov eax, g_aSystem
				mov word ptr lParam+2, ax
				invoke SendMessage, wParam, WM_DDE_ACK, m_hWnd, lParam
			.endif
			xor eax, eax

		.elseif (eax == WM_DDE_EXECUTE)

			DebugOut "WM_DDE_EXECUTE, wParam=%X, lParam=%X", wParam, lParam
			invoke UnpackDDElParam, WM_DDE_EXECUTE, lParam, addr lo, addr hi
			invoke GlobalLock, hi
;------------------------------- now a fantastic amount of work is to be done
			invoke OnDDEExecute, eax
;------------------------------- low word is a DDEACK
			.if (eax)
				mov lo, 8000h	;flag: command accepted
			.else
				mov lo, 0
			.endif

			invoke GlobalUnlock, hi
			DebugOut "Posting WM_DDE_ACK for WM_DDE_EXECUTE"
			invoke ReuseDDElParam, lParam, WM_DDE_EXECUTE, WM_DDE_ACK, lo, hi
			invoke PostMessage, wParam, WM_DDE_ACK, m_hWnd, eax

		.elseif (eax == WM_DDE_TERMINATE)

			DebugOut "WM_DDE_TERMINATE, wParam=%X", wParam
			.if (g_bOneInstance && (g_wDDEStatus == 1))
				invoke PostMessage, m_hWnd, WM_CLOSE, 0, 0
				invoke PostQuitMessage,0
			.else
				invoke PostMessage, wParam, WM_DDE_TERMINATE, m_hWnd, NULL
			.endif

		.elseif (eax == WM_DDE_ACK)

			DebugOut "WM_DDE_ACK, wParam=%X, lParam=%X", wParam, lParam
			.if (g_wDDEStatus == 0)
				mov eax, wParam
				mov g_hwndDDEServer, eax
				inc g_wDDEStatus
			.elseif (g_wDDEStatus == 1)
;---------------------------------- WM_DDE_EXECUTE
				invoke UnpackDDElParam, WM_DDE_ACK, lParam, addr lo, addr hi
				DebugOut "WM_DDE_ACK, free memory object=%X", hi
				invoke GlobalFree, hi
				invoke PostMessage, wParam, WM_DDE_TERMINATE, m_hWnd, NULL
			.endif
endif
		.elseif (eax == WM_TIMER)

			invoke KillTimer, m_hWnd, 1
			mov g_dwTimer, NULL
			invoke SetStatusPane2
			invoke UpdateMenu
if ?DRAWITEMSB
		.elseif (eax == WM_DRAWITEM)
			.if (wParam == IDC_STATUSBAR)
				push esi
				mov esi, lParam
				mov eax, [esi].DRAWITEMSTRUCT.itemData
				.if (dword ptr [eax] == "ysub")
					mov eax, 000000C0h
				.else
					mov eax, 00008000h
				.endif
				invoke SetTextColor, [esi].DRAWITEMSTRUCT.hDC, eax
				invoke SetBkMode, [esi].DRAWITEMSTRUCT.hDC, TRANSPARENT
				add [esi].DRAWITEMSTRUCT.rcItem.left, 4
				invoke DrawTextEx, [esi].DRAWITEMSTRUCT.hDC,\
					[esi].DRAWITEMSTRUCT.itemData, -1, addr [esi].DRAWITEMSTRUCT.rcItem,\
					DT_LEFT or DT_SINGLELINE or DT_VCENTER, NULL
				pop esi
			.endif
endif
		.else

			xor eax,eax ;indicates "no processing"

		.endif
done:
        ret
		align 4

MainDialog endp

;--- tries to load a file
;--- 1. with CreateFileMoniker and BindToFile -> view object dialog
;--- 2. with OpenStorage -> view storage dialog
;--- 3. create a typelib dialog (eax = 1 if successful)

SmartLoad@CMainDlg proc public uses __this this_:ptr CMainDlg, pszFilename:LPSTR

local	wszFile[MAX_PATH]:word

		mov __this, this_
		invoke LoadFile@CMainDlg, __this, pszFilename, FALSE
		.if (eax != S_OK)
			invoke MultiByteToWideChar,CP_ACP,MB_PRECOMPOSED,
					pszFilename, -1, addr wszFile, MAX_PATH
			invoke StgIsStorageFile, addr wszFile
			.if (eax == S_OK)
				invoke OpenStorage@CMainDlg, __this, pszFilename, NULL
			.else
				invoke Create2@CTypeLibDlg, pszFilename, NULL, TRUE
				.if (eax)
					invoke Show@CTypeLibDlg, eax, m_hWnd, FALSE
					mov eax, 1
				.else
					invoke OpenStream@CMainDlg, __this, pszFilename
					mov eax, 2
				.endif
			.endif
		.endif
		ret
		align 4
SmartLoad@CMainDlg endp

ToUpper proc uses eax
		.while (1)
			mov cl, [eax]
			.break .if ((cl == 0) || (cl == ':'))
			.if ((cl >= 'a') && (cl <= 'z'))
				sub cl, 'a'
				add cl, 'A'
				mov [eax], cl
			.endif
			inc eax
		.endw
		ret
		align 4
ToUpper endp

printf proc c public pszFormat:ptr BYTE, parms:VARARG

local	szOut[1024]:byte

		invoke wvsprintfA, addr szOut, pszFormat, addr parms
		push eax
		.if (g_hConOut != -1)
			.if (!g_bFirstOut)
				mov g_bFirstOut, TRUE
				push 0
				mov ecx, esp
				invoke WriteConsole, g_hConOut, CStr(13,10), 2, ecx, 0
				pop eax
				mov eax,[esp]
			.endif
			mov ecx, esp
			invoke WriteConsole, g_hConOut, addr szOut, eax, ecx, 0
		.else
			lea ecx, szOut
			.if (word ptr [ecx+eax-2] == 0A0Dh)
				mov byte ptr [ecx+eax-2],0
			.endif
			invoke MessageBox, m_hWnd, addr szOut, CStr("COMView"), MB_OK
		.endif
		pop eax
		ret
		align 4
printf endp

prAttachConsole typedef proto :dword

;*** CMainDlg constructor ***

Create@CMainDlg proc public uses esi __this


local	pszFileOut:LPSTR
local	bSmart:BOOLEAN
local	bBatchMode:BOOLEAN

		mov bSmart, FALSE
		invoke malloc, sizeof CMainDlg
		.if (eax == NULL)
			ret
		.endif
		mov __this,eax
		mov g_pMainDlg,eax

		mov m_pDlgProc,MainDialog

		invoke LoadAccelerators, g_hInstance, IDR_ACCELERATOR1
		mov g_hAccel,eax
		invoke LoadAccelerators, g_hInstance, IDR_ACCELERATOR2
		mov g_hAccel2,eax

if ?DDESUPPORT
		invoke GlobalAddAtom, CStr("COMView")
		movzx eax, ax
		mov g_aApplication, eax
		invoke GlobalAddAtom, CStr("system")
		movzx eax, ax
		mov g_aSystem, eax
endif
		invoke Create@CList, NULL
		mov g_pObjects, eax

		mov m_pszRoot,NULL
		mov m_iSortCol,-1

		lea esi, m_savedView
		mov ecx, LENGTHOF CMainDlg.savedView
		.while (ecx)
			mov [esi].CMainView.iSortCol, -1
			add esi, sizeof CMainView
			dec ecx
		.endw

		mov bBatchMode, FALSE
		mov pszFileOut, NULL

		mov ecx, g_argc
		mov esi, g_argv
		lodsd				;skip comview filename
		dec ecx
		.while (ecx)
			DebugOut "Parameter %s", dword ptr [esi+4]
			lodsd
			push ecx
			mov cl, [eax]
			.if ((cl == '/') || (cl == '-'))
				inc eax
				invoke ToUpper
				mov ecx, [eax]
				.if (cx == "B")
					mov bBatchMode, TRUE
				.elseif (cx == "S")
					mov bSmart, TRUE
				.elseif (ecx == ":TUO")
					lea ecx, [eax+4]
					.if ((byte ptr [ecx] == 0) && (dword ptr [esp]))
						lodsd
						mov ecx, eax
						dec dword ptr [esp]
					.endif
					mov pszFileOut, ecx
				.else
					invoke MessageBox, m_hWnd, CStr("Unknown option"), 0, MB_OK
				.endif
			.elseif (!g_pszFilename)
				mov g_pszFilename, eax
			.else
				invoke MessageBox, m_hWnd, CStr("Too many parameters"), 0, MB_OK
			.endif
			pop ecx
			dec ecx
		.endw

		.if (bBatchMode)
			invoke GetModuleHandle, CStr("KERNEL32")
			.if (eax)
				invoke GetProcAddress, eax, CStr("AttachConsole")
				.if (eax)
					assume eax:ptr prAttachConsole
					invoke eax, -1
					assume eax:nothing
					.if (eax)
						invoke GetStdHandle, STD_OUTPUT_HANDLE
						mov g_hConOut, eax
					.endif
				.endif
			.endif
			.if (!g_pszFilename)
				invoke printf, CStr("Filename parameter missing",13,10)
			.else
				invoke Create2@CCreateInclude, g_pszFilename
				.if (eax)
					mov esi, eax
					.if (pszFileOut)
						invoke SetOutputFile@CCreateInclude, esi, pszFileOut
					.endif
					invoke Run@CCreateInclude, esi, 0, INCMODE_BASIC
					.if (eax)
						invoke Run@CCreateInclude, esi, 0, INCMODE_DISPHLP
					.endif
					invoke Destroy@CCreateInclude, esi
				.endif
			.endif
			invoke PostQuitMessage,0
			jmp exit
		.elseif (pszFileOut)
			invoke MessageBox, m_hWnd, CStr("/OUT option valid in batch mode only"), 0, MB_OK
		.endif
		.if (g_pszFilename)
			.if (bSmart || g_bBindIsDefault)
				mov m_iMode, MODE_OBJECT
			.else
				mov m_iMode, MODE_TYPELIB
			.endif
		.endif

		invoke CreateDialogParam, g_hInstance, IDD_MAINDLG, 0, classdialogproc, __this

		.if (g_pszFilename)
			.if (bSmart || g_bBindIsDefault)
				invoke SmartLoad@CMainDlg, __this, g_pszFilename
				.if (eax == 1)
					invoke RefreshView, MODE_TYPELIB
				.endif
			.else
				invoke Create2@CTypeLibDlg, g_pszFilename, NULL, FALSE
				invoke Show@CTypeLibDlg, eax, m_hWnd, FALSE
			.endif
		.endif
exit:
		return __this
		align 4

Create@CMainDlg endp


;*** CMainDlg destructor ***


Destroy@CMainDlg proc public uses __this esi thisarg

		mov __this,this@

if ?DDESUPPORT
		.if (g_aApplication)
			invoke GlobalDeleteAtom, g_aApplication
			invoke GlobalDeleteAtom, g_aSystem
		.endif
endif
;------------------------- destroy created objects
		.while (1)
			invoke GetItem@CList, g_pObjects, 0
			.break .if (!eax)
			push eax
			invoke vf(eax, IObjectItem, IsLocked)
			pop ecx
			.if (eax)
				invoke vf(ecx, IObjectItem, Unlock)
			.else
				invoke vf(ecx, IObjectItem, Release)
			.endif
		.endw
		invoke Destroy@CList, g_pObjects

		invoke Reset@CMainDlg, __this
if ?HTMLHELP
		.if (g_pfnHtmlHelp)
			invoke g_pfnHtmlHelp, NULL, NULL, HH_UNINITIALIZE, g_dwHelpCookie
		.endif
endif
		invoke free, __this
		ret
		align 4

Destroy@CMainDlg endp

RefreshView@CMainDlg proc public uses __this thisarg, dwMode:DWORD

		mov __this,this@
		mov eax, dwMode
if ?MULTIDOC
		.if (g_bMultiDoc)
			mov ecx, sizeof CMainView
			mul ecx
			lea eax, m_savedView[eax]
			xor edx, edx
			xchg edx, [eax].CMainView.pDoc
			.if (edx)
				invoke Destroy@CDocument, edx
			.endif
			mov eax, dwMode
			.if (eax == m_iMode)
				invoke RefreshView, eax
			.endif
		.else
endif
			.if (eax == m_iMode)
				invoke PostMessage, m_hWnd, WM_COMMAND, IDM_REFRESH, 0
			.endif
if ?MULTIDOC
		.endif
endif
		ret
		align 4
RefreshView@CMainDlg endp

Reset@CMainDlg proc public uses esi __this thisarg

		mov __this,this@
if ?HDRBMPS
		invoke ResetHeaderBitmap, m_hWndLV
endif
if ?MULTIDOC
		.if (g_bMultiDoc)
			xor esi, esi
			.while (esi <= MODE_MAX)
				push esi
				mov eax, esi
				mov ecx, sizeof CMainView
				mul ecx
				lea esi, m_savedView[eax]
				.if ([esi].CMainView.pDoc)
					invoke Destroy@CDocument, [esi].CMainView.pDoc
					mov [esi].CMainView.pDoc, NULL
					mov [esi].CMainView.iTopIndex, 0
					mov [esi].CMainView.iSortCol, -1
					mov [esi].CMainView.iSortDir, 0
if ?SAVECOLORDER
					invoke free, [esi].CMainView.pdwColOrder
					mov [esi].CMainView.pdwColOrder, NULL
endif
				.endif
				pop esi
				inc esi
			.endw
		.else
endif
			.if (m_pDoc)
				invoke Destroy@CDocument, m_pDoc
			.endif
if ?MULTIDOC
		.endif
endif
		mov m_pDoc, NULL
		ret
		align 4
Reset@CMainDlg endp

UserColChanged@CMainDlg proc public uses __this  thisarg, dwMode:DWORD
		mov g_bColumnsChanged, TRUE
		invoke RefreshView@CMainDlg, this@, dwMode
		ret
		align 4
UserColChanged@CMainDlg endp

if 0
GetListView@CMainDlg proc public thisarg
		mov ecx,this@
		mov eax, [ecx].CMainDlg.hWndLV
		ret
		align 4
GetListView@CMainDlg endp

GetMode@CMainDlg  proc public thisarg
		mov ecx,this@
		mov eax, [ecx].CMainDlg.iMode
		ret
		align 4
GetMode@CMainDlg  endp

GetMenu@CMainDlg  proc public thisarg
		mov ecx,this@
		invoke GetMenu, [ecx].CMainDlg.hWnd
		ret
		align 4
GetMenu@CMainDlg  endp
endif

;*** end of CMainDlg methods

	end
