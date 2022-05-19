
;*** definition of COptionDlg and COptPageDlg methods

	.386
	.model flat,stdcall
	option casemap :none   ; case sensitive
	option proc:private

	include COMView.inc
INSIDE_COPTIONDLG equ 1
	include classes.inc
	include rsrc.inc
	include debugout.inc

?MODELESS	equ 1

BEGIN_CLASS COptPageDlg, CDlg
hWndLV		HWND ?
hImageList	HANDLE ?
iDragItem	DWORD ?
iDropItem	DWORD ?
bDrag		BOOLEAN ?
bDragHidden	BOOLEAN ?
bChanged	BOOLEAN ?
bError		BOOLEAN ?
END_CLASS

BEGIN_CLASS COptionDlg, CDlg
hWndTab		HWND ?
tab			COptPageDlg <>
END_CLASS

	.const

UserColumnCLSID label LPSTR
	dd CStr("control")
	dd CStr("insertable")
	dd CStr("programmable")
	dd CStr("printable")
	dd CStr("MiscStatus")
	dd CStr("DefaultIcon")
	dd CStr("ToolboxBitmap32")
	dd CStr("Implemented Categories")
	dd CStr("Required Categories")
	dd CStr("Verb")
	dd CStr("VersionIndependentProgID")
	dd CStr("shellex")
	dd CStr("ShellFolder")
	dd CStr("[AppID]")
	dd CStr("OLE DB Provider")
	dd CStr("OLE DB Enumerator")
	dd CStr("Ole1Class")
	dd CStr("DocObject")
	dd CStr("TreatAs")
	dd CStr("InProcHandler32")
	dd CStr("PersistentAddinsRegistered")
	dd CStr("InProcServer32\[Assembly]")
	dd CStr("InProcServer32\[Class]")
	dd CStr("InProcServer32\[RuntimeVersion]")
	dd CStr("InProcServer32\[ThreadingModel]")
NUMUSERCOLUMNS_CLSID textequ %($ - UserColumnCLSID) / sizeof LPSTR

UserColumnHKCR label LPSTR
	dd CStr("shell\open\ddeexec")
	dd CStr("shell\print\command")
	dd CStr("shellex")
	dd CStr("shellex\DataHandler")
	dd CStr("shellex\DropHandler")
	dd CStr("shellex\ContextMenuHandlers")
	dd CStr("shellex\DragDropHandlers")
	dd CStr("shellex\ExtShellFolderViews")
	dd CStr("shellex\PropertySheetHandlers")
	dd CStr("DefaultIcon")
	dd CStr("Insertable")
	dd CStr("PersistentHandler")
	dd CStr("protocol\StdFileEditing\server")
	dd CStr("HTML Handler")
	dd CStr("[EditFlags]")
	dd CStr("[AlwaysShowExt]")
	dd CStr("[NeverShowExt]")
	dd CStr("[FriendlyTypeName]")
	dd CStr("[Content Type]")
NUMUSERCOLUMNS_HKCR textequ %($ - UserColumnHKCR) / sizeof LPSTR

UserColumnInterface label LPSTR
	dd CStr("NumMethods")
	dd CStr("ProxyStubClsid")
	dd CStr("TypeLib\[Version]")
	dd CStr("BaseInterface")
	dd CStr("Distributor")
	dd CStr("OLEViewerIViewerCLSID")
NUMUSERCOLUMNS_INTERFACE textequ %($ - UserColumnInterface) / sizeof LPSTR

public UserColumnTypeInfo
public UserFormatTypeInfo

UserColumnTypeInfo label LPSTR
	dd CStr("Mops")
	dd CStr("HelpContext")
	dd CStr("HelpFile")
	dd CStr("SizeInst")
	dd CStr("SizeVft")
	dd CStr("Alignment")
	dd CStr("LCID")
	dd CStr("Constructor")
	dd CStr("Destructor")
	dd CStr("IDLFlags")
NUMUSERCOLUMNS_TYPEINFO equ ($ - UserColumnTypeInfo) / sizeof LPSTR
UserFormatTypeInfo label DWORD
	dd LVCFMT_LEFT
	dd LVCFMT_RIGHT
	dd LVCFMT_LEFT
	dd LVCFMT_RIGHT
	dd LVCFMT_RIGHT
	dd LVCFMT_RIGHT
	dd LVCFMT_RIGHT
	dd LVCFMT_RIGHT
	dd LVCFMT_RIGHT
	dd LVCFMT_RIGHT

;*** tabbed dialog definitions (tab name, dialog resource id, dialog proc)

TabDlgPages label CTabDlgPage
	CTabDlgPage {CStr("General"),	IDD_OPTPAGE_GENERAL,	OptionGeneralDialog}
	CTabDlgPage {g_szCLSID,			IDD_OPTPAGE_CLSID,		OptionCLSIDDialog}
	CTabDlgPage {g_szTypeLib,		IDD_OPTPAGE_TYPELIB,	OptionTypelibDialog}
	CTabDlgPage {g_szInterface,		IDD_OPTPAGE_INTERFACE,	OptionInterfaceDialog}
	CTabDlgPage {CStr("HKCR"),		IDD_OPTPAGE_HKCR,		OptionHKCRDialog}
	CTabDlgPage {CStr("Object"),	IDD_OPTPAGE_OBJECT,		OptionObjectDialog}
	CTabDlgPage {CStr("View Control"),IDD_OPTPAGE_CONTROL,	OptionViewControlDialog}
	CTabDlgPage {CStr("Properties"),IDD_OPTPAGE_PROPERTIES,	OptionPropertiesDialog}
NUMOPTDLGS textequ %($ - TabDlgPages) / sizeof CTabDlgPage

if 0;?HTMLHELP
HtmlHelpTab label dword
	dd CStr("OptionsGeneral")
	dd CStr("OptionsCLSID")
	dd CStr("OptionsTypeLib")
	dd CStr("OptionsInterface")
	dd CStr("OptionsHKCR")
	dd CStr("OptionsObject")
	dd CStr("OptionsViewControl")
	dd CStr("OptionsProperties")
endif

	.data

g_dwOptPage		dd 0
g_hCsrYes		HCURSOR NULL
g_hCsrNo		HCURSOR NULL

;*** table of entries to search for "module path" for CLSID entries
;*** prefer 32bit paths. Must be in .DATA (modified by ArrageTypes)

	public NUMREGKEYSEARCHES

pRegKeys label LPSTR
	dd CStr("InProcServer32")
	dd CStr("LocalServer32")
	dd CStr("TreatAs")
	dd CStr("InProcServer")
	dd CStr("LocalServer")
	dd CStr("AutoConvertTo")
	dd CStr("Ole1Class")
	dd CStr("RTFClassName")
	dd CStr("InProcHandler32")
	dd CStr("InProcHandler")
NUMREGKEYSEARCHES_ textequ %($ - pRegKeys) / sizeof LPSTR
NUMREGKEYSEARCHES equ NUMREGKEYSEARCHES_


	.code

__this	textequ <edi>
_this	textequ <[__this].COptPageDlg>
thisarg	textequ <this@:ptr COptPageDlg>

	MEMBER hWnd, pDlgProc
	MEMBER hWndLV, hImageList, iDragItem, iDropItem
	MEMBER bDrag, bDragHidden, bChanged, bError


SetChildDlgPos proc uses ebx __this thisarg

local	rect:RECT
local	point:POINT

		mov __this,this@

		invoke GetParent,m_hWnd
		invoke GetWindowLong,eax,DWL_USER
		mov ebx,eax

		invoke GetChildPos,[ebx].COptionDlg.hWndTab
		movzx ecx,ax
		mov point.x,ecx
		shr eax,16
		mov point.y,eax
		invoke GetClientRect,[ebx].COptionDlg.hWndTab,addr rect
		invoke TabCtrl_AdjustRect( [ebx].COptionDlg.hWndTab, FALSE, addr rect )
		mov eax,point.x
		add rect.left,eax
		mov eax,point.y
		add rect.top,eax
if 1
		invoke SetWindowPos,m_hWnd,NULL,rect.left,rect.top,0,0,\
					SWP_SHOWWINDOW or SWP_NOSIZE or SWP_NOZORDER
else
		invoke SetWindowPos,m_hWnd,HWND_TOP,rect.left,rect.top,0,0,\
					SWP_SHOWWINDOW or SWP_NOSIZE
endif
		ret
		align 4
		assume ebx:nothing

SetChildDlgPos endp


CalcLVLine proc uses ebx lpPoint:ptr POINT

local	hti:LVHITTESTINFO

		invoke GetChildPos, m_hWndLV
		movzx edx,ax		;xPos
		shr eax,16
		mov ecx,eax			;yPos

		mov ebx,lpPoint
		mov eax,[ebx].POINT.x
		sub eax,edx
		mov hti.pt.x,eax
		mov [ebx].POINT.x,eax

		mov eax,[ebx].POINT.y
		sub eax,ecx
		mov hti.pt.y,eax
		mov [ebx].POINT.y,eax
		
		invoke ListView_HitTest( m_hWndLV,addr hti)
		ret
		align 4

CalcLVLine endp

InitCheckButtons proc uses ebx pTab:LPVOID, numButtons:DWORD

		mov ebx, pTab
		mov ecx, numButtons
		.while (ecx)
			push ecx
			mov edx, [ebx+0]
			mov eax, [ebx+4]
			movzx eax, byte ptr [eax]
			invoke CheckDlgButton, m_hWnd, edx, eax
			pop ecx
			add ebx,8
			dec ecx
		.endw
		ret
		align 4

InitCheckButtons endp

GetCheckButton proc uses ebx pTab:LPVOID, numButtons:DWORD, iButton:DWORD

		mov ebx, pTab
		mov ecx, numButtons
		mov eax, iButton
		.while (ecx)
			.if (eax == [ebx+0])
				mov edx,[ebx+4]
				xor byte ptr [edx],1
				.break
			.endif
			add ebx,8
			dec ecx
		.endw
		ret
		align 4

GetCheckButton endp

;----------------------------------------------------------------------
;*** Dialog Proc for "Properties" options dialog

	.const

checktabProperties label dword
	dd IDC_CHECK1, offset g_bOwnWndForPropDlg
	dd IDC_CHECK2, offset g_bNewDlgForMethods
	dd IDC_CHECK3, offset g_bDepTypeLibDlg
	dd IDC_CHECK4, offset g_bTranslateUDTs
	dd IDC_CHECK5, offset g_bNoDispatchPropScan
	dd IDC_CHECK6, offset g_bDispUserCalls
	dd IDC_CHECK7, offset g_bShowForceTypeInfo
	dd IDC_CHECK8, offset g_bPropDlgAsTopLevelWnd
	dd IDC_CHECK9, offset g_bCollDlgAsTopLevelWnd
	dd IDC_CHECK10,offset g_bCloseCollDlgOnDlbClk
NUMCHECKSPROPERTIES textequ %($ - offset checktabProperties) / 8

	.code

OptionPropertiesDialog proc uses __this thisarg, message:dword, wParam:WPARAM, lParam:LPARAM

local	bTranslated:BOOL

		mov __this,this@

		mov eax,message
		.if (eax == WM_INITDIALOG)
			
			invoke InitCheckButtons, offset checktabProperties, NUMCHECKSPROPERTIES

			invoke SetDlgItemInt, m_hWnd, IDC_EDIT1, g_MaxCollItems, FALSE

			invoke SetChildDlgPos,__this
			mov eax,1
		.elseif (eax == WM_COMMAND)
			movzx eax,word ptr wParam+0
			push eax
			invoke GetCheckButton, offset checktabProperties, NUMCHECKSPROPERTIES, eax
			pop eax
			.if (eax == IDC_EDIT1)
				movzx eax, word ptr wParam+2
				.if (eax == EN_CHANGE)
					invoke GetDlgItemInt, m_hWnd, IDC_EDIT1, addr bTranslated, FALSE
					.if (bTranslated)
						mov g_MaxCollItems, eax
					.endif
				.endif
			.elseif (eax == IDM_REFRESH)
				invoke InitCheckButtons, offset checktabProperties, NUMCHECKSPROPERTIES
			.endif
		.else
			xor eax,eax
		.endif
		ret
		align 4

OptionPropertiesDialog endp


;----------------------------------------------------------------------
;*** Dialog Proc for "View Control" options dialog

	.const

checktabViewControl label dword
	dd IDC_CHECK1, offset g_bViewDlgAsTopLevelWnd
	dd IDC_CHECK2, offset g_bInPlaceSiteExSupp
	dd IDC_CHECK3, offset g_bAllowWindowless
	dd IDC_CHECK4, offset g_bDispatchSupp
	dd IDC_CHECK5, offset g_bDocumentSiteSupp
	dd IDC_CHECK6, offset g_bCommandTargetSupp
	dd IDC_CHECK7, offset g_bDrawIfNotActive
	dd IDC_CHECK8, offset g_bConfirmSaveReq
	dd IDC_CHECK9, offset g_bUseIQuickActivate
	dd IDC_CHECK10, offset g_bUseIPersistPropBag
	dd IDC_CHECK11, offset g_bUseIPersistStream
	dd IDC_CHECK12, offset g_bUseIPersistFile
	dd IDC_CHECK13, offset g_bDispQueryIFCalls
	dd IDC_CHECK14, offset g_bDispContainerCalls
	dd IDC_CHECK15, offset g_bServiceProviderSupp
	dd IDC_CHECK16, offset g_bUseIPointerInactive
NUMCHECKSVIEWCONTROL textequ %($ - offset checktabViewControl) / 8

	.code

OptionViewControlDialog proc uses __this thisarg, message:dword, wParam:WPARAM, lParam:LPARAM

		mov __this,this@

		mov eax,message
		.if (eax == WM_INITDIALOG)

			invoke InitCheckButtons, offset checktabViewControl, NUMCHECKSVIEWCONTROL

			invoke SetChildDlgPos,__this
			mov eax,1
		.elseif (eax == WM_COMMAND)
			movzx eax,word ptr wParam+0
			invoke GetCheckButton, offset checktabViewControl, NUMCHECKSVIEWCONTROL, eax
		.else
			xor eax,eax
		.endif
		ret
		align 4

OptionViewControlDialog endp


;----------------------------------------------------------------------
;*** Dialog Proc for "Object" options dialog

	.const

checktabObject label dword
	dd IDC_CHECK1, offset g_bObjDlgsAsTopLevelWnd
	dd IDC_CHECK2, offset g_bFreeLibs
	dd IDC_CHECK3, offset g_bQueryMI
	dd IDC_CHECK4, offset g_bUseEnumCPs
	dd IDC_CHECK5, offset g_bGrayClrforIF
NUMCHECKSOBJECT textequ %($ - offset checktabObject) / 8

	.code

OptionObjectDialog proc uses __this thisarg, message:dword, wParam:WPARAM, lParam:LPARAM

		mov __this,this@

		mov eax,message
		.if (eax == WM_INITDIALOG)

			invoke InitCheckButtons, offset checktabObject, NUMCHECKSOBJECT

			invoke SetChildDlgPos,__this
			mov eax,1
		.elseif (eax == WM_COMMAND)
			movzx eax,word ptr wParam+0
			push eax
			invoke GetCheckButton, offset checktabObject, NUMCHECKSOBJECT, eax
			pop eax
			.if ((eax == IDC_CHECK2) && g_bFreeLibs)
				invoke CoFreeUnusedLibraries
			.endif
		.else
			xor eax,eax
		.endif
		ret
		align 4

OptionObjectDialog endp


;----------------------------------------------------------------------
;*** Dialog Proc for "General" options dialog

	.const

checktabGeneral label dword
	dd IDC_CHECK1, offset g_bConfirmDelete
	dd IDC_CHECK2, offset g_bBindIsDefault
	dd IDC_CHECK3, offset g_bOneInstance
	dd IDC_CHECK4, offset g_bMultiDoc
	dd IDC_CHECK5, offset g_bLogToDebugWnd
NUMCHECKSGENERAL textequ %($ - offset checktabGeneral) / 8

	.code

OptionGeneralDialog proc uses __this thisarg, message:dword, wParam:WPARAM, lParam:LPARAM

local szText[32]:byte

		mov __this,this@

		mov eax,message
		.if (eax == WM_INITDIALOG)
			
			invoke InitCheckButtons, offset checktabGeneral, NUMCHECKSGENERAL

			invoke wsprintf, addr szText, CStr("0x%X"), g_LCID
			invoke SetDlgItemText, m_hWnd, IDC_EDIT1, addr szText

			invoke wsprintf, addr szText, CStr("%u"), g_dwFontWidth
			invoke SetDlgItemText, m_hWnd, IDC_EDIT2, addr szText

			invoke SetChildDlgPos,__this
			mov eax,1
		.elseif (eax == WM_COMMAND)
			movzx eax,word ptr wParam
			.if (eax == IDC_CHECK4)
				push eax
				invoke Reset@CMainDlg, g_pMainDlg
				pop eax
			.endif
			push eax
			invoke GetCheckButton, offset checktabGeneral, NUMCHECKSGENERAL, eax
			pop eax
			.if (eax == IDC_EDIT1)
				invoke GetDlgItemText, m_hWnd, IDC_EDIT1, addr szText, sizeof szText
				mov g_LCID, LOCALE_SYSTEM_DEFAULT
				.if (eax)
					invoke String2DWord, addr szText, addr g_LCID
				.endif
			.elseif (eax == IDC_EDIT2)
				invoke GetDlgItemText, m_hWnd, IDC_EDIT2, addr szText, sizeof szText
				.if (eax)
					invoke String2DWord, addr szText, addr g_dwFontWidth
				.endif
			.elseif (eax == IDC_CHECK4)
				mov ecx, g_pMainDlg
				invoke SendMessage, [ecx].CDlg.hWnd, WM_COMMAND, IDM_REFRESH, 0
			.endif
			invoke UpdateLogSwitch@CLogWindow
		.else
			xor eax,eax
		.endif
		ret
		align 4

OptionGeneralDialog endp


;----------------------------------------------------------------------
;*** Dialog Proc for "Typelib" options dialog


	.const

checktabTypeLib label dword
	dd IDC_CHECK1, offset g_bUseQueryPath
	dd IDC_CHECK2, offset g_bCreateMaxDispHlp
	dd IDC_CHECK3, offset g_bTypeFlagsAsNumber
	dd IDC_CHECK4, offset g_bMemIdInDecimal
	dd IDC_CHECK5, offset g_bWriteClipBoard
	dd IDC_CHECK6, offset g_bTLibDlgAsTopLevelWnd
	dd IDC_CHECK7, offset g_bValueInDecimal
NUMCHECKSTYPELIB textequ %($ - offset checktabTypeLib) / 8

	.code

OptionTypelibDialog proc uses ebx __this thisarg, message:dword, wParam:WPARAM, lParam:LPARAM

local	szStr[64]:byte

		mov __this,this@

		mov eax,message
		.if (eax == WM_INITDIALOG)

			invoke InitCheckButtons, offset checktabTypeLib, NUMCHECKSTYPELIB
			mov m_bChanged,FALSE
			invoke SetChildDlgPos,__this
			mov eax,1
		.elseif (eax == WM_COMMAND)
			movzx eax,word ptr wParam
			invoke GetCheckButton, offset checktabTypeLib, NUMCHECKSTYPELIB, eax
		.else
			xor eax,eax
		.endif
		ret
		align 4

OptionTypelibDialog endp


;----------------------------------------------------------------------
;*** handle WM_INITDIALOG of HKCR options dialog


COptionHKCROnInitDialog proc uses ebx

local	rect:RECT
local	hWndCB:HWND
local	bFlags[4]:dword

		invoke GetDlgItem,m_hWnd,IDC_COMBO1
		mov hWndCB,eax
		.if (eax == NULL)
			ret
		.endif
		invoke ComboBox_LimitText( hWndCB, sizeof g_szUserColHKCR - 1)

		invoke MakeUDColumnList, hWndCB, MODE_HKCR, 0
		ret
		align 4

COptionHKCROnInitDialog endp


;*** handle WM_COMMAND of HKCR options dialog


COptionHKCROnCommand proc wParam:WPARAM, lParam:LPARAM

local	szText[64]:byte
local	pNewStr:LPSTR
local	hWndCB:HWND

		movzx eax,word ptr wParam
		.if (eax == IDOK)

			invoke GetDlgItem,m_hWnd,IDC_COMBO1
			mov hWndCB,eax
			invoke SendMessage,hWndCB,CB_GETCURSEL,0,0
			.if (eax == CB_ERR)
				invoke GetWindowText,hWndCB,addr szText,sizeof szText
				.if (eax != 0)
					lea eax,szText
				.endif
			.else
				mov eax,[eax*sizeof dword + offset UserColumnHKCR]
			.endif
			mov pNewStr,eax
			.if (eax)
				invoke lstrcmp, eax, addr g_szUserColHKCR
				.if (eax)
					DebugOut "New Usercol for HKCR"
					invoke lstrcpy, addr g_szUserColHKCR, pNewStr
					invoke UserColChanged@CMainDlg, g_pMainDlg, MODE_HKCR
				.endif
			.else
				mov g_szUserColHKCR,0
			.endif

		.elseif (eax == IDC_COMBO1)
		.endif
		ret
		align 4

COptionHKCROnCommand endp


;--- Dialog Proc for "HKCR" option dialog


OptionHKCRDialog proc uses __this thisarg, message:dword, wParam:WPARAM, lParam:LPARAM

		mov __this,this@

		mov eax,message
		.if (eax == WM_INITDIALOG)
			invoke COptionHKCROnInitDialog
			invoke SetChildDlgPos,__this
			mov eax,1
		.elseif (eax == WM_COMMAND)
			invoke COptionHKCROnCommand, wParam, lParam
		.else
			xor eax,eax
		.endif
		ret
		align 4

OptionHKCRDialog endp

;--------------------------------------------------------------

;*** handle WM_INITDIALOG of Interface options dialog

COptionInterfaceOnInitDialog proc uses ebx

local	hWndCB:HWND

		movzx eax, g_bExcludeProxy
		invoke CheckDlgButton,m_hWnd,IDC_CHECK1,eax
		movzx eax, g_bExcludeTypeLib
		invoke CheckDlgButton,m_hWnd,IDC_CHECK2,eax

		invoke GetDlgItem, m_hWnd, IDC_COMBO1
		mov hWndCB,eax
		.if (eax == NULL)
			ret
		.endif
		invoke ComboBox_LimitText( hWndCB, sizeof g_szUserColInterface - 1)
		invoke MakeUDColumnList, hWndCB, MODE_INTERFACE, 0
		ret
		align 4

COptionInterfaceOnInitDialog endp


;*** handle WM_COMMAND of Interface options dialog

COptionInterfaceOnCommand proc wParam:WPARAM,lParam:LPARAM

local	pNewStr:LPSTR
local	hWndCB:HWND
local	szText[64]:byte

		movzx eax,word ptr wParam
		.if (eax == IDOK)

			invoke GetDlgItem,m_hWnd,IDC_COMBO1
			mov hWndCB,eax
			invoke SendMessage,hWndCB,CB_GETCURSEL,0,0
			.if (eax == CB_ERR)
				invoke GetWindowText,hWndCB,addr szText,sizeof szText
				.if (eax != 0)
					lea eax,szText
				.endif
			.else
				mov eax,[eax*sizeof dword + offset UserColumnInterface]
			.endif
			mov pNewStr,eax
			.if (eax)
				invoke lstrcmp, eax, addr g_szUserColInterface
				.if (eax)
					invoke lstrcpy, addr g_szUserColInterface, pNewStr
					invoke UserColChanged@CMainDlg, g_pMainDlg, MODE_INTERFACE
				.endif
			.else
				mov g_szUserColInterface,0
			.endif

		.elseif (eax == IDC_CHECK1)
			xor g_bExcludeProxy,1
		.elseif (eax == IDC_CHECK2)
			xor g_bExcludeTypeLib,1
		.elseif (eax == IDC_COMBO1)
		.endif
		ret
		align 4

COptionInterfaceOnCommand endp

;--- Dialog Proc for Interface option dialog


OptionInterfaceDialog proc uses __this thisarg, message:dword, wParam:WPARAM, lParam:LPARAM

		mov __this,this@

		mov eax,message
		.if (eax == WM_INITDIALOG)
			invoke COptionInterfaceOnInitDialog
			invoke SetChildDlgPos,__this
			mov eax,1
		.elseif (eax == WM_COMMAND)
			invoke COptionInterfaceOnCommand, wParam, lParam
		.else
			xor eax,eax
		.endif
		ret
		align 4

OptionInterfaceDialog endp


;------------------------------------------------------------

;--- to handle drag operations for a listview control,
;--- functions ArrangeTypes, XXXOnMouseMove and XXXOnLButtonUp are used

ArrangeTypes proc uses ebx esi iNewPos:dword

local	dwTemp[NUMREGKEYSEARCHES]:ptr byte

		mov esi,offset pRegKeys
		mov eax,m_iDragItem
		mov ebx,[eax*sizeof dword + esi]

		lea edx,dwTemp
		xor ecx,ecx
		.while (ecx < NUMREGKEYSEARCHES)
			lodsd
			.if (eax == ebx)
				lodsd
			.endif
			.if (ecx == iNewPos)
				mov eax,ebx
				sub esi,sizeof dword
			.endif
			mov [edx],eax
			add edx,sizeof dword
			inc ecx
		.endw
		mov edi,offset pRegKeys
		lea esi,dwTemp
		mov ecx,NUMREGKEYSEARCHES
		rep movsd
		ret
		align 4

ArrangeTypes endp



COptionCLSIDOnMouseMove proc lParam:dword

local	point:POINT

		mov ecx,lParam
		movzx eax,cx
		mov point.x,eax
		shr ecx,16
		mov point.y,ecx
		invoke CalcLVLine, addr point
		.if (eax != -1)
			push eax
			.if (m_bDragHidden)
				invoke ImageList_DragEnter, m_hWndLV,point.x,point.y
				mov m_bDragHidden,FALSE
			.endif
			.if (!g_hCsrYes)
				invoke LoadCursor,NULL,IDC_ARROW
				mov g_hCsrYes,eax
			.endif
			invoke SetCursor,g_hCsrYes
			invoke ImageList_DragMove,point.x,point.y
			pop eax
			.if (eax != m_iDropItem)
				mov m_iDropItem,eax
				invoke ImageList_DragLeave, m_hWndLV
				ListView_SetItemState m_hWndLV,m_iDropItem,LVIS_DROPHILITED,LVIS_DROPHILITED
				invoke UpdateWindow, m_hWndLV
				invoke ImageList_DragEnter, m_hWndLV,point.x,point.y
			.endif
		.else
			.if (eax != m_iDropItem)
				mov m_iDropItem,eax
				invoke ImageList_DragLeave, m_hWndLV
				mov m_bDragHidden,TRUE
				ListView_SetItemState m_hWndLV,m_iDropItem,0,LVIS_DROPHILITED
				.if (!g_hCsrNo)
					invoke LoadCursor,NULL,IDC_NO
					mov g_hCsrNo,eax
				.endif
				invoke SetCursor, g_hCsrNo
			.endif
		.endif
		ret
		align 4

COptionCLSIDOnMouseMove endp


COptionCLSIDOnLButtonUp proc lParam:dword

local	point:POINT

		.if (m_bDragHidden == FALSE)
			invoke ImageList_DragLeave, m_hWndLV
		.endif
		invoke ImageList_EndDrag
		invoke ImageList_Destroy, m_hImageList
		invoke ReleaseCapture
		mov m_bDrag,FALSE
		mov ecx,lParam
		movzx eax,cx
		mov point.x,eax
		shr ecx,16
		mov point.y,ecx
		invoke CalcLVLine, addr point
		.if (eax != -1)
			push eax
			ListView_SetItemState m_hWndLV,m_iDropItem,\
					LVIS_SELECTED or LVIS_FOCUSED,\
					LVIS_DROPHILITED or LVIS_SELECTED or LVIS_FOCUSED
			pop eax
			invoke ArrangeTypes, eax
			invoke InvalidateRect, m_hWndLV, 0, 1
		.endif
		ret
		align 4

COptionCLSIDOnLButtonUp endp


;*** handle WM_NOTIFY of CLSID options dialog


COptionCLSIDOnNotify proc uses ebx pNMHDR:ptr NMHDR

local	point:POINT
local	hImage:HANDLE

		mov ebx,pNMHDR
		assume ebx:ptr NMLISTVIEW

		mov eax,[ebx].hdr.code
		.if (eax == LVN_GETDISPINFO)
			assume ebx:ptr NMLVDISPINFO
			.if ([ebx].item.mask_ & LVIF_TEXT)
				mov eax,[ebx].item.iItem
				mov eax,[eax*sizeof LPSTR+offset pRegKeys]
				mov [ebx].item.pszText,eax
			.endif
		.elseif (eax == LVN_BEGINDRAG)
			assume ebx:ptr NMLISTVIEW
			invoke SetCapture,m_hWnd
			mov eax,[ebx].iItem
			mov m_iDragItem,eax
			mov m_iDropItem,-1
			invoke ListView_CreateDragImage( m_hWndLV, [ebx].iItem, addr point)
			mov m_hImageList,eax
			mov m_bDrag,TRUE
			invoke ImageList_BeginDrag, m_hImageList, 0, [ebx].ptAction.x,8
			invoke ImageList_DragEnter, m_hWndLV,0, [ebx].ptAction.y
			mov m_bDragHidden,FALSE
		.elseif (eax == NM_DBLCLK)
			invoke MessageBox,m_hWnd,CStr("use drag and drop to change search order"), addr g_szHint, MB_OK
		.else
			xor eax,eax
		.endif

		ret
		assume ebx:nothing
		align 4

COptionCLSIDOnNotify endp


;*** handle WM_COMMAND of CLSID options dialog


COptionCLSIDOnCommand proc wParam:WPARAM,lParam:LPARAM

local	szText[64]:byte
local	pNewStr:LPSTR
local	hWndCB:HWND

		movzx eax,word ptr wParam
		.if (eax == IDCANCEL)

			invoke PostMessage,m_hWnd,WM_CLOSE,0,0

		.elseif (eax == IDOK)

			DebugOut "IDOK, CLSID option page"
			invoke GetDlgItem, m_hWnd, IDC_COMBO1
			mov hWndCB,eax
			invoke SendMessage, hWndCB, CB_GETCURSEL,0,0
			.if (eax == CB_ERR)
				invoke GetWindowText, hWndCB, addr szText, sizeof szText
				.if (eax != 0)
					lea eax,szText
				.endif
			.else
				mov eax,[eax*sizeof dword + offset UserColumnCLSID]
			.endif
			mov pNewStr,eax
			.if (eax)
				invoke lstrcmp, eax, addr g_szUserColCLSID
				.if (eax)
					invoke lstrcpy, addr g_szUserColCLSID, pNewStr
					invoke UserColChanged@CMainDlg, g_pMainDlg, MODE_CLSID
				.endif
			.else
				mov g_szUserColCLSID,0
			.endif
;;			invoke PostMessage,m_hWnd,WM_CLOSE,0,0

		.elseif (eax == IDC_CHKHANDLER)
			xor g_bCtxInProcHandler,1
		.elseif (eax == IDC_CHKSERVER)
			xor g_bCtxInProcServer,1
		.elseif (eax == IDC_CHKLOCAL)
			xor g_bCtxLocalServer,1
		.elseif (eax == IDC_CHKREMOTE)
			xor g_bCtxRemoteServer,1
		.elseif (eax == IDC_CHECK3)
			xor g_bUseClassFactory2,1
		.elseif (eax == IDC_CHECK1)
			xor g_bAddAutoTreatEntries,1
		.elseif (eax == IDC_COMBO1)
		.endif
		ret
		align 4

COptionCLSIDOnCommand endp

;--- if dwIDStart is 0, a combobox is filled instead of a menu

MakeUDColumnList proc public uses ebx esi edi handle:HANDLE, mode:DWORD, dwIDStart:DWORD

local bIsDefault:BOOL

		xor ebx, ebx
		.if (mode == MODE_CLSID)
			mov esi, offset UserColumnCLSID
			mov edi, offset g_szUserColCLSID
			mov ecx, NUMUSERCOLUMNS_CLSID
		.elseif (mode == MODE_HKCR)
			mov esi, offset UserColumnHKCR
			mov edi, offset g_szUserColHKCR
			mov ecx, NUMUSERCOLUMNS_HKCR
		.elseif (mode == MODE_INTERFACE)
			mov esi, offset UserColumnInterface
			mov edi, offset g_szUserColInterface
			mov ecx, NUMUSERCOLUMNS_INTERFACE
		.else
			mov esi, offset UserColumnTypeInfo
			mov edi, offset g_szNull
			mov ecx, NUMUSERCOLUMNS_TYPEINFO
		.endif

		mov bIsDefault, FALSE
		.while (ebx < ecx)
			push ecx
			.if (dwIDStart)
				invoke lstrcmpi, edi, [ebx*sizeof dword+esi]
				.if (!eax)
					mov bIsDefault, TRUE
					mov edx, MF_STRING or MF_CHECKED
				.else
					mov edx, MF_STRING
				.endif
				mov ecx, dwIDStart
				add ecx, ebx
				invoke AppendMenu, handle, edx, ecx, [ebx*sizeof dword+esi]
			.else
				invoke SendMessage, handle, CB_ADDSTRING,0,[ebx*sizeof dword+esi]
			.endif
			pop ecx
			inc ebx
		.endw
		.if (byte ptr [edi])
			.if (dwIDStart)
				.if (bIsDefault == FALSE)
					inc ebx
					invoke AppendMenu, handle, MF_STRING or MF_CHECKED, ebx, edi
				.endif
			.else
				invoke SendMessage,handle, CB_SELECTSTRING, -1, edi
				invoke SetWindowText, handle, edi
			.endif
		.else
			.if (!dwIDStart)
				invoke SendMessage, handle, CB_SETCURSEL, ebx, 0
			.endif
		.endif
		ret
		align 4

MakeUDColumnList endp


;*** handle WM_INITDIALOG of CLSID options dialog


COptionCLSIDOnInitDialog proc uses ebx

local	lvc:LVCOLUMN
local	rect:RECT
local	hWndCB:HWND
local	bFlags[4]:dword

;---------------------------------- "type search order" listview
		invoke GetDlgItem,m_hWnd,IDC_LIST1
		mov m_hWndLV,eax
		mov m_bDrag,FALSE

		invoke GetClientRect, m_hWndLV,addr rect

		mov lvc.mask_,LVCF_TEXT or LVCF_WIDTH
		mov eax,rect.right
		mov lvc.cx_,eax
		mov lvc.pszText,CStr("Type")
		invoke ListView_InsertColumn( m_hWndLV,0,addr lvc)
		invoke ListView_SetItemCount( m_hWndLV,NUMREGKEYSEARCHES)

;---------------------------------- "user defined column" combobox
		invoke GetDlgItem,m_hWnd,IDC_COMBO1
		mov hWndCB,eax
		.if (eax)
			invoke ComboBox_LimitText( hWndCB, sizeof g_szUserColCLSID - 1)
			invoke MakeUDColumnList, hWndCB, MODE_CLSID, 0
		.endif

;---------------------------------- buttons

		movzx eax, g_bCtxInProcHandler
		invoke CheckDlgButton, m_hWnd, IDC_CHKHANDLER, eax

		movzx eax, g_bCtxInProcServer
		invoke CheckDlgButton, m_hWnd, IDC_CHKSERVER, eax

		movzx eax, g_bCtxLocalServer
		invoke CheckDlgButton, m_hWnd, IDC_CHKLOCAL, eax

		movzx eax, g_bCtxRemoteServer
		invoke CheckDlgButton, m_hWnd, IDC_CHKREMOTE, eax

		movzx eax, g_bUseClassFactory2
		invoke CheckDlgButton, m_hWnd, IDC_CHECK3, eax

		movzx eax, g_bAddAutoTreatEntries
		invoke CheckDlgButton, m_hWnd, IDC_CHECK1, eax

		ret
		align 4

COptionCLSIDOnInitDialog endp


;*** Dialog Proc for "options CLSID" dialog


OptionCLSIDDialog proc uses ebx __this thisarg, message:dword, wParam:WPARAM, lParam:LPARAM

		mov __this,this@

		mov eax,message
		.if (eax == WM_INITDIALOG)
			invoke COptionCLSIDOnInitDialog
			invoke SetChildDlgPos,__this
			mov eax,1
		.elseif (eax == WM_COMMAND)
			invoke COptionCLSIDOnCommand, wParam, lParam
		.elseif (eax == WM_NOTIFY)
			invoke COptionCLSIDOnNotify, lParam
		.elseif (eax == WM_MOUSEMOVE)
			.if (m_bDrag == TRUE)
				invoke COptionCLSIDOnMouseMove, lParam
			.endif
		.elseif (eax == WM_LBUTTONUP)
			.if (m_bDrag == TRUE)
				invoke COptionCLSIDOnLButtonUp, lParam
			.endif
		.else
			xor eax,eax ;indicates "no processing"
		.endif
		ret
		align 4

OptionCLSIDDialog endp

;-------------------------------------- new class -------------------------------

__this	textequ <edi>
_this	textequ <[__this].COptionDlg>
thisarg	textequ <this@:ptr COptionDlg>

		MEMBER hWndTab, tab

SelectTabDialog proc iIndex:dword

		mov eax, iIndex
		mov g_dwOptPage,eax

		mov ecx,sizeof CTabDlgPage
		mul ecx
		add eax,offset TabDlgPages
		mov ecx,[eax].CTabDlgPage.pDlgProc
		mov m_tab.pDlgProc,ecx
		mov ecx,[eax].CTabDlgPage.dwResID
		invoke CreateDialogParam,g_hInstance,ecx,m_hWnd,classdialogproc,addr m_tab
if 1
		invoke SetWindowPos, eax, m_hWndTab, 0, 0, 0, 0,
					SWP_NOMOVE or SWP_NOSIZE
endif
		invoke SetFocus, m_tab.hWnd
		ret
		align 4

SelectTabDialog endp


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
			.if (m_tab.hWnd != NULL)
				mov m_tab.bError,FALSE
				invoke SendMessage,m_tab.hWnd,WM_COMMAND,IDOK,0
				.if (m_tab.bError == FALSE)
					invoke DestroyWindow,m_tab.hWnd
					mov m_tab.hWnd,NULL
				.else
					invoke SetWindowLong,m_hWnd,DWL_MSGRESULT,TRUE
					mov eax,1
				.endif
			.endif
		.endif

		ret
		assume ebx:nothing
		align 4

OnNotify endp


;*** Dialog Proc for "option" dialog


OptionDialog proc uses ebx __this thisarg, message:dword, wParam:WPARAM, lParam:LPARAM

local	tci:TC_ITEM

		mov __this,this@

		mov eax,message
		.if (eax == WM_INITDIALOG)
			mov eax, m_hWnd
			mov g_hWndOption, eax
if ?MODELESS
			invoke GetParent, m_hWnd
			invoke EnableWindow, eax, FALSE
endif

			invoke GetDlgItem, m_hWnd, IDC_TAB1
			mov m_hWndTab,eax

			mov m_tab.hWnd,NULL

			mov tci.mask_,TCIF_TEXT or TCIF_PARAM
			mov ebx,offset TabDlgPages
			mov ecx,0
			.while (ecx < NUMOPTDLGS)
				push ecx
				mov tci.lParam,ebx
				mov eax,[ebx].CTabDlgPage.pTabName
				mov tci.pszText,eax
				invoke TabCtrl_InsertItem( m_hWndTab,ecx,addr tci)
				add ebx,sizeof CTabDlgPage
				pop ecx
				inc ecx
			.endw

			invoke CenterWindow, m_hWnd

			mov eax, g_dwOptPage
			push eax	
			invoke TabCtrl_SetCurSel( m_hWndTab, eax)
			pop eax
			invoke SelectTabDialog, eax

			mov eax,1

		.elseif (eax == WM_CLOSE)
if ?MODELESS
			invoke GetParent, m_hWnd
			invoke EnableWindow, eax, TRUE
			invoke DestroyWindow, m_hWnd
else
			invoke EndDialog, m_hWnd, 0
endif
			mov eax,1

		.elseif (eax == WM_DESTROY)

			invoke Destroy@COptionDlg, __this
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
		.elseif (eax == WM_COMMAND)
			movzx eax,word ptr wParam
			.if (eax == IDCANCEL)
				invoke PostMessage,m_hWnd,WM_CLOSE,0,0
			.elseif (eax == IDOK)
				mov m_tab.bError,FALSE
				.if (m_tab.hWnd)
					invoke SendMessage,m_tab.hWnd,WM_COMMAND,IDOK,0
				.endif
				.if (!m_tab.bError)
					invoke PostMessage,m_hWnd,WM_CLOSE,0,0
				.endif
			.elseif (eax == IDM_REFRESH)
				.if (m_tab.hWnd)
					invoke SendMessage,m_tab.hWnd,WM_COMMAND,eax,0
				.endif
			.endif
		.elseif (eax == WM_NOTIFY)
			invoke OnNotify, lParam
if ?HTMLHELP
		.elseif (eax == WM_HELP)

			mov eax, g_dwOptPage
			.if (eax == 6)
				mov ecx, CStr("ViewControl")
			.else
				mov ecx,sizeof CTabDlgPage
				mul ecx
				add eax,offset TabDlgPages
				mov ecx,[eax].CTabDlgPage.pTabName
			.endif
			sub esp, 128
			mov edx, esp
			invoke wsprintf, edx, CStr("optionsdialog.htm#Options%s"), ecx
			mov edx, esp
			invoke DoHtmlHelp, HH_DISPLAY_TOPIC, edx
			add esp, 128
endif
		.else
			xor eax,eax ;indicates "no processing"
		.endif
		ret
		align 4

OptionDialog endp


;*** constructor


Create@COptionDlg proc public

		invoke malloc, sizeof COptionDlg
		ret
		align 4
Create@COptionDlg endp

Show@COptionDlg proc public uses __this thisarg, hWnd:HWND
		
		mov __this,this@
		mov m_pDlgProc,OptionDialog
if ?MODELESS
		invoke CreateDialogParam, g_hInstance, IDD_OPTIONS, hWnd, classdialogproc,__this
else
		invoke DialogBoxParam, g_hInstance, IDD_OPTIONS, hWnd, classdialogproc,__this
endif
		ret
		align 4
Show@COptionDlg endp

Destroy@COptionDlg proc public thisarg

		mov ecx, this@
		invoke SetWindowLong, [ecx].COptionDlg.tab.hWnd, DWL_USER, NULL
		invoke free, this@
		mov g_hWndOption, NULL
		ret
		align 4
Destroy@COptionDlg endp

    end
