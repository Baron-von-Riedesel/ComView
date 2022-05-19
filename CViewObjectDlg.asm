

;*** definition of class CViewObjectDlg
;*** CViewObjectDlg implements a simple dialog hosting a CContainer object

	.386
	.model flat,stdcall
	option casemap :none   ; case sensitive
	option proc:private

	include COMView.inc
	include statusbar.inc
INSIDE_CVIEWOBJECTDLG equ 1
	include classes.inc
	include rsrc.inc
	include debugout.inc

?LOADFILE	equ 1		;show menu item "Load from File"

OLEIVERB_PRIMARY		equ 0
OLEIVERB_SHOW			equ -1
OLEIVERB_OPEN			equ -2
OLEIVERB_HIDE			equ -3
OLEIVERB_UIACTIVATE		equ -4
OLEIVERB_INPLACEACTIVATE equ -5
OLEIVERB_PROPERTIES		equ -7
DISPID_AMBIENT_USERMODE equ -709
DISPID_AMBIENT_UIDEAD	equ -710

OLEVERB_START equ 50000

BEGIN_CLASS CViewObjectDlg, CDlg
hWndSB			HWND ?
pObjectItem		LPOBJECTITEM ?
pUnknown		LPUNKNOWN ?
pContainer		pCContainer ?
pItem			pCInterfaceItem ?
pTypeInfoDlg	pCTypeInfoDlg ?
hIcon			HICON ?
dwUserVerbMax	DWORD ?
bViewObject		BOOLEAN ?
bObjectWithSite	BOOLEAN ?
bFilename		BOOLEAN ?
END_CLASS

STATUSCLASSNAME textequ <CStr("msctls_statusbar32")>

	.data

public g_dwBorder

g_rect			RECT {0,0,480,360}
g_dwBorder		DWORD 8
g_aViewControlClass	WORD NULL
g_bFilename		BOOLEAN FALSE
g_hIconCtrl		HICON NULL

if 0
externdef g_bLoadFile:BOOLEAN
endif

	.data?

g_szCaption		DB 64 dup (?)

	.const

szClassName db "controlcontainer",0

	.code

;--------------------------------------------------------------
;--- class CViewObjectDlg
;--------------------------------------------------------------

__this	textequ <ebx>
_this	textequ <[__this].CViewObjectDlg>
thisarg	textequ <this@:ptr CViewObjectDlg>


	MEMBER hWnd, pDlgProc, hWndSB, pObjectItem, pUnknown, pItem, pContainer, bViewObject
	MEMBER pItem, pTypeInfoDlg
	MEMBER bObjectWithSite, hIcon, dwUserVerbMax, bFilename

Register proc

local wc:WNDCLASS
		
		mov wc.style, 0
		mov wc.lpfnWndProc, OFFSET wndproc
		mov wc.cbClsExtra,NULL
		mov wc.cbWndExtra, sizeof DWORD
		push  g_hInstance
		pop wc.hInstance
		mov wc.hbrBackground, NULL
		mov wc.lpszMenuName, IDR_MENU4
		mov wc.lpszClassName, offset szClassName
		mov eax, g_hIconCtrl
		invoke LoadIcon,g_hInstance,IDI_CONTAINER
		mov g_hIconCtrl, eax
		mov wc.hIcon, eax
		invoke LoadCursor,NULL,IDC_ARROW
		mov wc.hCursor,eax
		invoke RegisterClass, addr wc
		mov	g_aViewControlClass,ax
		ret
		align 4
Register endp


Create@CViewObjectDlg proc public uses esi __this pObjectItem:LPOBJECTITEM, pItem:pCInterfaceItem

		invoke malloc, sizeof CViewObjectDlg
		.if (!eax)
			ret
		.endif

		mov __this,eax

		.if (!g_aViewControlClass)
			invoke Register
			mov g_szCaption, 0
		.endif

		mov eax, pObjectItem
		mov m_pObjectItem, eax
		.if (eax)
			invoke vf(eax, IObjectItem, AddRef)
			invoke GetUnknown@CObjectItem, m_pObjectItem
			mov m_pUnknown, eax
		.endif

		mov al, g_bFilename
		mov m_bFilename, al

		mov esi, pItem
		mov m_pItem, esi
		.while (esi)
			invoke IsEqualGUID,addr [esi].CInterfaceItem.iid, addr IID_IViewObject
			.if (eax)
				mov m_bViewObject, TRUE
				.break
			.endif
			invoke IsEqualGUID,addr [esi].CInterfaceItem.iid, addr IID_IViewObject2
			.if (eax)
				mov m_bViewObject, TRUE	
				.break
			.endif
			invoke IsEqualGUID,addr [esi].CInterfaceItem.iid, addr IID_IObjectWithSite
			.if (eax)
				mov m_bObjectWithSite, TRUE	
				.break
			.endif
			.break
		.endw

		invoke vf(m_pObjectItem, IObjectItem, SetViewObjectDlg), __this

		mov eax, __this
		ret
		align 4

Create@CViewObjectDlg endp

Show@CViewObjectDlg proc public thisarg, hWnd:HWND

if 1
		invoke CreateWindowEx, WS_EX_STATICEDGE, addr szClassName, CStr("View Control : %s"),\
			WS_OVERLAPPEDWINDOW or WS_CLIPCHILDREN, g_rect.left, g_rect.top, g_rect.right, g_rect.bottom,\
			hWnd, NULL, g_hInstance, this@
		mov ecx, this@
		invoke ShowWindow, [ecx].CViewObjectDlg.hWnd, SW_SHOWNORMAL
else
		invoke CreateDialogParam, g_hInstance, IDD_VIEWOBJECTDLG, ecx, classdialogproc, eax
endif
		ret
Show@CViewObjectDlg endp

Destroy@CViewObjectDlg proc uses __this thisarg

		mov __this,this@

		.if (m_pContainer)
			invoke Close@CContainer, m_pContainer
			invoke vf(m_pContainer, IUnknown, Release)
			.if (eax)
				invoke printf@CLogWindow, CStr("Destroy@CViewObjectDlg: Container::Release returned %u",10), eax
				DebugOut "Destroy@CViewObjectDlg: Container::Release returned %u", eax
			.endif
			mov m_pContainer, NULL
		.endif
		.if (m_pObjectItem)
			invoke vf(m_pObjectItem, IObjectItem, SetViewObjectDlg), NULL
			invoke vf(m_pObjectItem, IObjectItem, Release)
		.endif
		.if (m_hIcon)
			invoke DestroyIcon, m_hIcon
		.endif
		invoke SetWindowLong, m_hWnd, 0, 0
		invoke free, __this
		ret
		align 4

Destroy@CViewObjectDlg endp

OnCommand proc wParam:WPARAM, lParam:LPARAM

local	hMenu:HMENU
local	pPCI:LPPROVIDECLASSINFO
local	hWndParent:HWND

		invoke GetMenu, m_hWnd
		mov hMenu, eax

		invoke SetStatusText@CViewObjectDlg, __this, 0, addr g_szNull

		movzx eax,word ptr wParam
		.if ((eax == IDCANCEL) || (eax == IDM_EXITVIEWDLG))
			invoke PostMessage,m_hWnd,WM_CLOSE,0,0
		.elseif (eax == IDM_SAVESTREAM)
			invoke Save@CContainer, m_pContainer, SAVE_STREAM, TRUE
		.elseif (eax == IDM_SAVESTORAGE)
			invoke Save@CContainer, m_pContainer, SAVE_STORAGE, TRUE
		.elseif (eax == IDM_SAVEFILE)
			invoke Save@CContainer, m_pContainer, SAVE_FILE, TRUE
		.elseif (eax == IDM_SAVEPROPBAG)
			invoke Save@CContainer, m_pContainer, SAVE_PROPBAG, TRUE
if ?LOADFILE
		.elseif (eax == IDM_LOADFROMFILE)
			invoke Load@CContainer, m_pContainer
			.if (eax == S_OK)
				invoke PostMessage, m_hWnd, WM_COMMAND, IDM_UPDATEWINDOWCAPTION, TRUE
			.endif
endif
		.elseif (eax == IDM_PRIMARY)
			invoke DoVerb@CContainer, m_pContainer, OLEIVERB_PRIMARY
		.elseif (eax == IDM_SHOW)
			invoke DoVerb@CContainer, m_pContainer, OLEIVERB_SHOW
		.elseif (eax == IDM_OPEN)
			invoke DoVerb@CContainer, m_pContainer, OLEIVERB_OPEN
		.elseif (eax == IDM_HIDE)
			invoke DoVerb@CContainer, m_pContainer, OLEIVERB_HIDE
		.elseif (eax == IDM_UIACTIVATE)
			invoke DoVerb@CContainer, m_pContainer, OLEIVERB_UIACTIVATE
		.elseif (eax == IDM_INPLACEACTIVATE)
			invoke DoVerb@CContainer, m_pContainer, OLEIVERB_INPLACEACTIVATE
		.elseif (eax == IDM_VERBPROPERTIES)
			invoke DoVerb@CContainer, m_pContainer, OLEIVERB_PROPERTIES
		.elseif ((eax >= OLEVERB_START) && (eax <= m_dwUserVerbMax))
			sub eax, OLEVERB_START
			invoke DoVerb@CContainer, m_pContainer, eax
		.elseif (eax == IDM_CLOSEOBJECT)

			invoke Close2@CContainer, m_pContainer

		.elseif (eax == IDM_INPLACEDEACTIVATE)

			invoke InPlaceDeactivate@CContainer, m_pContainer

		.elseif (eax == IDM_UIDEACTIVATE)

			invoke UIDeactivate@CContainer, m_pContainer

		.elseif (eax == IDM_UPDATE)

			invoke Update@CContainer, m_pContainer

		.elseif (eax == IDM_ADVISE)

			invoke GetMenuState, hMenu, IDM_ADVISE, MF_BYCOMMAND
			push eax
			invoke Advise@CContainer, m_pContainer, 0
			mov ecx, eax
			pop eax
			and eax, MF_CHECKED
			xor al, MF_CHECKED
			.if (ecx == S_OK)
				invoke CheckMenuItem, hMenu, IDM_ADVISE, eax
			.endif

		.elseif (eax == IDM_VIEWADVISE)

			invoke GetMenuState, hMenu, IDM_VIEWADVISE, MF_BYCOMMAND
			push eax
			invoke Advise@CContainer, m_pContainer, 1
			mov ecx, eax
			pop eax
			and eax, MF_CHECKED
			xor al, MF_CHECKED
			.if (ecx == S_OK)
				invoke CheckMenuItem, hMenu, IDM_VIEWADVISE, eax
			.endif

		.elseif (eax == IDM_DATAADVISE)

			invoke GetMenuState, hMenu, IDM_DATAADVISE, MF_BYCOMMAND
			push eax
			invoke Advise@CContainer, m_pContainer, 2
			mov ecx, eax
			pop eax
			and eax, MF_CHECKED
			xor al, MF_CHECKED
			.if (ecx == S_OK)
				invoke CheckMenuItem, hMenu, IDM_DATAADVISE, eax
			.endif
if 0
		.elseif (eax == IDM_USERMODE)

			xor g_bUserMode, 1
;------------------------------- notify control of change
			.if (m_pContainer)
				invoke OnAmbientPropertyChange@CContainer, m_pContainer, DISPID_AMBIENT_USERMODE
			.endif

		.elseif (eax == IDM_UIDEAD)

			xor g_bUIDead, 1
;------------------------------- notify control of change
			.if (m_pContainer)
				invoke OnAmbientPropertyChange@CContainer, m_pContainer, DISPID_AMBIENT_UIDEAD
			.endif
else
		.elseif (eax == IDM_PROPCHANGED)

			.if (m_pContainer)
				invoke OnAmbientPropertyChange@CContainer, m_pContainer, lParam
			.endif
endif
		.elseif (eax == IDM_FILENAME)

			xor m_bFilename, 1
			mov al, m_bFilename
			mov g_bFilename, al
			invoke PostMessage, m_hWnd, WM_COMMAND, IDM_UPDATEWINDOWCAPTION, FALSE

		.elseif (eax == IDM_UPDATEWINDOWCAPTION)

			invoke SetWindowText, m_hWnd, addr g_szCaption
			invoke vf(m_pObjectItem, IObjectItem, SetWindowText_), m_hWnd
			.if (m_bFilename)
				invoke vf(m_pObjectItem, IObjectItem, AddFilename), m_hWnd, lParam
			.endif

		.elseif (eax == IDM_TYPEINFODLG)

			.if (m_pTypeInfoDlg)
				mov ecx, m_pTypeInfoDlg
				invoke RestoreAndActivateWindow, [ecx].CDlg.hWnd
			.else
				mov hWndParent, NULL
				invoke vf(m_pUnknown, IUnknown, QueryInterface), addr IID_IProvideClassInfo, addr pPCI
				.if (eax == S_OK)
					invoke vf(pPCI, IUnknown, Release)
					invoke TypeInfoDlgFromIProvideClassInfo, hWndParent, m_pUnknown
				.else
					invoke TypeInfoDlgFromIDispatch, hWndParent, m_pUnknown
				.endif
				mov m_pTypeInfoDlg, eax
			.endif

		.elseif (eax == IDM_OBJECTDLG)

			invoke vf(m_pObjectItem, IObjectItem, GetFlags)
			or eax, OBJITEMF_IGNOREOV
			invoke vf(m_pObjectItem, IObjectItem, SetFlags), eax
			mov ecx, g_pMainDlg
			invoke vf(m_pObjectItem, IObjectItem, ShowObjectDlg), [ecx].CDlg.hWnd

		.elseif (eax == IDM_PROPERTIESDLG)

			invoke vf(m_pObjectItem, IObjectItem, ShowPropertiesDlg), NULL

		.elseif (eax == IDM_VIEWSTORAGEDLG)

			invoke vf(m_pObjectItem, IObjectItem, ShowViewStorageDlg), NULL

		.endif
		ret
OnCommand endp

;--- get rect of control

GetControlRect proc pRect:ptr RECT

local	rect2:RECT

		invoke GetClientRect, m_hWnd, pRect
		invoke GetWindowRect, m_hWndSB, addr rect2
        mov edx, pRect
		mov ecx, rect2.bottom
		sub ecx, rect2.top
		sub [edx].RECT.bottom, ecx

		mov eax, g_dwBorder
		add [edx].RECT.left, eax
		add [edx].RECT.top, eax
		sub [edx].RECT.right, eax
		sub [edx].RECT.bottom, eax
		ret
GetControlRect endp        

ifdef @StackBase
	option stackbase:ebp
endif
	option prologue:@sehprologue
	option epilogue:@sehepilogue

OnPaint proc wParam:WPARAM

local	bActive:BOOL
local	pViewObject:LPVIEWOBJECT
local	ps:PAINTSTRUCT
local	rect:RECT
local	rect2:RECT
local	this@:LPVOID
local	szText[64]:byte

;;		DebugOut "OnPaint enter"
		invoke BeginPaint, m_hWnd, addr ps

		invoke GetSysColorBrush, COLOR_APPWORKSPACE
		invoke FillRect, ps.hdc, addr ps.rcPaint, eax

		xor eax, eax
		.if (m_pContainer)
			invoke IsActive@CContainer, m_pContainer
			mov bActive, eax
			.if (eax)
				invoke IsWindowless@CContainer, m_pContainer
			.elseif (g_bDrawIfNotActive)
				mov eax, 1
			.endif
		.elseif (m_bViewObject)
			mov eax, 1
		.else
			xor eax, eax
		.endif
		.if (eax)
			invoke vf(m_pUnknown, IUnknown, QueryInterface), addr IID_IViewObject, addr pViewObject
			.if (eax == S_OK)
				
				invoke SetMapMode, ps.hdc, MM_TEXT
				.if (m_bViewObject || (bActive == FALSE))
                	invoke GetControlRect, addr rect
					lea eax, rect
				.else
					xor eax, eax
				.endif
				mov this@, __this
				.try
					invoke vf(pViewObject, IViewObject, Draw), DVASPECT_CONTENT,\
							-1, NULL, NULL, 0, ps.hdc, eax, NULL, NULL, NULL
					.if (eax != S_OK)
						push eax
						StatusBar_GetTextLength m_hWndSB, 0
						pop ecx
						.if (!eax)
							invoke wsprintf, addr szText, CStr("IViewObject::Draw failed [%X]"), ecx
							invoke SetStatusText@CViewObjectDlg, __this, 0, addr szText
						.endif
					.endif
				.exceptfilter
					mov eax,_exception_info()
					mov eax, [eax].EXCEPTION_POINTERS.ExceptionRecord
					mov ecx, [eax].EXCEPTION_RECORD.ExceptionCode
					mov edx, [eax].EXCEPTION_RECORD.ExceptionAddress
					mov __this,this@
					invoke wsprintf, addr szText, CStr("Exception 0x%08X in IViewObject::Draw at 0x%08X"), ecx, edx
					invoke SetStatusText@CViewObjectDlg, __this, 0, addr szText
					mov eax,EXCEPTION_EXECUTE_HANDLER
				.except
					mov __this,this@
				.endtry
				invoke vf(pViewObject, IUnknown, Release)
			.endif
		.endif
		invoke EndPaint, m_hWnd, addr ps
		xor eax, eax
;;		DebugOut "OnPaint exit"
		ret
OnPaint endp

	option prologue: prologuedef
	option epilogue: epiloguedef
ifdef @StackBase
	option stackbase:esp
endif

OnSize proc dwWidth:DWORD, dwHeight:DWORD

local	rect:RECT
local	rectSB:RECT
local	dwWidths[3]:DWORD

		invoke GetWindowRect, m_hWndSB, addr rectSB
		mov eax, rectSB.bottom
		sub eax, rectSB.top
		mov ecx, dwHeight
		sub ecx, eax
		sub dwHeight, eax
		.if (CARRY?)
			@mov dwHeight, 0
		.endif
		invoke SetWindowPos, m_hWndSB, NULL, 0, ecx,\
				dwWidth, eax, SWP_NOZORDER

		mov dwWidths[2*type dwWidths],-1
		mov eax, dwWidth
		sub eax, 16*8
		mov dwWidths[1*type dwWidths], eax
		sub eax, 16*8
		mov dwWidths[0*type dwWidths], eax
		StatusBar_SetParts m_hWndSB, 3, addr dwWidths


if 0
		invoke GetWindowRect, m_hWndStatic, addr rect
		invoke ScreenToClient, m_hWnd, addr rect.left
else
		invoke GetClientRect, m_hWnd, addr rect
		mov eax, g_dwBorder
		add rect.left, eax
		add rect.top, eax
		sub rect.right, eax
		sub rect.bottom, eax
endif
		mov ecx,dwWidth
		mov eax,rect.left
		shl eax,1
		sub ecx, eax			;subtract 2*border from width
		.if (CARRY?)
			xor ecx, ecx
		.endif
		mov edx, ecx
		add ecx,rect.left
		mov rect.right, ecx

		mov ecx,dwHeight
		mov eax,rect.top
		shl eax,1
		sub ecx, eax			;subtract 2*border from height
		.if (CARRY?)
			xor ecx, ecx
		.endif
		mov eax, ecx
		add ecx,rect.top
		mov rect.bottom, ecx
if 0
		invoke SetWindowPos, m_hWndStatic, NULL, rect.left, rect.top,\
			edx, eax, SWP_NOMOVE or SWP_NOZORDER or SWP_NOACTIVATE
endif
		mov eax, m_pContainer
if 0
		.if (eax)
			invoke IsActive@CContainer, eax
		.endif
endif
		.if (eax)
			invoke SetRect@CContainer, m_pContainer, addr rect
		.endif
		invoke InvalidateRect, m_hWnd, 0, TRUE
		ret
OnSize endp


GetUserVerbs proc

local	pOleObject:LPOLEOBJECT
local	pEnumVerbs:LPENUMOLEVERB
local	hPopupMenu:HMENU
local	dwFetched:dword
local	oleverb:OLEVERB
local	szVerb[64]:BYTE
local	szText[128]:BYTE

		invoke GetMenu, m_hWnd
		invoke GetSubMenu, eax, 1
		invoke GetSubMenu, eax, 0
		mov hPopupMenu, eax
		.if (!eax)
			jmp done
		.endif

		invoke vf(m_pUnknown, IUnknown, QueryInterface), addr IID_IOleObject, addr pOleObject
		.if (eax != S_OK)
			jmp done
		.endif

		invoke vf(pOleObject, IOleObject, EnumVerbs), addr pEnumVerbs
		.if (eax == S_OK)
			mov m_dwUserVerbMax, OLEVERB_START
			.while (1)
				invoke vf(pEnumVerbs, IEnumOLEVERB, Next), 1, addr oleverb, addr dwFetched
				.break .if ((eax != S_OK) || (dwFetched == 0))
				.if ((oleverb.lVerb > 0) && (oleverb.lpszVerbName != NULL))
					invoke wsprintf, addr szText, CStr("%S (%d)"),oleverb.lpszVerbName,oleverb.lVerb
					mov ecx, OLEVERB_START
					add ecx,oleverb.lVerb
					.if (ecx > m_dwUserVerbMax)
						mov m_dwUserVerbMax, ecx
					.endif
					invoke AppendMenu, hPopupMenu, oleverb.fuFlags, ecx, addr szText
				.endif
			.endw
			invoke vf(pEnumVerbs, IEnumOLEVERB, Release)
		.endif
		invoke vf(pOleObject, IUnknown, Release)
done:
		ret
GetUserVerbs endp

OnEnterMenuLoop proc uses esi

local	hMenu:HMENU
local	pDispatch:LPDISPATCH
local	pPCI:LPPROVIDECLASSINFO
local	pPersistStorage:LPPERSISTSTORAGE
local	pPersistStreamInit:LPPERSISTSTREAMINIT
local	pPersistFile:LPPERSISTFILE
local	pPersistPropertyBag:LPPERSISTPROPERTYBAG

		invoke GetMenu, m_hWnd
		mov hMenu, eax

;-------------------------------------- options menu
if 0
		mov ecx, MF_UNCHECKED
		.if (g_bUserMode)
			mov cl, MF_CHECKED
		.endif
		invoke CheckMenuItem, hMenu, IDM_USERMODE, ecx

		mov ecx, MF_UNCHECKED
		.if (g_bUIDead)
			mov cl, MF_CHECKED
		.endif
		invoke CheckMenuItem, hMenu, IDM_UIDEAD, ecx
endif
		mov ecx, MF_UNCHECKED
		.if (m_bFilename)
			mov cl, MF_CHECKED
		.endif
		invoke CheckMenuItem, hMenu, IDM_FILENAME, ecx

		mov esi, MF_ENABLED or MF_BYCOMMAND
		invoke vf(m_pUnknown, IUnknown, QueryInterface), addr IID_IDispatch, addr pDispatch
		.if (eax == S_OK)
			invoke vf(pDispatch, IUnknown, Release)
		.else
			invoke vf(m_pUnknown, IUnknown, QueryInterface), addr IID_IProvideClassInfo, addr pPCI
			.if (eax == S_OK)
				invoke vf(pPCI, IUnknown, Release)
			.else
				mov esi, MF_GRAYED or MF_BYCOMMAND
			.endif
		.endif
		invoke EnableMenuItem, hMenu, IDM_TYPEINFODLG, esi
		invoke EnableMenuItem, hMenu, IDM_PROPERTIESDLG, esi

		mov esi, MF_ENABLED or MF_BYCOMMAND
		invoke vf(m_pObjectItem, IObjectItem, GetStorage)
		.if (!eax)
			mov esi, MF_GRAYED or MF_BYCOMMAND
		.endif
		invoke EnableMenuItem, hMenu, IDM_VIEWSTORAGEDLG, esi

;-------------------------------------- file menu
		mov esi, MF_GRAYED or MF_BYCOMMAND
		invoke vf(m_pUnknown, IUnknown, QueryInterface), addr IID_IPersistStreamInit, addr pPersistStreamInit
		.if (eax == S_OK && m_pContainer)
			mov esi, MF_ENABLED
			invoke vf(pPersistStreamInit, IUnknown, Release)
		.endif
		invoke EnableMenuItem, hMenu, IDM_SAVESTREAM, esi

		mov esi, MF_GRAYED or MF_BYCOMMAND
		invoke vf(m_pUnknown, IUnknown, QueryInterface), addr IID_IPersistStorage, addr pPersistStorage
		.if (eax == S_OK && m_pContainer)
			mov esi, MF_ENABLED
			invoke vf(pPersistStorage, IUnknown, Release)
		.endif
		invoke EnableMenuItem, hMenu, IDM_SAVESTORAGE, esi

		mov esi, MF_GRAYED or MF_BYCOMMAND
		invoke vf(m_pUnknown, IUnknown, QueryInterface), addr IID_IPersistFile, addr pPersistFile
		.if (eax == S_OK && m_pContainer)
			mov esi, MF_ENABLED
			invoke vf(pPersistFile, IUnknown, Release)
		.endif
		invoke EnableMenuItem, hMenu, IDM_SAVEFILE, esi
if ?LOADFILE
		invoke EnableMenuItem, hMenu, IDM_LOADFROMFILE, esi
endif

		mov esi, MF_GRAYED or MF_BYCOMMAND
		invoke vf(m_pUnknown, IUnknown, QueryInterface), addr IID_IPersistPropertyBag, addr pPersistPropertyBag
		.if (eax == S_OK && m_pContainer)
			mov esi, MF_ENABLED
			invoke vf(pPersistPropertyBag, IUnknown, Release)
		.endif
		invoke EnableMenuItem, hMenu, IDM_SAVEPROPBAG, esi
done:
		ret

OnEnterMenuLoop endp

OnCreate proc

local	hMenu:HMENU
local	pOleObject:LPOLEOBJECT
local	rect:RECT
local	clsid:CLSID
local	dwWidths[3]:DWORD

		invoke GetMenu, m_hWnd
		mov hMenu, eax

if 0
		mov eax, g_hIconCtrl
		.if (!eax)
			invoke LoadIcon,g_hInstance,IDI_CONTAINER
			mov g_hIconCtrl, eax
		.endif
		.if (eax)
			invoke SendMessage, m_hWnd, WM_SETICON, ICON_SMALL, g_hIconCtrl
			invoke SendMessage, m_hWnd, WM_SETICON, ICON_BIG, g_hIconCtrl
		.endif
endif

		.if (!g_szCaption)
			invoke GetWindowText, m_hWnd, addr g_szCaption, sizeof g_szCaption
		.endif

		invoke CreateWindowEx, 0, STATUSCLASSNAME,\
					NULL, WS_CHILD or WS_VISIBLE or SBARS_SIZEGRIP or CCS_BOTTOM,\
					0,0,0,0, m_hWnd, IDC_STATUSBAR,	g_hInstance, NULL
		mov m_hWndSB, eax
		mov dwWidths[0*type dwWidths],-1
		mov dwWidths[1*type dwWidths],-1
		mov dwWidths[2*type dwWidths],-1
		StatusBar_SetParts m_hWndSB, 3, addr dwWidths

		.if (g_rect.left)
			invoke SetWindowPos, m_hWnd, NULL, g_rect.left, g_rect.top,\
				g_rect.right, g_rect.bottom, SWP_NOZORDER or SWP_NOACTIVATE
		.else
			invoke CenterWindow, m_hWnd
		.endif

		invoke GetGUID@CObjectItem, m_pObjectItem, addr clsid
		invoke SetWindowIcon, m_hWnd, addr clsid
		mov m_hIcon, eax
		invoke vf(m_pObjectItem, IObjectItem, SetWindowText_), m_hWnd
		.if (m_bFilename)
			invoke vf(m_pObjectItem, IObjectItem, AddFilename), m_hWnd, FALSE
		.endif
			

		.if (m_bViewObject == FALSE)
if 0
			movzx eax, g_bUserMode
			.if (eax)
				invoke CheckMenuItem, hMenu, IDM_USERMODE, MF_CHECKED
			.else
				invoke CheckMenuItem, hMenu, IDM_USERMODE, MF_UNCHECKED 
			.endif
endif

			invoke Create@CContainer, m_pObjectItem, __this, NULL
			mov m_pContainer, eax

			.if (!eax)
				invoke PostMessage, m_hWnd, WM_CLOSE, 0, 0
			.endif

			invoke GetUserVerbs

			invoke vf(m_pUnknown, IUnknown, QueryInterface), addr IID_IOleObject, addr pOleObject
			.if (eax == S_OK)
				invoke vf(pOleObject, IUnknown, Release)
			.else
				jmp step1
			.endif


		.else
step1:
			invoke EnableMenuItem, hMenu, 1, MF_GRAYED or MF_BYPOSITION
		.endif
		ret
		align 4

OnCreate endp

CViewObjectWndProc proc uses __this thisarg, message:dword, wParam:WPARAM, lParam:LPARAM

local	wp:WINDOWPLACEMENT
local	rect:RECT
local	point:POINT

		mov __this,this@

;;		DebugOut "ViewObjectWndProc, msg=%X, wParam=%X", message, wParam

		mov eax,message
		.if (eax == WM_CREATE)

			invoke OnCreate
			xor eax, eax

		.elseif (eax == WM_COMMAND)

			invoke OnCommand, wParam, lParam

		.elseif (eax == WM_PAINT)

			invoke OnPaint, wParam
			xor eax, eax

		.elseif (eax == WM_MOVE)

			mov wp.length_, sizeof WINDOWPLACEMENT
			invoke GetWindowPlacement, m_hWnd, addr wp
			invoke SystemParametersInfo, SPI_GETWORKAREA,\
				NULL, addr rect, NULL
			push esi
			mov eax, wp.rcNormalPosition.left
			mov edx, wp.rcNormalPosition.top
			mov ecx, wp.rcNormalPosition.right
			mov esi, wp.rcNormalPosition.bottom
			.if ((ecx > rect.right) || (esi > rect.bottom))
				xor eax, eax
				xor edx, edx
			.endif
			pop esi
			add eax, 20
			add edx, 20
			mov g_rect.left, eax
			mov g_rect.top, edx

		.elseif (eax == WM_SIZE)

			.if (wParam != SIZE_MINIMIZED)
				movzx ecx, word ptr lParam+0	;new width 
				movzx edx, word ptr lParam+2	;new height
				invoke OnSize, ecx, edx
				xor eax, eax
				.if (wParam == SIZE_RESTORED)
					invoke GetWindowRect, m_hWnd, addr g_rect
					mov eax, g_rect.right
					sub eax, g_rect.left
					mov g_rect.right, eax
					mov eax, g_rect.bottom
					sub eax, g_rect.top
					mov g_rect.bottom, eax
					add g_rect.top, 20
					add g_rect.left, 20
				.endif
			.endif

		.elseif (eax == WM_CLOSE)

if 0
			mov wp.iLength, sizeof WINDOWPLACEMENT
			invoke GetWindowPlacement, m_hWnd, addr wp
			invoke CopyRect, addr g_rect, addr wp.rcNormalPosition
			mov eax, g_rect.right
			sub eax, g_rect.left
			mov g_rect.right, eax
			mov eax, g_rect.bottom
			sub eax, g_rect.top
			mov g_rect.bottom, eax
endif
			push m_hWnd
			invoke Destroy@CViewObjectDlg, __this
			pop ecx
			invoke DestroyWindow, ecx

			xor eax, eax

		.elseif (eax == WM_DESTROY)
;---------------------------- we shouldn't receive this message here
;---------------------------- if yes, someone has destroyed our window without WM_CLOSE
			invoke MessageBox, 0, CStr("unexpected WM_DESTROY received"), 0, MB_OK
			invoke Destroy@CViewObjectDlg, __this

if 0
		.elseif (eax == WM_ACTIVATEAPP)
			DebugOut "WM_ACTIVATEAPP"
		.elseif (eax == WM_ACTIVATE)
			DebugOut "WM_ACTIVATE %X", wParam
else
		.elseif (eax == WM_ACTIVATE)
endif

			movzx eax,word ptr wParam
			.if (eax == WA_INACTIVE)
				mov g_hWndDlg, NULL
				mov g_pViewObjectDlg, NULL
			.else
				mov eax, m_hWnd
				mov g_hWndDlg, eax
				mov g_pViewObjectDlg, __this
			.endif
			.if (m_pContainer)
				movzx eax,word ptr wParam
				.if (eax == WA_INACTIVE)
					mov eax, FALSE
				.else
					mov eax, TRUE
				.endif
				invoke OnActivate@CContainer, m_pContainer, eax
			.endif

    	.elseif (eax == WM_ENTERMENULOOP)

			invoke OnEnterMenuLoop
    		StatusBar_SetSimpleMode m_hWndSB, TRUE

    	.elseif (eax == WM_EXITMENULOOP)

		    StatusBar_SetSimpleMode m_hWndSB, FALSE

	    .elseif (eax == WM_MENUSELECT)

	    	movzx ecx, word ptr wParam+0
			invoke DisplayStatusBarString, m_hWndSB, ecx

		.elseif (eax == WM_MOUSEMOVE)

			.if (m_pContainer)        
            	invoke GetControlRect, addr rect
	        	movsx ecx, word ptr lParam+0
   		    	movsx edx, word ptr lParam+2
				mov point.x, ecx
				mov point.y, edx
				DebugOut "WM_MOUSEMOVE, X=%d, Y=%d", ecx, edx
				push ecx
				push edx
				invoke PtInRect, addr rect, point
				pop edx
				pop ecx
           	    .if (eax)
        		    invoke OnMouseMove@CContainer, m_pContainer, ecx, edx, wParam
                .endif
            .endif
        
		.elseif (eax == WM_LBUTTONDOWN)

			.if (m_pContainer)        
            	invoke GetControlRect, addr rect
	        	movsx ecx, word ptr lParam+0
   		    	movsx edx, word ptr lParam+2
				mov point.x, ecx
				mov point.y, edx
				push ecx
				push edx
				invoke PtInRect, addr rect, point
				pop edx
				pop ecx
           	    .if (eax)
        		    invoke OnMouseClick@CContainer, m_pContainer
                .endif
            .endif
        
		.elseif (eax == WM_SETCURSOR)

			push esi
            mov esi, FALSE
			.if (m_pContainer)        
            	invoke GetControlRect, addr rect
				invoke GetMessagePos
				movsx ecx, ax
				mov point.x, ecx
				shr eax,16
				movsx edx, ax
				mov point.y, edx
				invoke ScreenToClient, m_hWnd, addr point
				DebugOut "WM_SETCURSOR, X=%d, Y=%d", point.x, point.y
				invoke PtInRect, addr rect, point
				.if (eax)
                    movzx eax, word ptr lParam+2
					invoke OnSetCursor@CContainer, m_pContainer, point.x, point.y, eax
					.if (eax == S_OK)
						mov esi, TRUE
					.endif
                .endif
            .endif
            mov eax, esi
			pop esi        
			.if (eax == FALSE)
				invoke DefWindowProc, m_hWnd, message, wParam, lParam
			.endif
        
		.elseif (eax == WM_KEYDOWN)
if 1
			.if (wParam == VK_F6)
				.if (m_pContainer)
					invoke Create@CObjectItem, m_pContainer, NULL
					.if (eax)
						push eax
						invoke vf(eax, IObjectItem, ShowObjectDlg), m_hWnd
						pop eax
						invoke vf(eax, IObjectItem, Release)
					.endif
				.endif
			.endif
endif
		.elseif (eax == WM_WNDDESTROYED)
			
			mov eax, lParam
			mov edx, m_pTypeInfoDlg
			.if (edx && (eax == [edx].CDlg.hWnd))
				mov m_pTypeInfoDlg, NULL
			.endif

if ?HTMLHELP
		.elseif (eax == WM_HELP)

			invoke DoHtmlHelp, HH_DISPLAY_TOPIC, CStr("viewdialog.htm")
endif
		.else
			.if (m_pContainer)
				invoke IsWindowless@CContainer, m_pContainer
				.if (eax)
					push 0
					invoke OnMessage@CContainer, m_pContainer, message, wParam, lParam, esp
					pop ecx
					.if (eax == S_OK)
						mov eax, ecx
						ret
					.endif
				.endif
			.endif
			invoke DefWindowProc, m_hWnd, message, wParam, lParam
		.endif
		ret
		align 4

CViewObjectWndProc endp

wndproc proc hWnd:HWND, uMsg:DWORD, wParam:WPARAM, lParam:LPARAM
		
		mov eax, uMsg
		.if (eax == WM_CREATE)
			mov eax, lParam
			invoke SetWindowLong, hWnd, 0, [eax].CREATESTRUCT.lpCreateParams
			mov eax, lParam
			mov eax, [eax].CREATESTRUCT.lpCreateParams
			mov ecx, hWnd
			mov [eax].CViewObjectDlg.hWnd, ecx
		.else
			invoke GetWindowLong, hWnd, 0
		.endif
		.if (eax)
			invoke CViewObjectWndProc, eax, uMsg, wParam, lParam
		.else
			invoke DefWindowProc, hWnd, uMsg, wParam, lParam
		.endif
		ret

wndproc endp


SetStatusText@CViewObjectDlg proc public uses __this thisarg, iPart:DWORD, pszText:LPSTR
		mov __this, this@
		StatusBar_SetText m_hWndSB, iPart, pszText
		ret
SetStatusText@CViewObjectDlg endp

SetMenu@CViewObjectDlg proc public uses __this thisarg, iCmd:DWORD, dwFlags:DWORD
		mov __this, this@
		invoke GetMenu, m_hWnd
		invoke EnableMenuItem, eax, iCmd, dwFlags
		ret
SetMenu@CViewObjectDlg endp

TranslateAccelerator@CViewObjectDlg proc public uses __this thisarg, pMsg:ptr MSG
		mov __this, this@
		xor eax, eax
		.if (m_pContainer)
			invoke TranslateAccelerator@CContainer, m_pContainer, pMsg
		.endif
		ret
TranslateAccelerator@CViewObjectDlg endp

    end

