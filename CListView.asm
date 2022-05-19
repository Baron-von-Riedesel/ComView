
;*** implements a listview class with in place editing in report view
;*** IMPORTANT: keep this file independant from COMView

	.386
	.model flat,stdcall
	option casemap :none   ; case sensitive
	option proc:private

	.nolist
	.nocref
WIN32_LEAN_AND_MEAN	equ 1
	include windows.inc
	include commctrl.inc
	include windowsx.inc
	include statusbar.inc
	include macros.inc

	include CListView.inc
	include debugout.inc
	.list
	.cref

?EDITCHILD	equ 1		;1 = edit/combobox control is child of listview
?COMBOBOX	equ 1		;1 = support comboboxes
?EDITCTRLID	equ 1001	;ID of edit control in combo box

IDC_EDIT	equ 100


	.data

externdef g_oldListViewProc: WNDPROC
externdef g_hInstance:HINSTANCE

g_hWndEdit		HWND 0
;g_oldEditProc	LPVOID 0
g_oldEditProc	WNDPROC 0
if ?COMBOBOX
g_hWndCB		HWND 0
;;g_oldCBProc		LPVOID 0
g_oldCBProc		WNDPROC 0
g_fPushed		DWORD 0
endif
;g_oldListViewProc LPVOID 0
g_oldListViewProc WNDPROC 0
g_iItem			DWORD 0
g_iSubItem		DWORD 0
if ?EDITCHILD eq 0
g_hWndLV		HWND 0
endif
g_dwHeight		DWORD 0
	.code

;--- returns 0 if listview item updated

EndEditMode proc hWnd:HWND, fCancel:BOOL

local dwItem:DWORD
local rect:RECT
local rect2:RECT
local hWndParent:HWND
local dispinfo:NMLVDISPINFO
local szText[MAX_PATH]:byte

		DebugOut "EndEditMode"
		invoke GetParent, hWnd		;get parent of listview for LVN_xxx notification
		mov hWndParent, eax

		.if (fCancel)
			mov dispinfo.item.pszText, NULL
		.else
if ?COMBOBOX
			.if (g_hWndCB)
				invoke GetWindowText, g_hWndCB, addr szText, sizeof szText
			.else
endif
				invoke GetWindowText, g_hWndEdit, addr szText, sizeof szText
if ?COMBOBOX
			.endif
endif
			lea eax, szText
			mov dispinfo.item.pszText, eax
			mov dispinfo.item.cchTextMax, sizeof szText
		.endif

		mov dispinfo.item.mask_, LVIF_TEXT
		mov dispinfo.item.lParam, CB_ERR
if ?COMBOBOX		
		.if (g_hWndCB)
			invoke ComboBox_FindStringExact( g_hWndCB, -1, dispinfo.item.pszText)
			.if (eax != CB_ERR)
				invoke ComboBox_GetItemData( g_hWndCB, eax)
				mov dispinfo.item.lParam, eax
			.endif
;--------------------------- first clear global var to avoid reentrance
			xor eax, eax
			xchg eax, g_hWndCB
			invoke DestroyWindow, eax
		.endif
endif
		.if (g_hWndEdit)
;--------------------------- first clear global var to avoid reentrance
			xor eax, eax
			xchg eax, g_hWndEdit
			invoke DestroyWindow, eax
		.endif
		mov eax, hWnd
		mov dispinfo.hdr.hwndFrom, eax
		invoke GetWindowLong, hWnd, GWL_ID
		mov dispinfo.hdr.idFrom, eax
		mov dispinfo.hdr.code, LVN_ENDLABELEDIT
		mov eax, g_iItem
		mov dispinfo.item.iItem, eax
		mov eax, g_iSubItem
		mov dispinfo.item.iSubItem, eax
		invoke SendMessage, hWndParent, WM_NOTIFY, hWnd, addr dispinfo
		DebugOut "EndEditMode: WM_NOTIFY, LVN_ENDLABELEDIT returned %X", eax
		.if (eax && dispinfo.item.pszText)
			mov dispinfo.item.mask_, LVIF_TEXT
			invoke ListView_SetItem( hWnd, addr dispinfo.item)
			mov eax, 1
		.endif
		xor al, 1
		ret
EndEditMode endp

;--- edit control wndproc (subclassed)

myEditProc proc hWnd:HWND, uMsg:DWORD, wParam:WPARAM, lParam:LPARAM

local	hWndLV:HWND

;;	DebugOut "EditWndproc, msg=%X", uMsg

	mov eax, uMsg
	.if (eax == WM_GETDLGCODE)

		mov eax,DLGC_WANTALLKEYS or DLGC_HASSETSEL

	.elseif (eax == WM_DESTROY)

		mov g_hWndEdit, NULL

	.elseif (eax == WM_KEYDOWN)

		mov eax, wParam
		.if ((eax == VK_RETURN) || (eax == VK_TAB) || (eax == VK_ESCAPE))
if ?EDITCHILD
if ?COMBOBOX
			.if (g_hWndCB)
				invoke GetParent, g_hWndCB	;parent of COMBOBOX -> listview
			.else
				invoke GetParent, hWnd		;parent of EDIT -> listview
			.endif
else
			invoke GetParent, hWnd		;parent of EDIT -> listview
endif
else
			mov eax, g_hWndLV
endif
			mov hWndLV, eax

			.if (wParam == VK_ESCAPE)
				invoke EndEditMode, hWndLV, TRUE
			.else
				invoke EndEditMode, hWndLV, FALSE
			.endif
			invoke GetFocus
			.if (!eax)
				invoke SetFocus, hWndLV
			.endif
		.else
			invoke CallWindowProc, g_oldEditProc, hWnd, uMsg, wParam, lParam
		.endif

	.elseif (eax == WM_KILLFOCUS)

		invoke CallWindowProc, g_oldEditProc, hWnd, uMsg, wParam, lParam
;---------------------- terminating edit mode during CBN_KILLFOCUS doesnt work
;---------------------- for win9x. so do it here
		.if (g_hWndCB)
			push eax
			invoke GetParent, g_hWndCB
			invoke EndEditMode, eax, FALSE
			pop eax
		.endif

	.else
		invoke CallWindowProc, g_oldEditProc, hWnd, uMsg, wParam, lParam
	.endif
	ret
myEditProc endp


;--- called from message LVM_EDITLABEL


StartEditMode proc hWnd:HWND, iItem:DWORD, iSubItem:DWORD

local dwItem:DWORD
local rect:RECT
local rect2:RECT
local hWndParent:HWND
local dispinfo:NMLVDISPINFO
local szText[MAX_PATH]:byte


		DebugOut "StartEditMode"
		.if (g_hWndCB || g_hWndEdit)
			ret
		.endif
		lea edx, rect
		invoke ListView_GetSubItemRect( hWnd, iItem, iSubItem, LVIR_BOUNDS, edx)
		inc rect.left
		inc rect.left
;		inc rect.top
		dec rect.bottom
		dec rect.right

		mov eax, rect.right
		sub eax, rect.left
		mov rect.right, eax

		mov eax, rect.bottom
		sub eax, rect.top
		mov rect.bottom, eax

		invoke GetParent, hWnd
		mov hWndParent, eax

if ?EDITCHILD
		invoke CreateWindowEx, 0, CStr("edit"), CStr(""),\
			WS_CHILD or ES_LEFT or ES_AUTOHSCROLL, rect.left, rect.top,\
			rect.right, rect.bottom,\
			hWnd, IDC_EDIT, g_hInstance, NULL
else
		invoke GetClientRect, hWnd, addr rect2
		invoke ClientToScreen, hWnd, addr rect2
		invoke ScreenToClient, hWndParent, addr rect2

		mov eax, rect2.left
		add rect.left, eax
		mov eax, rect2.top
		add rect.top, eax
		mov eax, hWnd
		mov g_hWndLV, eax
		invoke CreateWindowEx, 0, CStr("edit"), CStr(""),\
			WS_CHILD or ES_LEFT or ES_AUTOHSCROLL, rect.left, rect.top,\
			rect.right, rect.bottom,\
			hWndParent, IDC_EDIT, g_hInstance, NULL
endif

		.if (eax)
			mov g_hWndEdit, eax
			invoke SetWindowLong, g_hWndEdit, GWL_WNDPROC, myEditProc
			mov g_oldEditProc, eax

			mov dispinfo.item.mask_, LVIF_TEXT
			mov eax, iItem
			mov dispinfo.item.iItem, eax
			mov g_iItem, eax
			mov eax, iSubItem
			mov dispinfo.item.iSubItem, eax
			mov g_iSubItem, eax
			lea eax, szText
			mov dispinfo.item.pszText, eax
			mov dispinfo.item.cchTextMax, sizeof szText
			invoke ListView_GetItem( hWnd, addr dispinfo.item)
			invoke SetWindowText, g_hWndEdit, addr szText
			mov eax, hWnd
			mov dispinfo.hdr.hwndFrom, eax
			invoke GetWindowLong, hWnd, GWL_ID
			mov dispinfo.hdr.idFrom, eax
			mov dispinfo.hdr.code, LVN_BEGINLABELEDIT
			invoke SendMessage, hWndParent, WM_NOTIFY, hWnd, addr dispinfo
			.if (eax)
				invoke SendMessage, hWnd, WM_CANCELMODE, 0, 0
			.else
				invoke SendMessage, hWnd, WM_GETFONT, 0, 0
				invoke SendMessage, g_hWndEdit, WM_SETFONT, eax, 0
;				invoke SendMessage, g_hWndEdit, EM_SETMARGINS,\
;					EC_LEFTMARGIN or EC_RIGHTMARGIN or EC_USEFONTINFO, 0
				invoke SendMessage, g_hWndEdit, EM_SETMARGINS,\
					EC_LEFTMARGIN or EC_RIGHTMARGIN, 00040004h
				invoke SetWindowPos, g_hWndEdit, HWND_TOP, 0,0,0,0,\
					SWP_SHOWWINDOW or SWP_NOMOVE or SWP_NOSIZE
				invoke SetFocus, g_hWndEdit
			.endif
		.endif
		ret
StartEditMode endp


if ?COMBOBOX

DrawDownArrowBtn proc hWnd:HWND

LOCAL rc :RECT
LOCAL rcBtn :RECT
LOCAL rcEdit :RECT
LOCAL hdc :DWORD
local hwndEdit:HWND

		invoke GetDC, hWnd
		mov hdc, eax
		invoke GetClientRect,hWnd, addr rc
		invoke GetDlgItem, hWnd, ?EDITCTRLID
		mov hwndEdit, eax
		invoke GetClientRect,hwndEdit,addr rcEdit
		invoke SetRect,addr rcBtn,rcEdit.right,0,rc.right,rc.bottom
		.if (g_fPushed == TRUE)
			invoke DrawFrameControl,hdc,addr rcBtn,DFC_SCROLL,DFCS_SCROLLDOWN or DFCS_PUSHED
		.else
			invoke DrawFrameControl,hdc,addr rcBtn,DFC_SCROLL,DFCS_SCROLLDOWN
		.endif 
		invoke ReleaseDC, hWnd, hdc
		ret
DrawDownArrowBtn endp

;--- combobox control wndproc (subclassed)

myCBProc proc hWnd:HWND,uMsg:UINT,wParam:WPARAM,lParam:LPARAM

LOCAL rc :RECT
local hwndEdit:HWND

;;	DebugOut "ComboboxWndproc, msg=%X", uMsg
	mov eax, uMsg
	.if (eax == WM_PAINT)

		invoke CallWindowProc,g_oldCBProc,hWnd,uMsg,wParam,lParam
		invoke DrawDownArrowBtn, hWnd

	.elseif (eax == WM_DESTROY)

		mov g_hWndCB, NULL

	.elseif (eax == WM_LBUTTONDOWN)

		invoke CallWindowProc, g_oldCBProc, hWnd, uMsg, wParam, lParam
		mov g_fPushed, TRUE
		invoke DrawDownArrowBtn, hWnd

	.elseif (eax == WM_LBUTTONUP)

		invoke CallWindowProc, g_oldCBProc, hWnd, uMsg, wParam, lParam
		mov g_fPushed, FALSE
		invoke DrawDownArrowBtn, hWnd

	.elseif (eax == WM_SIZE)

		invoke CallWindowProc, g_oldCBProc, hWnd, uMsg, wParam, lParam
		invoke GetDlgItem, hWnd, ?EDITCTRLID
		mov hwndEdit, eax
		invoke GetWindowRect,hwndEdit,addr rc
		invoke ScreenToClient, hWnd, addr rc.left
		invoke ScreenToClient, hWnd, addr rc.right
		mov ecx, rc.bottom
		add ecx, rc.top
		invoke SetWindowPos, hwndEdit, NULL, 0, 0, rc.right, ecx, SWP_NOZORDER
		invoke SendMessage, hwndEdit, EM_SETMARGINS,\
					EC_LEFTMARGIN or EC_RIGHTMARGIN, 00020002h

	.else
		invoke CallWindowProc, g_oldCBProc, hWnd, uMsg, wParam, lParam
	.endif
	ret

myCBProc endp


;--- called from message LVM_COMBOBOXMODE, start edit mode


StartComboBoxMode proc hWnd:HWND, iItem:DWORD, iSubItem:DWORD

local dwItem:DWORD
local rect:RECT
local rect2:RECT
local hWndParent:HWND
local hWndEdit:HWND
local dispinfo:NMLVDISPINFO
local szText[MAX_PATH]:byte


		DebugOut "StartComboBoxMode"
		.if (g_hWndCB || g_hWndEdit)
			ret
		.endif
		lea edx, rect
		invoke ListView_GetSubItemRect( hWnd, iItem, iSubItem, LVIR_BOUNDS, edx)
		inc rect.left
		inc rect.left
		inc rect.left
;		inc rect.top
;		dec rect.bottom
;		dec rect.right

		mov eax, rect.right
		sub eax, rect.left
		mov rect.right, eax

		mov eax, rect.bottom
		sub eax, rect.top
		mov rect.bottom, eax

		invoke GetParent, hWnd
		mov hWndParent, eax

if ?EDITCHILD
		mov ecx, rect.bottom
		shl ecx, 3
		invoke CreateWindowEx, 0, CStr("ComboBox"), CStr(""),\
			WS_CHILD or WS_VSCROLL or CBS_DROPDOWN or CBS_AUTOHSCROLL, rect.left, rect.top,\
			rect.right, ecx,\
			hWnd, IDC_EDIT, g_hInstance, NULL
else
		invoke GetClientRect, hWnd, addr rect2
		invoke ClientToScreen, hWnd, addr rect2
		invoke ScreenToClient, hWndParent, addr rect2

		mov eax, rect2.left
		add rect.left, eax
		mov eax, rect2.top
		add rect.top, eax
		mov eax, hWnd
		mov g_hWndLV, eax
		invoke CreateWindowEx, 0, CStr("ComboBox"), CStr(""),\
			WS_CHILD or WS_VSCROLL or CBS_DROPDOWN or CBS_AUTOHSCROLL, rect.left, rect.top,\
			rect.right, rect.bottom,\
			hWndParent, IDC_EDIT, g_hInstance, NULL
endif

		.if (eax)
			mov g_hWndCB, eax
			invoke SetWindowLong, g_hWndCB, GWL_WNDPROC, myCBProc
			mov g_oldCBProc, eax
if 1
			invoke GetDlgItem, g_hWndCB, ?EDITCTRLID
			mov hWndEdit, eax
			.if (eax)
				invoke SetWindowLong, hWndEdit, GWL_WNDPROC, myEditProc
				mov g_oldEditProc, eax
			.endif
endif
			invoke SendMessage, hWnd, WM_GETFONT, 0, 0
			invoke SendMessage, g_hWndCB, WM_SETFONT, eax, 0
			mov ecx, rect.bottom
			sub ecx, 7
			mov g_dwHeight, ecx
			invoke ComboBox_SetItemHeight( g_hWndCB, -1, ecx)

			mov dispinfo.item.mask_, LVIF_TEXT
			mov eax, iItem
			mov dispinfo.item.iItem, eax
			mov g_iItem, eax
			mov eax, iSubItem
			mov dispinfo.item.iSubItem, eax
			mov g_iSubItem, eax
			lea eax, szText
			mov dispinfo.item.pszText, eax
			mov dispinfo.item.cchTextMax, sizeof szText
			invoke ListView_GetItem( hWnd, addr dispinfo.item)
			invoke SetWindowText, g_hWndCB, addr szText
			mov eax, hWnd
			mov dispinfo.hdr.hwndFrom, eax
			invoke GetWindowLong, hWnd, GWL_ID
			mov dispinfo.hdr.idFrom, eax
			mov dispinfo.hdr.code, LVN_BEGINLABELEDIT
			invoke SendMessage, hWndParent, WM_NOTIFY, hWnd, addr dispinfo
			.if (eax)
				invoke SendMessage, hWnd, WM_CANCELMODE, 0, 0
			.else
				invoke SetWindowPos, g_hWndCB, HWND_TOP, 0,0,0,0,\
					SWP_SHOWWINDOW or SWP_NOMOVE or SWP_NOSIZE
				invoke SetFocus, g_hWndCB
			.endif
		.endif
		ret
StartComboBoxMode endp

endif

;--- listview wndproc (subclassed)

myListViewProc proc hWnd:HWND, uMsg:DWORD, wParam:WPARAM, lParam:LPARAM

;;	DebugOut "listviewproc, msg=%X, wParam=%X, lParam=%X", uMsg, wParam, lParam

	mov eax, uMsg
	.if (eax == LVM_EDITLABEL)
		invoke StartEditMode, hWnd, wParam, lParam
		jmp exit
	.elseif (eax == LVM_GETEDITCONTROL)
		.if (g_hWndCB)
			invoke GetDlgItem, g_hWndCB, ?EDITCTRLID
		.else
			mov eax, g_hWndEdit
		.endif
		jmp exit
if ?COMBOBOX
	.elseif (eax == LVM_COMBOBOXMODE)
		invoke StartComboBoxMode, hWnd, wParam, lParam
		jmp exit
	.elseif (eax == LVM_GETCOMBOBOXCONTROL)
		mov eax, g_hWndCB
		jmp exit
	.elseif (g_hWndEdit || g_hWndCB)
else
	.elseif (g_hWndEdit)
endif
		.if (eax == WM_CANCELMODE)
			invoke EndEditMode, hWnd, TRUE
		.elseif ((eax == WM_VSCROLL) || (eax == WM_HSCROLL) || (eax == WM_MOUSEWHEEL) || (eax == LVM_ENDEDITMODE))
			invoke EndEditMode, hWnd, FALSE
			jmp exit
		.elseif (eax == WM_COMMAND)
			DebugOut "listviewproc, WM_COMMAND in edit mode, wParam=%X, lParam=%X", wParam, lParam
			movzx eax, word ptr wParam+0
			movzx edx, word ptr wParam+2
			.if (eax == IDC_EDIT)
				.if (g_hWndEdit && (edx == EN_KILLFOCUS))
					invoke EndEditMode, hWnd, FALSE
				.endif
			.endif
		.elseif (eax == WM_NOTIFY)
			mov eax, lParam
			DebugOut "listviewproc, WM_NOTIFY in edit mode, wParam=%X, lParam=%X, code=%d, idFrom=%X, hwndFrom=%X",\
				wParam, eax, [eax].NMHDR.code, [eax].NMHDR.idFrom, [eax].NMHDR.hwndFrom
			mov ecx, [eax].NMHDR.code
;----------------------------- dont care about listview character mode
			.if ((ecx == HDN_ITEMCLICK) || (ecx == HDN_ITEMCLICKW))
				invoke EndEditMode, hWnd, FALSE
				jmp exit
			.endif
		.endif
	.endif
	invoke CallWindowProc, g_oldListViewProc, hWnd, uMsg, wParam, lParam
exit:
;;	DebugOut "listviewproc end, msg=%X, wParam=%X, lParam=%X", uMsg, wParam, lParam
	ret
myListViewProc endp

;--- this is the API, called by clients

CreateEditListView proc public hWndLV:HWND

	invoke SetWindowLong, hWndLV, GWL_WNDPROC, myListViewProc
	mov g_oldListViewProc, eax
	ret

CreateEditListView endp

	end
