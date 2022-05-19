
;*** definition of CAmbientPropDlg methods

	.386
	.model flat,stdcall
	option casemap :none   ; case sensitive
	option proc:private

	include COMView.inc
INSIDE_CAMBIENTPROPDLG equ 1
	include classes.inc
	include rsrc.inc
	include debugout.inc
	include CListView.inc

?MODELESS	equ 1

BEGIN_CLASS CAmbientPropDlg, CDlg
hWndLV	HWND ?
END_CLASS

	MEMBER hWnd, pDlgProc, hWndLV

PropItem	struct
id		DWORD ?
pValue	PVOID ?
wType	WORD ?
wFlags	BYTE ?
PropItem	ends

PROPITEM_FACTIVE	equ 1
PROPITEM_FHEX		equ 2

	.const

ColumnsListView label CColHdr
		CColHdr <CStr("Name")	, 40>
		CColHdr <CStr("ID")		, 10>
		CColHdr <CStr("Type")	, 20>
		CColHdr <CStr("Value")	, 30>
NUMCOLS_PROPERTIES textequ %($ - ColumnsListView) / sizeof CColHdr

COLUMN_VALUE	equ 3

	.data

g_pAmbientPropDlg	LPVOID	NULL
g_dwBackColor		DWORD	0FFFFFFh
g_pFont				LPDISPATCH	NULL
g_bstrDisplayName	BSTR	NULL
g_dwForeColor		DWORD	0h
g_bstrScaleUnits	BSTR	NULL
g_wTextAlign		DWORD	0
g_wAppearance		DWORD	1
g_LogFont			LOGFONT <-13,0,0,0,400,FALSE,FALSE,FALSE,\
						DEFAULT_CHARSET,OUT_DEFAULT_PRECIS,CLIP_DEFAULT_PRECIS,\
						DEFAULT_QUALITY,DEFAULT_PITCH, "MS Sans Serif">
g_bMessageReflect	BOOLEAN FALSE
g_bAutoClip			BOOLEAN FALSE
g_bShowGrabHandles	BOOLEAN TRUE
g_bShowHatching		BOOLEAN TRUE
g_bDisplayAsDefault	BOOLEAN FALSE
g_bSupportsMnemonics	BOOLEAN FALSE
g_dwCodePage		DWORD 0
g_hPalette			DWORD 0
g_bstrCharSet		BSTR NULL
g_dwTransferPriority	DWORD 0
g_bRightToLeft		BOOLEAN FALSE
g_bTopToBottom		BOOLEAN TRUE

	align 4

PropTab	label PropItem
	PropItem {DISPID_AMBIENT_BACKCOLOR,			offset g_dwBackColor,		VT_I4,	PROPITEM_FHEX}
	PropItem {DISPID_AMBIENT_DISPLAYNAME,		offset g_bstrDisplayName,	VT_BSTR}
FontItem label PropItem
	PropItem {DISPID_AMBIENT_FONT,				offset g_pFont,				VT_DISPATCH}
	PropItem {DISPID_AMBIENT_FORECOLOR,			offset g_dwForeColor,		VT_I4,	PROPITEM_FHEX}
	PropItem {DISPID_AMBIENT_LOCALEID,			offset g_LCID,				VT_I4,	PROPITEM_FACTIVE or PROPITEM_FHEX}
	PropItem {DISPID_AMBIENT_MESSAGEREFLECT,	offset g_bMessageReflect,	VT_BOOL}
	PropItem {DISPID_AMBIENT_SCALEUNITS,		offset g_bstrScaleUnits,	VT_BSTR}
	PropItem {DISPID_AMBIENT_TEXTALIGN,			offset g_wTextAlign,		VT_I2}
	PropItem {DISPID_AMBIENT_USERMODE,			offset g_bUserMode,			VT_BOOL,PROPITEM_FACTIVE}
	PropItem {DISPID_AMBIENT_UIDEAD,			offset g_bUIDead,			VT_BOOL,PROPITEM_FACTIVE}
	PropItem {DISPID_AMBIENT_SHOWGRABHANDLES,	offset g_bShowGrabHandles,	VT_BOOL}
	PropItem {DISPID_AMBIENT_SHOWHATCHING,		offset g_bShowHatching,		VT_BOOL}
	PropItem {DISPID_AMBIENT_DISPLAYASDEFAULT,	offset g_bDisplayAsDefault,	VT_BOOL}
	PropItem {DISPID_AMBIENT_SUPPORTSMNEMONICS,	offset g_bSupportsMnemonics,VT_BOOL}
	PropItem {DISPID_AMBIENT_AUTOCLIP,			offset g_bAutoClip,			VT_BOOL}
	PropItem {DISPID_AMBIENT_APPEARANCE,		offset g_wAppearance,		VT_I2}
	PropItem {DISPID_AMBIENT_CODEPAGE,			offset g_dwCodePage,		VT_I4}
	PropItem {DISPID_AMBIENT_PALETTE,			offset g_hPalette,			VT_I4,	PROPITEM_FHEX}
	PropItem {DISPID_AMBIENT_CHARSET,			offset g_bstrCharSet,		VT_BSTR}
	PropItem {DISPID_AMBIENT_TRANSFERPRIORITY,	offset g_dwTransferPriority,VT_I4}
	PropItem {DISPID_AMBIENT_RIGHTTOLEFT,		offset g_bRightToLeft,		VT_BOOL}
	PropItem {DISPID_AMBIENT_TOPTOBOTTOM,		offset g_bTopToBottom,		VT_BOOL}
NUMPROPS	equ ($ - offset PropTab) / sizeof PropItem

	.code

__this	textequ <ebx>
_this	textequ <[__this].CAmbientPropDlg>
thisarg	textequ <this@:ptr CAmbientPropDlg>

SetItem proto pPropItem:ptr PropItem, iItem:DWORD, bNew:BOOL

OLECREATEFONTINDIRECT typedef proto :ptr FONDDESC, :ptr IID, :ptr LPFONT

CreateFontFromLogFont proc iPointSize:DWORD

local lpfnOleCreateFontIndirect:ptr OLECREATEFONTINDIRECT
local pFont:LPFONT
local wlfFaceName[LF_FACESIZE]:word
local fontdesc:FONTDESC

		invoke GetModuleHandle, CStr("OLEAUT32")
		invoke GetProcAddress, eax, CStr("OleCreateFontIndirect")
		.if (!eax)
			jmp exit
		.endif
		mov lpfnOleCreateFontIndirect, eax
		mov fontdesc.cbSizeofstruct, sizeof FONTDESC
		invoke MultiByteToWideChar,CP_ACP,MB_PRECOMPOSED,
				addr g_LogFont.lfFaceName, LF_FACESIZE, addr wlfFaceName, LF_FACESIZE
		lea eax, wlfFaceName
		mov fontdesc.lpstrName, eax
		mov eax, iPointSize
		mov ecx, 1000
		mul ecx
		mov fontdesc.cySize.Lo, eax
		mov fontdesc.cySize.Hi, edx
		mov eax, g_LogFont.lfWeight
		mov fontdesc.sWeight, ax
		movzx ax, g_LogFont.lfCharSet
		mov fontdesc.sCharset, ax
		movzx eax, g_LogFont.lfItalic
		mov fontdesc.fItalic, eax
		movzx eax, g_LogFont.lfUnderline
		mov fontdesc.fUnderline, eax
		movzx eax, g_LogFont.lfStrikeOut
		mov fontdesc.fStrikethrough, eax
		invoke lpfnOleCreateFontIndirect, addr fontdesc, addr IID_IFont, addr pFont
		.if (eax == S_OK)
			.if (g_pFont)
				invoke vf(g_pFont, IUnknown, Release)
				mov g_pFont, NULL
			.endif
			mov eax, pFont
			mov g_pFont, eax
		.else
			invoke OutputMessage, m_hWnd, eax, CStr("OleCreateFontIndirect"), 0
		.endif
		mov eax, 1
exit:
		ret
		align 4

CreateFontFromLogFont endp

_SelectFont proc

local cf:CHOOSEFONT

		mov cf.lStructSize, sizeof CHOOSEFONT
		mov eax, m_hWnd
		mov cf.hwndOwner, eax
		mov cf.lpLogFont, offset g_LogFont
		mov cf.Flags, CF_SCREENFONTS or CF_INITTOLOGFONTSTRUCT or CF_EFFECTS
		invoke ChooseFont, addr cf
		.if (eax)
			invoke CreateFontFromLogFont, cf.iPointSize
		.endif
		ret
		align 4
_SelectFont endp


OnBeginLabelEdit proc  uses esi edi pNMLVDI:ptr NMLVDISPINFO

local dwRet:DWORD
local dwRC:DWORD
local hWndCB:HWND
local lvi:LVITEM

		mov dwRC, FALSE
		mov edi, pNMLVDI
		
		mov eax, [edi].NMLVDISPINFO.item.iItem
		mov lvi.iItem, eax
		@mov lvi.iSubItem, 0
		@mov lvi.mask_, LVIF_PARAM
		invoke ListView_GetItem( m_hWndLV, addr lvi)
		mov esi, lvi.lParam
		invoke SendMessage, m_hWndLV, LVM_GETCOMBOBOXCONTROL, 0, 0
		mov hWndCB, eax
		.if ([esi].PropItem.wType == VT_BOOL)
			invoke ComboBox_AddString( hWndCB, CStr("False"))
			invoke ComboBox_SetItemData( hWndCB, eax, FALSE)
			invoke ComboBox_AddString( hWndCB, CStr("True"))
			invoke ComboBox_SetItemData( hWndCB, eax, TRUE)
		.elseif ([esi].PropItem.id == DISPID_AMBIENT_PALETTE)
			invoke MessageBox, m_hWnd, CStr("This property cannot be edited"), 0, MB_OK
			invoke SetWindowLong, m_hWnd, DWL_MSGRESULT, 1
			mov dwRC, TRUE
		.elseif ([esi].PropItem.wType == VT_DISPATCH)
			.if ([esi].PropItem.id == DISPID_AMBIENT_FONT)
				invoke _SelectFont
				.if (eax)
					invoke SetItem, esi, lvi.iItem, FALSE
				.endif
			.endif
			invoke SetWindowLong, m_hWnd, DWL_MSGRESULT, 1
			mov dwRC, TRUE
		.endif
		return dwRC
		align 4

OnBeginLabelEdit endp

NotifyContainerWindows proc uses edi
		xor edi, edi
		.while (g_aViewControlClass)
			movzx eax, g_aViewControlClass
			invoke FindWindowEx, NULL, edi, eax, NULL
			.break .if (!eax)
			mov edi, eax
			invoke PostMessage, edi, WM_COMMAND, IDM_PROPCHANGED, [esi].PropItem.id
		.endw
		ret
		align 4
NotifyContainerWindows endp

OnEndLabelEdit proc  uses esi edi pNMLVDI:ptr NMLVDISPINFO

local dwValue:DWORD
local dwRC:DWORD
local dwMode:DWORD
local bError:BOOLEAN
local bChanged:BOOLEAN
local lvi:LVITEM

		mov dwRC, FALSE
		mov edi, pNMLVDI
		
		.if (![edi].NMLVDISPINFO.item.pszText)
			jmp done
		.endif

		mov eax, [edi].NMLVDISPINFO.item.iItem
		mov lvi.iItem, eax
		@mov lvi.iSubItem, 0
		@mov lvi.mask_, LVIF_PARAM
		invoke ListView_GetItem( m_hWndLV, addr lvi)

		mov dwMode, LVM_EDITLABEL
		mov bError, FALSE
		mov bChanged, FALSE
		mov esi, lvi.lParam
		.if ([esi].PropItem.wType == VT_BOOL)
			mov eax, [edi].NMLVDISPINFO.item.lParam
			.if (eax != CB_ERR)
				mov ecx, [esi].PropItem.pValue
				.if (al != byte ptr [ecx])
					mov bChanged, TRUE
					mov [ecx], al
				.endif
			.else
				mov bError, TRUE
				mov dwMode, LVM_COMBOBOXMODE
			.endif
		.elseif ([esi].PropItem.wType == VT_BSTR)
			mov ecx, [esi].PropItem.pValue
			invoke SysFreeString, dword ptr [ecx]
			invoke SysStringFromLPSTR, [edi].NMLVDISPINFO.item.pszText, 0
			mov ecx, [esi].PropItem.pValue
			mov [ecx], eax
		.else
			invoke String2DWord, [edi].NMLVDISPINFO.item.pszText, addr dwValue
			.if (eax)
				mov eax, dwValue
				.if ((eax & 0FFFF0000h) && ([esi].PropItem.wType == VT_I2))
					mov bError, TRUE
				.else
					mov ecx, [esi].PropItem.pValue
					.if (eax != dword ptr [ecx])
						mov [ecx], eax
						mov bChanged, TRUE
					.endif
				.endif
			.else
				mov bError, TRUE
			.endif
		.endif
		.if (bError)
			invoke MessageBeep, MB_OK
			invoke SendMessage, m_hWndLV, dwMode, [edi].NMLVDISPINFO.item.iItem, [edi].NMLVDISPINFO.item.iSubItem
			invoke SendMessage, m_hWndLV, LVM_GETEDITCONTROL, 0, 0
	 		invoke SetWindowText, eax, [edi].NMLVDISPINFO.item.pszText
		.else
			.if (bChanged)
				invoke NotifyContainerWindows
			.endif
			invoke SetWindowLong, m_hWnd, DWL_MSGRESULT, 1
			mov dwRC, TRUE
		.endif
done:
		return dwRC
		align 4

OnEndLabelEdit endp


StartEditMode proc dwItem:DWORD

local lvi:LVITEM

		invoke ListView_GetItemCount( m_hWndLV)
		.if (eax > dwItem)
			mov eax, dwItem
			mov lvi.iItem, eax
			@mov lvi.iSubItem, 0
			@mov lvi.mask_, LVIF_PARAM
			invoke ListView_GetItem( m_hWndLV, addr lvi)
			mov ecx, lvi.lParam
			.if ([ecx].PropItem.wType == VT_BOOL)
				invoke SendMessage, m_hWndLV, LVM_COMBOBOXMODE, dwItem, COLUMN_VALUE
			.else
				invoke SendMessage, m_hWndLV, LVM_EDITLABEL, dwItem, COLUMN_VALUE
			.endif
		.endif
		ret
		align 4

StartEditMode endp

ShowContextMenu proc uses esi bMouse:BOOL

local hMenu:HMENU
local dwCnt:DWORD
local dwDefault:DWORD
local pFont:LPUNKNOWN
local hti:LVHITTESTINFO
local pt:POINT
local bTmp:BOOLEAN
local lvi:LVITEM

	invoke GetItemPosition, m_hWndLV, bMouse, addr pt
	mov eax, pt.x
	mov ecx, pt.y
	mov hti.pt.x, eax
	mov hti.pt.y, ecx
	invoke ScreenToClient, m_hWndLV, addr hti.pt
	invoke ListView_HitTest( m_hWndLV, addr hti)
	.if (eax == -1)
		jmp done
	.endif

	mov lvi.iItem, eax
	@mov lvi.iSubItem, 0
	@mov lvi.mask_, LVIF_PARAM
	invoke ListView_GetItem( m_hWndLV, addr lvi)
	mov esi, lvi.lParam

	invoke ListView_GetSelectedCount( m_hWndLV)
	mov dwCnt, eax

	invoke CreatePopupMenu
	mov hMenu, eax
	mov edx, MF_STRING
	mov ecx, IDM_EDIT
	.if (dwCnt != 1)
		or edx, MF_GRAYED
		mov ecx, IDM_ENABLE
	.endif
	mov dwDefault, ecx
	invoke AppendMenu, hMenu, edx, IDM_EDIT, CStr("&Edit")
	mov edx, MF_STRING
	.if (!([esi].PropItem.wFlags & PROPITEM_FACTIVE))
		or edx, MF_CHECKED 
	.endif
	invoke AppendMenu, hMenu, edx, IDM_ENABLE, CStr("&Disabled")

	.if (([esi].PropItem.wType == VT_DISPATCH) && (dwCnt == 1))
		mov eax, [esi].PropItem.pValue
		mov eax, [eax]
		.if (eax)
			invoke AppendMenu, hMenu, MF_STRING, IDM_PROPERTIES, CStr("&Properties")
		.endif
	.endif

	invoke SetMenuDefaultItem, hMenu, dwDefault, FALSE

	invoke TrackPopupMenu, hMenu, TPM_LEFTALIGN or TPM_LEFTBUTTON or TPM_RETURNCMD,
			pt.x, pt.y, 0, m_hWnd, NULL
	push eax
	invoke DestroyMenu, hMenu
	pop eax
	.if (eax == IDM_EDIT)

		invoke StartEditMode, lvi.iItem

	.elseif (eax == IDM_ENABLE)

		mov bTmp, TRUE
		.if ([esi].PropItem.wFlags & PROPITEM_FACTIVE)
			mov bTmp, FALSE
		.endif
		@mov lvi.iItem, -1
		@mov lvi.iSubItem, 0
		@mov lvi.mask_, LVIF_PARAM
		.while (1)
			invoke ListView_GetNextItem( m_hWndLV, lvi.iItem, LVNI_SELECTED)
			.break .if (eax == -1)
			mov lvi.iItem, eax
			invoke ListView_GetItem( m_hWndLV, addr lvi)
			mov ecx, lvi.lParam
			and [ecx].PropItem.wFlags, NOT PROPITEM_FACTIVE
			.if (bTmp)
				or [ecx].PropItem.wFlags, PROPITEM_FACTIVE
			.endif
			invoke SetItem, ecx, lvi.iItem, FALSE
		.endw

	.elseif (eax == IDM_PROPERTIES)

		mov eax, [esi].PropItem.pValue
		mov eax, [eax]
		.if (eax)
			mov pFont, eax
			invoke Find@CObjectItem, pFont
			.if (eax)
				invoke vf(eax, IObjectItem, ShowPropertiesDlg), m_hWnd
			.else
				invoke Create@CObjectItem, pFont, NULL
				.if (eax)
					push eax
					invoke vf(eax, IObjectItem, ShowPropertiesDlg), m_hWnd
					pop eax
					invoke vf(eax, IObjectItem, Release)
				.endif
			.endif
		.endif
	.endif
done:
	ret
	align 4

ShowContextMenu endp

OnNotify proc uses esi edi pNMHDR:ptr NMHDR

local hr:DWORD
local hti:LVHITTESTINFO
local lvi:LVITEM
local pt:POINT

		@mov hr, 0
		mov edi, pNMHDR

		.if ([edi].NMHDR.code == NM_CLICK)

			invoke GetCursorPos, addr hti.pt
			invoke ScreenToClient, m_hWndLV, addr hti.pt
			invoke ListView_SubItemHitTest( m_hWndLV, addr hti)
			.if (hti.iSubItem == COLUMN_VALUE)
				invoke StartEditMode, hti.iItem
			.endif

		.elseif ([edi].NMHDR.code == NM_DBLCLK)

			invoke GetCursorPos, addr hti.pt
			invoke ScreenToClient, m_hWndLV, addr hti.pt
			invoke ListView_SubItemHitTest( m_hWndLV, addr hti)
			.if (hti.iSubItem == 0)
				mov eax, hti.iItem
				mov lvi.iItem, eax
				@mov lvi.iSubItem, 0
				@mov lvi.mask_, LVIF_PARAM
				invoke ListView_GetItem( m_hWndLV, addr lvi)
				mov esi, lvi.lParam
				.if ([esi].PropItem.wType == VT_BOOL)
					mov eax, [esi].PropItem.pValue
					xor byte ptr [eax], 1
					invoke SetItem, esi, hti.iItem, FALSE
					invoke NotifyContainerWindows
				.else
					invoke StartEditMode, hti.iItem
				.endif
			.endif

		.elseif ([edi].NMHDR.code == NM_RCLICK)

			invoke ShowContextMenu, TRUE

		.elseif ([edi].NMHDR.code == NM_CUSTOMDRAW)

			.if ([edi].NMLVCUSTOMDRAW.nmcd.dwDrawStage == CDDS_PREPAINT)

				invoke SetWindowLong, m_hWnd, DWL_MSGRESULT, CDRF_NOTIFYITEMDRAW
				mov hr, 1

			.elseif ([edi].NMLVCUSTOMDRAW.nmcd.dwDrawStage == CDDS_ITEMPREPAINT)

				mov esi, [edi].NMLVCUSTOMDRAW.nmcd.lItemlParam
				.if (!([esi].PropItem.wFlags & PROPITEM_FACTIVE))
;;					invoke GetSysColor, COLOR_BTNFACE
					mov eax, 0A7A7A7h
					mov [edi].NMLVCUSTOMDRAW.clrText, eax
				.endif

			.endif

		.elseif ([edi].NMHDR.code == LVN_KEYDOWN)

			.if ([edi].NMLVKEYDOWN.wVKey == VK_APPS)
				invoke ShowContextMenu, FALSE
			.endif

		.elseif ([edi].NMHDR.code == LVN_BEGINLABELEDIT)

			invoke OnBeginLabelEdit, edi
			mov hr, eax

		.elseif ([edi].NMHDR.code == LVN_ENDLABELEDIT)

			invoke OnEndLabelEdit, edi
			mov hr, eax

		.endif

		return hr
		align 4

OnNotify endp

SetItem proc uses esi pPropItem:ptr PropItem, iItem:DWORD, bNew:BOOL

local	dispparams:DISPPARAMS
local	varResult:VARIANT
local	pDispatch:LPDISPATCH
local	lvi:LVITEM
local	szText[MAX_PATH]:BYTE

	mov esi, pPropItem
	mov eax, iItem
	mov lvi.iItem, eax
	@mov lvi.iSubItem, 0
	invoke GetStdDispIdStr, [esi].PropItem.id
	mov lvi.pszText, eax
	.if (bNew)
		@mov lvi.mask_, LVIF_TEXT or LVIF_PARAM
		mov lvi.lParam, esi
		invoke ListView_InsertItem( m_hWndLV, addr lvi)
	.else
		mov lvi.mask_, LVIF_TEXT
		invoke ListView_SetItem( m_hWndLV, addr lvi)
	.endif
	inc lvi.iSubItem
	@mov lvi.mask_, LVIF_TEXT
	invoke wsprintf, addr szText, CStr("%d"), [esi].PropItem.id
	lea eax, szText
	mov lvi.pszText, eax
	invoke ListView_SetItem( m_hWndLV, addr lvi)
	inc lvi.iSubItem
	movzx eax, [esi].PropItem.wType
	invoke GetVarType, eax
	mov lvi.pszText, eax
	invoke ListView_SetItem( m_hWndLV, addr lvi)
	inc lvi.iSubItem

	mov eax, [esi].PropItem.pValue
	.if ([esi].PropItem.wType == VT_BOOL)
		movzx eax, byte ptr [eax]
		.if (eax)
			mov ecx, CStr("True")
		.else
			mov ecx, CStr("False")
		.endif
		invoke wsprintf, addr szText, CStr("%s"), ecx
	.elseif ([esi].PropItem.wType == VT_BSTR)
		mov eax, [eax]
		.if (eax)
			invoke wsprintf, addr szText, CStr("%S"), eax
		.else
			mov szText, 0
		.endif
	.elseif ([esi].PropItem.wType == VT_DISPATCH)
		mov ecx, [eax]
		.if (ecx)
			mov szText, 0
			invoke vf(ecx, IUnknown, QueryInterface), addr IID_IDispatch, addr pDispatch
			.if (eax == S_OK)
				xor eax, eax
				mov dispparams.rgvarg, eax
				mov dispparams.rgdispidNamedArgs, eax
				mov dispparams.cArgs, eax
				mov dispparams.cNamedArgs, eax
				invoke VariantInit, addr varResult
				invoke vf(pDispatch, IDispatch, Invoke_), DISPID_VALUE, addr IID_NULL,
					g_LCID, DISPATCH_PROPERTYGET, addr dispparams, addr varResult, NULL, NULL
				.if (eax == S_OK && varResult.vt == VT_BSTR)
					invoke wsprintf, addr szText, CStr("%S"), varResult.bstrVal
				.endif
				invoke VariantClear, addr varResult
				invoke vf(pDispatch, IUnknown, Release)
			.endif
		.else
			invoke lstrcpy, addr szText, CStr("NULL")
		.endif
	.elseif ([esi].PropItem.id == DISPID_AMBIENT_PALETTE)
		.if (!g_hPalette)
			invoke lstrcpy, addr szText, CStr("NULL")
		.else
			invoke wsprintf, addr szText, CStr("0x%X"), g_hPalette
		.endif
	.else
		.if ([esi].PropItem.wFlags & PROPITEM_FHEX)
			mov ecx, CStr("0x%X")
		.else
			mov ecx, CStr("%d")
		.endif
		invoke wsprintf, addr szText, ecx, dword ptr [eax]
	.endif
	lea eax, szText
	mov lvi.pszText, eax
	invoke ListView_SetItem( m_hWndLV, addr lvi)
	ret
	align 4

SetItem endp

RefreshList proc uses esi edi

	mov esi, offset PropTab
	mov ecx, NUMPROPS
	xor edi, edi
	.while (ecx)
		push ecx
		invoke SetItem, esi, edi, TRUE
		inc edi
		add esi, sizeof PropItem
		pop ecx
		dec ecx
	.endw
	ret
	align 4
RefreshList endp


;*** Dialog Proc for "ambient properties" dialog


AmbientPropDialog proc uses __this thisarg, message:dword, wParam:WPARAM, lParam:LPARAM

		mov __this,this@

		mov eax,message
		.if (eax == WM_INITDIALOG)
			mov g_pAmbientPropDlg, __this

			invoke GetDlgItem, m_hWnd, IDC_LIST1
			mov m_hWndLV,eax
			invoke ListView_SetExtendedListViewStyle( m_hWndLV, LVS_EX_GRIDLINES or LVS_EX_INFOTIP or LVS_EX_SUBITEMIMAGES)

			invoke SetLVColumns, m_hWndLV, NUMCOLS_PROPERTIES, addr ColumnsListView

			invoke CreateEditListView, m_hWndLV

			invoke RefreshList

			mov eax,1

		.elseif (eax == WM_CLOSE)
if ?MODELESS
			invoke DestroyWindow, m_hWnd
else
			invoke EndDialog, m_hWnd, 0
endif
			mov eax,1

		.elseif (eax == WM_DESTROY)

			invoke Destroy@CAmbientPropDlg, __this
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

			movzx eax, word ptr wParam
			.if (eax == IDCANCEL)
				invoke PostMessage, m_hWnd, WM_CLOSE, 0, 0
			.elseif (eax == IDOK)
				invoke ListView_GetNextItem( m_hWndLV, -1, LVNI_FOCUSED)
				.if (eax != -1)
					invoke StartEditMode, eax
				.endif
			.endif

		.elseif (eax == WM_NOTIFY)
			invoke OnNotify, lParam
		.else
			xor eax,eax ;indicates "no processing"
		.endif
		ret
		align 4

AmbientPropDialog endp

Init proc

	.if ((!(FontItem.wFlags & PROPITEM_FACTIVE)) && (!g_pFont))
		invoke CreateFontFromLogFont, 8*10
		.if (g_pFont)
			or FontItem.wFlags, PROPITEM_FACTIVE
		.endif
	.endif
	ret
	align 4
Init endp

;*** constructor


Create@CAmbientPropDlg proc public

		invoke Init
		mov eax, g_pAmbientPropDlg
		.if (eax)
			jmp done
		.endif
		invoke malloc, sizeof CAmbientPropDlg
done:
		ret
		align 4

Create@CAmbientPropDlg endp

Show@CAmbientPropDlg proc public uses __this thisarg, hWnd:HWND
	
		mov eax, g_pAmbientPropDlg
		.if (eax)
			invoke RestoreAndActivateWindow, [eax].CDlg.hWnd
			jmp done
		.endif
		mov __this,this@
		mov m_pDlgProc, AmbientPropDialog
if ?MODELESS
		invoke CreateDialogParam, g_hInstance, IDD_AMBIENTPROPDLG, NULL, classdialogproc,__this
else
		invoke DialogBoxParam, g_hInstance, IDD_AMBIENTPROPDLG, hWnd, classdialogproc,__this
endif
done:
		ret
		align 4

Show@CAmbientPropDlg endp

Destroy@CAmbientPropDlg proc public thisarg

		invoke free, this@
		mov g_pAmbientPropDlg, NULL
		ret
		align 4
Destroy@CAmbientPropDlg endp

GetAmbientProp proc public dwDispId:DWORD, pVariant:ptr VARIANT

		invoke Init
		mov eax, dwDispId
		mov ecx, NUMPROPS
		mov edx, offset PropTab
		.while (ecx)
			.if (eax == [edx].PropItem.id)
				.break .if (!([edx].PropItem.wFlags & PROPITEM_FACTIVE))
				mov ecx, pVariant
				mov ax, [edx].PropItem.wType
				mov [ecx].VARIANT.vt, ax
				mov eax, [edx].PropItem.pValue
				.if ([edx].PropItem.wType == VT_BOOL)
					movzx eax,byte ptr [eax]
					neg ax
					mov [ecx].VARIANT.boolVal, ax
				.elseif ([edx].PropItem.wType == VT_BSTR)
					mov eax, [eax]
					push ecx
					invoke SysAllocString, eax
					pop ecx
					mov [ecx].VARIANT.bstrVal, eax
				.elseif ([edx].PropItem.wType == VT_DISPATCH)
					mov eax, [eax]
					.if (eax)
						pushad
						invoke vf(eax, IUnknown, AddRef)
						popad
					.endif
					mov [ecx].VARIANT.pdispVal, eax
				.else
					mov eax,[eax]
					mov [ecx].VARIANT.lVal, eax
				.endif
				mov edx, eax
				return S_OK
			.endif
			add edx, sizeof PropItem
			dec ecx
		.endw
		return DISP_E_MEMBERNOTFOUND

GetAmbientProp endp


    end
