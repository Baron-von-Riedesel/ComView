
;*** definition of class CPropertyStorageDlg

	.386
	.model flat,stdcall
	option casemap :none   ; case sensitive
	option proc:private

	include COMView.inc
	include statusbar.inc
INSIDE_CPROPERTYSTORAGEDLG equ 1
	include classes.inc
	include rsrc.inc
	include debugout.inc


BEGIN_CLASS CPropertyStorageDlg, CDlg
hWndLV			HWND		?		;hWnd of listview
hWndSB			HWND		?
pPropertyStorage	LPPROPERTYSTORAGE ?
pPropertySetStorage	LPPROPERTYSETSTORAGE ?
END_CLASS

__this	textequ <ebx>
_this	textequ <[__this].CPropertyStorageDlg>
thisarg	textequ <this@:ptr CPropertyStorageDlg>

	MEMBER hWnd, pDlgProc
	MEMBER hWndLV, hWndSB, pPropertyStorage, pPropertySetStorage

	.const

ColumnsEnumProp label CColHdr
		CColHdr <CStr("Name")	, 30>
		CColHdr <CStr("Value")	, 40>
		CColHdr <CStr("Type")	, 30>
NUMCOLS_ENUMPROP textequ %($ - ColumnsEnumProp) / sizeof CColHdr

	.code

propviewdetailproc proc uses __this esi hWnd:HWND, message:DWORD, wParam:WPARAM, lParam:LPARAM

local	hWndEdit:HWND
local	lvi:LVITEM

		mov eax, message
		.if (eax == WM_INITDIALOG)
			mov __this, lParam
			invoke SetWindowLong, hWnd, DWL_USER, __this
			invoke GetDlgItem, hWnd, IDC_EDIT1
			mov hWndEdit, eax
			invoke ListView_GetNextItem( m_hWndLV, -1, LVNI_SELECTED)
			.if (eax != -1)
				mov lvi.iItem, eax
				mov lvi.iSubItem, 1
				mov lvi.mask_, LVIF_TEXT
				invoke malloc, 10000h
				mov lvi.pszText, eax
				mov lvi.cchTextMax, 10000h
				invoke ListView_GetItem( m_hWndLV, addr lvi)
				invoke SetWindowText, hWndEdit, lvi.pszText
				invoke free, lvi.pszText
			.endif
			mov eax, 1
		.elseif (eax == WM_CLOSE)
			invoke EndDialog, hWnd, 0
		.elseif (eax == WM_COMMAND)
			movzx eax, word ptr wParam+0
			.if (eax == IDCANCEL)
				invoke EndDialog, hWnd, 0
			.elseif (eax == IDOK)
				invoke EndDialog, hWnd, 0
			.endif
		.else
			xor eax, eax
		.endif
		ret
		align 4

propviewdetailproc endp

OnNotify proc uses esi pNMHdr:ptr NMHDR

local	pt:POINT

		mov ecx, pNMHdr
		.if ([ecx].NMHDR.idFrom == IDC_LIST1)
			.if ([ecx].NMHDR.code == NM_RCLICK)
				invoke CreatePopupMenu
				mov esi, eax
				invoke AppendMenu, esi, MF_STRING, IDM_VIEW, CStr("&View")
				invoke SetMenuDefaultItem, esi, 0, TRUE
				invoke GetCursorPos, addr pt
				invoke TrackPopupMenu, esi, TPM_LEFTALIGN or TPM_LEFTBUTTON,
						pt.x,pt.y,0, m_hWnd, NULL
				invoke DestroyMenu, esi
			.endif
		.endif
		ret
		align 4
OnNotify endp

if 0
InitVarFormat proc
		.data
protoVarFormatDateTime typedef proto :ptr VARIANT, :DWORD, :DWORD, :ptr BSTR
LPVARFORMATDATETIME typedef ptr protoVarFormatDateTime
g_pfnVarFormatDateTime LPVARFORMATDATETIME NULL
		.code
		.if (!g_pfnVarFormatDateTime)
			invoke GetModuleHandle, CStr("OLEAUT32")
			.if (eax)
				invoke GetProcAddress, eax, CStr("VarFormatDateTime")
				mov g_pfnVarFormatDateTime, eax
			.endif
		.endif
		ret
		align 4
InitVarFormat endp
endif

SetValue proc uses esi edi hWndLV:HWND, plvi:ptr LVITEM, ppropvar:ptr PROPVARIANT

local	pszSuffix:LPSTR
local	var:VARIANT
local	pv:PROPVARIANT
local	dwSize:DWORD
local	pMem:LPVOID
local	systime:SYSTEMTIME
local	szValue[256]:byte

		mov esi, ppropvar
		assume esi:ptr PROPVARIANT

		.if ([esi].vt & VT_VECTOR)
			invoke VariantInit, addr pv
			mov ax, [esi].vt
			and ax, VT_TYPEMASK
			mov pv.vt, ax
			mov ecx, [esi].caub.cElems
			mov edi, [esi].caub.pElems
			.while (ecx)
				push ecx
				push eax
				.if (ax == VT_LPSTR)
					mov eax, [edi]
					mov pv.pszVal, eax
					add edi, 4
					invoke SetValue, hWndLV, plvi, addr pv
					clc
				.elseif (ax == VT_VARIANT)
					invoke SetValue, hWndLV, plvi, edi
					add edi, sizeof VARIANT
					clc
				.else
					stc
				.endif
				pop eax
				pop ecx
				jc next
				dec ecx
			.endw
			jmp done
		.endif
next:
		mov pMem, NULL
		mov szValue, 0
		lea eax, szValue
		mov ecx, plvi
		mov [ecx].LVITEM.pszText, eax
		.if ([esi].vt == VT_LPSTR)
			invoke lstrcpyn, addr szValue, [esi].pszVal, sizeof szValue
		.elseif ([esi].vt == VT_FILETIME)
			invoke FileTimeToLocalFileTime, addr [esi].filetime, addr [esi].filetime
			invoke FileTimeToSystemTime, addr [esi].filetime, addr systime
			movzx ecx, systime.wYear
			invoke wsprintf, addr szValue, CStr("%02u/%02u/%u %02u:%02u:%02u"),\
				SWORD ptr systime.wMonth, SWORD ptr systime.wDay, ecx,\
				SWORD ptr systime.wHour, SWORD ptr systime.wMinute, SWORD ptr systime.wSecond
;;			invoke wsprintf, addr szValue, CStr("%08X:%08X"),\
;;				[esi].filetime.dwHighDateTime, [esi].filetime.dwLowDateTime
		.else
			invoke VariantChangeType, esi, esi, 0, VT_BSTR
			.if (eax == S_OK)
				invoke SysStringLen, [esi].bstrVal
				inc eax
				mov dwSize, eax
				.if (eax >= sizeof szValue)
					push eax
					invoke malloc, eax
					mov pMem, eax
					pop edx
					mov ecx, plvi
					mov [ecx].LVITEM.pszText, eax
				.endif
				mov ecx, plvi
				invoke WideCharToMultiByte,CP_ACP,0, [esi].bstrVal,-1,[ecx].LVITEM.pszText,dwSize,0,0
			.else
				movzx ecx, [esi].vt
				invoke wsprintf, addr szValue, CStr("VariantChangeType failed [%X] for %X"), eax, ecx
			.endif
		.endif
		mov ecx, plvi
		invoke ListView_SetItem( hWndLV, ecx)
		.if (!eax)
			mov ecx, plvi
			push [ecx].LVITEM.mask_
			push [ecx].LVITEM.iSubItem
			mov [ecx].LVITEM.mask_, 0
			mov [ecx].LVITEM.iSubItem, 0
			invoke ListView_InsertItem( hWndLV, plvi)
			mov ecx, plvi
			pop [ecx].LVITEM.iSubItem
			pop [ecx].LVITEM.mask_
			invoke ListView_SetItem( hWndLV, plvi)
		.endif
		.if (pMem)
			invoke free, pMem
		.endif
		mov ecx, plvi
		inc [ecx].LVITEM.iItem
done:
		ret
		assume esi:nothing
		align 4

SetValue endp


	.const
PSGUID_USERDEFINEDINFORMATION GUID {0d5cdd505h, 02e9ch, 0101bh, {093h, 97h, 008h, 000h, 02bh, 02ch, 0f9h, 0aeh}}
	.code

PROPVARIANTCLEAR typedef proto stdcall :ptr PROPVARIANT

OnInitDialog proc

local	pEnumSTATPROPSTG:LPENUMSTATPROPSTG
local	bSecondPass:BOOL
local	filler:DWORD
local	sps:STATPROPSTG
local	propspec:PROPSPEC
local	propvar:PROPVARIANT
local	lpfnPropVariantClear:ptr PROPVARIANTCLEAR
local	statpropsetstg:STATPROPSETSTG
local	pszType:LPSTR
local	szTypeEx[32]:byte
local	szType[64]:byte
local	szName[64]:byte
local	wszGUID[40]:word
local	szGUID[40]:byte
local	lvi:LVITEM
local	szText[128]:byte

		invoke GetDlgItem, m_hWnd, IDC_LIST1
		mov m_hWndLV, eax
		invoke GetDlgItem, m_hWnd, IDC_STATUSBAR
		mov m_hWndSB, eax
		invoke ListView_SetExtendedListViewStyle( m_hWndLV, LVS_EX_FULLROWSELECT or LVS_EX_INFOTIP)
		invoke SetLVColumns, m_hWndLV, NUMCOLS_ENUMPROP, addr ColumnsEnumProp
		invoke GetModuleHandle, CStr("OLE32")
		invoke GetProcAddress, eax, CStr("PropVariantClear")
		mov lpfnPropVariantClear, eax

		invoke SetWindowText, m_hWnd, CStr("View IPropertyStorage")

		@mov lvi.iItem, 0
		mov filler, -1	
		mov bSecondPass, FALSE	
next:
		mov lvi.mask_, LVIF_TEXT

		invoke vf(m_pPropertyStorage, IPropertyStorage, Enum), addr pEnumSTATPROPSTG
		.if (eax != S_OK)
			jmp done
		.endif

		.while (1)
			invoke vf(pEnumSTATPROPSTG, IEnumSTATPROPSTG, Next), 1, addr sps, NULL
			.break .if (eax != S_OK)
			.if (sps.lpwstrName)
				invoke WideCharToMultiByte,CP_ACP,0, sps.lpwstrName,-1,addr szName,sizeof szName,0,0
				invoke CoTaskMemFree, sps.lpwstrName
			.else
				invoke wsprintf, addr szName, CStr("PropID(%d)"), sps.propid
			.endif
			mov lvi.mask_, LVIF_TEXT or LVIF_PARAM

			lea eax, szName
			mov lvi.pszText, eax
			mov lvi.iSubItem, 0
			mov eax, sps.propid
			mov lvi.lParam, eax
			invoke ListView_InsertItem( m_hWndLV, addr lvi)

			mov lvi.mask_, LVIF_TEXT

			mov propspec.ulKind, PRSPEC_PROPID
			mov eax, sps.propid
			mov propspec.propid, eax

			mov szType, 0

			invoke vf(m_pPropertyStorage, IPropertyStorage, ReadMultiple), 1, addr propspec, addr propvar
			.if (eax == S_OK)

				movzx ecx, propvar.vt
				and cx, VT_TYPEMASK
				invoke GetVarType, ecx
				mov pszType, eax
				mov ecx, offset g_szNull
				mov edx, ecx
				mov eax, ecx
				.if (propvar.vt & VT_ARRAY)
					mov edx, CStr("Array of ")
				.endif
				.if (propvar.vt & VT_VECTOR)
					pushad
					invoke wsprintf, addr szTypeEx, CStr("Vector[%u] of "), propvar.caub.cElems
					popad
					lea ecx, szTypeEx
				.endif
				.if (propvar.vt & VT_BYREF)
					mov eax, CStr("ptr of ")
				.endif
				invoke wsprintf, addr szType, CStr("%s%s%s%s"), edx, ecx, eax, pszType

				.if (szType)
					lea eax, szType
					mov lvi.pszText, eax
					mov lvi.iSubItem, 2
					invoke ListView_SetItem( m_hWndLV, addr lvi)
				.endif

				mov lvi.iSubItem, 1
				invoke SetValue, m_hWndLV, addr lvi, addr propvar
				.if (lpfnPropVariantClear)
					invoke lpfnPropVariantClear, addr propvar
				.else
					invoke VariantClear, addr propvar
				.endif

			.else
				invoke wsprintf, addr szName, CStr("IPropertyStorage::ReadMultiple failed [%X]"), eax
				lea eax, szName
				mov lvi.pszText, eax
				mov lvi.iSubItem, 1
				invoke ListView_SetItem( m_hWndLV, addr lvi)
				inc lvi.iItem
			.endif
		.endw

		invoke vf(pEnumSTATPROPSTG, IUnknown, Release)

		invoke vf(m_pPropertyStorage, IPropertyStorage, Stat), addr statpropsetstg
		.if (eax == S_OK)
			invoke StringFromGUID2, addr statpropsetstg.fmtid, addr wszGUID, 40
			invoke WideCharToMultiByte,CP_ACP,0, addr wszGUID,-1,addr szGUID,sizeof szGUID,0,0
			invoke wsprintf, addr szText, CStr("FmtId=%s"), addr szGUID
			push eax
			invoke StringFromGUID2, addr statpropsetstg.clsid, addr wszGUID, 40
			invoke WideCharToMultiByte,CP_ACP,0, addr wszGUID,-1,addr szGUID,sizeof szGUID,0,0
			pop eax
			lea ecx, [eax+szText]
			invoke wsprintf, ecx, CStr(" ClsId=%s"), addr szGUID
			StatusBar_SetText m_hWndSB, 0, addr szText
		.endif

;-------------------------------- this is a special case
;-------------------------------- with propertyset IDocumentSummary
		.if (m_pPropertySetStorage && (!bSecondPass))
;-------------------------------- before we can open second section, 
;-------------------------------- the first has to be released (totally)
			invoke vf(m_pPropertyStorage, IUnknown, Release)
			invoke vf(m_pPropertyStorage, IUnknown, Release)
			mov m_pPropertyStorage, NULL
			invoke vf(m_pPropertySetStorage, IPropertySetStorage, Open), addr PSGUID_USERDEFINEDINFORMATION,\
					STGM_READ or STGM_SHARE_EXCLUSIVE, addr m_pPropertyStorage
			.if (eax == S_OK)
				mov lvi.iSubItem, 0
				mov lvi.pszText, CStr("****************************")
				invoke ListView_InsertItem( m_hWndLV, addr lvi)
				mov lvi.iSubItem, 2
				invoke ListView_SetItem( m_hWndLV, addr lvi)
				mov lvi.pszText, CStr("Userdefined Property Set")
				mov lvi.iSubItem, 1
				invoke ListView_SetItem( m_hWndLV, addr lvi)
				inc lvi.iItem
				mov bSecondPass, TRUE
				jmp next
			.endif
		.endif
done:
		ret
		align 4
OnInitDialog endp

CPropertyStorageDialog proc uses __this thisarg, message:DWORD, wParam:WPARAM, lParam:LPARAM

	mov __this,this@

	mov eax, message
	.if (eax == WM_INITDIALOG)

		invoke OnInitDialog
		mov eax, 1

	.elseif (eax == WM_CLOSE)

		invoke EndDialog, m_hWnd, 0

	.elseif (eax == WM_NOTIFY)

		invoke OnNotify, lParam

	.elseif (eax == WM_COMMAND)

		movzx eax, word ptr wParam+0
		.if (eax == IDCANCEL)
			invoke EndDialog, m_hWnd, 0
		.elseif (eax == IDM_VIEW)
			invoke DialogBoxParam, g_hInstance, IDD_VIEWDETAIL, m_hWnd, propviewdetailproc, __this
		.endif

	.else
		xor eax, eax
	.endif
	ret
	align 4

CPropertyStorageDialog endp


Destroy@CPropertyStorageDlg proc uses __this this_:ptr CPropertyStorageDlg
	mov __this, this_
	.if (m_pPropertySetStorage)
		invoke vf(m_pPropertySetStorage, IUnknown, Release)
	.endif
	.if (m_pPropertyStorage)
		invoke vf(m_pPropertyStorage, IUnknown, Release)
	.endif
	invoke free, __this
	ret
	align 4
Destroy@CPropertyStorageDlg endp

Create@CPropertyStorageDlg proc public uses __this pPropertyStorage:LPPROPERTYSTORAGE, pPropertySetStorage:LPPROPERTYSETSTORAGE

		invoke malloc, sizeof CPropertyStorageDlg
		.if (!eax)
			ret
		.endif
		mov __this, eax
		mov m_pDlgProc, CPropertyStorageDialog
		mov eax, pPropertyStorage
		mov m_pPropertyStorage, eax
		invoke vf(eax, IUnknown, AddRef)
		mov eax, pPropertySetStorage
		mov m_pPropertySetStorage, eax
		.if (eax)
			invoke vf(eax, IUnknown, AddRef)
		.endif
		return __this
		align 4

Create@CPropertyStorageDlg endp

Show@CPropertyStorageDlg proc public this_:ptr CPropertyStorageDlg, hWnd:HWND 
		invoke DialogBoxParam, g_hInstance, IDD_ENUMFORMATETCDLG,
			hWnd, classdialogproc, this_
		invoke Destroy@CPropertyStorageDlg, this_
		ret
		align 4
Show@CPropertyStorageDlg endp

	end
