
;*** definition of class CEnumFORMATETCDlg
;*** will handle IEnumFORMATETC objects

	.386
	.model flat,stdcall
	option casemap :none   ; case sensitive
	option proc:private

	include COMView.inc
INSIDE_CENUMFORMATETCDLG equ 1
	include classes.inc
	include rsrc.inc
	include debugout.inc

?MODELESS			equ 1


BEGIN_CLASS CEnumFORMATETCDlg, CDlg
hWndLV			HWND		?		;hWnd of listview
pEnumFORMATETC	LPENUMFORMATETC	?
pDataObject		LPDATAOBJECT	?
END_CLASS

__this	textequ <ebx>
_this	textequ <[__this].CEnumFORMATETCDlg>
thisarg	textequ <this@:ptr CEnumFORMATETCDlg>

	MEMBER hWnd, pDlgProc
	MEMBER hWndLV, pEnumFORMATETC, pDataObject

	.const

ColumnsEnumFORMATETC label CColHdr
		CColHdr <CStr("ClipFormat")	, 30>
		CColHdr <CStr("TargetDevice"), 15>
		CColHdr <CStr("Aspect")		, 20>
		CColHdr <CStr("Index")		, 10, FCOLHDR_RDX10>
		CColHdr <CStr("Tymed")		, 25>
NUMCOLS_ENUMFORMATETC textequ %($ - ColumnsEnumFORMATETC) / sizeof CColHdr

	.code

;--------------------------------------------------------------
;--- class CEnumFORMATETCDlg
;--------------------------------------------------------------

Destroy@CEnumFORMATETCDlg proc uses __this thisarg

		mov __this,this@
		.if (m_pEnumFORMATETC)
			invoke vf(m_pEnumFORMATETC, IUnknown, Release)
		.endif
		.if (m_pDataObject)
			invoke vf(m_pDataObject, IUnknown, Release)
		.endif
		invoke free, __this
		ret
		align 4
Destroy@CEnumFORMATETCDlg endp


GetDVAspectString proc dwAspect:DWORD, pStrOut:LPSTR
	mov ecx, dwAspect
	.if (ecx == DVASPECT_CONTENT)
		mov eax, CStr("Content")
	.elseif (ecx == DVASPECT_THUMBNAIL)
		mov eax, CStr("Thumbnail")
	.elseif (ecx == DVASPECT_ICON)
		mov eax, CStr("Icon")
	.elseif (ecx == DVASPECT_DOCPRINT)
		mov eax, CStr("Docprint")
	.else
		invoke wsprintf, pStrOut, CStr("%X"), ecx
		xor eax, eax
	.endif
	.if (eax)
		invoke lstrcpy, pStrOut, eax
	.endif
	ret

GetDVAspectString endp

TymedValue label dword
	dd TYMED_HGLOBAL,CStr("HGlobal")
	dd TYMED_FILE,CStr("File")
	dd TYMED_ISTREAM,CStr("IStream")
	dd TYMED_ISTORAGE,CStr("IStorage")
	dd TYMED_GDI,CStr("GDI")
	dd TYMED_MFPICT,CStr("MFPict")
	dd TYMED_ENHMF,CStr("EnhMF")
	dd TYMED_NULL,CStr("Null")
NUMTYMEDVALUE equ ($ - TymedValue) / (sizeof DWORD * 2)

GetTymedString proc uses esi edi ebx tymed:DWORD, pStrOut:LPSTR

		mov ebx,tymed
		mov edi, pStrOut
		mov byte ptr [edi], 0
		mov ecx,NUMTYMEDVALUE
		mov esi,offset TymedValue
		.while (ecx)
			push ecx
			lodsd
			test ebx, eax
			.if (!ZERO?)
				.if (byte ptr [edi])
					invoke lstrcat, edi, CStr("|")
				.endif
				invoke lstrcat, edi, dword ptr [esi]
			.endif
			lodsd
			pop ecx
			dec ecx
		.endw
		.if (!(byte ptr [edi]))
			invoke wsprintf, pStrOut, CStr("%X"), ebx
		.endif
		ret

GetTymedString endp

pCFValue label dword
	dd CF_TEXT
	dd CF_BITMAP
	dd CF_METAFILEPICT
	dd CF_SYLK
	dd CF_DIF
	dd CF_TIFF
	dd CF_OEMTEXT
	dd CF_DIB
	dd CF_PALETTE
	dd CF_PENDATA
	dd CF_RIFF
	dd CF_WAVE
	dd CF_UNICODETEXT
	dd CF_ENHMETAFILE
	dd CF_HDROP
	dd CF_LOCALE
	dd CF_MAX
NUMCFVALUE equ ($ - pCFValue) / sizeof DWORD
	dd CStr("CF_TEXT")
	dd CStr("CF_BITMAP")
	dd CStr("CF_METAFILEPICT")
	dd CStr("CF_SYLK")
	dd CStr("CF_DIF")
	dd CStr("CF_TIFF")
	dd CStr("CF_OEMTEXT")
	dd CStr("CF_DIB")
	dd CStr("CF_PALETTE")
	dd CStr("CF_PENDATA")
	dd CStr("CF_RIFF")
	dd CStr("CF_WAVE")
	dd CStr("CF_UNICODETEXT")
	dd CStr("CF_ENHMETAFILE")
	dd CStr("CF_HDROP")
	dd CStr("CF_LOCALE")
	dd CStr("CF_MAX")

GetCFString proc uses edi cf:DWORD, pStrOut:LPSTR

		mov eax,cf
		mov ecx,NUMCFVALUE
		mov edi,offset pCFValue
		repnz scasd
		.if (ZERO?)
			add edi, (NUMCFVALUE-1)*4
			invoke lstrcpy, pStrOut, dword ptr [edi]
		.else
			invoke GetClipboardFormatName, eax, pStrOut, 128
			.if (!eax)
				invoke wsprintf, pStrOut, CStr("%X"), cf
			.endif
		.endif
		ret

GetCFString endp

viewdetailproc proc uses __this hWnd:HWND, message:DWORD, wParam:WPARAM, lParam:LPARAM

local	dwSize:DWORD
local	hWndEdit:HWND
local	rect:RECT
local	pStgMedium:ptr STGMEDIUM
local	pMetaFilePict:ptr METAFILEPICT
local	ps:PAINTSTRUCT
local	pt1:POINT
local	pt2:POINT

		mov eax, message
		.if (eax == WM_INITDIALOG)
			invoke SetWindowLong, hWnd, DWL_USER, lParam
			invoke GetDlgItem, hWnd, IDC_EDIT1
			mov hWndEdit, eax
			invoke ShowWindow, hWndEdit, SW_HIDE
			mov eax, 1
		.elseif (eax == WM_CLOSE)
			invoke EndDialog, hWnd, 0
		.elseif (eax == WM_PAINT)
			invoke BeginPaint, hWnd, addr ps
			invoke GetDlgItem, hWnd, IDC_EDIT1
			mov hWndEdit, eax
			invoke GetWindowRect, hWndEdit, addr rect
			invoke ScreenToClient, hWnd, addr rect
			invoke ScreenToClient, hWnd, addr rect.right
			invoke GetWindowLong, hWnd, DWL_USER
			mov pStgMedium, eax
			mov ecx, eax
			.if ([eax].STGMEDIUM.tymed == TYMED_ENHMF)
				invoke PlayEnhMetaFile, ps.hdc, [ecx].STGMEDIUM.hEnhMetaFile, addr rect
			.elseif([eax].STGMEDIUM.tymed == TYMED_MFPICT) 
				mov eax, [ecx].STGMEDIUM.hMetaFilePict
				invoke GlobalLock, eax
				DebugOut "hMetaFilePict.mm=%X, xExt=%X, yExt=%X, hMF=%X",\
					[eax].METAFILEPICT.mm_, [eax].METAFILEPICT.xExt,\
					[eax].METAFILEPICT.yExt, [eax].METAFILEPICT.hMF
				mov pMetaFilePict, eax
				invoke SetMapMode, ps.hdc, [eax].METAFILEPICT.mm_
				push eax
				invoke SetViewportOrgEx, ps.hdc, rect.left, rect.top, addr pt1
				mov ecx, rect.right
				sub ecx, rect.left
				mov edx, rect.bottom
				sub edx, rect.top
				invoke SetViewportExtEx, ps.hdc, ecx, edx, addr pt2
				mov eax, pMetaFilePict
				invoke PlayMetaFile, ps.hdc, [eax].METAFILEPICT.hMF
				mov eax, pStgMedium
				invoke GlobalUnlock, [eax].STGMEDIUM.hMetaFilePict
				pop eax
				invoke SetMapMode, ps.hdc, eax
				invoke SetViewportOrgEx, ps.hdc, pt1.x, pt1.y, NULL
				invoke SetViewportExtEx, ps.hdc, pt2.x, pt2.y, NULL
			.endif
			invoke GetStockObject, WHITE_BRUSH
			invoke FrameRect, ps.hdc, addr rect, eax
			invoke EndPaint, hWnd, addr ps
			mov eax, 1
		.elseif (eax == WM_COMMAND)
			movzx eax, word ptr wParam+0
			.if (eax == IDCANCEL)
				invoke EndDialog, hWnd, 0
			.elseif (eax == IDOK)
				invoke EndDialog, hWnd, eax
			.endif
		.else
			xor eax, eax
		.endif
		ret
		align 4

viewdetailproc endp

;--- displays context menu

ShowContextMenu proc uses esi bMouse:BOOL

local	pt:POINT
local	hPopupMenu:HMENU
local	stgmedium:STGMEDIUM
local	pFormatETC:ptr FORMATETC
local	pStream:LPSTREAM
local	lvi:LVITEM

		invoke ListView_GetNextItem( m_hWndLV, -1, LVIS_SELECTED)
		.if (eax != -1)
			mov lvi.iItem, eax
			mov lvi.iSubItem,0
			mov lvi.mask_, LVIF_PARAM
			invoke ListView_GetItem( m_hWndLV, addr lvi)
			mov esi, lvi.lParam
			mov pFormatETC, esi
			.if (eax && esi)
				invoke CreatePopupMenu
				mov hPopupMenu, eax

				.if (m_pDataObject && (([esi].FORMATETC.tymed & (TYMED_HGLOBAL or TYMED_ISTREAM or TYMED_ISTORAGE))))
					invoke AppendMenu, hPopupMenu, MF_STRING, IDM_VIEW, CStr("&View")
				.endif

				.if (m_pDataObject && (([esi].FORMATETC.tymed & (TYMED_ENHMF or TYMED_MFPICT))))
					invoke AppendMenu, hPopupMenu, MF_STRING, IDM_PLAY, CStr("&Play")
				.endif

				invoke SetMenuDefaultItem, hPopupMenu, 0, TRUE
				invoke GetCursorPos, addr pt
				invoke TrackPopupMenu, hPopupMenu, TPM_LEFTALIGN or TPM_LEFTBUTTON,\
						pt.x,pt.y,0,m_hWnd,NULL
				invoke DestroyMenu, hPopupMenu
			.endif
		.endif
		ret
		align 4

ShowContextMenu endp


;--- RefreshList will 


RefreshList proc uses esi thisarg

local fe:FORMATETC
local lvi:LVITEM
local szText[128]:byte


		invoke ListView_DeleteAllItems( m_hWndLV)

		@mov lvi.iItem, 0
		xor esi, esi

;--- get variants, put them on stack

		.while (1)
			invoke vf(m_pEnumFORMATETC, IEnumFORMATETC, Next), 1, addr fe, NULL
			.break .if (eax != S_OK)
			inc esi

			invoke malloc, sizeof FORMATETC
			mov lvi.lParam, eax
			.if (eax)
				mov ecx, eax
				invoke CopyMemory, ecx, addr fe, sizeof FORMATETC
			.endif
			mov lvi.mask_, LVIF_TEXT or LVIF_PARAM
			@mov lvi.iSubItem, 0
			lea eax, szText
			mov lvi.pszText, eax
	
	        movzx   ecx, fe.cfFormat
			invoke GetCFString, ecx, addr szText
			invoke ListView_InsertItem( m_hWndLV, addr lvi)
			inc lvi.iSubItem
			mov lvi.mask_, LVIF_TEXT

	        mov   eax, fe.ptd
			invoke wsprintf, addr szText, CStr("%X"), eax
			invoke ListView_SetItem( m_hWndLV, addr lvi)
			inc lvi.iSubItem

	        mov   ecx, fe.dwAspect
			invoke GetDVAspectString, ecx, addr szText
			invoke ListView_SetItem( m_hWndLV, addr lvi)
			inc lvi.iSubItem

	        mov   eax, fe.lindex
			invoke wsprintf, addr szText, CStr("%d"), eax
			invoke ListView_SetItem( m_hWndLV, addr lvi)
			inc lvi.iSubItem

	        mov   ecx, fe.tymed
			invoke GetTymedString, ecx, addr szText
			invoke ListView_SetItem( m_hWndLV, addr lvi)
			inc lvi.iSubItem

			inc lvi.iItem
		.endw
		ret

RefreshList endp


;--- WM_NOTIFY


OnNotify proc  pNMHdr:ptr NMHDR

local dwRC:DWORD
local varItemNew:VARIANT
local varItemOld:VARIANT
local dwSize:DWORD
local hti:LVHITTESTINFO

		mov dwRC, FALSE
		mov eax, pNMHdr
		.if ([eax].NMHDR.code == NM_RCLICK)

			invoke ShowContextMenu, TRUE

		.elseif ([eax].NMHDR.code == NM_DBLCLK)

			invoke PostMessage, m_hWnd, WM_COMMAND, IDM_VIEW, 0

		.elseif ([eax].NMHDR.code == LVN_DELETEITEM)

			invoke free, [eax].NMLISTVIEW.lParam

		.elseif ([eax].NMHDR.code == LVN_DELETEALLITEMS)

			invoke SetWindowLong, m_hWnd, DWL_MSGRESULT, 0
			mov eax, TRUE

		.endif

		return dwRC

OnNotify endp


OnViewPlay proc uses esi iCmd:DWORD

local	stgmedium:STGMEDIUM
local	pFormatETC:ptr FORMATETC
local	pStream:LPSTREAM
local	lvi:LVITEM

		invoke ListView_GetNextItem( m_hWndLV, -1, LVIS_SELECTED)
		.if (eax == -1)
			jmp done
		.endif
		mov lvi.iItem, eax
		mov lvi.iSubItem,0
		mov lvi.mask_, LVIF_PARAM
		invoke ListView_GetItem( m_hWndLV, addr lvi)
		mov esi, lvi.lParam
		mov pFormatETC, esi

		invoke vf(m_pDataObject, IDataObject, GetData), esi, addr stgmedium
		.if (eax == S_OK)
			mov eax, E_FAIL
			.if (iCmd == IDM_PLAY)
				invoke DialogBoxParam, g_hInstance, IDD_VIEWDETAIL, m_hWnd, viewdetailproc, addr stgmedium
				invoke ReleaseStgMedium, addr stgmedium
				jmp done
			.endif

			.if (stgmedium.tymed == TYMED_HGLOBAL)
				invoke CreateStreamOnHGlobal, stgmedium.hGlobal, TRUE, addr pStream
			.elseif (stgmedium.tymed == TYMED_ISTREAM)
				mov eax, stgmedium.pstm
				mov pStream, eax
				invoke vf(pStream, IUnknown, AddRef)
				mov eax, S_OK
			.elseif (stgmedium.tymed == TYMED_ISTORAGE)
				mov eax, stgmedium.pstg
				mov pStream, eax
				invoke vf(pStream, IUnknown, AddRef)
				mov eax, S_OK
			.endif
			.if (eax == S_OK)
				invoke Create@CViewStorageDlg, pStream, NULL, NULL
				.if (eax)
					invoke Show@CViewStorageDlg, eax, NULL
				.endif
			.endif
			.if (stgmedium.tymed != TYMED_HGLOBAL)
				invoke ReleaseStgMedium, addr stgmedium
			.endif
		.else
			invoke OutputMessage, m_hWnd, eax, CStr("IDataObject::GetData"), 0
		.endif
done:
		ret

OnViewPlay endp

;--- WM_COMMAND processing


OnCommand proc wParam:WPARAM, lParam:LPARAM

		movzx eax, word ptr wParam+0
		.if (eax == IDCANCEL)

			invoke PostMessage, m_hWnd, WM_CLOSE, 0, 0

		.elseif (eax == IDOK)

;;			invoke PostMessage, m_hWnd, WM_CLOSE, 0, 0

		.elseif (eax == IDM_VIEW)

			invoke OnViewPlay, eax

		.elseif (eax == IDM_PLAY)

			invoke OnViewPlay, eax

		.endif
		ret
		align 4

OnCommand endp


;--- enum a FORMATETC in a simple dialog


CEnumFORMATETCDialog proc uses esi __this thisarg, message:DWORD, wParam:WPARAM, lParam:LPARAM


		mov __this,this@

		mov eax, message
		.if (eax == WM_INITDIALOG)

			invoke GetDlgItem, m_hWnd, IDC_LIST1
			mov m_hWndLV, eax

			invoke SetLVColumns, m_hWndLV, NUMCOLS_ENUMFORMATETC, addr ColumnsEnumFORMATETC
		
			invoke RefreshList, __this
			.if (!eax)
				invoke PostMessage, m_hWnd, WM_CLOSE, 0, 0
			.else
				invoke ListView_SetExtendedListViewStyle( m_hWndLV,LVS_EX_FULLROWSELECT or LVS_EX_INFOTIP)
if ?MODELESS
				invoke ShowWindow, m_hWnd, SW_SHOWNORMAL
endif
			.endif

			mov eax, 1

		.elseif (eax == WM_CLOSE)

if ?MODELESS
			invoke DestroyWindow, m_hWnd
else
			invoke EndDialog, m_hWnd, 0
endif
		.elseif (eax == WM_DESTROY)

			invoke ListView_DeleteAllItems( m_hWndLV)

			invoke Destroy@CEnumFORMATETCDlg, __this

		.elseif (eax == WM_COMMAND)

			invoke OnCommand, wParam, lParam

		.elseif (eax == WM_NOTIFY)

			invoke OnNotify, lParam

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
		.else
			xor eax, eax
		.endif
		ret
		align 4

CEnumFORMATETCDialog endp


;--- constructor


Create@CEnumFORMATETCDlg proc public uses esi __this pEnumFORMATETC:LPENUMFORMATETC, pDataObject:LPDATAOBJECT

		invoke malloc, sizeof CEnumFORMATETCDlg
		.if (!eax)
			ret
		.endif

		mov __this,eax
		mov m_pDlgProc, CEnumFORMATETCDialog

		mov eax, pEnumFORMATETC
		mov m_pEnumFORMATETC, eax
		.if (eax)
			invoke vf(eax, IUnknown, AddRef)
		.endif
		mov eax, pDataObject
		mov m_pDataObject, eax
		.if (eax)
			invoke vf(eax, IUnknown, AddRef)
		.endif

		return __this
error:
		invoke Destroy@CEnumFORMATETCDlg, __this
		return 0

Create@CEnumFORMATETCDlg endp

		end
