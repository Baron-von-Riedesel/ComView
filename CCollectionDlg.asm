
;*** definition of class CCollectionDlg
;*** will handle IEnumVARIANT + IEnumUnknown objects

	.386
	.model flat,stdcall
	option casemap :none   ; case sensitive
	option proc:private

	include COMView.inc
INSIDE_CCOLLECTIONDLG equ 1
	include classes.inc
	include rsrc.inc
	include debugout.inc
	include CListView.inc

;?MAXITEMS			equ 2000		;cancel enumeration if more than 2000 items
?TYPEFROMTYPEINFO	equ 1
?MODELESS			equ 1

VALUE_COLUMN		equ 1

BEGIN_CLASS CCollectionDlg, CDlg
hWndLV			HWND		?		;hWnd of listview
pDispatch		LPDISPATCH	?
memid			MEMBERID	?
pEnumVariant	LPENUMVARIANT	?	;used for showing a collection
pObjectItem		pCObjectItem	?
pVariants		LPVARIANT	?
iNumVariants	DWORD		?
pDispItem		LPDISPATCH	?		;IDispatch of an item or NULL
dwDispId		DWORD		?		;dispid of property to show
pParamReturn	LPPARAMRETURN ?
iItem			DWORD		?
variant			VARIANT		<>		;
dwNumVars		DWORD		?
hWndFrom		HWND		?		;properties dialog which has created us
iSortCol		DWORD		?		;sort column index
iSortDir		DWORD		?		;sort column direction
bUnknown		BOOLEAN		?		;is actually a IEnumUnknown!
bExplore		BOOLEAN		?
bInlineEdit		BOOLEAN		?		;allow inline edit
bDirty			BOOLEAN		?
bNumeric		BOOLEAN		?		;format of value column
bCol1Alpha		BOOLEAN		?		;format of first column
END_CLASS

RefreshList		proto

__this	textequ <ebx>
_this	textequ <[__this].CCollectionDlg>
thisarg	textequ <this@:ptr CCollectionDlg>

	MEMBER hWnd, pDlgProc
	MEMBER hWndLV, pDispatch, memid, pEnumVariant, pObjectItem, ParamReturn, iItem, bExplore
	MEMBER pVariants, iNumVariants, bUnknown, variant, dwNumVars, hWndFrom, pDispItem
	MEMBER iSortCol, iSortDir, bNumeric, bCol1Alpha
	MEMBER pParamReturn, bInlineEdit, bDirty, dwDispId

	.data

g_bExplore	BOOLEAN	FALSE
g_rect		RECT <0,0,0,0>

	.const

ColumnsCollections label CColHdr
		CColHdr <CStr("Item")		, 15,		FCOLHDR_RDX10>
		CColHdr <CStr("Value")		, 50>
		CColHdr <CStr("Type")		, 35>
NUMCOLS_COLLECTIONS textequ %($ - ColumnsCollections) / sizeof CColHdr

	.code

ClearArray proc uses esi

local	hCsrOld:HCURSOR

	invoke SetCursor, g_hCsrWait
	mov hCsrOld, eax
	mov ecx, m_iNumVariants
	mov esi, m_pVariants
	.while (ecx)
		push ecx
		invoke VariantClear, esi
		pop ecx
		add esi, sizeof VARIANT
		dec ecx
	.endw
	invoke free, m_pVariants
	mov m_pVariants, NULL
	@mov m_iNumVariants, 0
	invoke SetCursor, hCsrOld
	ret
	align 4
ClearArray endp

;--------------------------------------------------------------
;--- class CCollectionDlg
;--------------------------------------------------------------

Destroy@CCollectionDlg proc uses __this thisarg

	mov __this,this@

	.if (m_pObjectItem)
		invoke vf(m_pObjectItem, IObjectItem, Release) 
	.endif
	.if (m_pDispItem)
		invoke vf(m_pDispItem, IUnknown, Release)
	.endif
	.if (m_pDispatch)
		invoke vf(m_pDispatch, IUnknown, Release)
	.endif
	.if (m_pEnumVariant)
		invoke vf(m_pEnumVariant, IUnknown, Release)
	.endif
	invoke ClearArray
	invoke VariantClear, addr m_variant
	invoke free, __this
	ret
	align 4
Destroy@CCollectionDlg endp


DoSort proc
	mov eax, m_iSortCol
	xor ecx, ecx
;---------------------------------- Value column is special case
	.if ((eax == 0) && m_bCol1Alpha)
		;
	.elseif (eax == VALUE_COLUMN)
		movzx ecx, m_bNumeric
	.elseif ([eax * sizeof CColHdr + offset ColumnsCollections].CColHdr.wFlags & FCOLHDR_RDXMASK)
		inc ecx
	.endif
	invoke LVSort, m_hWndLV, m_iSortCol, m_iSortDir, ecx
	ret
	align 4
DoSort endp


;--- show context menu


ShowContextMenu proc bMouse:BOOL

local	pt:POINT
local	hPopupMenu:HMENU

	invoke ListView_GetSelectedCount( m_hWndLV)
	.if (eax)
		invoke GetSubMenu,g_hMenu, ID_SUBMENU_COLLECTIONDLG
		.if (eax != 0)
			mov hPopupMenu, eax
			.if (m_bExplore)
				mov ecx, MF_CHECKED
			.else
				mov ecx, MF_UNCHECKED
			.endif
			invoke CheckMenuItem, hPopupMenu, IDM_EXPLORE, ecx
if 1
			invoke ListView_GetNextItem( m_hWndLV, -1, LVIS_SELECTED)
			mov ecx, eax
			mov edx, m_pVariants
			shl ecx, 4
			add ecx, edx
			.if (([ecx].VARIANT.vt == VT_DISPATCH) || ([ecx].VARIANT.vt == VT_UNKNOWN))
				mov ecx, MF_ENABLED
			.else
				mov ecx, MF_GRAYED
			.endif
			push ecx
			invoke EnableMenuItem, hPopupMenu, IDOK, ecx
			pop ecx
			invoke EnableMenuItem, hPopupMenu, IDM_EXPLORE, ecx
endif

			invoke SetMenuDefaultItem, hPopupMenu, IDOK, FALSE
			invoke GetItemPosition, m_hWndLV, bMouse, addr pt
			invoke TrackPopupMenu, hPopupMenu, TPM_LEFTALIGN or TPM_LEFTBUTTON,
					pt.x,pt.y,0,m_hWnd,NULL
		.endif
	.endif
	ret
	align 4

ShowContextMenu endp

;--- get a typeinfo of an item

ShowContextMenuHdr proc uses esi edi pNMHdr:ptr NMHDR

local dwIndex:DWORD
local pt:POINT
local pTypeInfo:LPTYPEINFO
local pTypeAttr:ptr TYPEATTR
local pFuncDesc:ptr FUNCDESC
local bstr:BSTR
local lvc:LVCOLUMN
local szText[64]:byte

	mov pTypeInfo, NULL
	mov pTypeAttr, NULL
	invoke vf(m_pDispItem, IDispatch, GetTypeInfo), 0, g_LCID, addr pTypeInfo
	.if (eax == S_OK)
		invoke vf(pTypeInfo, ITypeInfo, GetTypeAttr), addr pTypeAttr
	.endif

	.if (pTypeAttr)

		invoke CreatePopupMenu
		mov esi, eax

		mov edi, pTypeAttr
		movzx ecx, [edi].TYPEATTR.cFuncs
		mov dwIndex, 0
		.while (ecx)
			push ecx
			invoke vf(pTypeInfo, ITypeInfo, GetFuncDesc), dwIndex, addr pFuncDesc
			.if (eax == S_OK)
				mov edi, pFuncDesc
				.if ([edi].FUNCDESC.invkind == INVOKE_PROPERTYGET)
					invoke vf(pTypeInfo, ITypeInfo, GetDocumentation), [edi].FUNCDESC.memid, addr bstr, NULL, NULL, NULL
					.if (eax == S_OK)
						invoke WideCharToMultiByte, CP_ACP, 0, bstr, -1, addr szText, sizeof szText, 0, 0
						invoke SysFreeString, bstr
					.else
						invoke wsprintf, addr szText, CStr("MemId %X"), [edi].FUNCDESC.memid
					.endif
					mov ecx, MF_STRING
					mov edx, [edi].FUNCDESC.memid
					.if (edx == m_dwDispId)
						or ecx, MF_CHECKED
					.endif
					mov edx, dwIndex
					inc edx
					invoke AppendMenu, esi, ecx, edx, addr szText
				.endif
				invoke vf(pTypeInfo, ITypeInfo, ReleaseFuncDesc), edi
			.endif
			pop ecx
			inc dwIndex
			dec ecx
		.endw

		invoke GetCursorPos, addr pt
		invoke TrackPopupMenu, esi, TPM_LEFTALIGN or TPM_LEFTBUTTON or TPM_RETURNCMD,
			pt.x, pt.y, 0, m_hWnd, NULL
		.if (eax)
			mov ecx, eax
			dec ecx
			invoke vf(pTypeInfo, ITypeInfo, GetFuncDesc), ecx, addr pFuncDesc
			.if (eax == S_OK)
				mov edi, pFuncDesc
				mov eax, [edi].FUNCDESC.memid
				.if (eax != m_dwDispId)
					mov m_dwDispId, eax
					invoke vf(pTypeInfo, ITypeInfo, GetDocumentation), [edi].FUNCDESC.memid, addr bstr, NULL, NULL, NULL
					.if (eax == S_OK)
						invoke WideCharToMultiByte, CP_ACP, 0, bstr, -1, addr szText, sizeof szText, 0, 0
						invoke SysFreeString, bstr
						mov lvc.mask_, LVCF_TEXT
						lea eax, szText
						mov lvc.pszText,eax
						invoke ListView_SetColumn( m_hWndLV, VALUE_COLUMN, addr lvc)
					.endif
					invoke PostMessage, m_hWnd, WM_COMMAND, IDC_REFRESH, 0
				.endif
				invoke vf(pTypeInfo, ITypeInfo, ReleaseFuncDesc), edi
			.endif
		.endif

		invoke vf(pTypeInfo, ITypeInfo, ReleaseTypeAttr), pTypeAttr
		invoke vf(pTypeInfo, ITypeInfo, Release)

		invoke DestroyMenu, esi

	.endif
	ret
	align 4

ShowContextMenuHdr endp


if ?TYPEFROMTYPEINFO

GetDispatchType proc pDispatch:LPDISPATCH, pszTextOut:LPSTR, dwMax:DWORD

local dwRC:DWORD
local pTypeInfo:LPTYPEINFO
local bstr:BSTR

	mov dwRC, 0
	invoke vf(pDispatch, IDispatch, GetTypeInfo), 0, g_LCID, addr pTypeInfo
	.if (eax == S_OK)
		invoke vf(pTypeInfo, ITypeInfo, GetDocumentation), MEMBERID_NIL, addr bstr, NULL, NULL, NULL
		.if (eax == S_OK)
			invoke WideCharToMultiByte, CP_ACP, 0, bstr, -1, pszTextOut, dwMax, 0, 0
			mov eax, pszTextOut
			mov dwRC, eax
			invoke SysFreeString, bstr
		.endif
		invoke vf(pTypeInfo, ITypeInfo, Release)
	.endif
	return dwRC
	align 4

GetDispatchType endp

IsNumeric proc
	.const
NumericVariantTypes label WORD
	dw VT_I2, VT_I4, VT_R4, VT_R8, VT_I1, VT_UI1, VT_UI2, VT_UI4, VT_I8, VT_UI8, VT_INT, VT_UINT
NUMVARTYPES equ ($ - NumericVariantTypes) / sizeof WORD
	.code
	mov edx, edi
	mov edi, offset NumericVariantTypes
	mov ecx, NUMVARTYPES
	repnz scasw
	mov edi, edx
	ret
	align 4
IsNumeric endp

GetDispatchProperty proc pDispatch:LPDISPATCH, pszTextOut:LPSTR, dwMax:DWORD, pdwNumCount:ptr DWORD

local dwDispId:DWORD
local pwszNames:ptr WORD
local pTypeInfo:LPTYPEINFO
local vtOrg:WORD
local dispparams:DISPPARAMS
local varResult:VARIANT
local varResult2:VARIANT

	xor eax, eax
	mov dispparams.rgvarg,eax
	mov dispparams.rgdispidNamedArgs,eax
	mov dispparams.cArgs,eax
	mov dispparams.cNamedArgs,eax
	invoke VariantInit, addr varResult
	invoke vf(pDispatch, IDispatch, Invoke_), m_dwDispId, addr IID_NULL,
			g_LCID, DISPATCH_PROPERTYGET, addr dispparams, addr varResult, NULL, NULL
	.if (eax == S_OK)
		movzx eax, varResult.vt
		mov vtOrg, ax
		invoke IsNumeric
		.if (ZERO?)
			mov ecx, pdwNumCount
			inc dword ptr [ecx]
		.endif
		invoke VariantChangeType, addr varResult, addr varResult, 0, VT_BSTR
		.if (eax == S_OK)
			invoke WideCharToMultiByte, CP_ACP, 0, varResult.bstrVal, -1, pszTextOut, dwMax, 0, 0
			.if (vtOrg == VT_BOOL)
				mov ecx, pszTextOut
				.if (word ptr [ecx] == "0")
					invoke lstrcpy, ecx, CStr("False")
				.else
					invoke lstrcpy, ecx, CStr("True")
				.endif
			.endif
		.else
			movzx eax, varResult.vt
			.if ((eax == VT_DISPATCH) && varResult.pdispVal)
				mov pwszNames, CStrW(L("_NewEnum"))
				invoke vf(varResult.pdispVal, IDispatch, GetIDsOfNames), addr IID_NULL,
					addr pwszNames, 1, g_LCID, addr dwDispId
				.if ((eax == S_OK) && (dwDispId == DISPID_NEWENUM))
					mov pwszNames, CStrW(L("Count"))
					invoke vf(varResult.pdispVal, IDispatch, GetIDsOfNames), addr IID_NULL,
						addr pwszNames, 1, g_LCID, addr dwDispId
					.if (eax == S_OK)
						invoke VariantInit, addr varResult2
						invoke vf(varResult.pdispVal, IDispatch, Invoke_), dwDispId, addr IID_NULL,
							g_LCID, DISPATCH_PROPERTYGET, addr dispparams, addr varResult2, NULL, NULL
						.if (eax == S_OK)
							invoke VariantChangeType, addr varResult2, addr varResult2, 0, VT_BSTR
							.if (eax == S_OK)
								invoke WideCharToMultiByte, CP_ACP, 0, varResult2.bstrVal, -1, pszTextOut, dwMax, 0, 0
							.endif
							invoke VariantClear, addr varResult2
						.endif
					.endif
				.endif
			.endif
		.endif
		invoke VariantClear, addr varResult
	.endif
	ret
	align 4
GetDispatchProperty endp

endif

;--- RefreshList will set m_pVariants + m_iNumVariants

RefreshList proc uses esi

local bRC:BOOL
local varItem:VARIANT
local pCurVar:ptr VARIANT
local dwESP:DWORD
local dwSize:DWORD
local dwOrgVt:DWORD
local rect:RECT
local hCsrOld:HCURSOR
local dwNumCnt:DWORD
local lvi:LVITEM
local lvc:LVCOLUMN
local szText[128]:byte

	invoke SetCursor, g_hCsrWait
	mov hCsrOld, eax

	invoke ClearArray
	invoke ListView_DeleteAllItems( m_hWndLV)
	mov dwESP, esp

	.if (m_pDispItem)
		invoke vf(m_pDispItem, IUnknown, Release)
		mov m_pDispItem, NULL
	.endif

	@mov dwNumCnt, 0
	@mov lvi.iItem, 0
	lea eax, szText
	mov lvi.pszText, eax
	xor esi, esi

;--- get variants, put them on stack

	.while (esi < g_MaxCollItems)
		sub esp, sizeof VARIANT
		invoke VariantInit, esp
		mov ecx, esp
		.if (m_pEnumVariant)
			invoke vf(m_pEnumVariant, IEnumVARIANT, Next), 1, ecx, NULL
		.else
			.if (esi < m_dwNumVars)
				mov ecx, m_variant.parray
				mov eax, esi
				shl eax, 4
				add eax, [ecx].SAFEARRAY.pvData
				mov edx, esp
				invoke VariantCopy, edx, eax
				mov eax, S_OK
			.else
				mov eax, S_FALSE
			.endif
		.endif
		.break .if (eax != S_OK)
		.if (m_bUnknown)
			mov eax, VT_UNKNOWN
			xchg eax, [esp]
			mov [esp].VARIANT.punkVal, eax
		.endif
		inc esi

		@mov lvi.iSubItem, 0
		lea eax, szText
		mov lvi.pszText, eax
		.if (m_pEnumVariant)
			invoke wsprintf, addr szText,CStr("%u"), esi
		.else
			push edi
			push esi
			push ebx
			mov ecx, m_variant.parray
			movzx edi, [ecx].SAFEARRAY.cDims
			dec esi
			mov szText, 0
			.while (edi)
				mov edx, m_variant.parray
				lea edx, [edx].SAFEARRAY.rgsabound
				lea ecx, [edi-1]
				lea ecx, [edx+ecx*8]
				mov eax, esi
				cdq
				push [ecx].SAFEARRAYBOUND.lLbound
				mov ecx,[ecx].SAFEARRAYBOUND.cElements
				div ecx
				add [esp], edx
				mov esi, eax
				dec edi
			.endw
			mov ecx, m_variant.parray
			movzx esi, [ecx].SAFEARRAY.cDims
			.if (esi == 1)
				mov ebx, CStr("%u")
			.else
				mov m_bCol1Alpha, TRUE
				mov ebx, CStr("[%u]")
			.endif
			lea edi, szText
			.while (esi)
				pop eax
				invoke wsprintf, edi, ebx, eax
				add edi, eax
				dec esi
			.endw
			pop ebx
			pop esi
			pop edi
		.endif

;------------------------- ok, variant on the stack now
		mov eax, lvi.iItem
		mov lvi.lParam, eax
		mov lvi.mask_, LVIF_TEXT or LVIF_PARAM
		invoke ListView_InsertItem( m_hWndLV, addr lvi)

		.if ((!m_pDispItem) && ([ESP].VARIANT.vt == VT_DISPATCH))
			mov edx, [esp].VARIANT.pdispVal
			mov m_pDispItem, edx
			invoke vf(m_pDispItem, IUnknown, AddRef)
		.endif
;------------------------- copy variant to varItem, change it to BSTR

		mov szText, 0
		invoke VariantInit, addr varItem
		movzx eax, [ESP].VARIANT.vt
		mov dwOrgVt, eax
		.if ((eax == VT_DISPATCH) && m_dwDispId)
			mov edx, [esp].VARIANT.pdispVal
			.if (edx)
				invoke GetDispatchProperty, edx, addr szText, sizeof szText, addr dwNumCnt
			.endif
		.else
			movzx eax, [esp].VARIANT.vt
			invoke IsNumeric
			.if (ZERO?)
				inc dwNumCnt
			.endif
			mov ecx, esp
			invoke VariantChangeType, addr varItem, ecx, 0, VT_BSTR
			.if (eax == S_OK)
				invoke WideCharToMultiByte, CP_ACP, 0, varItem.bstrVal, -1, addr szText, sizeof szText, 0, 0
			.endif
		.endif
		mov lvi.mask_, LVIF_TEXT
		mov lvi.iSubItem, 1
		lea eax, szText
		mov lvi.pszText, eax
		invoke ListView_SetItem( m_hWndLV, addr lvi)
		inc lvi.iSubItem

;------------------------- get type of item

if ?TYPEFROMTYPEINFO
		xor eax, eax
		.if (dwOrgVt == VT_DISPATCH)
			mov edx, [esp].VARIANT.pdispVal
			invoke GetDispatchType, edx, addr szText, sizeof szText
		.endif
		.if (!eax)
			invoke GetVarType, dwOrgVt
		.endif
else
		invoke GetVarType, dwOrgVt
endif
		mov lvi.pszText, eax
		invoke ListView_SetItem( m_hWndLV, addr lvi)
		invoke VariantClear, addr varItem
		inc lvi.iItem
	.endw
	add esp, sizeof VARIANT
;------------------------- if all items are numeric, set bNumeric member
;------------------------- and listview column format
	mov al, FALSE
	.if (esi == dwNumCnt)
		mov al, TRUE
	.endif
	.if (al != m_bNumeric)
		mov m_bNumeric, al
		.if (al)
			mov ecx, LVCFMT_RIGHT
		.else
			mov ecx, LVCFMT_LEFT
		.endif
		mov lvc.fmt, ecx
		mov lvc.mask_, LVCF_FMT
		invoke ListView_SetColumn( m_hWndLV, VALUE_COLUMN, addr lvc)
	.endif
;------------------------- ok, we have the variants onto the stack
;------------------------- esp -> start of array
;------------------------- esi -> # of variants
;------------------------- now copy array to a heap object

	.if (esi == 0)
		mov bRC, esi
	.else
		mov m_iNumVariants, esi
		mov eax, esi
		mov ecx, sizeof VARIANT
		mul ecx
		mov dwSize, eax
		invoke malloc, eax
		.if (eax)
			mov m_pVariants, eax

			pushad
			mov ecx, esi
			mov esi, dwESP
			sub esi, sizeof VARIANT
			mov edi, eax
			.while (ecx)
				movsd
				movsd
				movsd
				movsd
				sub esi, 2 * sizeof VARIANT
				dec ecx
			.endw
			popad
			mov bRC, TRUE
		.else
;--------------------------------- we couldnt get a heap object, so clear
;--------------------------------- all 
			mov edx, esp
			.while (esi)
				push edx
				invoke VariantClear, edx
				pop edx
				add edx, sizeof VARIANT
				dec esi
			.endw
			mov bRC, FALSE
		.endif
	.endif
	mov esp, dwESP
	.if (m_iSortCol != -1)
		invoke DoSort
	.endif
	invoke SetCursor, hCsrOld
	return bRC
	align 4

RefreshList endp

;--- explore mode is on, item has changed, so update properties dialog

UpdateExploreView proc iItem:DWORD

local pCurVar:ptr VARIANT
local varItem:VARIANT
local rect:RECT
local pDispatch:LPDISPATCH
local pPropertiesDlg:ptr CPropertiesDlg
local lvi:LVITEM

	invoke VariantInit, addr varItem
	mov eax, m_pVariants
	mov ecx, iItem
	shl ecx, 4
	add ecx, eax
	mov pCurVar, ecx
	invoke VariantCopy, addr varItem, ecx
	.if (varItem.vt == VT_DISPATCH)
dodispatch:
		mov pPropertiesDlg, NULL
		mov eax, m_pObjectItem
		.if (eax)
			invoke vf(eax, IObjectItem, GetPropDlg)
			mov pPropertiesDlg, eax
			invoke vf(m_pObjectItem, IObjectItem, Release) 
			mov m_pObjectItem, NULL
		.endif
		.if (pPropertiesDlg)
if ?POSTSETDISP
			invoke vf(varItem.pdispVal, IUnknown, AddRef)
endif
			invoke SetDispatch@CPropertiesDlg, pPropertiesDlg, varItem.pdispVal, FALSE
			.if (!eax)
				mov pPropertiesDlg, NULL
			.endif
		.else
			invoke Create2@CPropertiesDlg, varItem.punkVal, NULL
			.if (eax)
				mov pPropertiesDlg, eax
				invoke GetWindowRect, m_hWnd, addr rect
				mov eax, rect.right
				sub eax, rect.left
				.if (eax > rect.left)
					mov eax, rect.right
					mov rect.left, eax
				.else
					sub rect.left, eax
				.endif
				invoke SetWindowPos@CPropertiesDlg, pPropertiesDlg, addr rect
				invoke Show@CPropertiesDlg, pPropertiesDlg, m_hWnd
			.endif
		.endif
		.if (pPropertiesDlg)
			invoke GetObjectItem@CPropertiesDlg, pPropertiesDlg
			mov m_pObjectItem, eax
			invoke vf(eax, IObjectItem, AddRef)
		.else
			invoke VariantClear, pCurVar
			mov lvi.mask_, LVIF_TEXT
			mov ecx, iItem
			mov lvi.iItem, ecx
			@mov lvi.iSubItem, 0
			mov lvi.pszText, CStr("***")
			invoke ListView_SetItem( m_hWndLV, addr lvi)
		.endif
	.elseif (varItem.vt == VT_UNKNOWN)
		invoke vf(varItem.punkVal, IUnknown, QueryInterface), addr IID_IDispatch, addr pDispatch
		.if (eax == S_OK)
			invoke vf(pDispatch, IUnknown, Release)
			jmp dodispatch
		.else
			invoke MessageBeep, MB_OK
		.endif
	.endif
	invoke VariantClear, addr varItem
	ret
	align 4

UpdateExploreView endp


OnBeginLabelEdit proc uses edi pNMLVDI:ptr NMLVDISPINFO

local dwRC:DWORD

	mov dwRC, FALSE
	return dwRC
	align 4

OnBeginLabelEdit endp


;--- WM_NOTIFY for listview


OnNotifyLV proc  pNMLV:ptr NMLISTVIEW

local dwRC:DWORD
local varItemNew:VARIANT
local varItemOld:VARIANT
local hti:LVHITTESTINFO
local lvi:LVITEM

	mov dwRC, FALSE
	mov eax, pNMLV
	.if ([eax].NMLISTVIEW.hdr.code == NM_RCLICK)

		invoke ShowContextMenu, TRUE

	.elseif ([eax].NMLISTVIEW.hdr.code == NM_DBLCLK)

		invoke PostMessage, m_hWnd, WM_COMMAND, IDOK, 0

	.elseif ([eax].NMLISTVIEW.hdr.code == NM_CLICK)

		.if (m_bInlineEdit)
			invoke GetCursorPos, addr hti.pt
			invoke ScreenToClient, m_hWndLV, addr hti.pt
			invoke ListView_SubItemHitTest( m_hWndLV, addr hti)
			.if (hti.iSubItem == 1)
				invoke ListView_GetItemCount( m_hWndLV)
				.if (eax > hti.iItem)
					invoke SendMessage, m_hWndLV, LVM_EDITLABEL, hti.iItem, hti.iSubItem
				.endif
			.endif
		.endif

	.elseif ([eax].NMLISTVIEW.hdr.code == LVN_BEGINLABELEDIT)

		invoke OnBeginLabelEdit, eax
		mov dwRC, eax

	.elseif ([eax].NMLISTVIEW.hdr.code == LVN_ENDLABELEDIT)

		.if ([eax].NMLVDISPINFO.item.pszText)
			push esi
			mov esi, eax
			invoke VariantInit, addr varItemNew
			invoke VariantInit, addr varItemOld
			invoke SysStringFromLPSTR, [esi].NMLVDISPINFO.item.pszText, 0
			mov varItemNew.bstrVal, eax
			mov varItemNew.vt, VT_BSTR
			mov lvi.mask_, LVIF_PARAM
			mov eax, [esi].NMLVDISPINFO.item.iItem
			mov lvi.iItem, eax
			mov lvi.iSubItem, 0
			invoke ListView_GetItem( m_hWndLV, addr lvi)
			mov edx, lvi.lParam
			shl edx, 4
			mov ecx, m_variant.parray
			add edx, [ecx].SAFEARRAY.pvData
			mov esi, edx

			invoke VariantChangeType, addr varItemOld, esi, 0, VT_BSTR
			invoke _strcmpW, varItemOld.bstrVal, varItemNew.bstrVal
			.if (eax)
				movzx ecx, [esi].VARIANT.vt
				.if (ecx == VT_EMPTY)
					mov ecx, VT_BSTR
				.endif
				invoke VariantChangeType, esi, addr varItemNew, 0, ecx
				mov m_bDirty, TRUE
			.endif
			pop esi
			invoke VariantClear, addr varItemNew
			invoke VariantClear, addr varItemOld
			invoke SetWindowLong, m_hWnd, DWL_MSGRESULT, 1
			mov dwRC, TRUE
		.endif

	.elseif ([eax].NMLISTVIEW.hdr.code == LVN_ITEMCHANGED)
	
		mov ecx, [eax].NMLISTVIEW.iItem
		mov edx, m_pVariants
		shl ecx, 4
		add ecx, edx
		.if (([ecx].VARIANT.vt == VT_DISPATCH) || ([ecx].VARIANT.vt == VT_UNKNOWN))
			.if (m_bExplore && ([eax].NMLISTVIEW.uNewState & LVIS_SELECTED))
;;				invoke UpdateExploreView, [eax].NMLISTVIEW.iItem
				invoke UpdateExploreView, [eax].NMLISTVIEW.lParam
			.endif
		.endif
	
	.elseif ([eax].NMLISTVIEW.hdr.code == LVN_KEYDOWN)

		.if ([eax].NMLVKEYDOWN.wVKey == VK_APPS)

			invoke ShowContextMenu, FALSE

		.elseif (([eax].NMLVKEYDOWN.wVKey == VK_F6) && (m_pEnumVariant))

			invoke Create@CObjectItem, m_pEnumVariant, NULL
			.if (eax)
				push eax
				invoke vf(eax, IObjectItem, ShowObjectDlg), m_hWnd
				pop eax
				invoke vf(eax, IObjectItem, Release)
			.endif

		.endif

	.elseif ([eax].NMLISTVIEW.hdr.code == LVN_COLUMNCLICK)

		mov eax,[eax].NMLISTVIEW.iSubItem
		.if (eax == m_iSortCol)
			xor m_iSortDir,1
		.else
			mov m_iSortCol,eax
			mov m_iSortDir,0
		.endif
		invoke DoSort

	.endif
	return dwRC
	align 4

OnNotifyLV endp

;--- WM_NOTIFY

OnNotify proc uses esi pNMHdr:ptr NMHDR

	mov esi, pNMHdr
;;	DebugOut "WM_NOTIFY, code=%d, idFrom=%d", [esi].NMHDR.code, [esi].NMHDR.idFrom
	xor eax, eax
	.if ([esi].NMHDR.idFrom == IDC_LIST1)
		invoke OnNotifyLV, pNMHdr
	.else
		invoke ListView_GetHeader( m_hWndLV)
		.if (eax == [esi].NMHDR.hwndFrom)
			.if (([esi].NMHDR.code == NM_RCLICK) && m_pDispItem)
				invoke ShowContextMenuHdr, esi
			.endif
		.endif
	.endif
	ret
	align 4

OnNotify endp

;--- HResult error occurred, set statusline 

PrepareInvokeErrorReturn proc uses esi HResult:DWORD, pExcepInfo:ptr EXCEPINFO, dwArgErr:DWORD

local	dwESP:DWORD
local	pszHResult:LPSTR
local	pszFlags:LPSTR
local	szText[256]:byte

	mov pszFlags, CStr("PropPut")
	.if (HResult == DISP_E_EXCEPTION)
		mov dwESP, esp
		mov esi, pExcepInfo
		assume esi:ptr EXCEPINFO
		.if ([esi].bstrDescription)
			invoke SysStringLen, [esi].bstrDescription
			add eax, 4
			and al, 0FCh
			sub esp, eax
			mov edx, esp
			invoke WideCharToMultiByte, CP_ACP, 0, [esi].bstrDescription, -1, edx, eax, 0, 0
			mov ecx, esp
		.else
			mov ecx,CStr("")
		.endif
		invoke wsprintf, addr szText, CStr("'%.192s'[%X] Exception at Invoke(%s)"), ecx, [esi].scode, pszFlags
		mov esp, dwESP
		invoke SysFreeString, [esi].bstrSource
		invoke SysFreeString, [esi].bstrDescription
		invoke SysFreeString, [esi].bstrHelpFile
	.else
		mov eax, HResult
		.if ((eax == DISP_E_TYPEMISMATCH) || (eax == DISP_E_PARAMNOTFOUND))
			invoke wsprintf, addr szText, CStr("[%X] Error at Invoke(%s) [uArgErr=%u]"), HResult, pszFlags, dwArgErr
		.else
			invoke wsprintf, addr szText, CStr("[%X] Error at Invoke(%s)"), HResult, pszFlags
		.endif
	.endif
	invoke MessageBox, m_hWnd, addr szText, 0, MB_OK
	ret
	assume esi:nothing
	align 4

PrepareInvokeErrorReturn endp


;--- rewrite array


PutProperty proc

local	ExcepInfo:EXCEPINFO
local	dwArgErr:DWORD
local	dispid:DWORD
local	DispParams:DISPPARAMS
;local	varResult:VARIANT

	mov ExcepInfo.bstrSource, NULL
	mov ExcepInfo.bstrDescription, NULL
	mov ExcepInfo.bstrHelpFile, NULL

	lea edx, m_variant
	mov DispParams.rgvarg, edx
	mov dispid, DISPID_PROPERTYPUT
	lea eax, dispid
	mov DispParams.rgdispidNamedArgs, eax
	mov DispParams.cNamedArgs, 1
	mov DispParams.cArgs, 1

;	invoke VariantInit, addr varResult
	invoke vf(m_pDispatch, IDispatch, Invoke_), m_memid, addr IID_NULL,
			g_LCID, DISPATCH_PROPERTYPUT, addr DispParams, NULL,
			addr ExcepInfo, addr dwArgErr
	.if (eax == S_OK)
		mov m_bDirty, FALSE
	.else
		lea ecx, ExcepInfo
		invoke PrepareInvokeErrorReturn, eax, ecx, dwArgErr
	.endif
	ret
	align 4
PutProperty endp


;--- WM_COMMAND processing


OnCommand proc wParam:WPARAM, lParam:LPARAM

local	pVariant:ptr VARIANT
local	rect:RECT
local	pPropertiesDlg:ptr CPropertiesDlg
local	lvi:LVITEM

	movzx eax, word ptr wParam+0
	.if (eax == IDCANCEL)

		.if (m_bDirty)
			invoke MessageBox, m_hWnd, CStr("Throw away changes?"), addr g_szWarning, MB_YESNO
			.if (eax == IDNO)
				ret
			.endif
		.endif
		invoke PostMessage, m_hWnd, WM_CLOSE, 0, 0

	.elseif (eax == IDOK)

		invoke ListView_GetNextItem( m_hWndLV, -1, LVNI_SELECTED)
		.if (eax != -1)
			mov lvi.iItem, eax
			mov lvi.iSubItem, 0
			mov lvi.mask_, LVIF_PARAM
			invoke ListView_GetItem( m_hWndLV, addr lvi)
			mov eax, lvi.lParam
			mov ecx, eax
			shl eax, 4				;eax * sizeof VARIANT
			add eax, m_pVariants
			mov pVariant, eax
;;			.if ([eax].VARIANT.vt != VT_EMPTY)
			.if (([eax].VARIANT.vt == VT_DISPATCH) || \
				([eax].VARIANT.vt == VT_UNKNOWN))
ife ?MODELESS
				mov eax, m_pParamReturn
				mov [eax].PARAMRETURN.iCurItem, ecx
				xor ecx, ecx
				xchg ecx, m_pVariants
				mov [eax].PARAMRETURN.pVariants, ecx
				xor ecx, ecx
				xchg ecx, m_iNumVariants
				mov [eax].PARAMRETURN.iNumVariants, ecx
else
				.if ([eax].VARIANT.vt == VT_DISPATCH)
					invoke Create2@CPropertiesDlg, [eax].VARIANT.pdispVal, NULL
					.if (eax)
						mov pPropertiesDlg, eax
						invoke GetWindowRect, m_hWnd, addr rect
						.if (!g_bCloseCollDlgOnDlbClk)
							add rect.left, 20
							add rect.top, 20
						.endif
						invoke SetWindowPos@CPropertiesDlg, pPropertiesDlg, addr rect
						invoke Show@CPropertiesDlg, pPropertiesDlg, NULL
						.if (g_bCloseCollDlgOnDlbClk)
							invoke PostMessage, m_hWnd, WM_CLOSE, 0, 0
						.endif
						jmp done
					.endif
				.endif
				mov eax, pVariant
				invoke Create@CObjectItem, [eax].VARIANT.punkVal, NULL
				.if (eax)
					push eax
					invoke vf(eax, IObjectItem, ShowObjectDlg), NULL
					pop eax
					invoke vf(eax, IObjectItem, Release)
					.if (g_bCloseCollDlgOnDlbClk)
						invoke PostMessage, m_hWnd, WM_CLOSE, 0, 0
					.endif
				.endif
endif
			.else
				invoke MessageBeep, MB_OK
			.endif
		.endif

	.elseif (eax == IDM_EXPLORE)

		mov eax, m_pObjectItem
		.if (eax)
			invoke vf(eax, IObjectItem, GetPropDlg)
		.endif
		mov ecx, eax
		xor m_bExplore, 1
		mov al, m_bExplore
		mov g_bExplore, al
		.if (!al && ecx)
			invoke PostMessage, [ecx].CDlg.hWnd, WM_CLOSE, 0, 0
		.elseif (al && (!ecx))
			invoke ListView_GetNextItem( m_hWndLV, -1, LVNI_SELECTED)
			.if (eax != -1)
				mov lvi.iItem, eax
				@mov lvi.iSubItem, 0
				mov lvi.mask_, LVIF_PARAM
				invoke ListView_GetItem( m_hWndLV, addr lvi)
				invoke UpdateExploreView, lvi.lParam
			.endif
		.endif

	.elseif (eax == IDC_REFRESH)

		.if (m_pEnumVariant)
			invoke vf(m_pEnumVariant, IEnumVARIANT, Reset)
		.endif
		invoke RefreshList

	.elseif (eax == IDC_PUTPROP)

		invoke PutProperty

	.endif
done:
	ret
	align 4

OnCommand endp

;--- WM_SIZE

	.const
BtnTab dd IDC_REFRESH, IDC_PUTPROP, IDCANCEL
NUMBUTTONS textequ %($ - BtnTab) / sizeof DWORD
	.code


OnSize proc uses esi edi dwType:dword, dwWidth:dword, dwHeight:dword

local dwRim:DWORD
local dwHeightBtn:DWORD
local dwWidthBtn:DWORD
local dwXPos:DWORD
local dwYPos:DWORD
local dwAddX:DWORD
local dwHeightSB:DWORD
local dwHeightLV:DWORD
local rect:RECT

;	invoke GetWindowRect, m_hWndSB, addr rect
;	mov eax, rect.bottom
;	sub eax, rect.top
;	mov dwHeightSB, eax
	mov dwHeightSB, 0

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

;;	invoke BeginDeferWindowPos, 2 + NUMBUTTONS
	invoke BeginDeferWindowPos, 1 + NUMBUTTONS
	mov edi, eax

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
	invoke DeferWindowPos, edi, m_hWndLV, NULL, 0, 0, dwWidth, dwHeightLV, SWP_NOMOVE or SWP_NOZORDER or SWP_NOACTIVATE

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
		invoke DeferWindowPos, edi, eax, NULL, dwXPos, dwYPos, 0, 0, SWP_NOSIZE or SWP_NOZORDER or SWP_NOACTIVATE
		mov eax, rect.right
		sub eax, rect.left
		add eax, dwAddX
		add dwXPos, eax
		pop ecx
		dec ecx
	.endw

;;	invoke DeferWindowPos, edi, m_hWndSB, NULL, 0, 0, 0, 0, SWP_NOZORDER or SWP_NOACTIVATE

	invoke EndDeferWindowPos, edi

	ret
	align 4
OnSize endp

;--- WM_INITDIALOG

OnInitDialog proc

local bstr:BSTR
local pTypeInfo:LPTYPEINFO
local rect:RECT
local szText[128]:byte

	invoke GetDlgItem, m_hWnd, IDC_LIST1
	mov m_hWndLV, eax

	.if (m_hWndFrom)
		invoke GetWindowRect, m_hWndFrom, addr rect
		mov ecx, rect.left
		mov eax, rect.top
		add ecx, 20
		add eax, 20
		mov g_rect.left, ecx
		mov g_rect.top, eax
	.endif
	invoke MySetWindowPos, m_hWnd, addr g_rect

	invoke SetLVColumns, m_hWndLV, NUMCOLS_COLLECTIONS, addr ColumnsCollections

	invoke CreateEditListView, m_hWndLV

	invoke RefreshList
	.if (!eax)
		invoke PostMessage, m_hWnd, WM_CLOSE, 0, 0
		jmp done
	.endif
	.if ((!m_pEnumVariant))
		invoke SetWindowText, m_hWnd, CStr("Array of Variants")
		mov ecx, m_pVariants
		movzx ecx,[ecx].VARIANT.vt
		.if ((ecx != VT_UNKNOWN) &&  (ecx != VT_DISPATCH) && (ecx != VT_PTR))
			invoke ListView_SetExtendedListViewStyle( m_hWndLV, LVS_EX_GRIDLINES or LVS_EX_INFOTIP)
			mov m_bInlineEdit, TRUE
		.else
			invoke ListView_SetExtendedListViewStyle( m_hWndLV,LVS_EX_FULLROWSELECT or LVS_EX_INFOTIP)
		.endif
	.else
		.if (m_pDispatch)
			mov word ptr szText," :"
			invoke GetDispatchType, m_pDispatch, addr szText+2, sizeof szText-2
			.if (eax)
				sub esp, 256
				mov edx, esp
				invoke GetWindowText, m_hWnd, edx, 256
				mov edx, esp
				invoke lstrcat, edx, addr szText
				invoke SetWindowText, m_hWnd, esp
				add esp, 256
			.endif
		.endif
		invoke ListView_SetExtendedListViewStyle( m_hWndLV,LVS_EX_FULLROWSELECT or LVS_EX_INFOTIP)
	.endif
	.if (!m_bInlineEdit)
		invoke GetDlgItem, m_hWnd, IDC_PUTPROP
		invoke EnableWindow, eax, FALSE
	.endif
if ?MODELESS
	invoke ShowWindow, m_hWnd, SW_SHOWNORMAL
	mov eax, m_pParamReturn
	mov [eax].PARAMRETURN.iNumVariants, 1
endif
done:
	ret
	align 4

OnInitDialog endp


;--- enum a collection in a simple dialog


CCollectionDialog proc uses esi __this thisarg, message:DWORD, wParam:WPARAM, lParam:LPARAM

local	rect:RECT

	mov __this,this@

	mov eax, message
	.if (eax == WM_INITDIALOG)

		invoke OnInitDialog
		mov eax, 1

	.elseif (eax == WM_CLOSE)

		mov ecx, m_pObjectItem
		.if (ecx)
			invoke vf(ecx, IObjectItem, GetPropDlg)
			.if (eax)
				invoke PostMessage, [eax].CDlg.hWnd, WM_CLOSE, 0, 0
			.endif
		.endif
		invoke SaveNormalWindowPos, m_hWnd, addr g_rect
if ?MODELESS
		invoke DestroyWindow, m_hWnd
else
		ListView_GetItemCount m_hWndLV
		invoke EndDialog, m_hWnd, eax
endif
	.elseif (eax == WM_DESTROY)

		invoke Destroy@CCollectionDlg, __this

	.elseif (eax == WM_COMMAND)

		invoke OnCommand, wParam, lParam

	.elseif (eax == WM_NOTIFY)

		invoke OnNotify, lParam

	.elseif (eax == WM_SIZE)

		.if (wParam != SIZE_MINIMIZED)
			movzx eax, word ptr lParam+0
			movzx ecx, word ptr lParam+2
			invoke OnSize, wParam, eax, ecx
		.endif
if ?HTMLHELP
	.elseif (eax == WM_HELP)

		invoke DoHtmlHelp, HH_DISPLAY_TOPIC, CStr("CollectionDialog.htm")
endif
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

CCollectionDialog endp


;--- constructor


Create@CCollectionDlg proc public uses esi __this hWndFrom:HWND, pDispatch:LPDISPATCH, memid:MEMBERID, pVariant:ptr VARIANT, pParamReturn:ptr PARAMRETURN

	invoke malloc, sizeof CCollectionDlg
	.if (!eax)
		ret
	.endif

	mov __this,eax
	mov m_pDlgProc, CCollectionDialog

	mov m_iSortCol, -1

;--------------------------- we need pDispatch+memid in case of updates
	mov eax, pDispatch
	mov m_pDispatch, eax
	.if (eax)
		invoke vf(eax, IUnknown, AddRef)
	.endif
	mov eax, memid
	mov m_memid, eax

	mov eax, pParamReturn
	mov m_pParamReturn, eax
	mov [eax].PARAMRETURN.iNumVariants, 0
	mov [eax].PARAMRETURN.pVariants, NULL
	mov [eax].PARAMRETURN.iCurItem, -1

	mov m_iItem, -1
	mov al, g_bExplore
	mov m_bExplore, al

	invoke VariantInit, addr m_variant

	mov eax, pVariant
	mov esi, [eax].VARIANT.punkVal
;;	.if ([eax].VARIANT.vt == VT_UNKNOWN)
	.if (([eax].VARIANT.vt == VT_UNKNOWN) || ([eax].VARIANT.vt == VT_DISPATCH))
		invoke vf(esi, IUnknown, QueryInterface), addr IID_IEnumVARIANT, addr m_pEnumVariant
		.if (eax != S_OK)
			invoke vf(esi, IUnknown, QueryInterface), addr IID_IEnumUnknown, addr m_pEnumVariant
			.if (eax != S_OK)
				jmp error
			.endif
			mov m_bUnknown, TRUE
		.endif
	.elseif ([eax].VARIANT.vt == (VT_ARRAY or VT_VARIANT))
		invoke VariantCopy, addr m_variant, eax
		mov ecx, m_variant.parray
		movzx edx, [ecx].SAFEARRAY.cDims
		lea ecx, [ecx].SAFEARRAY.rgsabound
		mov esi, 1
		.while (edx)
			mov eax,[ecx].SAFEARRAYBOUND.cElements
			push edx
			mul esi
			pop edx
			mov esi, eax
			mov m_dwNumVars, esi
			add ecx, sizeof SAFEARRAYBOUND
			dec edx
		.endw
	.else
		jmp error
	.endif

	mov eax, hWndFrom
	mov m_hWndFrom, eax

	return __this
error:
	invoke Destroy@CCollectionDlg, __this
	return 0
	align 4

Create@CCollectionDlg endp

	end
