

;*** definition of class CCPropertiesDlg
;*** CCPropertiesDlg implements a dialog showing/editing object properties

	.386
	.model flat,stdcall
	option casemap :none   ; case sensitive
	option proc:private

	include COMView.inc
	include statusbar.inc
INSIDE_CPROPERTIESDLG equ 1
	include classes.inc
	include rsrc.inc
	include CListView.inc
	include debugout.inc


?MODELESS		equ 1		;properties dialog is modeless
?MAXTEXTINPLACE	equ 128		;size of BSTR forcing COMView to show item in dialog
?EDITBUG		equ 1		;handles bug that edit control receives no WM_KILLFOCUS
?PUREINDIRECTION equ 1		;in vtbl mode: decrease ret arguments indirection
?CODEGEN		equ 1		;unfinished yet
?USEDISPINVOKE	equ 0		;dont use, doesnt work
?FORCETYPEINFO	equ 0		;use explicit typeinfo selected from object dialog
?COLLECTIONMODELESS	equ 1	;CCollectionDlg is modeless
?VARSUPP		equ 1		;display VARDESC entries in typeinfo
?SIMFUNCDESC	equ 1		;simulate FUNCDESC mode for VARDESC
?QUERYTI2		equ 0		;use QueryInterface to get an ITypeInfo2 pointer
?SHOWAMBIENT	equ 0		;show "Ambient" tab
?COMBOBOX		equ 1		;show combo boxed for UDT enums
?VARDESCRELEASE	equ 0		;save VARDESC ptr in simulated FUNCDESC structure
?PROPICON		equ 1		;properties dialogs have own icon

;;LPDATASOURCE	typedef ptr DataSource
;;externdef IID_DataSource:IID

WM_SETDISPATCH	equ WM_USER+101	;used by ?POSTSETDISP

FLAG_VARDESC	equ 80000000h

TAB_PROPERTIES	equ 0
TAB_METHODS		equ 1
if ?SHOWAMBIENT
TAB_AMBIENT		equ 2
endif
TAB_UNDEFINED	equ -1

ResetDispParams macro DispParams
	xor eax, eax
	mov DispParams.rgvarg, eax
	mov DispParams.rgdispidNamedArgs, eax
	mov DispParams.cArgs, eax
	mov DispParams.cNamedArgs, eax
	endm

;--- member ParamReturn is used to transfer VARIANTS between dialogs
;--- it is used:
;--- 1. transfer variants from IEnumVariant from enumvariantproc to main dialog
;--- 2. Get parameters needed for GETPROPERTY in GetPropertyWithParam
;--- 3. Inside ExecuteMethod to get parameters before Invoke_ call

;--- problem: to avoid forcing user to enter parameters several times
;--- in the sequence a. GETPROPERTY for Edit b. GETPROPTERTY for Compare
;--- c. PUTPROPERTY to change Property d: GETPROPERTY to refresh list view
;--- this sequence must be sure!

BEGIN_CLASS CPropertiesDlg, CDlg
hWndTab			HWND		?		;hWnd of tab control
hWndLV			HWND		?		;hWnd of listview
hWndSB			HWND		?		;hWnd of status bar
hIcon			HICON		?
iTab			DWORD		?		;currently selected tab (0=Properties,1=Methods)
pObjectItem		LPOBJECTITEM ?
pDispatch		LPDISPATCH	?		;current object ptr (neednt be a IDispatch!)
pTypeInfo		LPTYPEINFO	?		;current ITypeInfo ptr
pTypeLibDlg		pCTypeLibDlg ?		;ptr typelib dialog started from here
ptPos			POINT		<>		;start position of window
pvarResult		LPVARIANT	?
ParamReturn		PARAMRETURN <>		;structure for communication with CParamsDlg
ExcepInfo		EXCEPINFO	<>
dwArgErr		DWORD		?
wLastFlags		DWORD		?
LastHResult		DWORD		?
pszText			LPSTR		?		;text buffer (holds current property value)
dwTextMax		DWORD		?		;size of text buffer
dwTextSize		DWORD		?		;size of last read string (with term 0)
dwListIdx		DWORD		?		;current listview index during refresh
iSortCol		DWORD		?		;sort column index
iSortDir		DWORD		?		;sort column direction
dwRim			DWORD		?
wTypeFlags		WORD		?		;flags from TYPEATTR
bEdit			BOOLEAN		?		;allow editing in detail view
bRefresh		BOOLEAN		?		;update listview content (do not insert)
bException		BOOLEAN		?		;exception has occured
bScanMode		BOOLEAN		?		;property scan mode
bVtblMode		BOOLEAN		?		;typeinfo is TKIND_INTERFACE
bUseTIInvoke	BOOLEAN		?		;use ITypeInfo::Invoke for method calls
bHexadecimal	BOOLEAN		?
END_CLASS

__this	textequ <ebx>
_this	textequ <[__this].CPropertiesDlg>
thisarg	textequ <this@:ptr CPropertiesDlg>

	MEMBER hWnd, pDlgProc, hWndTab, hWndLV, hWndSB, iTab, hIcon
	MEMBER pObjectItem, pDispatch, pTypeInfo, pTypeLibDlg
	MEMBER ptPos, ParamReturn, dwListIdx, ExcepInfo, dwArgErr, wLastFlags
	MEMBER LastHResult, pszText, dwTextMax, dwTextSize, iSortCol, iSortDir
	MEMBER bEdit, bRefresh, bException, bScanMode, bVtblMode, bHexadecimal
	MEMBER wTypeFlags, bUseTIInvoke
	MEMBER dwRim, pvarResult

GetProperty			proto :ptr FUNCDESC, :BOOL
GetTypeInfo			proto
GetRetType			proto :ptr FUNCDESC
GetRetElemDesc		proto :ptr FUNCDESC
GetNumArgs			proto :ptr FUNCDESC
GetTypeInfoFromIProvideClassInfo proto pObject:LPUNKNOWN, bSource:BOOL
RefreshView			proto iTabIdx:DWORD
Variant2String		proto pFuncDesc:ptr FUNCDESC, pvarResult:ptr VARIANT, :BOOL
StartNewDialog		proto memid:DWORD, pvarResult:ptr VARIANT
PrepareInvokeErrorReturn proto hWndSB:HWND

	.data

g_dwCount		DWORD 0
g_pTypeInfoDlg	pCTypeInfoDlg NULL
if ?PROPICON
g_hIconProp		HICON NULL
endif
g_bHexadecimal	BYTE 0		;display prop values in hexadecimal
g_bSkipCodeGen	BYTE 0		;skip source code generation for next ExecuteInvoke

	.const

ColumnsProperties label CColHdr
		CColHdr <CStr("Property")	, 45>
		CColHdr <CStr("Value")		, 55>
NUMCOLS_PROPERTIES textequ %($ - ColumnsProperties) / sizeof CColHdr

ColumnsMethods label CColHdr
		CColHdr <CStr("Method")		, 45>
		CColHdr <CStr("Returns")	, 30>
		CColHdr <CStr("req/opt Params")	, 25>
NUMCOLS_METHODS textequ %($ - ColumnsMethods) / sizeof CColHdr


TabDlgPages label CTabDlgPage
	CTabDlgPage {CStr("Properties")}
	CTabDlgPage {CStr("Methods")}
if ?SHOWAMBIENT
	CTabDlgPage {CStr("Ambient")}
endif
NUMDLGS textequ %($ - TabDlgPages) / sizeof CTabDlgPage

;--- translation for HResults returned from IDispatch::Invoke

HResultTab label dword
	dd DISP_E_BADPARAMCOUNT
	dd DISP_E_BADVARTYPE
	dd DISP_E_MEMBERNOTFOUND 
	dd DISP_E_NONAMEDARGS 
	dd DISP_E_OVERFLOW 
	dd DISP_E_PARAMNOTFOUND 
	dd DISP_E_TYPEMISMATCH 
	dd DISP_E_UNKNOWNINTERFACE 
	dd DISP_E_UNKNOWNLCID 
 	dd DISP_E_PARAMNOTOPTIONAL 
 	dd E_UNEXPECTED 
NUMHRESULT equ ($ - HResultTab) / sizeof dword
HResultStr label dword
	dd CStr("BADPARAMCOUNT")
	dd CStr("BADVARTYPE")
	dd CStr("MEMBERNOTFOUND")
	dd CStr("NONAMEDARGS")
	dd CStr("OVERFLOW")
	dd CStr("PARAMNOTFOUND")
	dd CStr("TYPEMISMATCH")
	dd CStr("UNKNOWNINTERFACE")
	dd CStr("UNKNOWNLCID")
 	dd CStr("PARAMNOTOPTIONAL")
 	dd CStr("E_UNEXPECTED")

;;IID_DataSource sIID_DataSource

	.code

GetFuncDesc proc uses edi index:DWORD, ppFuncDesc:ptr ptr FUNCDESC

local pVarDesc:ptr VARDESC

		.if (index & FLAG_VARDESC)
			and index, 7FFFFFFFh
			invoke vf(m_pTypeInfo, ITypeInfo, GetVarDesc), index, addr pVarDesc
			.if (eax == S_OK)
				invoke malloc, sizeof FUNCDESC
				mov ecx, ppFuncDesc
				mov [ecx], eax
				mov edi, eax
				mov ecx, pVarDesc
				mov eax, [ecx].VARDESC.memid
				mov [edi].FUNCDESC.memid, eax
				mov [edi].FUNCDESC.invkind, INVOKE_PROPERTYGET

				.if ([ecx].VARDESC.varkind == VAR_DISPATCH)
					mov [edi].FUNCDESC.funckind, FUNC_DISPATCH
				.else
					mov [edi].FUNCDESC.funckind, FUNC_STATIC
				.endif
				mov [edi].FUNCDESC.cParams, 0
				mov [edi].FUNCDESC.oVft, 0
				.if ([ecx].VARDESC.wVarFlags & VARFLAG_FRESTRICTED)
					mov [edi].FUNCDESC.wFuncFlags, FUNCFLAG_FRESTRICTED
				.endif
if ?VARDESCRELEASE
				mov ax, [ecx].VARDESC.elemdescVar.tdesc.vt
				mov [edi].FUNCDESC.elemdescFunc.tdesc.vt, ax
				invoke vf(m_pTypeInfo, ITypeInfo, ReleaseVarDesc), pVarDesc
else
				invoke CopyMemory, addr [edi].FUNCDESC.elemdescFunc, addr [ecx].VARDESC.elemdescVar, sizeof ELEMDESC
				mov eax, pVarDesc
				mov [edi].FUNCDESC.lprgscode, eax
endif
				mov eax, S_OK
			.endif
		.else
			invoke vf(m_pTypeInfo, ITypeInfo, GetFuncDesc), index, ppFuncDesc
		.endif
		ret
		align 4

GetFuncDesc endp

ReleaseFuncDesc proc index:DWORD, pFuncDesc:ptr FUNCDESC

		.if (index & FLAG_VARDESC)
if ?VARDESCRELEASE eq 0
			mov ecx, pFuncDesc
			invoke vf(m_pTypeInfo, ITypeInfo, ReleaseVarDesc), [ecx].FUNCDESC.lprgscode
endif
			invoke free, pFuncDesc
		.else
			invoke vf(m_pTypeInfo, ITypeInfo, ReleaseFuncDesc), pFuncDesc
		.endif
		ret
		align 4

ReleaseFuncDesc endp

;--- get func desc (with listview iItem)

GetFuncDesc2 proc iItem:DWORD, ppFuncDesc:ptr ptr FUNCDESC

local lvi:LVITEM

		mov eax, iItem
		mov lvi.iItem, eax
		@mov lvi.iSubItem, 0
		mov lvi.mask_, LVIF_PARAM
		invoke ListView_GetItem( m_hWndLV, addr lvi)
		.if (eax)
			invoke GetFuncDesc, lvi.lParam, ppFuncDesc
		.else
			mov ecx, ppFuncDesc
			mov dword ptr [ecx], NULL
			mov eax, E_FAIL
		.endif
		ret
		align 4

GetFuncDesc2 endp

;--- release funcdesc (with listview iItem)

ReleaseFuncDesc2 proc iItem:DWORD, pFuncDesc:ptr FUNCDESC

local lvi:LVITEM

		mov eax, iItem
		mov lvi.iItem, eax
		@mov lvi.iSubItem, 0
		mov lvi.mask_, LVIF_PARAM
		invoke ListView_GetItem( m_hWndLV, addr lvi)
		.if (eax)
			invoke ReleaseFuncDesc, lvi.lParam, pFuncDesc
		.endif
		ret
		align 4

ReleaseFuncDesc2 endp


UpdateTypeInfoDlg	proc dwIndex:DWORD

local pFuncDesc:ptr FUNCDESC
local lvi:LVITEM

		.if ((!g_bSyncTypeInfoAndProp) && g_pTypeInfoDlg)
			mov ecx, g_pTypeInfoDlg
			invoke PostMessage, [ecx].CDlg.hWnd, WM_CLOSE, 0, 0
		.elseif ((g_bSyncTypeInfoAndProp) && (!g_pTypeInfoDlg))
			invoke Create2@CTypeInfoDlg, m_pTypeInfo
			mov g_pTypeInfoDlg, eax
			.if (eax)
				push eax
				invoke GetWindow, m_hWnd, GW_OWNER
				pop ecx
				invoke Show@CTypeInfoDlg, ecx, eax, 0
				invoke SetWindowPos, eax, m_hWnd, 0, 0, 0, 0,
					SWP_NOMOVE or SWP_NOSIZE or SWP_NOACTIVATE or SWP_SHOWWINDOW
				invoke SetFocus, m_hWndLV
			.endif
		.endif
		.if (g_pTypeInfoDlg)
			.if (dwIndex == -1)
				invoke ListView_GetNextItem( m_hWndLV, -1, LVNI_SELECTED)
				.if (eax != -1)
					mov lvi.iItem, eax
					@mov lvi.iSubItem, 0
					mov lvi.mask_, LVIF_PARAM
					invoke ListView_GetItem( m_hWndLV, addr lvi)
					mov eax, lvi.lParam
					mov dwIndex,eax
				.endif
			.endif
			invoke GetFuncDesc, dwIndex, addr pFuncDesc
			.if (eax == S_OK)
				mov ecx, pFuncDesc
				.if (dwIndex & FLAG_VARDESC)
					invoke FindVar@CTypeInfoDlg, g_pTypeInfoDlg, [ecx].FUNCDESC.memid
				.else
					invoke FindFunc@CTypeInfoDlg, g_pTypeInfoDlg, [ecx].FUNCDESC.memid, [ecx].FUNCDESC.invkind
				.endif
				invoke ReleaseFuncDesc, dwIndex, pFuncDesc
			.endif
		.endif
		ret
		align 4

UpdateTypeInfoDlg endp

;--------------------------------------------------------------
;--- class CPropertiesDlg
;--------------------------------------------------------------

Destroy@CPropertiesDlg proc uses __this thisarg

		mov __this,this@

		invoke ParamReturnClear, addr m_ParamReturn
		.if (m_pszText)
			invoke free, m_pszText
			mov m_pszText, NULL
		.endif
		.if (m_pTypeInfo)
			invoke vf(m_pTypeInfo, IUnknown, Release)
		.endif
		.if (m_pDispatch)
			invoke vf(m_pDispatch, IUnknown, Release)
		.endif
		.if (g_bDepTypeLibDlg && m_pTypeLibDlg)
			mov ecx, m_pTypeLibDlg
			invoke PostMessage, [ecx].CDlg.hWnd, WM_CLOSE, 0, 0
		.endif
		.if (m_pObjectItem)
			invoke vf( m_pObjectItem, IObjectItem, SetPropDlg), NULL
			invoke vf( m_pObjectItem, IObjectItem, Release)
		.endif
		.if (m_hIcon)
			invoke DestroyIcon, m_hIcon
		.endif
if 1
		dec g_dwCount
		.if ((g_dwCount == 0) && g_pTypeInfoDlg)
			mov ecx, g_pTypeInfoDlg
			invoke PostMessage, [ecx].CDlg.hWnd, WM_CLOSE, 0, 0
			mov g_pTypeInfoDlg, NULL
		.endif
endif
		invoke free, __this

		ret
		align 4

Destroy@CPropertiesDlg endp


;--- pItem MUST point to an IDispatch interface!


Create@CPropertiesDlg proc public uses esi __this pObjectItem:LPOBJECTITEM, pItem:ptr CInterfaceItem

local	pTypeLib:LPTYPELIB
local	pUnknown:LPUNKNOWN
local	pDispatch:LPDISPATCH
ifdef _DEBUG
local	this_:ptr CPropertiesDlg
endif

		invoke malloc, sizeof CPropertiesDlg
		.if (!eax)
			ret
		.endif

		mov __this,eax
ifdef _DEBUG
		mov this_, __this
endif
		mov m_pDlgProc, CPropertiesDialog
		mov m_iTab, TAB_UNDEFINED
		mov m_iSortCol, -1
		mov al, g_bHexadecimal
		mov m_bHexadecimal, al

		inc g_dwCount

		mov eax, pObjectItem
		mov m_pObjectItem, eax
		invoke vf(eax, IObjectItem, AddRef)
		invoke GetUnknown@CObjectItem, pObjectItem
		mov pUnknown, eax
		mov ecx, pItem
		invoke vf(pUnknown, IUnknown, QueryInterface), addr [ecx].CInterfaceItem.iid, addr m_pDispatch
		.if (eax != S_OK)
			invoke Destroy@CPropertiesDlg, __this
			return 0
		.endif
if ?FORCETYPEINFO eq 0
;-------------------------------- if it is an IDispatch, get typeinfo from there
		invoke vf(m_pDispatch, IUnknown, QueryInterface), addr IID_IDispatch, addr pDispatch
		.if (eax == S_OK)
			invoke vf(pDispatch, IUnknown, Release)
			mov ecx, pDispatch
			.if (ecx == m_pDispatch)
				invoke GetTypeInfo
			.endif
		.endif
endif
;-------------------------- try to load typeinfo from typelib infos 
		mov ecx, pItem
		.if ((m_pTypeInfo == NULL) && ([ecx].CInterfaceItem.dwVerMajor != -1))
			invoke LoadRegTypeLib, addr [ecx].CInterfaceItem.TypelibGUID,
				[ecx].CInterfaceItem.dwVerMajor,
				[ecx].CInterfaceItem.dwVerMinor,
				g_LCID, addr pTypeLib
			.if (eax == S_OK)
				mov ecx, pItem
				invoke vf(pTypeLib, ITypeLib, GetTypeInfoOfGuid),
					addr [ecx].CInterfaceItem.iid, addr m_pTypeInfo
				invoke vf(pTypeLib, IUnknown, Release)
			.endif
		.endif
		.if (!m_pTypeInfo)
			invoke Destroy@CPropertiesDlg, __this
			return 0
		.endif

		invoke vf(m_pObjectItem, IObjectItem, SetPropDlg), __this

		return __this
		align 4

Create@CPropertiesDlg endp

SetWindowPos@CPropertiesDlg proc public uses __this thisarg, pptPos:ptr POINT

		mov __this, this@
		mov ecx, pptPos
		.if (ecx)
			mov eax,[ecx].POINT.x
			mov m_ptPos.x, eax
			mov eax,[ecx].POINT.y
			mov m_ptPos.y, eax
		.endif
		ret
		align 4

SetWindowPos@CPropertiesDlg endp

Create2@CPropertiesDlg proc public uses __this pUnknown:LPUNKNOWN, pTypeInfo:LPTYPEINFO

local	pUnknown2:LPUNKNOWN
local	pTypeAttr:ptr TYPEATTR

		invoke malloc, sizeof CPropertiesDlg
		.if (!eax)
			ret
		.endif

		mov __this,eax

		inc g_dwCount

		mov m_pDlgProc, CPropertiesDialog

		mov m_iTab, TAB_UNDEFINED
		mov m_iSortCol, -1
		mov al, g_bHexadecimal
		mov m_bHexadecimal, al

;--- if we got a typeinfo, make sure its the correct GUID

		mov eax, pTypeInfo
		.if (eax)
			mov m_pTypeInfo, eax
			invoke vf(eax, IUnknown, AddRef)
			invoke vf(m_pTypeInfo, ITypeInfo, GetTypeAttr), addr pTypeAttr
			.if (eax != S_OK)
				jmp error
			.endif
			mov ecx, pTypeAttr
			invoke vf(pUnknown, IUnknown, QueryInterface), addr [ecx].TYPEATTR.guid, addr m_pDispatch
			push eax
			invoke vf(m_pTypeInfo, ITypeInfo, ReleaseTypeAttr), pTypeAttr
			pop eax
			.if ((eax != S_OK) || (!m_pDispatch))
				.if (!m_pDispatch)
					invoke vf(pUnknown, IUnknown, QueryInterface), addr IID_IDispatch, addr m_pDispatch
				.endif
				.if (eax != S_OK)
					jmp error
				.endif
			.endif
		.else

;--- else got a IDispatch

			invoke vf(pUnknown, IUnknown, QueryInterface), addr IID_IDispatch, addr m_pDispatch
			.if (eax != S_OK)
				jmp error
			.endif
			invoke GetTypeInfo
			.if (!m_pTypeInfo)
				jmp error
			.endif
		.endif

		invoke vf(pUnknown, IUnknown, QueryInterface), addr IID_IUnknown, addr pUnknown2
		invoke Find@CObjectItem, pUnknown2
		.if (eax)
			mov m_pObjectItem, eax
			invoke vf(eax, IObjectItem, AddRef)
		.else
			invoke Create@CObjectItem, pUnknown2, NULL
			mov m_pObjectItem, eax
		.endif
		invoke vf(m_pObjectItem, IObjectItem, SetPropDlg), __this
		invoke vf(pUnknown2, IUnknown, Release)

		.if (g_pTypeInfoDlg)
			invoke SetTypeInfo@CTypeInfoDlg, g_pTypeInfoDlg, m_pTypeInfo
		.endif

		return __this
error:
		invoke Destroy@CPropertiesDlg, __this
		return 0
		align 4

Create2@CPropertiesDlg endp

Show@CPropertiesDlg proc public thisarg, hWnd:HWND

		invoke CreateDialogParam, g_hInstance, IDD_PROPERTIESDLG, hWnd, classdialogproc, this@
		ret
		align 4

Show@CPropertiesDlg endp

GetObjectItem@CPropertiesDlg proc public thisarg

		mov ecx, this@
		mov eax, [ecx].CPropertiesDlg.pObjectItem
		ret
		align 4

GetObjectItem@CPropertiesDlg endp

;--- this function is used if all property dialogs use same window

SetDispatch@CPropertiesDlg proc public uses __this thisarg, pNewDisp:LPDISPATCH, bRefresh:BOOL

local tmp1:dword
local tmp2:dword

		mov __this,this@

		mov eax, m_pTypeInfo
		mov tmp1, eax
		;push m_pTypeInfo
		mov eax, m_pDispatch
		mov tmp2, eax
		;push m_pDispatch

		mov m_pTypeInfo, NULL
		mov eax, pNewDisp
		mov m_pDispatch, eax

		.if (bRefresh)
			mov m_iTab, TAB_UNDEFINED
			invoke TabCtrl_SetCurSel( m_hWndTab, TAB_PROPERTIES)
			invoke RefreshView, TAB_PROPERTIES
		.else
			invoke RefreshView, m_iTab
		.endif

		.if (eax)
if ?POSTSETDISP eq 0
			invoke vf(m_pDispatch, IUnknown, AddRef)
endif
			.if tmp2
				invoke vf(tmp2, IUnknown, Release)
			.endif
			.if (tmp1)
				invoke vf(tmp1, IUnknown, Release)
			.endif

			.if (m_pObjectItem)
				invoke vf(m_pObjectItem, IObjectItem, Release) 
			.endif
			invoke Create@CObjectItem, m_pDispatch, NULL
			mov m_pObjectItem, eax
			invoke vf( eax, IObjectItem, SetPropDlg), __this

			.if (g_pTypeInfoDlg && m_pTypeInfo)
				invoke SetTypeInfo@CTypeInfoDlg, g_pTypeInfoDlg, m_pTypeInfo
				.if (eax)
					invoke UpdateTypeInfoDlg, -1
				.endif
			.endif
			mov eax, 1
		.else
			mov eax, tmp2
			mov m_pDispatch, eax
			mov eax, tmp1
			mov m_pTypeInfo, eax
		.endif
		ret
		align 4

SetDispatch@CPropertiesDlg endp

;--- called by IPropertyNotifySink::OnChanged

OnChanged@CPropertiesDlg proc public uses __this thisarg, dispId:DWORD

local	pTypeInfo2:LPTYPEINFO2
local	dwIndex:DWORD
local	pFuncDesc:ptr FUNCDESC
local	typekind:TYPEKIND
local	lvfi:LVFINDINFO
local	lvi:LVITEM

;--- only dispId of property is available
;--- in listviews lParam is FUNCDESC index only!
;--- so the following has to be done:
;--- 1. check if properties are displayed
;--- 2. get FUNCDESC index from dispid
;--- 3. do LVM_FINDITEM to find FUNCDESC index
;--- 4. get FUNCDESC
;--- 5. call GetProperty with FUNCDESC
;--- 6. set listview item
;--- 7. release FUNCDESC

		mov __this,this@
		.if (m_iTab == TAB_PROPERTIES)
;------------------------------ no need to query for ITypeInfo2, just do a cast
if ?QUERYTI2
			invoke vf(m_pTypeInfo, ITypeInfo, QueryInterface), addr IID_ITypeInfo2, addr pTypeInfo2
else
			mov eax, m_pTypeInfo
			mov pTypeInfo2, eax
			mov eax, S_OK
endif
			.if (eax == S_OK)
				invoke vf(pTypeInfo2, ITypeInfo2, GetFuncIndexOfMemId), dispId, INVOKE_PROPERTYGET, addr dwIndex
				.if (eax == S_OK)
					mov lvfi.flags, LVFI_PARAM
;------------------------------ adjust index returned by # of IDispatch methods
;------------------------------ that's possibly a bug in ITypeInfo2 interface,
;------------------------------ since TYPEKIND is TKIND_DISPATCH, not TKIND_INTERFACE
					add dwIndex, 7
					mov eax, dwIndex
					mov lvfi.lParam, eax
					invoke ListView_FindItem( m_hWndLV, -1, addr lvfi)
					.if (eax != -1)
						mov lvi.iItem, eax
						invoke vf(pTypeInfo2, ITypeInfo2, GetFuncDesc), dwIndex, addr pFuncDesc
						.if (eax == S_OK)
							invoke GetProperty, pFuncDesc, TRUE
							mov lvi.iSubItem, 1
							mov lvi.mask_, LVIF_TEXT
							mov eax, m_pszText
							mov lvi.pszText, eax
							invoke ListView_SetItem( m_hWndLV, addr lvi)
							invoke vf(pTypeInfo2, ITypeInfo2, ReleaseFuncDesc), pFuncDesc
						.endif
					.endif
				.endif
if ?QUERYTI2
				invoke vf(pTypeInfo2, ITypeInfo2, Release)
endif
			.endif
		.endif
		ret
		align 4

OnChanged@CPropertiesDlg endp

;--- show context menu


ShowContextMenu proc uses esi bMouse:BOOL

local	dwFlags:DWORD
local	pt:POINT
local	iItem:DWORD
local	pFuncDesc:ptr FUNCDESC
local	bstr:BSTR
local	dwContext:DWORD

		invoke GetSubMenu,g_hMenu, ID_SUBMENU_PROPERTIESDLG
		.if (eax)
			mov esi, eax
			invoke ListView_GetNextItem( m_hWndLV, -1, LVNI_SELECTED)
			mov iItem, eax
			mov ecx, MF_ENABLED
			inc eax
			.if (!eax)
				mov ecx, MF_GRAYED or MF_DISABLED
				dec eax
			.else
				mov eax, IDM_EDITITEM
			.endif
			mov dwFlags, ecx
			invoke SetMenuDefaultItem, esi, eax, FALSE
			invoke EnableMenuItem, esi, IDM_EDITITEM, dwFlags
			.if (m_iTab == TAB_PROPERTIES)
				mov eax, dwFlags
			.else
				mov eax, MF_GRAYED or MF_DISABLED
			.endif
			invoke EnableMenuItem, esi, IDM_COPYVALUE, eax

			.if (g_bSyncTypeInfoAndProp)
				mov ecx, MF_CHECKED
			.else
				mov ecx, MF_UNCHECKED
			.endif
			invoke CheckMenuItem, esi, IDM_TYPEINFO, ecx
			.if (g_bShowAllMembers)
				mov ecx, MF_CHECKED
			.else
				mov ecx, MF_UNCHECKED
			.endif
			invoke CheckMenuItem, esi, IDM_SHOWALL, ecx
			.if (m_bHexadecimal)
				mov ecx, MF_CHECKED
			.else
				mov ecx, MF_UNCHECKED
			.endif
			invoke CheckMenuItem, esi, IDM_HEXADECIMAL, ecx

			mov dwFlags, MF_GRAYED or MF_DISABLED
			.if (iItem != -1)
				invoke GetFuncDesc2, iItem, addr pFuncDesc
				.if (eax == S_OK)
					mov ecx, pFuncDesc
					invoke vf(m_pTypeInfo, ITypeInfo, GetDocumentation), [ecx].FUNCDESC.memid, NULL, NULL, addr dwContext, addr bstr
					.if (eax == S_OK && bstr)
						invoke SysFreeString, bstr
						.if (dwContext)
							mov dwFlags, MF_ENABLED
						.endif
					.endif
					invoke ReleaseFuncDesc2, iItem, pFuncDesc
				.endif
			.endif
			invoke EnableMenuItem, esi, IDM_CONTEXTHELP, dwFlags


			mov ecx, MF_UNCHECKED
			.if (m_bUseTIInvoke || m_bVtblMode)
				mov ecx, MF_CHECKED
			.endif
			invoke CheckMenuItem, esi, IDM_USETIINVOKE, ecx
			mov ecx, MF_ENABLED
			.if (m_bVtblMode)
				mov ecx, MF_GRAYED or MF_DISABLED
			.endif
			invoke EnableMenuItem, esi, IDM_USETIINVOKE, ecx

			invoke DeleteMenu, esi, IDM_FORCETYPEINFO, MF_BYCOMMAND
			.if (g_bShowForceTypeInfo)
				invoke AppendMenu, esi, MF_ENABLED or MF_BYCOMMAND, IDM_FORCETYPEINFO, CStr("Set TypeInfo")
			.endif

			.if (!(dwFlags & MF_DISABLED))
				invoke SetMenuDefaultItem, esi, IDM_EDITITEM, FALSE
			.endif
			invoke GetItemPosition, m_hWndLV, bMouse, addr pt
			invoke TrackPopupMenu, esi, TPM_LEFTALIGN or TPM_LEFTBUTTON,
					pt.x,pt.y,0,m_hWnd,NULL
		.endif
		ret
		align 4

ShowContextMenu endp


;--- get error string for errors returned by IDispatch::Invoke


GetInvokeErrorString proc uses edi hResult:DWORD

		mov eax,hResult
		mov ecx,NUMHRESULT
		mov edi,offset HResultTab
		repnz scasd
		.if (ZERO?)
			sub ecx,NUMHRESULT-1
			neg ecx
			mov eax,[ecx*4+offset HResultStr]
		.else
			mov eax, CStr("")
		.endif
		ret
		align 4

GetInvokeErrorString endp


;--- ensure text buffer pointer is large enough


GetTextBufferPtr proc dwSize:DWORD

		mov eax, m_pszText
		mov ecx, dwSize
		.if (ecx < MAX_PATH)
			mov ecx, MAX_PATH
		.endif
		.if ((ecx > MAX_PATH) || (eax == NULL))
			push ecx
			.if (eax)
				invoke free, eax
				mov m_pszText, NULL
			.endif
			pop eax
			mov m_dwTextMax, eax
			invoke malloc, eax
			.if (!eax)
				ret
			.endif
			mov m_pszText, eax
		.endif
		ret
		align 4

GetTextBufferPtr endp


if ?CODEGEN

GetVarAsString proc pVariant:ptr VARIANT, pStrOut:LPSTR, dwMax:DWORD, pStrOut2:LPSTR
		mov edx, pVariant
		movzx eax, [edx].VARIANT.vt
		.if ((eax == VT_ERROR) || (eax & VT_BYREF))
			mov eax, [edx].VARIANT.lVal
			invoke wsprintf, pStrOut, CStr("0%Xh"), eax
		.elseif (eax == VT_I1 || eax == VT_I2 || eax == VT_I4 || eax == VT_UI1 || eax == VT_UI2 || eax == VT_UI4)
			mov eax, [edx].VARIANT.lVal
			invoke wsprintf, pStrOut, CStr("%u"), eax
		.elseif (eax == VT_BSTR)
			mov eax, [edx].VARIANT.bstrVal
			.if (eax)
				invoke wsprintf, pStrOut, CStr("BStr(""%S"")"), eax
			.else
				invoke wsprintf, pStrOut, CStr("NULL")
			.endif
		.elseif (eax == VT_BOOL)
			movsx eax, [edx].VARIANT.boolVal
			invoke wsprintf, pStrOut, CStr("%d"), eax
		.elseif ((eax == VT_R8) || (eax == VT_DATE) || (eax == VT_CY) || (eax == VT_I8))
			push edx
			invoke wsprintf, pStrOut, CStr("0%Xh"), dword ptr [edx].VARIANT.dblVal+0
			pop edx
			invoke wsprintf, pStrOut2, CStr("0%Xh"), dword ptr [edx].VARIANT.dblVal+4
		.else
			invoke lstrcpy, pStrOut, CStr("0")
		.endif
		ret
		align 4

GetVarAsString endp

GenerateCode proc uses edi esi pFuncDesc:ptr FUNCDESC, wFlags:DWORD, pDispParams:ptr DISPPARAMS,
					pvarResult:ptr VARIANT

local	pszFlags:LPSTR
local	dwParams:DWORD
local	szType[64]:byte
local	szType2[64]:byte
local	szDest[64]:byte
local	szDest2[64]:byte

		.if (g_bLogActive && g_bDispUserCalls)
			mov esi, pDispParams
			mov ecx, [esi].DISPPARAMS.cArgs
			mov dwParams, ecx
			.if (ecx)
				invoke printf@CLogWindow, CStr("sub esp, %u * sizeof VARIANT",10), ecx
			.endif
			mov eax, dwParams
			mov ecx, eax
			mov esi, [esi].DISPPARAMS.rgvarg
			shl eax, 4
			add esi, eax
			.while (ecx)
				sub esi, sizeof VARIANT
				push ecx
				dec ecx
				push ecx
				invoke wsprintf, addr szDest, CStr("mov [esp + %u * sizeof VARIANT].VARIANT."), ecx
				pop ecx
				invoke wsprintf, addr szDest2, CStr("mov dword ptr [esp + %u * sizeof VARIANT].VARIANT."), ecx
				movzx eax,[esi].VARIANT.vt
				and eax, NOT VT_BYREF
				invoke GetVarType, eax
				lea edi, szType
				invoke lstrcpy, edi, eax
				invoke CharUpper, edi
				test [esi].VARIANT.vt, VT_BYREF
				.if (ZERO?)
					invoke printf@CLogWindow, CStr("%svt, VT_%s",10), addr szDest, edi
				.else
					invoke printf@CLogWindow, CStr("%svt, VT_%s or VT_BYREF",10), addr szDest, edi
				.endif
				invoke GetVarAsString, esi, addr szType, sizeof szType, addr szType2
				movzx eax,[esi].VARIANT.vt
				.if (ax & VT_BYREF)
					invoke printf@CLogWindow, CStr("%sbyref, %s",10), addr szDest, addr szType
				.elseif (eax == VT_CY)
					invoke printf@CLogWindow, CStr("%scyVal+0, %s",10), addr szDest2, addr szType
					invoke printf@CLogWindow, CStr("%scyVal+4, %s",10), addr szDest2, addr szType2
				.elseif (eax == VT_DATE)
					invoke printf@CLogWindow, CStr("%sdate+0, %s",10), addr szDest2, addr szType
					invoke printf@CLogWindow, CStr("%sdate+4, %s",10), addr szDest2, addr szType2
				.elseif (eax == VT_R8)
					invoke printf@CLogWindow, CStr("%sdblVal+0, %s",10), addr szDest2, addr szType
					invoke printf@CLogWindow, CStr("%sdblVal+4, %s",10), addr szDest2, addr szType2
				.elseif (eax == VT_BSTR)
					invoke printf@CLogWindow, CStr("%sbstrVal, %s",10), addr szDest, addr szType
				.elseif (eax == VT_BOOL)
					invoke printf@CLogWindow, CStr("%sboolVal, %s",10), addr szDest, addr szType
				.elseif (eax == VT_ERROR)
					invoke printf@CLogWindow, CStr("%sscode, DISP_E_PARAMNOTFOUND",10), addr szDest
				.else
					invoke printf@CLogWindow, CStr("%slVal, %s",10), addr szDest, addr szType
				.endif
				
				pop ecx
				dec ecx
			.endw
			.if (dwParams)
				invoke printf@CLogWindow, CStr("mov dispparams.cArgs, %u",10), dwParams
				invoke printf@CLogWindow, CStr("mov dispparams.rgvarg, esp",10)
			.else
				invoke printf@CLogWindow, CStr("mov dispparams.cArgs, 0",10)
				invoke printf@CLogWindow, CStr("mov dispparams.rgvarg, NULL",10)
			.endif
			.if ((wFlags == DISPATCH_PROPERTYPUT) || (wFlags == DISPATCH_PROPERTYPUTREF))
				invoke printf@CLogWindow, CStr("mov dispparams.cNamedArgs, 1",10)
				invoke printf@CLogWindow, CStr("push DISPID_PROPERTYPUT",10)
				invoke printf@CLogWindow, CStr("mov dispparams.rgdispidNamedArgs, esp",10)
			.else
				invoke printf@CLogWindow, CStr("mov dispparams.cNamedArgs, 0",10)
				invoke printf@CLogWindow, CStr("mov dispparams.rgdispidNamedArgs, NULL",10)
			.endif
			mov ecx, wFlags
			.if (ecx == DISPATCH_METHOD)
				mov ecx, CStr("DISPATCH_METHOD")
			.elseif (ecx == DISPATCH_PROPERTYGET)
				mov ecx, CStr("DISPATCH_PROPERTYGET")
			.elseif (ecx == DISPATCH_PROPERTYPUT)
				mov ecx, CStr("DISPATCH_PROPERTYPUT")
			.elseif (ecx == DISPATCH_PROPERTYPUTREF)
				mov ecx, CStr("DISPATCH_PROPERTYPUTREF")
			.else
				invoke wsprintf, addr szType, CStr("0%Xh"), ecx
				lea ecx, szType
			.endif
			.if (pvarResult)
				mov edx, CStr("addr vtResult")
			.else
				mov edx, CStr("NULL")
			.endif
			mov edi, pFuncDesc
			.if (m_bVtblMode || (m_bUseTIInvoke && (m_wTypeFlags & TYPEFLAG_FDUAL)))
				invoke printf@CLogWindow, CStr("invoke vf(pTypeInfo, ITypeInfo, Invoke_), pDispatch, 0%Xh, %s, addr dispparams, %s, NULL, NULL",10),
					[edi].FUNCDESC.memid, ecx, edx
			.else
				invoke printf@CLogWindow, CStr("invoke vf(pDispatch, IDispatch, Invoke_), 0%Xh, addr IID_NULL, LOCALE_SYSTEM_DEFAULT, %s, addr dispparams, %s, NULL, NULL",10),
					[edi].FUNCDESC.memid, ecx, edx
			.endif
			.if ((dwParams) || (wFlags == DISPATCH_PROPERTYPUT) || (wFlags == DISPATCH_PROPERTYPUTREF))
				mov eax, dwParams
				shl eax, 4
				.if ((wFlags == DISPATCH_PROPERTYPUT) || (wFlags == DISPATCH_PROPERTYPUTREF))
					add eax, 4
				.endif
				invoke printf@CLogWindow, CStr("add esp, %u",10), eax
			.endif
		.endif
		ret
		align 4

GenerateCode endp

endif

;--- execute IDispatch:Invoke in a guarded code section

ifdef @StackBase
	option stackbase:ebp
endif
	option prologue:@sehprologue
	option epilogue:@sehepilogue

ExecuteInvoke proc uses esi edi __this memid:MEMBERID, wFlags:DWORD,
			pDispParams:ptr DISPPARAMS, pvarResult:ptr VARIANT

local	this@:ptr CPropertiesDlg
local	szText[128]:byte

		mov this@, __this		;save it
		mov m_ExcepInfo.bstrSource, NULL
		mov m_ExcepInfo.bstrDescription, NULL
		mov m_ExcepInfo.bstrHelpFile, NULL
		mov m_ExcepInfo.pfnDeferredFillIn, NULL
		mov m_dwArgErr, 0
		mov eax, wFlags
		mov m_wLastFlags, eax
		.try
			invoke SetBusyState@CMainDlg, TRUE
if ?USEDISPINVOKE
			invoke DispInvoke, m_pDispatch, m_pTypeInfo, memid,
					wFlags, pDispParams, pvarResult, addr m_ExcepInfo, addr m_dwArgErr
else
			.if (m_bVtblMode || (m_bUseTIInvoke && (m_wTypeFlags & TYPEFLAG_FDUAL)))
				invoke vf(m_pTypeInfo, ITypeInfo, Invoke_), m_pDispatch, memid,
					wFlags, pDispParams, pvarResult, addr m_ExcepInfo, addr m_dwArgErr
			.else
				invoke vf(m_pDispatch, IDispatch, Invoke_), memid, addr IID_NULL,
					g_LCID, wFlags, pDispParams, pvarResult, addr m_ExcepInfo, addr m_dwArgErr
			.endif
endif
			invoke SetBusyState@CMainDlg, FALSE
		.exceptfilter
			invoke SetBusyState@CMainDlg, FALSE
			mov __this,this@
			mov eax,_exception_info()
			mov eax,(EXCEPTION_POINTERS ptr [eax]).ExceptionRecord
			mov ecx,(EXCEPTION_RECORD ptr [eax]).ExceptionCode
			mov edx,ecx
			mov ecx,(EXCEPTION_RECORD ptr [eax]).ExceptionAddress
			.if (wFlags & DISPATCH_PROPERTYGET)
				mov esi, CStr("PropertyGet")
			.elseif (wFlags & DISPATCH_PROPERTYPUT)
				mov esi, CStr("PropertyPut")
			.else
				mov esi, CStr("Method")
			.endif
			.if (m_bScanMode)
				mov eax, CStr("Continue?")
				mov edi, MB_YESNO
			.else
				mov eax, CStr("")
				mov edi, MB_OK
			.endif
			invoke wsprintf, addr szText, CStr("Exception 0x%08X occured at 0x%08X.",10,"MemberId: %d(%Xh), Invoke Type: %s",10,"Function aborted. %s"), edx, ecx, memid, memid, esi, eax
			.if (m_bVtblMode || (m_bUseTIInvoke && (m_wTypeFlags & TYPEFLAG_FDUAL)))
				mov ecx, CStr("Error executing ITypeInfo::Invoke()")
			.else
				mov ecx, CStr("Error executing IDispatch::Invoke()")
			.endif
			invoke MessageBox, m_hWnd, addr szText, ecx, edi
			.if (eax == IDNO)
				mov m_bScanMode, FALSE
			.endif
			mov m_bException, TRUE
			mov eax, EXCEPTION_EXECUTE_HANDLER
		.except
			mov __this,this@
			mov eax, E_UNEXPECTED
		.endtry
		mov m_LastHResult, eax
		ret
		align 4

ExecuteInvoke endp

	option prologue: prologuedef
	option epilogue: epiloguedef
ifdef @StackBase
	option stackbase:esp
endif


;--- get real number of parameters

GetNumArgs proc uses edi pFuncDesc:ptr FUNCDESC

		mov edi, pFuncDesc
		assume edi:ptr FUNCDESC

		movzx eax, [edi].FUNCDESC.cParams
		.if (m_bVtblMode)
			movzx ecx, [edi].cParams
			mov edx, [edi].lprgelemdescParam
			.while (ecx)
				.if ([edx].ELEMDESC.paramdesc.wParamFlags & (PARAMFLAG_FRETVAL or PARAMFLAG_FLCID))
					dec eax
				.endif
				add edx, sizeof ELEMDESC
				dec ecx
			.endw
		.endif
		ret
		assume edi:nothing
		align 4

GetNumArgs endp

;--- get real type of return

GetRetType proc uses edi pFuncDesc:ptr FUNCDESC

		mov edi, pFuncDesc
		assume edi:ptr FUNCDESC

		movzx eax, [edi].elemdescFunc.tdesc.vt
		.if (m_bVtblMode)
			movzx ecx, [edi].cParams
			mov edx, [edi].lprgelemdescParam
			mov eax, VT_VOID
			.while (ecx)
				.if ([edx].ELEMDESC.paramdesc.wParamFlags & PARAMFLAG_FRETVAL)
if ?PUREINDIRECTION
					.if ([edx].ELEMDESC.tdesc.vt == VT_PTR)
						mov edx, [edx].ELEMDESC.tdesc.lptdesc
						movzx eax, [edx].ELEMDESC.tdesc.vt
						.break
					.endif
else
					movzx eax, [edx].ELEMDESC.tdesc.vt
					.break
endif
				.endif
				add edx, sizeof ELEMDESC
				dec ecx
			.endw
		.endif
		ret
		assume edi:nothing
		align 4

GetRetType endp

;--- get ELEMDESC of return type
;--- edi -> FUNCDESC

GetRetElemDesc proc uses edi pFuncDesc:ptr FUNCDESC

		mov edi, pFuncDesc
		assume edi:ptr FUNCDESC

		lea eax, [edi].elemdescFunc
		.if (m_bVtblMode)
			movzx ecx, [edi].cParams
			mov edx, [edi].lprgelemdescParam
			.while (ecx)
				.if ([edx].ELEMDESC.paramdesc.wParamFlags & PARAMFLAG_FRETVAL)
if ?PUREINDIRECTION
					.if ([edx].ELEMDESC.tdesc.vt == VT_PTR)
						mov edx, [edx].ELEMDESC.tdesc.lptdesc
						mov eax, edx
						.break
					.endif
else
					mov eax, edx
					.break
endif
				.endif
				add edx, sizeof ELEMDESC
				dec ecx
			.endw
		.endif
		ret
		assume edi:nothing
		align 4

GetRetElemDesc endp

FreeErrorStrings proc
		invoke SysFreeString, m_ExcepInfo.bstrSource
		invoke SysFreeString, m_ExcepInfo.bstrDescription
		invoke SysFreeString, m_ExcepInfo.bstrHelpFile
		ret
		align 4
FreeErrorStrings endp

;--- callback from CParamsDlg: user has entered params and pressed OK
;--- now call the method and return S_OK to close the params dialog 

GetPropertyWithParamCB proc uses edi __this this_:ptr CPropertiesDlg, hWnd:HWND, pFuncDesc:ptr FUNCDESC,
				iNumVariants:DWORD, pVariants:ptr VARIANT, hWndSB:HWND

local DispParams:DISPPARAMS

		mov __this, this_
		mov edi, pFuncDesc
		assume edi:ptr FUNCDESC

		ResetDispParams DispParams

		mov eax, pVariants
		mov DispParams.rgvarg, eax
		mov ecx, iNumVariants
		mov DispParams.cArgs, ecx
if ?CODEGEN
		.if (!g_bSkipCodeGen)
			invoke GenerateCode, edi, DISPATCH_PROPERTYGET, addr DispParams, m_pvarResult
		.endif
		mov g_bSkipCodeGen, 0
endif
		invoke ExecuteInvoke, [edi].memid, DISPATCH_PROPERTYGET,
			addr DispParams, m_pvarResult
		.if (eax == S_OK)
			mov edi, pVariants
			.if (edi && (!m_ParamReturn.pVariants))
				mov eax, iNumVariants
				shl eax, 4
				invoke malloc, eax
				mov m_ParamReturn.pVariants, eax
				.if (eax)
					mov ecx, iNumVariants
					mov m_ParamReturn.iNumVariants, ecx
					.while (ecx)
						push ecx
						push eax
						invoke VariantCopy, eax, edi
						pop eax
						add eax, sizeof VARIANT
						add edi, sizeof VARIANT
						pop ecx
						dec ecx
					.endw
				.endif
			.endif
			@mov eax, S_OK
		.elseif (hWndSB)
			invoke PrepareInvokeErrorReturn, hWndSB
		.endif
		ret
		align 4

GetPropertyWithParamCB endp


;--- get a property into a VARIANT


GetPropertyWithParam proc pFuncDesc:ptr FUNCDESC, pvarResult:ptr VARIANT

		mov eax, pvarResult
		mov m_pvarResult, eax
		invoke VariantInit, eax
		invoke GetNumArgs, pFuncDesc
		.if (eax && (!m_ParamReturn.pVariants))
			invoke Create@CParamsDlg, m_pTypeInfo, pFuncDesc, __this, offset GetPropertyWithParamCB
			invoke DialogBoxParam, g_hInstance, IDD_PARAMSDLG, m_hWnd, classdialogproc, eax
		.else
;--------------------- dont call PrepareInvokeErrorReturn in this case
			invoke GetPropertyWithParamCB, __this, m_hWnd, pFuncDesc,
				m_ParamReturn.iNumVariants, m_ParamReturn.pVariants, NULL
		.endif
		ret
		align 4

GetPropertyWithParam endp


;--- translate a returned HResult to a string


HResult2String proc uses esi edi pFuncDesc:ptr FUNCDESC, HResult:DWORD

;local dwESP:DWORD	;use EDI to save/restore ESP

		invoke GetTextBufferPtr, MAX_PATH
		mov eax, HResult
		mov ecx, pFuncDesc
		.if (eax == DISP_E_EXCEPTION)
;			mov dwESP, esp
			mov edi, esp
			lea esi, m_ExcepInfo
			.if ([esi].EXCEPINFO.bstrDescription)
				invoke SysStringLen, [esi].EXCEPINFO.bstrDescription
				add eax, 4
				and al, 0FCh
				sub esp, eax
				mov edx, esp
				invoke WideCharToMultiByte, CP_ACP, 0, [esi].EXCEPINFO.bstrDescription, -1, edx, eax, 0, 0
				mov edx, esp
			.else
				mov edx, CStr("")
			.endif
			invoke wsprintf, m_pszText, CStr("[%X] '%.128s' Exception"), [esi].EXCEPINFO.scode, edx
;			mov esp, dwESP
			mov esp, edi

;;			DebugOut "HResult2String, excepinfo=%X, %X, %X", [esi].EXCEPINFO.bstrSource, [esi].EXCEPINFO.bstrDescription, [esi].EXCEPINFO.bstrHelpFile
			invoke FreeErrorStrings

		.elseif ((eax == DISP_E_BADPARAMCOUNT) && ecx && ([ecx].FUNCDESC.cParams))
			invoke wsprintf, m_pszText, CStr("[Parameter required]")
		.else
			invoke GetInvokeErrorString, eax
			invoke wsprintf, m_pszText, CStr("Invoke Error[%X] %s"), HResult, eax
		.endif
		ret
		align 4

HResult2String endp

;--- if return type is VT_USERDEFINED (Enum), get translation

UserType2String proc  uses esi edi pFuncDesc:ptr FUNCDESC, pvarResult:ptr VARIANT

local	pTypeInfoRef:LPTYPEINFO
local	pTypeAttr:ptr TYPEATTR
local	pVarDesc:ptr VARDESC
local	varTemp:VARIANT
local	dwIndex:DWORD
local	bstrVar:BSTR
local	dwTmp:DWORD
local	dwSize:DWORD
local	bFound:BOOL
local	hr:DWORD

		mov hr, E_FAIL
		mov edi, pFuncDesc
		assume edi:ptr FUNCDESC
		.if (!edi)
			jmp exit
		.endif
		invoke GetRetType, edi
		.if (eax != VT_USERDEFINED)
			jmp exit
		.endif

		invoke GetRetElemDesc, edi
		mov esi, eax
		assume esi:ptr ELEMDESC

;		DebugOut "UserType2String memid=%X", [edi].FUNCDESC.memid

		invoke vf(m_pTypeInfo, ITypeInfo, GetRefTypeInfo), [esi].tdesc.hreftype, addr pTypeInfoRef
		.if (eax == S_OK)
			invoke vf(pTypeInfoRef, ITypeInfo, GetTypeAttr), addr pTypeAttr
			.if (eax == S_OK)
				mov dwIndex, 0
				mov bFound, FALSE
				invoke VariantInit, addr varTemp
				.while (bFound == FALSE)
					mov edx, pTypeAttr
					mov ecx, dwIndex
					.break .if (cx > [edx].TYPEATTR.cVars)
					invoke vf(pTypeInfoRef, ITypeInfo, GetVarDesc), dwIndex, addr pVarDesc
					.if (eax == S_OK)
						mov esi, pVarDesc
						assume esi:ptr VARDESC
						mov ecx, [esi].lpvarValue
						movzx eax, [ecx].VARIANT.vt
						.if (ax != varTemp.vt)
							push ecx
							invoke VariantChangeType, addr varTemp, pvarResult, 0, eax
							pop ecx
						.endif
						mov edx, varTemp.lVal
						.if (edx == [ecx].VARIANT.lVal)
							invoke vf(pTypeInfoRef, ITypeInfo, GetNames), [esi].memid, addr bstrVar, 1, addr dwTmp
							.if (eax == S_OK)
								invoke SysStringLen, bstrVar
								inc eax
								mov m_dwTextSize, eax
							    invoke WideCharToMultiByte,CP_ACP,0, bstrVar,-1, m_pszText, m_dwTextSize,0,0 
								invoke SysFreeString, bstrVar
                                mov hr, S_OK
							.endif
							mov bFound, TRUE
						.endif
						invoke vf(pTypeInfoRef, ITypeInfo, ReleaseVarDesc), pVarDesc
					.endif
					inc dwIndex
				.endw
				invoke VariantClear, addr varTemp
				invoke vf(pTypeInfoRef, ITypeInfo, ReleaseTypeAttr), pTypeAttr
			.endif
			invoke vf(pTypeInfoRef, ITypeInfo, Release)
		.endif
exit:
		mov eax, hr
		ret
		align 4
		assume edi:nothing
		assume esi:nothing

UserType2String endp


;--- translate a returned VARIANT to a string
;--- sets m_pszText + m_dwTextSize


Variant2String proc pFuncDesc:ptr FUNCDESC, pvarResult:ptr VARIANT, bErrTransform:BOOL

local dwRC:DWORD
;local dwError:DWORD
local varResult2:VARIANT

		invoke VariantInit, addr varResult2
		invoke VariantChangeType, addr varResult2, pvarResult, 0, VT_BSTR
		mov dwRC, eax
		.if (eax == S_OK)
			invoke SysStringLen, varResult2.bstrVal
			inc eax
			mov m_dwTextSize, eax
			invoke GetTextBufferPtr, eax
			.if (!eax)
				mov dwRC, E_OUTOFMEMORY
				jmp exit
			.endif
			invoke WideCharToMultiByte, CP_ACP, 0, varResult2.bstrVal, -1, m_pszText, m_dwTextSize, 0, 0
			mov ecx, pvarResult
			.if ([ecx].VARIANT.vt == VT_BOOL)
				mov ecx, m_pszText
				.if (word ptr [ecx] == "0")
					invoke lstrcpy, ecx, CStr("False")
				.else
					invoke lstrcpy, ecx, CStr("True")
				.endif
			.elseif (bErrTransform && g_bTranslateUDTs)
				invoke UserType2String, pFuncDesc, ecx
				.if ((eax != S_OK) && m_bHexadecimal)
					invoke GetVariant, pvarResult, m_pszText, m_dwTextMax, CStr("&H")
				.endif
			.elseif (m_bHexadecimal)
				invoke GetVariant, ecx, m_pszText, m_dwTextMax, CStr("&H")
			.endif
		.elseif (bErrTransform)
;			mov dwError, eax
			invoke GetTextBufferPtr, MAX_PATH
			mov eax, pvarResult
			movzx eax, [eax].VARIANT.vt
			.if (eax == VT_NULL)
				invoke lstrcpy, m_pszText, CStr("[VT_NULL]")
			.elseif (eax == VT_DISPATCH)
				invoke lstrcpy, m_pszText, CStr("[VT_DISPATCH]")
			.elseif (eax == VT_UNKNOWN)
				invoke lstrcpy, m_pszText, CStr("[VT_UNKNOWN]")
			.elseif (eax == (VT_ARRAY or VT_VARIANT))
				mov eax, pvarResult
				mov ecx, [eax].VARIANT.parray
				movzx edx, [ecx].SAFEARRAY.cDims
				push edx
				invoke wsprintf, m_pszText, CStr("[Array of Variants]")
				pop edx
				mov eax, pvarResult
				mov ecx, [eax].VARIANT.parray
				lea ecx, [ecx].SAFEARRAY.rgsabound
				.while (edx)
					pushad
					push [ecx].SAFEARRAYBOUND.cElements
					push [ecx].SAFEARRAYBOUND.lLbound
					invoke lstrlen, m_pszText
					mov ecx, m_pszText
					add ecx, eax
					pop eax
					pop edx
					invoke wsprintf, ecx, CStr("[%u-%u]"), eax, edx
					popad
					add ecx, sizeof SAFEARRAYBOUND
					dec edx
				.endw
			.elseif (eax == VT_ERROR)
				mov eax, pvarResult
				invoke wsprintf, m_pszText, CStr("[VT_ERROR(%X)]"), [eax].VARIANT.scode
			.else
				invoke GetVarType, eax
				invoke wsprintf, m_pszText, CStr("VariantChangeType Error[%X], Type=%s"), dwRC, eax
			.endif
			invoke lstrlen, m_pszText
			inc eax
			mov m_dwTextSize, eax
		.else
			invoke GetTextBufferPtr, MAX_PATH
			mov eax, m_pszText
			mov byte ptr [eax],0
			mov m_dwTextSize, 1
		.endif
exit:
		invoke VariantClear, addr varResult2
		return dwRC
		align 4

Variant2String endp

;--- HResult error occurred, set statusline 

ifdef @StackBase
	option stackbase:ebp
endif

PrepareInvokeErrorReturn proc uses esi hWndSB:HWND

local	dwESP:DWORD
local	dwSize:DWORD
local	pszHResult:LPSTR
local	pszFlags:LPSTR
local	pszDescription:LPSTR
local	pszText:LPSTR

		.if (m_wLastFlags & DISPATCH_PROPERTYGET)
			mov ecx, CStr("PropGet")
		.elseif (m_wLastFlags & DISPATCH_METHOD)
			mov ecx, CStr("Method")
		.elseif (m_wLastFlags & DISPATCH_PROPERTYPUT)
			mov ecx, CStr("PropPut")
		.else
			mov ecx, CStr("PropPutRef")
		.endif
		mov pszFlags, ecx
		mov dwESP, esp
		.if (m_LastHResult == DISP_E_EXCEPTION)
			lea esi, m_ExcepInfo
			.if ([esi].EXCEPINFO.bstrDescription)
				invoke SysStringLen, [esi].EXCEPINFO.bstrDescription
				mov dwSize, eax
				add eax, 4
				and al, 0FCh
				sub esp, eax
				mov pszDescription, esp
				invoke WideCharToMultiByte, CP_ACP, 0, [esi].EXCEPINFO.bstrDescription, -1, pszDescription, eax, 0, 0
			.else
				mov dwSize, 0
				mov pszDescription, offset g_szNull
			.endif
			mov ecx, dwSize
			add ecx, 64
			sub esp, ecx
			mov pszText, esp
			invoke wsprintf, pszText, CStr("'%s'[%X] Exception at Invoke(%s)"), pszDescription, [esi].EXCEPINFO.scode, pszFlags
			invoke FreeErrorStrings
		.else
			sub esp, 256
			mov pszText, esp
			invoke GetInvokeErrorString, m_LastHResult
			mov pszHResult, eax
			mov eax, m_LastHResult
			.if ((eax == DISP_E_TYPEMISMATCH) || (eax == DISP_E_PARAMNOTFOUND))
				invoke wsprintf, pszText, CStr("%s[%X] Error at Invoke(%s) [uArgErr=%u]"), pszHResult, m_LastHResult, pszFlags, m_dwArgErr
			.else
				invoke wsprintf, pszText, CStr("%s[%X] Error at Invoke(%s)"), pszHResult, m_LastHResult, pszFlags
			.endif
		.endif
		StatusBar_SetText hWndSB, 0, pszText
		StatusBar_SetTipText hWndSB, 0, pszText
		mov esp, dwESP
		invoke MessageBeep, MB_OK
		ret
		align 4

PrepareInvokeErrorReturn endp


;--- try to put a new value in property
;--- returns a HResult


PutProperty proc uses edi iItem:DWORD, pszText:LPSTR

local	varResult:VARIANT
local	varParm:VARIANT
local	dispid:DWORD
local	dwESP:DWORD
local	lvi:LVITEM
local	DispParams:DISPPARAMS
local	pFuncDesc:ptr FUNCDESC
local	dwRC:DWORD
local	szText[MAX_PATH]:byte

		DebugOut "PutProperty entered"

		invoke VariantInit, addr varParm
		invoke VariantInit, addr varResult

		mov eax, iItem
		mov lvi.iItem, eax
		mov lvi.mask_, LVIF_PARAM
		@mov lvi.iSubItem, 0
		invoke ListView_GetItem( m_hWndLV, addr lvi)

		invoke GetFuncDesc, lvi.lParam, addr pFuncDesc
		.if (eax != S_OK)
			mov dwRC, eax
			invoke MessageBeep, MB_OK
			jmp exit2
		.endif
		mov edi, pFuncDesc
		assume edi:ptr FUNCDESC

;---------------------------------- reload property without new request for parms
		mov g_bSkipCodeGen, 1
		invoke GetPropertyWithParam, edi, addr varResult
		.if (eax == S_OK)
			invoke Variant2String, edi, addr varResult, FALSE
			invoke VariantClear, addr varResult
			mov eax, S_OK
		.else
			push eax
			invoke FreeErrorStrings
			pop eax
		.endif

		mov dwESP, esp

		.if (eax == S_OK)
			invoke lstrcmp, m_pszText, pszText
;---------------------------------- item hasn't changed
			.if (eax == 0)
				mov dwRC, S_OK
;;				jmp exit3
				jmp setitem
			.endif
		.endif

		invoke SysStringFromLPSTR, pszText, 0
		mov varParm.bstrVal, eax
		mov varParm.vt, VT_BSTR

;---------------------------------- we have the property get funcdesc here!

		invoke GetRetType, edi
		invoke VariantChangeType, addr varParm, addr varParm, 0, eax

;---------------------------------- create temporarily a variant array on the stack
;---------------------------------- since we need 1 extra parameter for propertyput

		lea edx, varParm
		mov eax, m_ParamReturn.iNumVariants
		.if (eax && m_ParamReturn.pVariants)
			inc eax
			mov ecx, sizeof VARIANT
			mul ecx
			sub esp, eax
			mov edx, esp
			pushad
			mov edi, edx
			lea esi, varParm
			mov ecx, sizeof VARIANT/4
			rep movsd
			mov esi, m_ParamReturn.pVariants
			mov ecx, m_ParamReturn.iNumVariants
			shl ecx, 2
			rep movsd
			popad
		.endif
		mov DispParams.rgvarg, edx
		mov dispid, DISPID_PROPERTYPUT
		lea eax,dispid
		mov DispParams.rgdispidNamedArgs, eax
		mov DispParams.cNamedArgs, 1
		invoke GetNumArgs, edi
		inc eax
		mov DispParams.cArgs, eax
if ?CODEGEN
		invoke GenerateCode, edi, DISPATCH_PROPERTYPUT, addr DispParams, addr varResult
endif
		invoke ExecuteInvoke, [edi].memid, DISPATCH_PROPERTYPUT,
				addr DispParams, addr varResult
		mov dwRC, eax
		.if (eax == S_OK)

;---------------------------------- reload property without new request for parms

setitem:
			mov g_bSkipCodeGen, 1
			invoke GetPropertyWithParam, edi, addr varResult
			.if (eax == S_OK)
				invoke Variant2String, edi, addr varResult, TRUE
				invoke VariantClear, addr varResult
;;setitem:
				mov lvi.mask_, LVIF_TEXT
				mov lvi.iSubItem, 1
;;				invoke lstrcpyn, addr szText, m_pszText, sizeof szText
;;				lea eax, szText
				mov eax, m_pszText
				mov lvi.pszText, eax
				invoke ListView_SetItem( m_hWndLV, addr lvi)
				DebugOut "PutProperty: listview item set to %s", lvi.pszText
			.else
				invoke PrepareInvokeErrorReturn, m_hWndSB
			.endif
		.else
			invoke PrepareInvokeErrorReturn, m_hWndSB
		.endif
		mov esp, dwESP
		invoke VariantClear, addr varParm
		invoke VariantClear, addr varResult
exit3:
		invoke ReleaseFuncDesc, lvi.lParam, edi
exit2:
		return dwRC
		assume edi:nothing
		align 4

PutProperty endp

ifdef @StackBase
	option stackbase:esp
endif

;--- small dialog to show large BSTRs (property/return value of a function)
;--- (there is still the wsprintf limitation of 1024 characters)


viewdetailproc proc uses __this hWnd:HWND, message:DWORD, wParam:WPARAM, lParam:LPARAM

local	dwSize:DWORD
local	hWndEdit:HWND

		mov eax, message
		.if (eax == WM_INITDIALOG)
			mov __this, lParam
			invoke SetWindowLong, hWnd, DWL_USER, __this
			invoke GetDlgItem, hWnd, IDC_EDIT1
			mov hWndEdit, eax
			.if (m_bEdit)
				invoke SendMessage, hWndEdit, EM_SETREADONLY,  0, 0
			.endif
			invoke SetWindowText, hWndEdit, m_pszText
			mov eax, 1
		.elseif (eax == WM_CLOSE)
			invoke EndDialog, hWnd, 0
		.elseif (eax == WM_COMMAND)
			movzx eax, word ptr wParam+0
			.if (eax == IDCANCEL)
				invoke EndDialog, hWnd, 0
			.elseif (eax == IDOK)
				invoke GetDlgItem, hWnd, IDC_EDIT1
				mov hWndEdit, eax
				invoke GetWindowTextLength, hWndEdit
				inc eax
				mov dwSize, eax
				invoke malloc, eax
				.if (eax)
					push eax
					invoke GetWindowText, hWndEdit, eax, dwSize
					pop eax
				.endif
				invoke EndDialog, hWnd, eax
			.endif
		.else
			xor eax, eax
		.endif
		ret
		align 4

viewdetailproc endp

;--- check if typeinfo is available for IUnknown. If yes, open properties dlg

StartNewPropertiesDialog proc pTypeDesc:ptr TYPEDESC, pUnknown:LPUNKNOWN

local	pTypeInfo:LPTYPEINFO
local	rect:RECT
local	pPropertiesDlg:ptr CPropertiesDlg

		mov ecx, pTypeDesc
		invoke vf(m_pTypeInfo, ITypeInfo, GetRefTypeInfo), [ecx].TYPEDESC.hreftype, addr pTypeInfo
		.if (eax == S_OK)
			.if (g_bOwnWndForPropDlg)

				invoke Find@CObjectItem, pUnknown
				.if (eax)
					invoke vf(eax, IObjectItem, GetPropDlg)
					.if (eax)
						mov ecx, [eax].CDlg.hWnd
						.if (ecx == m_hWnd)
							StatusBar_SetText m_hWndSB, 0, CStr("nothing done (it's the same object)")
							invoke MessageBeep, MB_OK
						.else
							invoke RestoreAndActivateWindow, ecx
						.endif
						jmp exit
					.endif
				.endif

				invoke Create2@CPropertiesDlg, pUnknown, pTypeInfo
				.if (eax)
					mov pPropertiesDlg, eax
					invoke GetWindowRect, m_hWnd, addr rect
					add rect.left,20
					add rect.top,20
					invoke SetWindowPos@CPropertiesDlg, pPropertiesDlg, addr rect
					invoke GetWindow, m_hWnd, GW_OWNER
					invoke Show@CPropertiesDlg, pPropertiesDlg, eax
				.else
					invoke MessageBox, m_hWnd, CStr("No type info available. No dialog will be opened"), 0, MB_OK
				.endif
			.else
if ?POSTSETDISP
				invoke vf(pUnknown, IUnknown, AddRef)
				invoke PostMessage, m_hWnd, WM_SETDISPATCH, 0, pUnknown
else
				invoke SetDispatch@CPropertiesDlg, __this, pUnknown, TRUE
endif
			.endif
exit:
			invoke vf(pTypeInfo, IUnknown, Release)
		.endif
		ret
		align 4

StartNewPropertiesDialog endp


;--- display Variants of type VT_UNKNOWN + VT_SAFEARRAY.
;--- If object is IEnumVARIANT or SAFEARRAY, CCollectionDlg will handle
;--- this in a simple dialog, showing all item of collection/array.
;--- if VT_UNKNOWN isnt a collection, open a simple object dialog.

CheckIUnknown proc memId:DWORD, pVariant:ptr VARIANT

local varItem:VARIANT
local pr:PARAMRETURN
local szText[128]:byte

		mov ecx, pVariant
		.if ([ecx].VARIANT.punkVal)
			invoke Create@CCollectionDlg,  m_hWnd, m_pDispatch, memId, ecx, addr pr
			.if (g_bCollDlgAsTopLevelWnd)
				mov ecx, NULL
			.else
				mov ecx, m_hWnd
			.endif
			.if (eax)
if ?COLLECTIONMODELESS
				invoke CreateDialogParam, g_hInstance, IDD_COLLECTIONDLG, ecx, classdialogproc, eax
				.if (pr.iNumVariants == 0)
else
				invoke DialogBoxParam, g_hInstance, IDD_COLLECTIONDLG, ecx, classdialogproc, eax
				.if (!eax)
endif
					StatusBar_SetText m_hWndSB, 0, CStr("no items in collection")
					invoke MessageBeep, MB_OK
ife ?COLLECTIONMODELESS
				.else
					.if (pr.iCurItem != -1)
						invoke VariantInit, addr varItem
						mov eax, pr.pVariants
						mov ecx, pr.iCurItem
						shl ecx, 4
						add ecx, eax
						invoke VariantCopy, addr varItem, ecx
						invoke ParamReturnClear, addr pr
						.if ((varItem.vt == VT_DISPATCH) || (varItem.vt == VT_UNKNOWN))
							invoke StartNewDialog, -1, addr varItem
						.endif
						invoke VariantClear, addr varItem
					.else
						invoke ParamReturnClear, addr pr
					.endif
endif
				.endif
			.else
;-------------------------------------- object isnt a collection/array
				mov ecx, pVariant
				.if ([ecx].VARIANT.vt == VT_UNKNOWN)
if 0
					invoke vf([ecx].VARIANT.punkVal, IUnknown, QueryInterface), addr IID_DataSource, addr pDataSource
					.if (eax == S_OK)
						invoke wsprintf, addr szText, CStr("possibly a DataSource object, but that's not implemented yet")
						invoke vf(pDataSource, DataSource, Release)
					.else
						invoke wsprintf, addr szText, CStr("what to do with this IUnknown?")
					.endif
					StatusBar_SetText m_hWndSB, 0, addr szText
					invoke MessageBeep, MB_OK
else
					invoke GetRetElemDesc, edi
					.if ([eax].ELEMDESC.tdesc.vt == VT_PTR)
						mov eax, [eax].ELEMDESC.tdesc.lptdesc
					.endif
					.if ([eax].TYPEDESC.vt == VT_USERDEFINED)
						mov ecx, pVariant
						invoke StartNewPropertiesDialog, eax, [ecx].VARIANT.punkVal
					.else
						mov ecx, pVariant
						invoke Create@CObjectItem, [ecx].VARIANT.punkVal, NULL
						.if (eax)
							push eax
							invoke vf(eax, IObjectItem, ShowObjectDlg), m_hWnd
							pop eax
							invoke vf(eax, IObjectItem, Release)
						.endif
					.endif
endif
				.else
;--------------------------------- this shouldnt happen
					StatusBar_SetText m_hWndSB, 0, CStr("unexpected return from CCollectionDlg")
					invoke MessageBeep, MB_OK
				.endif
			.endif
		.else
			invoke wsprintf, addr szText, CStr("Method returned NULL")
			StatusBar_SetText m_hWndSB, 0, addr szText
			invoke MessageBeep, MB_OK
		.endif
		ret
		align 4

CheckIUnknown endp


;--- a IDispatch was returned. start new properties dialog. returns pCPropertiesDlg


StartNewDialog proc uses esi memId:DWORD, pVariant:ptr VARIANT

local pTypeInfo:LPTYPEINFO
local pTypeAttr:ptr TYPEATTR
local pDispatch:LPDISPATCH
local pUnknown:LPUNKNOWN
local rect:RECT
local pPropertiesDlg:ptr CPropertiesDlg
local szText[128]:byte

		mov pTypeInfo, NULL
		mov pTypeAttr, NULL
		mov pPropertiesDlg, NULL

		mov ecx, pVariant
		.if (([ecx].VARIANT.vt == VT_UNKNOWN) || (memId == DISPID_NEWENUM))
			invoke CheckIUnknown, memId, ecx
		.elseif ([ecx].VARIANT.vt == (VT_ARRAY or VT_VARIANT))
			invoke CheckIUnknown, memId, ecx
		.elseif ([ecx].VARIANT.vt == VT_DISPATCH)
			mov ecx, [ecx].VARIANT.pdispVal
			mov pDispatch, ecx
			.if (ecx == m_pDispatch)
				StatusBar_SetText m_hWndSB, 0, CStr("nothing done (it's the same object)")
				invoke MessageBeep, MB_OK
			.elseif (ecx)
				.if (g_bOwnWndForPropDlg)

					invoke vf(pDispatch, IUnknown, QueryInterface), addr IID_IUnknown, addr pUnknown
					.if (eax == S_OK)
						invoke Find@CObjectItem, pUnknown
						push eax
						invoke vf(pUnknown, IUnknown, Release)
						pop eax
						.if (eax)
							invoke vf(eax, IObjectItem, GetPropDlg)
							.if (eax)
								mov ecx, [eax].CDlg.hWnd
								.if (ecx == m_hWnd)
									StatusBar_SetText m_hWndSB, 0, CStr("nothing done (it's the same object)")
									invoke MessageBeep, MB_OK
								.else
									invoke RestoreAndActivateWindow, ecx
								.endif
								jmp exit
							.endif
						.endif
					.endif

					invoke Create2@CPropertiesDlg, pDispatch, pTypeInfo
					.if (eax)
						mov pPropertiesDlg, eax
						invoke GetWindowRect, m_hWnd, addr rect
						add rect.left,20
						add rect.top,20
						invoke SetWindowPos@CPropertiesDlg, pPropertiesDlg, addr rect
						invoke GetWindow, m_hWnd, GW_OWNER
						invoke Show@CPropertiesDlg, pPropertiesDlg, eax
					.else
						invoke MessageBox, m_hWnd, CStr("No type info available. No dialog will be opened"), 0, MB_OK
					.endif
				.else
if ?POSTSETDISP
					invoke vf(pDispatch, IUnknown, AddRef)
					invoke PostMessage, m_hWnd, WM_SETDISPATCH, 0, pDispatch
else
					invoke SetDispatch@CPropertiesDlg, __this, pDispatch, TRUE
endif
				.endif
			.else
				invoke wsprintf, addr szText, CStr("Method returned NULL")
				StatusBar_SetText m_hWndSB, 0, addr szText
				invoke MessageBeep, MB_OK
			.endif
		.else
;------------------------------------- shouldnt happen
			movzx eax, [ecx].VARIANT.vt
			invoke GetVarType, eax
			invoke wsprintf, addr szText, CStr("Method returned variant type VT_%s"), eax
			StatusBar_SetText m_hWndSB, 0, addr szText
			invoke MessageBeep, MB_OK
		.endif
exit:
		.if (pTypeInfo)
			.if (pTypeAttr)
				invoke vf(pTypeInfo, ITypeInfo, ReleaseTypeAttr), pTypeAttr
			.endif
			invoke vf(pTypeInfo, ITypeInfo, Release)
		.endif
		return pPropertiesDlg
		assume esi:nothing
		align 4

StartNewDialog endp


;--- execute a method, parameters are read into a VARIANT array
;--- this proc may be a callback for CParamsDialog if methods has params


ExecuteMethodCB proc uses esi edi __this this_:ptr CPropertiesDlg, hWnd:HWND, pFuncDesc:ptr FUNCDESC, iNumVariants:DWORD, pVariants:ptr VARIANT, hWndSB:HWND

local pResult:ptr VARIANT
local dwSize:DWORD
local DispParams:DISPPARAMS
local varResult:VARIANT
local varTemp:VARIANT
local szText[MAX_PATH]:byte


		mov __this, this_

		mov edi, pFuncDesc
		assume edi:ptr FUNCDESC

		invoke VariantInit, addr varResult
;------------------------------------------------------ if its a method, call it
		mov ecx, iNumVariants
		mov DispParams.cArgs, ecx
		mov eax, pVariants
		mov DispParams.rgvarg, eax
		@mov DispParams.cNamedArgs, 0
		@mov DispParams.rgdispidNamedArgs,NULL

		invoke GetRetType, edi
		.if (eax == VT_VOID)
			mov esi, NULL
		.else
			lea esi, varResult
		.endif
		mov pResult, esi

		StatusBar_SetText hWndSB, 0, CStr("Executing")

;------------------------------- this call may yield control!
;------------------------------- so disable window temporarily
		invoke EnableWindow, hWnd, FALSE

if ?CODEGEN
		invoke GenerateCode, edi, DISPATCH_METHOD, addr DispParams, pResult
endif
		invoke ExecuteInvoke, [edi].memid, DISPATCH_METHOD,
				addr DispParams, pResult

;------------------------------- reenable window
		push eax
		invoke EnableWindow, hWnd, TRUE
		invoke SetActiveWindow, hWnd
		StatusBar_SetText hWndSB, 0, CStr("")
		pop eax

		.if (eax == S_OK)
			movzx eax, varResult.vt
;----------------------------------------- opening new properties dialog
;----------------------------------------- due to methods return type may be
;----------------------------------------- annoying, so make it optional
			.if (eax == VT_UNKNOWN)
				invoke StartNewDialog, [edi].memid, addr varResult
			.elseif ((eax == VT_DISPATCH) && (g_bNewDlgForMethods))
				invoke StartNewDialog, [edi].memid, addr varResult
			.elseif ((eax == (VT_ARRAY or VT_VARIANT)) && (g_bNewDlgForMethods))
				invoke StartNewDialog, [edi].memid, addr varResult
			.elseif (eax != VT_EMPTY)
				;this function is not excluded from FPO, hence DO NOT
				;use stack variables once ESP has been modified!
				push esi
				mov esi, esp
				invoke Variant2String, edi, addr varResult, TRUE
				mov eax, m_dwTextSize
				add eax, 24
				and al, 0FCh
				sub esp, eax
				mov edx,esp
				invoke wsprintf, edx, CStr("'%s' returned"), m_pszText
				.if (eax > ?MAXTEXTINPLACE)
					invoke DialogBoxParam, g_hInstance, IDD_VIEWDETAIL, m_hWnd, viewdetailproc, __this
				.else
					StatusBar_SetText m_hWndSB, 0, esp
					StatusBar_SetTipText m_hWndSB, 0, esp
				.endif
				mov esp, esi
				pop esi
			.else
				StatusBar_SetText m_hWndSB, 0, CStr("Ok")
			.endif
if 1
			mov esi, pVariants
			mov ecx, DispParams.cArgs
			.while (ecx)
				push ecx
				.if ([esi].VARIANT.vt & VT_BYREF)
					.if ([esi].VARIANT.vt == (VT_BYREF or VT_DISPATCH))
						invoke VariantInit, addr varTemp
						mov edx, [esi].VARIANT.byref
						xor ecx, ecx
						xchg ecx, [edx]
						mov varTemp.pdispVal, ecx
						mov varTemp.vt, VT_DISPATCH
						.if (g_bNewDlgForMethods)
							invoke StartNewDialog, -1, addr varTemp
						.endif
						invoke VariantClear, addr varTemp
					.elseif ([esi].VARIANT.vt == (VT_BYREF or VT_UNKNOWN))
						mov edx, [esi].VARIANT.byref
						xor ecx, ecx
						xchg ecx, [edx]
						mov varTemp.pdispVal, ecx
						mov varTemp.vt, VT_UNKNOWN
						.if (g_bNewDlgForMethods)
							invoke StartNewDialog, -1, addr varTemp
						.endif
						invoke VariantClear, addr varTemp
					.elseif ([esi].VARIANT.vt == (VT_BYREF or VT_BSTR))
						mov edx, [esi].VARIANT.byref
						xor eax, eax
						xchg eax, [edx]
						push eax
						invoke printf@CLogWindow, CStr("%u. parameter returned BSTR: '%.128S'",10), ecx, eax
						pop eax
						invoke SysFreeString, eax
					.endif 
				.endif
				pop ecx
				add esi, sizeof VARIANT
				dec ecx
			.endw
endif
		.else
			invoke PrepareInvokeErrorReturn, hWndSB
		.endif

		invoke VariantClear, addr varResult
		return m_LastHResult
		align 4

ExecuteMethodCB endp


;--- try to execute a INVOKE_FUNC function description
;--- if the function needs parameters, create a CParamsDlg and get these


ExecuteMethod proc pFuncDesc:ptr FUNCDESC

		invoke GetNumArgs, pFuncDesc
		.if (eax)
			invoke Create@CParamsDlg, m_pTypeInfo, pFuncDesc, __this, offset ExecuteMethodCB
			invoke DialogBoxParam, g_hInstance, IDD_PARAMSDLG, m_hWnd, classdialogproc, eax
		.else
			invoke ExecuteMethodCB, __this, m_hWnd, pFuncDesc, 0, NULL, m_hWndSB
		.endif
		ret
		align 4

ExecuteMethod endp


GetDispatchPtrCB proc uses esi edi __this this_:ptr CPropertiesDlg, hWnd:HWND, pFuncDesc:ptr FUNCDESC, iNumVariants:DWORD, pVariants:ptr VARIANT

local dwRC:DWORD
local DispParams:DISPPARAMS
local varResult:VARIANT
local szText[128]:byte


		mov __this, this_
		ResetDispParams DispParams
		mov eax, pVariants
		mov DispParams.rgvarg, eax
		mov ecx, iNumVariants
		mov DispParams.cArgs, ecx

		mov edi, pFuncDesc
		assume edi:ptr FUNCDESC

		invoke VariantInit, addr varResult
;--------------------------------------- get the property

if ?CODEGEN
		invoke GenerateCode, edi, DISPATCH_PROPERTYGET, addr DispParams, addr varResult
endif
		invoke ExecuteInvoke, [edi].memid, DISPATCH_PROPERTYGET,
			addr DispParams, addr varResult
		.if (eax == S_OK)
			invoke StartNewDialog, [edi].memid, addr varResult
		.else
			invoke PrepareInvokeErrorReturn, m_hWndSB
		.endif
		mov dwRC, TRUE
exit:
		invoke VariantClear, addr varResult
		return dwRC
		assume edi:nothing
		align 4

GetDispatchPtrCB endp


;--- FUNCDESC returntype must be VT_PTR+VT_USERDEFINED/VT_DISPATCH/VT_UNKNOWN


GetDispatchPtr proc pFuncDesc:ptr FUNCDESC

		invoke GetNumArgs, pFuncDesc
		.if (eax)
			invoke Create@CParamsDlg, m_pTypeInfo, pFuncDesc, __this, offset GetDispatchPtrCB
			invoke DialogBoxParam, g_hInstance, IDD_PARAMSDLG, m_hWnd, classdialogproc, eax
		.else
			invoke GetDispatchPtrCB, __this, m_hWnd, pFuncDesc, 0, NULL
		.endif
		ret
		align 4

GetDispatchPtr endp


;--- recursive called: check VT_USERDEFINED type


CheckIt proc pTypeInfo:LPTYPEINFO, hRefType:HREFTYPE

local dwRC:DWORD
local pTypeInfoRef:LPTYPEINFO
local pTypeAttr:ptr TYPEATTR

		mov dwRC, FALSE

		invoke vf(pTypeInfo, ITypeInfo, GetRefTypeInfo),
				hRefType, addr pTypeInfoRef
		.if (eax == S_OK)
			invoke vf(pTypeInfoRef, ITypeInfo, GetTypeAttr), addr pTypeAttr
			.if (eax == S_OK)
				mov eax, pTypeAttr
				mov ecx, [eax].TYPEATTR.typekind
				.if (ecx == TKIND_ALIAS)
					movzx ecx, [eax].TYPEATTR.tdescAlias.vt
					.if (ecx == VT_USERDEFINED)
						invoke CheckIt, pTypeInfoRef, [eax].TYPEATTR.tdescAlias.hreftype
						mov dwRC, eax
					.elseif ((ecx == VT_DISPATCH) || (ecx == VT_UNKNOWN))
						mov dwRC, TRUE
					.endif
				.elseif ((ecx == TKIND_DISPATCH) || (ecx == TKIND_COCLASS) || (ecx == TKIND_INTERFACE))
					mov dwRC, TRUE
				.endif
				invoke vf(pTypeInfoRef, ITypeInfo, ReleaseTypeAttr), pTypeAttr
			.endif
			invoke vf(pTypeInfoRef, ITypeInfo, Release)
		.else
;------------------------------------ this shouldnt happen (but does in MSACCESS!)
			sub esp, MAX_PATH
			mov edx, esp
			invoke wsprintf, edx, CStr("Error %X at ITypeInfo::GetRefTypeInfo"), eax
			mov edx, esp
			StatusBar_SetText m_hWndSB, 0, edx
			add esp, MAX_PATH
			invoke MessageBeep, MB_OK
		.endif
		return dwRC
		align 4

CheckIt endp

;--- check if VT_PTR is a pointer to VT_DISPATCH/VT_UNKNOWN
;--- if function is virtual, check for 2 indirections (PTR PTR xxxx)

IsDispatchPtr proc uses esi edi pFuncDesc:ptr FUNCDESC

local dwRC:DWORD
local pTypeInfo:LPTYPEINFO
local pTypeAttr:ptr TYPEATTR

		mov dwRC, FALSE

		mov edi, pFuncDesc
		assume edi:ptr FUNCDESC

		invoke GetRetElemDesc, edi
		mov esi,[eax].ELEMDESC.tdesc.lptdesc
		assume esi:ptr ELEMDESC

		movzx eax, [esi].tdesc.vt
if ?PUREINDIRECTION eq 0
		.if (eax == VT_PTR)
			.if (m_bVtblMode)
				mov esi,[esi].tdesc.lptdesc
				movzx eax, [esi].tdesc.vt
			.endif
		.endif
endif
		.if ((eax == VT_DISPATCH) || (eax == VT_UNKNOWN))
			mov eax, TRUE
		.elseif (eax == VT_USERDEFINED)
			invoke CheckIt, m_pTypeInfo, [esi].tdesc.hreftype
		.endif
		ret
		assume edi:nothing
		assume esi:nothing
		align 4

IsDispatchPtr endp

;--- is type a enum type

IsEnumType proc iItem:DWORD

local pFuncDesc:ptr FUNCDESC
local pTypeInfoRef:LPTYPEINFO
local pTypeAttr:ptr TYPEATTR
local bRC:BOOL

		mov bRC, FALSE
		invoke GetFuncDesc2, iItem, addr pFuncDesc
		.if (eax == S_OK)
			invoke GetRetElemDesc, pFuncDesc
			.if ([eax].TYPEDESC.vt == VT_BOOL)
				mov bRC, TRUE
			.elseif ([eax].TYPEDESC.vt == VT_USERDEFINED)
				lea ecx, pTypeInfoRef
				invoke vf(m_pTypeInfo, ITypeInfo, GetRefTypeInfo), [eax].TYPEDESC.hreftype, ecx
				.if (eax == S_OK)
					invoke vf(pTypeInfoRef, ITypeInfo, GetTypeAttr), addr pTypeAttr
					.if (eax == S_OK)
						mov ecx, pTypeAttr
						.if ([ecx].TYPEATTR.typekind == TKIND_ENUM)
							mov bRC, TRUE
						.endif
						invoke vf(pTypeInfoRef, ITypeInfo, ReleaseTypeAttr), pTypeAttr
					.endif
					invoke vf(pTypeInfoRef, ITypeInfo, Release)
				.endif
			.endif
			invoke ReleaseFuncDesc2, iItem, pFuncDesc
		.endif
		return bRC
		align 4

IsEnumType endp


;--- WM_NOTIFY/NM_DBLCLK: doubleclick in listview

;--- dwMode:
;---  0 = NM_DBLCLK
;---  1 = Return pressed (IDOK)
;---  2 = Edit/Execute selected in context menu

OnEdit proc uses edi dwMode:DWORD

local pFuncDesc:ptr FUNCDESC
local pVarDesc:ptr FUNCDESC
local varResult:VARIANT
local bEnum:BOOL
local iItem:DWORD
local tmpFuncDesc:FUNCDESC
local szText[64]:byte

		StatusBar_SetText m_hWndSB, 0, CStr("")

		invoke ListView_GetNextItem( m_hWndLV, -1, LVNI_SELECTED)
		.if (eax != -1)
			mov iItem, eax
			invoke GetFuncDesc2, iItem, addr pFuncDesc
			.if (eax == S_OK)
				mov edi,pFuncDesc
				assume edi:ptr FUNCDESC
				invoke GetRetType, edi
				mov ecx, eax
;------------------------------------------------------ it its a method, call it
;------------------------------------------------------ else edit the property
				.if ([edi].invkind == INVOKE_FUNC)
					invoke ExecuteMethod, edi

				.elseif ((ecx == VT_DISPATCH) || (ecx == VT_UNKNOWN))
					invoke GetDispatchPtr, edi
				.elseif (ecx == (VT_ARRAY or VT_VARIANT))
					invoke GetDispatchPtr, edi
				.elseif ((ecx == VT_BOOL) && (dwMode == 0))
					invoke ParamReturnClear, addr m_ParamReturn
					invoke GetPropertyWithParam, edi, addr varResult
					.if (eax == S_OK)
;----------------------------------------- display cur value in message box
;----------------------------------------- since we dont have a UI for result
						invoke GetNumArgs, edi
						.if (eax)
							.if (!varResult.boolVal)
								mov ecx, CStr("Current value is False")
							.else
								mov ecx, CStr("Current value is True")
							.endif
							invoke MessageBox, m_hWnd, ecx, CStr("Result"), MB_OK
						.endif
						.if (varResult.boolVal == 0)
							mov ecx, CStr("-1")
						.else
							mov ecx, CStr("0")
						.endif
						invoke PutProperty, iItem, ecx
					.elseif (eax != -1)
						invoke PrepareInvokeErrorReturn, m_hWndSB
					.endif
					invoke VariantClear, addr varResult
				.else
					xor eax, eax
					.if (ecx == VT_PTR)
						invoke IsDispatchPtr, edi
						.if (eax)
							invoke GetDispatchPtr, edi
						.else
							jmp label1
						.endif
					.else
label1:
						invoke IsEnumType, iItem
						.if (eax)
							invoke SendMessage, m_hWndLV, LVM_COMBOBOXMODE, iItem, 1
						.else
							invoke SendMessage, m_hWndLV, LVM_EDITLABEL, iItem, 1
						.endif
					.endif
				.endif
				invoke ReleaseFuncDesc2, iItem, edi
			.else
				invoke wsprintf, addr szText, CStr("ITypeInfo::GetFuncDesc failed[%X]"), eax
				StatusBar_SetText m_hWndSB, 0, addr szText
				invoke MessageBeep, MB_OK
			.endif
		.endif
		ret
		assume edi:nothing
		align 4

OnEdit endp

;--- copy value of current property to the clipboard

OnCopyValue proc

local	szText[2048]:byte

		invoke ListView_GetNextItem( m_hWndLV, -1, LVNI_SELECTED)
		.if (eax != -1)
			lea ecx, szText
			ListView_GetItemText m_hWndLV, eax, 1, ecx, sizeof szText
			invoke CopyStringToClipboard, m_hWndLV, addr szText
		.endif
		ret
		align 4

OnCopyValue endp

;--- translate a string back into an UDT enum value 

TranslateUDT proc uses esi edi iItem:DWORD, pszText:LPSTR, pszOut:LPSTR, iMax:DWORD

local	pTypeInfoRef:LPTYPEINFO
local	pTypeAttr:ptr TYPEATTR
local	pVarDesc:ptr VARDESC
local	pFuncDesc:ptr FUNCDESC
local	dwIndex:DWORD
local	dwTmp:DWORD
local	varResult:VARIANT
local	bstrVar:BSTR
local	wszText[MAX_PATH]:WORD

		invoke GetFuncDesc2, iItem, addr pFuncDesc
		.if (eax != S_OK)
			jmp exit
		.endif
		invoke GetRetElemDesc, pFuncDesc
		mov esi, eax
		assume esi:ptr ELEMDESC
		.if ([esi].tdesc.vt == VT_BOOL)
			invoke lstrcmp, pszText, CStr("True")
			.if (!eax)
				mov ecx, pszOut
				mov dword ptr [ecx], "1-"
			.else
				invoke lstrcmp, pszText, CStr("False")
				.if (!eax)
					mov ecx, pszOut
					mov dword ptr [ecx], "0"
				.endif
			.endif
			jmp exit
		.endif

		invoke MultiByteToWideChar,CP_ACP,MB_PRECOMPOSED,
			pszText,-1,addr wszText, LENGTHOF wszText 

		invoke vf(m_pTypeInfo, ITypeInfo, GetRefTypeInfo), [esi].tdesc.hreftype, addr pTypeInfoRef
		.if (eax == S_OK)
			invoke vf(pTypeInfoRef, ITypeInfo, GetTypeAttr), addr pTypeAttr
			.if (eax == S_OK)
				mov dwIndex, 0
;------------------------------- scan the vartab , get the DISPID
				.while (1)
					mov eax, dwIndex
					mov edx, pTypeAttr
					.break .if (ax >= [edx].TYPEATTR.cVars)
					invoke vf(pTypeInfoRef, ITypeInfo, GetVarDesc), dwIndex, addr pVarDesc
					.if (eax == S_OK)
						mov ecx, pVarDesc
						invoke vf(pTypeInfoRef, ITypeInfo, GetNames), [ecx].VARDESC.memid, addr bstrVar, 1, addr dwTmp
						.if (eax == S_OK)
							invoke _strcmpW, addr wszText, bstrVar
							push eax
							invoke SysFreeString, bstrVar
							pop eax
						.else
							mov eax, 1
						.endif
						.if (eax == 0)
							invoke VariantInit, addr varResult
							mov ecx, pVarDesc
							invoke VariantChangeType, addr varResult, [ecx].VARDESC.lpvarValue, 0, VT_BSTR
							invoke WideCharToMultiByte, CP_ACP, 0, varResult.bstrVal, -1, pszOut, iMax, 0, 0
							invoke vf(pTypeInfoRef, ITypeInfo, ReleaseVarDesc), pVarDesc
							.break
						.endif
						invoke vf(pTypeInfoRef, ITypeInfo, ReleaseVarDesc), pVarDesc
					.endif
					inc dwIndex
				.endw
				invoke vf(pTypeInfoRef, ITypeInfo, ReleaseTypeAttr), pTypeAttr
			.endif
			invoke vf(pTypeInfoRef, ITypeInfo, Release)
		.endif
		invoke ReleaseFuncDesc2, iItem, pFuncDesc
exit:
		ret
		assume esi:nothing
		assume edi:nothing
		align 4

TranslateUDT endp

;--- fill a combobox

FillComboBox proc uses esi pFuncDesc:ptr FUNCDESC, hWndCB:HWND

local	pTypeInfoRef:LPTYPEINFO
local	pTypeAttr:ptr TYPEATTR
local	pVarDesc:ptr VARDESC
local	dwSize:DWORD
local	varResult:VARIANT
local	dwIndex:DWORD
local	dwTmp:DWORD
local	bstrVar:BSTR
local	szText[MAX_PATH]:byte

		invoke GetRetElemDesc, pFuncDesc
		mov esi, eax
		assume esi:ptr ELEMDESC

		.if ([esi].tdesc.vt == VT_BOOL)
			invoke ComboBox_AddString( hWndCB, CStr("False"))
			invoke ComboBox_AddString( hWndCB, CStr("True"))
			jmp done
		.endif

		DebugOut "FillComboBox, pFuncDesc=%X, hreftype=%X", pFuncDesc, [esi].tdesc.hreftype

		invoke vf(m_pTypeInfo, ITypeInfo, GetRefTypeInfo), [esi].tdesc.hreftype, addr pTypeInfoRef
		.if (eax == S_OK)
			invoke vf(pTypeInfoRef, ITypeInfo, GetTypeAttr), addr pTypeAttr
			.if (eax == S_OK)
				mov edx, pTypeAttr
				movzx ecx, [edx].TYPEATTR.cVars
				@mov dwIndex, 0
				.while (ecx)
					push ecx
					invoke vf(pTypeInfoRef, ITypeInfo, GetVarDesc), dwIndex, addr pVarDesc
					.if (eax == S_OK)
						mov ecx, pVarDesc
						invoke vf(pTypeInfoRef, ITypeInfo, GetNames), [ecx].VARDESC.memid, addr bstrVar, 1, addr dwTmp
						.if (eax == S_OK)
							invoke VariantInit, addr varResult
							invoke SysStringLen, bstrVar
							add eax, 4
							and al,0FCh
							mov dwSize, eax
							mov ecx, pVarDesc
							.if ([ecx].VARDESC.varkind == VAR_CONST)
								invoke VariantChangeType, addr varResult, [ecx].VARDESC.lpvarValue, 0, VT_BSTR
								.if (eax == S_OK)
									invoke SysStringLen, varResult.bstrVal
									add eax, 16
									and al, 0FCh
									add dwSize, eax
								.endif
							.endif
							sub esp, dwSize
							mov ecx, esp
							.if (g_bTranslateUDTs)
							    invoke wsprintf, ecx, CStr("%S"), bstrVar
							.else
							    invoke wsprintf, ecx, CStr("%S"), varResult.bstrVal
							.endif
							invoke ComboBox_AddString( hWndCB, esp)
							add esp, dwSize
							invoke VariantClear, addr varResult
							invoke SysFreeString, bstrVar
						.endif
						invoke vf(pTypeInfoRef, ITypeInfo, ReleaseVarDesc), pVarDesc
					.endif
					pop ecx
					inc dwIndex
					dec ecx
				.endw
				invoke vf(pTypeInfoRef, ITypeInfo, ReleaseTypeAttr), pTypeAttr
			.endif
			invoke vf(pTypeInfoRef, ITypeInfo, Release)
		.else
			invoke wsprintf, addr szText, CStr("ITypeInfo::GetRefTypeInfo failed [%X]"), eax
			StatusBar_SetText m_hWndSB, 0, addr szText
		.endif
done:
		ret
		assume esi:nothing
		align 4

FillComboBox endp

OnBeginLabelEdit proc uses edi pNMLVDI:ptr NMLVDISPINFO

local pFuncDesc:ptr FUNCDESC
local hWndEdit:HWND
local hWndCB:HWND
local dwRC:DWORD
local varResult:VARIANT
local iItem:DWORD

		mov dwRC, FALSE
		mov ecx, pNMLVDI
		mov ecx, [ecx].NMLVDISPINFO.item.iItem
		mov iItem, ecx
		invoke GetFuncDesc2, iItem, addr pFuncDesc
		DebugOut "OnBeginLabelEdit, iItem=%X, pFuncDesc=%X", iItem, pFuncDesc
		.if (eax == S_OK)
			invoke SendMessage, m_hWndLV, LVM_GETEDITCONTROL, 0, 0
			mov hWndEdit, eax
			mov edi, pFuncDesc
			assume edi:ptr FUNCDESC
			invoke GetRetType, edi
			mov ecx, eax
;----------------------------------------- if it is a function, dont allow editing
			.if ([edi].invkind == INVOKE_FUNC)
				invoke SetWindowLong, m_hWnd, DWL_MSGRESULT, 1
				mov dwRC, TRUE
;----------------------------------------- same for IDispatch/IUnknown properties
;;			.elseif ((ecx == VT_DISPATCH) || (ecx == VT_UNKNOWN) || (ecx == VT_BOOL))
			.elseif ((ecx == VT_DISPATCH) || (ecx == VT_UNKNOWN))
				invoke SetWindowLong, m_hWnd, DWL_MSGRESULT, 1
				mov dwRC, TRUE
;----------------------------------------- possibly same for ptr UDT 
			.elseif (ecx == VT_PTR)
				invoke IsDispatchPtr, edi
				.if (eax)
					invoke SetWindowLong, m_hWnd, DWL_MSGRESULT, 1
					mov dwRC, TRUE
if 1
				.else
					mov m_pszText, 0
					invoke SetWindowText, hWndEdit, m_pszText
endif
				.endif
			.else
;----------------------------------------- now get the property value
				invoke ParamReturnClear, addr m_ParamReturn
				invoke GetPropertyWithParam, edi, addr varResult
				.if (eax == S_OK)
;----------------------------------------- it may be a IDispatch/IUnknown in a VARIANT!
					movzx eax, varResult.vt
					.if ((eax == VT_DISPATCH) || (eax == VT_UNKNOWN))
						invoke StartNewDialog, [edi].memid, addr varResult
						invoke SetWindowLong, m_hWnd, DWL_MSGRESULT, 1
						mov dwRC, TRUE
					.elseif (eax == (VT_ARRAY or VT_VARIANT))
						invoke StartNewDialog, [edi].memid, addr varResult
						invoke SetWindowLong, m_hWnd, DWL_MSGRESULT, 1
						mov dwRC, TRUE
					.else
						invoke SendMessage, m_hWndLV, LVM_GETCOMBOBOXCONTROL, 0, 0
						mov hWndCB, eax
						.if (eax)
							invoke Variant2String, edi, addr varResult, TRUE
							invoke SetWindowText, hWndCB, m_pszText
							invoke FillComboBox, edi, hWndCB
							invoke ComboBox_SelectString( hWndCB, -1, m_pszText)
						.else
							invoke Variant2String, edi, addr varResult, FALSE
							mov edx, m_pszText
							mov ecx, m_dwTextSize
							.while (ecx)
								mov al,[edx]
								.break .if ((al == 13) || (al == 10))
								inc edx
								dec ecx
							.endw
							.if (ecx || (m_dwTextSize > ?MAXTEXTINPLACE))
								mov m_bEdit, TRUE
								invoke DialogBoxParam, g_hInstance, IDD_VIEWDETAIL, m_hWnd, viewdetailproc, __this
								mov m_bEdit, FALSE
								.if (eax)
									push eax
									invoke PutProperty, iItem, eax
									pop eax
									invoke free, eax
								.endif
								invoke SetWindowLong, m_hWnd, DWL_MSGRESULT, 1
								mov dwRC, TRUE
							.else
								invoke SetWindowText, hWndEdit, m_pszText
							.endif
						.endif
					.endif
				.elseif (eax == -1)		;has user canceled?
					invoke SetWindowLong, m_hWnd, DWL_MSGRESULT, 1
					mov dwRC, TRUE
				.else
					invoke PrepareInvokeErrorReturn, m_hWndSB
					mov m_pszText, 0
					invoke SetWindowText, hWndEdit, m_pszText
				.endif
				invoke VariantClear, addr varResult
			.endif
			invoke ReleaseFuncDesc2, iItem, edi
		.else
			invoke SetWindowLong, m_hWnd, DWL_MSGRESULT, 1
			mov dwRC, TRUE
		.endif
		return dwRC
		assume edi:nothing
		align 4

OnBeginLabelEdit endp


OnEndLabelEdit proc  pNMLVDI:ptr NMLVDISPINFO

local bEnum:BOOL
local dwRC:DWORD
local iItem:DWORD
local iSubItem:DWORD
local pszText:LPSTR
local szText[64]:byte

		mov dwRC, FALSE
		
		mov eax, pNMLVDI
		mov edx, [eax].NMLVDISPINFO.item.iSubItem
		mov ecx, [eax].NMLVDISPINFO.item.iItem
		mov iSubItem, edx
		mov iItem, ecx
		mov edx, [eax].NMLVDISPINFO.item.pszText
		mov pszText, edx

		.if (edx)
					
			invoke IsEnumType, iItem
			mov bEnum, eax
			.if (eax)
				mov szText, 0
				mov ecx, pNMLVDI
				.if (g_bTranslateUDTs)
					invoke TranslateUDT, iItem, pszText, addr szText, sizeof szText
				.endif
				.if (szText)
					lea edx, szText
				.else
					mov edx, pszText
				.endif
			.else
				mov edx, pszText
			.endif
			invoke PutProperty, iItem, edx
;--------------------------------- PutProperty should have set new value of listview item
;--------------------------------- so return FALSE
			.if (eax == S_OK)
				invoke SetWindowLong, m_hWnd, DWL_MSGRESULT, 0
				mov dwRC, TRUE
			.else
if ?COMBOBOX
				.if (bEnum)
					invoke SendMessage, m_hWndLV, LVM_COMBOBOXMODE,	iItem, iSubItem
				.else
					invoke SendMessage, m_hWndLV, LVM_EDITLABEL, iItem, iSubItem
				.endif
else
				invoke SendMessage, m_hWndLV, LVM_EDITLABEL, iItem, iSubItem
endif
				invoke SendMessage, m_hWndLV, LVM_GETEDITCONTROL, 0, 0
				invoke SetWindowText, eax, pszText
			.endif
		.endif
		return dwRC
		align 4

OnEndLabelEdit endp


;--- WM_NOTIFY


OnNotify proc uses edi pNMHdr:ptr NMHDR

local hti:LVHITTESTINFO
local pTypeAttr:ptr TYPEATTR
local pFuncDesc:ptr FUNCDESC
local dwRC:DWORD

;;		DebugOut "OnNotify start"
		mov dwRC, FALSE
		mov eax, pNMHdr
		.if ([eax].NMHDR.idFrom == IDC_TAB1)

			mov eax, [eax].NMHDR.code
			.if (eax == TCN_SELCHANGE)
				
				mov m_iSortCol, -1
				mov m_iSortDir, 0

				invoke TabCtrl_GetCurSel( m_hWndTab)
				invoke RefreshView, eax
			.elseif (eax == TCN_SELCHANGING)
if ?EDITBUG
;----------------------------- edit control for in-place-editing
;----------------------------- receives no WM_KILLFOCUS message if user
;----------------------------- clicks a tab. so we check if edit mode
;----------------------------- is active and force it to terminate
				invoke SendMessage, m_hWndLV, LVM_GETEDITCONTROL, 0, 0
				.if (eax)
					invoke SendMessage, m_hWndLV, LVM_ENDEDITMODE, 0, 0
					.if (eax)
						invoke SetWindowLong, m_hWnd, DWL_MSGRESULT, 1
						mov dwRC, TRUE
					.endif
				.endif
endif
			.endif

		.elseif ([eax].NMHDR.idFrom == IDC_LIST1)

			.if ([eax].NMHDR.code == NM_RCLICK)

				invoke ShowContextMenu, TRUE

			.elseif ([eax].NMHDR.code == NM_DBLCLK)

				invoke OnEdit, 0

			.elseif ([eax].NMHDR.code == NM_CLICK)

;-------------------------------------- if user clicked in second column
;-------------------------------------- try to enter in-place-edit mode
				.if (m_iTab == TAB_PROPERTIES)
					invoke GetCursorPos, addr hti.pt
					invoke ScreenToClient, m_hWndLV, addr hti.pt
					invoke ListView_SubItemHitTest( m_hWndLV, addr hti)
					.if (hti.iSubItem == 1)
						invoke ListView_GetItemCount( m_hWndLV)
						.if (eax > hti.iItem)
if ?COMBOBOX
							invoke IsEnumType, hti.iItem
							.if (eax)
								invoke SendMessage, m_hWndLV, LVM_COMBOBOXMODE, hti.iItem, hti.iSubItem
							.else
								invoke SendMessage, m_hWndLV, LVM_EDITLABEL, hti.iItem, hti.iSubItem
							.endif
else
							invoke SendMessage, m_hWndLV, LVM_EDITLABEL, hti.iItem, hti.iSubItem
endif
						.endif
					.endif
				.endif

			.elseif ([eax].NMHDR.code == LVN_KEYDOWN)

				invoke GetKeyState, VK_CONTROL
				and al, 80h
				mov eax, pNMHdr
				.if (ZERO?)
					.if ([eax].NMLVKEYDOWN.wVKey == VK_APPS)
						invoke ShowContextMenu, FALSE
					.elseif ([eax].NMLVKEYDOWN.wVKey == VK_F6)
						invoke Create@CObjectItem, m_pTypeInfo, NULL
						.if (eax)
							push eax
							invoke vf(eax, IObjectItem, ShowObjectDlg), m_hWnd
							pop eax
							invoke vf(eax, IObjectItem, Release)
						.endif
					.endif
				.else
					.if ([eax].NMLVKEYDOWN.wVKey == 'U')
						xor g_bTranslateUDTs, 1
						.if (g_hWndOption)
							invoke PostMessage, g_hWndOption, WM_COMMAND, IDM_REFRESH, 0
						.endif
						invoke PostMessage, m_hWnd, WM_COMMAND, IDM_REFRESH, 0
					.endif
				.endif

			.elseif (([eax].NMHDR.code == LVN_ITEMCHANGED) && \
					 ([eax].NMLISTVIEW.uNewState & LVIS_SELECTED))

				StatusBar_SetText m_hWndSB, 0, CStr("")
				mov eax, pNMHdr
				invoke UpdateTypeInfoDlg, [eax].NMLISTVIEW.lParam

			.elseif ([eax].NMHDR.code == LVN_BEGINLABELEDIT)

				invoke OnBeginLabelEdit, eax
				mov dwRC, eax

			.elseif ([eax].NMHDR.code == LVN_ENDLABELEDIT)

				invoke OnEndLabelEdit, eax
				mov dwRC, eax

			.elseif ([eax].NMHDR.code == LVN_COLUMNCLICK)

				mov eax,[eax].NMLISTVIEW.iSubItem
				.if (eax == m_iSortCol)
					xor m_iSortDir,1
				.else
					mov m_iSortCol,eax
					mov m_iSortDir,0
				.endif	
				invoke LVSort, m_hWndLV, m_iSortCol, m_iSortDir, 0

			.endif

		.elseif ([eax].NMHDR.idFrom == IDC_STATUSBAR)

			.if ([eax].NMHDR.code == NM_DBLCLK)
				StatusBar_GetTextLength m_hWndSB, 0
				.if (eax)
					inc eax
					movzx eax, ax
					invoke GetTextBufferPtr, eax
					StatusBar_GetText m_hWndSB, 0, m_pszText
					invoke DialogBoxParam, g_hInstance, IDD_VIEWDETAIL, m_hWnd, viewdetailproc, __this
				.endif
			.endif

		.endif
;;		DebugOut "OnNotify end"
		return dwRC
		align 4
		assume edi:nothing

OnNotify endp


;--- get a property in scan mode, translate error message
;--- in scan mode avoid message boxes


GetProperty proc uses edi pFuncDesc:ptr FUNCDESC, bTransform:BOOL

local varResult:VARIANT
local DispParams:DISPPARAMS
local dwRC:DWORD
if 0
local pFuncDesc2:ptr FUNCDESC
local iPut:DWORD
endif

		mov edi, pFuncDesc
		assume edi:ptr FUNCDESC
;;		.if (g_bNoDispatchPropScan)
		.if (g_bNoDispatchPropScan && ([edi].memid != DISPID_VALUE))
			mov varResult.vt, VT_EMPTY
			invoke GetRetType, edi
			.if (eax == VT_DISPATCH)
				mov varResult.vt, VT_DISPATCH
;----------------------------------------- possibly same for ptr UDT 
			.elseif (eax == VT_PTR)
				invoke IsDispatchPtr, edi
				.if (eax)
					mov varResult.vt, VT_DISPATCH
				.endif
			.endif
			.if (varResult.vt == VT_DISPATCH)
				invoke GetTextBufferPtr, MAX_PATH
				invoke lstrcpy, m_pszText, CStr("[VT_DISPATCH]")
				return S_OK
			.endif
		.endif
		invoke VariantInit, addr varResult
		ResetDispParams DispParams
		invoke ExecuteInvoke, [edi].memid, DISPATCH_PROPERTYGET,
				addr DispParams, addr varResult
		mov dwRC, eax
		.if (eax == S_OK)
if 0
;--- doesn't work because GetRetType doesnt work with INVOKE_PROPERTYPUT
			int 3
			mov pFuncDesc2, NULL
			invoke vf(m_pTypeInfo, ITypeInfo2, GetFuncIndexOfMemId),
				[edi].memid, INVOKE_PROPERTYPUT, addr iPut
			.if (eax == S_OK)
				invoke vf(m_pTypeInfo, ITypeInfo, GetFuncDesc), iPut, addr pFuncDesc2
				.if (eax == S_OK)
					mov edi, pFuncDesc2
				.endif
			.endif
endif
			invoke Variant2String, edi, addr varResult, bTransform
if 0
			.if (pFuncDesc2)
				invoke vf(m_pTypeInfo, ITypeInfo, ReleaseFuncDesc), edi
				mov edi, pFuncDesc
			.endif
endif
		.else
			invoke HResult2String, edi, eax
		.endif
		invoke VariantClear, addr varResult
		return dwRC
		align 4
		assume edi:nothing

GetProperty endp


;--- read 1 object property: set Name (subitem 0) + Value(subitem 1) in a listview item


CheckFuncDesc proc uses esi edi pFuncDesc:ptr FUNCDESC, dwIndex:DWORD, pszName:LPSTR

local dwNames:DWORD
local bstr:BSTR
local lvi:LVITEM
local szText[MAX_PATH]:byte

		mov edi, pFuncDesc
		assume edi:ptr FUNCDESC
		mov eax,[edi].invkind
;--------------------------------------- do some selection
		.if ((m_iTab == TAB_PROPERTIES) && (eax != INVOKE_PROPERTYGET))
			jmp exit
		.endif
		.if ((m_iTab == TAB_METHODS) && (eax != INVOKE_FUNC))
			jmp exit
		.endif

;;		DebugOut "CheckFuncDesc, memid=%X", [edi].FUNCDESC.memid

		.if (!g_bShowAllMembers)
			.if ((eax == INVOKE_FUNC) && ([edi].wFuncFlags & FUNCFLAG_FRESTRICTED))
				jmp exit
			.endif
		.endif

		mov szText, 0
		.if (!pszName)
			invoke vf(m_pTypeInfo, ITypeInfo, GetNames), [edi].memid, addr bstr, 1, addr dwNames
			.if (eax == S_OK)
				invoke WideCharToMultiByte, CP_ACP, 0, bstr, -1, addr szText, sizeof szText, 0, 0
				invoke SysFreeString, bstr
			.endif
			lea eax, szText
			mov pszName, eax
		.endif

		mov lvi.mask_, LVIF_TEXT or LVIF_PARAM
		mov eax, m_dwListIdx
		mov lvi.iItem, eax
		inc m_dwListIdx
		@mov lvi.iSubItem, 0
		mov eax, dwIndex
		mov lvi.lParam, eax
		mov eax, pszName
		mov lvi.pszText, eax
		.if (m_bRefresh)
			invoke ListView_SetItem( m_hWndLV, addr lvi)
		.else
			invoke ListView_InsertItem( m_hWndLV, addr lvi)
		.endif
if 0
		DebugOut "property scan(%u): %s", dwIndex, addr szText
endif
		lea eax, szText
		mov lvi.pszText, eax
		mov lvi.mask_, LVIF_TEXT
		inc lvi.iSubItem
		.if ([edi].invkind == INVOKE_PROPERTYGET)
			invoke GetProperty, edi, TRUE
			mov eax, m_pszText
			mov lvi.pszText, eax
			invoke ListView_SetItem( m_hWndLV, addr lvi)
		.else
			invoke GetRetType, edi
			.if (eax != VT_VOID)
				invoke GetRetElemDesc, edi
				mov ecx, eax
				invoke GetParameterType, m_pTypeInfo, ecx, addr szText, sizeof szText
				invoke ListView_SetItem( m_hWndLV, addr lvi)
			.endif
			inc lvi.iSubItem
			movsx esi, [edi].cParams
			.if (esi)
				mov edx, [edi].lprgelemdescParam
				xor eax, eax
				mov ecx, esi
				.while (ecx)
					.if ([edx].ELEMDESC.paramdesc.wParamFlags & PARAMFLAG_FOPT)
						inc eax
					.endif
					.if ([edx].ELEMDESC.paramdesc.wParamFlags & (PARAMFLAG_FLCID or PARAMFLAG_FRETVAL))
						dec esi
					.endif
					dec ecx
					add edx, sizeof ELEMDESC
				.endw
				sub esi, eax
				invoke wsprintf, addr szText, CStr("%d/%d"), esi, eax
				invoke ListView_SetItem( m_hWndLV, addr lvi)
			.endif
		.endif
exit:
		ret
		assume edi:nothing
		align 4

CheckFuncDesc endp

if ?VARSUPP

if ?SIMFUNCDESC eq 0

CheckVarDesc proc uses esi edi pVarDesc:ptr VARDESC, dwIndex:DWORD

local dwNames:DWORD
local bstr:BSTR
local lvi:LVITEM
local DispParams:DISPPARAMS
local varResult:VARIANT
local szText[MAX_PATH]:byte

		mov edi, pVarDesc
		assume edi:ptr VARDESC
;--------------------------------------- do some selection

		.if (([edi].varkind != VAR_DISPATCH) || (m_iTab == TAB_METHODS))
			jmp exit
		.endif
		.if (!g_bShowAllMembers)
			.if ([edi].wVarFlags & VARFLAG_FRESTRICTED)
				jmp exit
			.endif
		.endif

		invoke vf(m_pTypeInfo, ITypeInfo, GetNames), [edi].memid, addr bstr, 1, addr dwNames
		.if (eax == S_OK)
			invoke WideCharToMultiByte, CP_ACP, 0, bstr, -1, addr szText, sizeof szText, 0, 0 
			invoke SysFreeString, bstr
			
			mov lvi.imask, LVIF_TEXT or LVIF_PARAM
			mov eax, m_dwListIdx
			mov lvi.iItem, eax
			inc m_dwListIdx
			@mov lvi.iSubItem, 0
			mov eax, dwIndex
			or eax, FLAG_VARDESC
			mov lvi.lParam, eax
			lea eax, szText
			mov lvi.pszText, eax
			.if (m_bRefresh)
				ListView_SetItem m_hWndLV, addr lvi
			.else
				ListView_InsertItem m_hWndLV, addr lvi
			.endif
			lea eax, szText
			mov lvi.pszText, eax
			mov lvi.imask, LVIF_TEXT
			inc lvi.iSubItem

			ResetDispParams DispParams
			invoke VariantInit, addr varResult
			invoke ExecuteInvoke, [edi].memid, DISPATCH_PROPERTYGET,
					addr DispParams, addr varResult
			.if (eax == S_OK)
				invoke Variant2String, NULL, addr varResult, FALSE
				mov eax, m_pszText
				mov lvi.pszText, eax
				ListView_SetItem m_hWndLV, addr lvi
				invoke VariantClear, addr varResult
			.else
				invoke HResult2String, NULL, eax
			.endif
		.endif
exit:
		ret
		assume edi:nothing
		align 4

CheckVarDesc endp

endif

endif

;--- get typeinfo

GetTypeInfo proc

local	dwIndex:DWORD
local	guid:GUID
local	hRefType:HREFTYPE
local	pTypeInfoRef:LPTYPEINFO
local	pTypeAttr:ptr TYPEATTR
local	pTypeLib:LPTYPELIB
local	szText[128]:byte

if 0
;-------------------------- try IProvideClassInfo first. Thats better than
;-------------------------- IDispatch::GetTypeInfo, since from there one
;-------------------------- may get a TKIND_INTERFACE, which is - as far as
;-------------------------- properties are concerned - harder to live with
	.if (m_pTypeInfo == NULL)
		invoke GetTypeInfoFromIProvideClassInfo, m_pDispatch, FALSE
		mov m_pTypeInfo, eax
	.endif
;-------------------------- last try is IDispatch
	.if (m_pTypeInfo == NULL)
		invoke vf(m_pDispatch, IDispatch, GetTypeInfo), 0, g_LCID, addr m_pTypeInfo
		.if (eax != S_OK)
			invoke wsprintf, addr szText, CStr("No typeinfo available[%X]"), eax
			StatusBar_SetText m_hWndSB, 0, addr szText
			invoke MessageBeep, MB_OK
		.endif
	.endif
else
	.if (m_pTypeInfo == NULL)
;-------------------------- new strategy: get typeinfo from IDispatch.
;-------------------------- if we get an TKIND_INTERFACE, get containing
;-------------------------- type library and search for correspondant
;-------------------------- TKIND_DISPATCH
		invoke vf(m_pDispatch, IDispatch, GetTypeInfo), 0, g_LCID, addr m_pTypeInfo
		.if (eax != S_OK)
			.if (m_hWndSB)
				invoke wsprintf, addr szText, CStr("No typeinfo available[%X]"), eax
				StatusBar_SetText m_hWndSB, 0, addr szText
				invoke MessageBeep, MB_OK
			.else
				invoke OutputMessage, m_hWnd, eax, CStr("IDispatch::GetTypeInfo"), 0
			.endif
		.else
			invoke vf(m_pTypeInfo, ITypeInfo, GetTypeAttr), addr pTypeAttr
			.if (eax == S_OK)
				mov ecx, pTypeAttr
				push [ecx].TYPEATTR.typekind
				invoke CopyMemory, addr guid, addr [ecx].TYPEATTR.guid, sizeof GUID
				invoke vf(m_pTypeInfo, ITypeInfo, ReleaseTypeAttr), pTypeAttr
				pop eax
if ?USEDISPINVOKE
				.if (eax == TKIND_DISPATCH)
					invoke vf(m_pTypeInfo, ITypeInfo, GetRefTypeOfImplType), -1, addr hRefType
					.if (eax == S_OK)
						invoke vf(m_pTypeInfo, ITypeInfo, GetRefTypeInfo), hRefType, addr pTypeInfoRef
						.if (eax == S_OK)
							invoke vf(m_pTypeInfo, ITypeInfo, Release)
							mov eax, pTypeInfoRef
							mov m_pTypeInfo, eax
						.endif
					.endif
				.endif
else
				.if (eax == TKIND_INTERFACE)
					invoke vf(m_pTypeInfo, ITypeInfo, GetContainingTypeLib), addr pTypeLib, addr dwIndex
					.if (eax == S_OK)
						invoke vf(pTypeLib, ITypeLib, GetTypeInfoOfGuid), addr guid, addr pTypeInfoRef
						.if (eax == S_OK)
							invoke vf(m_pTypeInfo, ITypeInfo, Release)
							mov eax, pTypeInfoRef
							mov m_pTypeInfo, eax
						.endif
						invoke vf(pTypeLib, ITypeLib, Release)
					.endif
				.endif
endif
			.endif
		.endif
	.endif
endif
	ret
	align 4

GetTypeInfo endp

;--- RefreshView: returns TRUE=typeinfo found or FALSE=no typeinfo found

RefreshView proc uses esi edi iTab:DWORD

local	pTypeAttr:ptr TYPEATTR
local	pFuncDesc:ptr FUNCDESC
local	pDispatch:LPDISPATCH
if ?VARSUPP
local	pVarDesc:ptr VARDESC
endif
if 0
local	pTypeInfo:LPTYPEINFO
local	pTypeAttr2:ptr TYPEATTR
local	hRefType:DWORD
endif
local	pOleObject:LPOLEOBJECT
local	pUnknown:LPUNKNOWN
local	pszDispIdName:LPSTR
local	dwIndex:DWORD
local	dwFlags:DWORD
local	iTopIndex:DWORD
local	dwRC:BOOL
local	bstr:BSTR
local	hCsrOld:HCURSOR
local	szDocText[64]:byte
local	szDefItem[64]:byte
local	szText[128]:byte

	invoke GetTypeInfo
	.if (!m_pTypeInfo)
		mov dwRC, FALSE
		jmp exit
	.endif

	mov dwRC, TRUE

	invoke SetWindowRedraw( m_hWndLV, FALSE)
	invoke SetCursor, g_hCsrWait
	mov hCsrOld, eax

	mov eax, iTab
	.if (eax != m_iTab)
		mov m_iTab, eax
		mov iTopIndex, 0
		invoke ListView_DeleteAllItems( m_hWndLV)
		.repeat
			invoke ListView_DeleteColumn( m_hWndLV,0)
		.until (eax == 0)
if ?SHOWAMBIENT
		.if ((m_iTab == TAB_PROPERTIES) || (m_iTab == TAB_AMBIENT))
else
		.if (m_iTab == TAB_PROPERTIES)
endif
			invoke SetLVColumns, m_hWndLV, NUMCOLS_PROPERTIES, addr ColumnsProperties
		.else
			invoke SetLVColumns, m_hWndLV, NUMCOLS_METHODS, addr ColumnsMethods
		.endif
	.endif

	invoke ListView_GetItemCount( m_hWndLV)
	.if (eax)
		mov m_bRefresh, TRUE
	.else
		mov m_bRefresh, FALSE
	.endif

	mov m_dwListIdx, 0
	mov m_bException, FALSE
	mov szDefItem, 0
	mov m_bScanMode, TRUE
if ?SHOWAMBIENT
	.if (m_iTab == TAB_AMBIENT)
		mov m_bScanMode, FALSE
		xor esi, esi
		sub esp, sizeof FUNCDESC
		mov edi, esp
		.while (1)
			invoke GetAmbientDispId, esi, addr [edi].FUNCDESC.memid, addr pszDispIdName
			.break .if (!eax)
			mov [edi].FUNCDESC.invkind, INVOKE_PROPERTYGET
			mov [edi].FUNCDESC.wFuncFlags, 0
			mov [edi].FUNCDESC.elemdescFunc.tdesc.vt, VT_I4
			mov [edi].FUNCDESC.cParams, 0
			invoke CheckFuncDesc, edi, -1, pszDispIdName
			inc esi
		.endw
		add esp, sizeof FUNCDESC
	.endif
endif
	invoke vf(m_pTypeInfo, ITypeInfo, GetTypeAttr),addr pTypeAttr
	.if (eax == S_OK)
		mov esi, pTypeAttr
		mov ax, [esi].TYPEATTR.wTypeFlags
		mov m_wTypeFlags, ax
		.if ([esi].TYPEATTR.typekind == TKIND_INTERFACE)
			mov m_bVtblMode, TRUE
			invoke vf(m_pDispatch, IUnknown, QueryInterface), addr [esi].TYPEATTR.guid, addr pUnknown
			.if (eax == S_OK)
				invoke vf(m_pDispatch, IUnknown, Release)
				mov eax, pUnknown
				mov m_pDispatch, eax
if 0
				mov eax, m_pTypeInfo
				mov pTypeInfo, eax
				invoke vf(pTypeInfo, IUnknown, AddRef)
				.while (1)
					.break .if (![esi].TYPEATTR.cImplTypes)
					invoke vf(pTypeInfo, ITypeInfo, GetRefTypeOfImplType), 0, addr hRefType
					.break .if (eax != S_OK)
					push pTypeInfo
					invoke vf(pTypeInfo, ITypeInfo, GetRefTypeInfo), hRefType, addr pTypeInfo
					xchg eax, [esp]
					invoke vf(eax, IUnknown, Release)
					pop eax
					.break .if (eax != S_OK)
					invoke vf(pTypeInfo, ITypeInfo, GetTypeAttr),addr pTypeAttr2
					.break .if (eax != S_OK)
					mov esi, pTypeAttr2
					mov dwIndex, 0
					mov m_bScanMode, TRUE
					.while (m_bScanMode)	;may stop after exception
						mov eax, dwIndex
						.break .if (ax >= [esi].TYPEATTR.cFuncs)
						invoke vf(pTypeInfo, ITypeInfo, GetFuncDesc), dwIndex, addr pFuncDesc
						.if (eax == S_OK)
							mov eax, pFuncDesc
							invoke CheckFuncDesc, pFuncDesc, dwIndex, NULL
							invoke vf(pTypeInfo,ITypeInfo,ReleaseFuncDesc),pFuncDesc
						.endif
						inc dwIndex
					.endw
					invoke vf(pTypeInfo, ITypeInfo, ReleaseTypeAttr), esi
				.endw
				mov esi, pTypeAttr
				.if (pTypeInfo)
					invoke vf(pTypeInfo, IUnknown, Release)
				.endif
endif
			.endif
		.else
;----------------------- cleaner would be to query TYPEFLAG_FDUAL
;----------------------- but you cannot rely on that
			invoke vf(m_pDispatch, IUnknown, QueryInterface), addr IID_IDispatch, addr pDispatch
			.if (eax == S_OK)
				invoke vf(pDispatch, IUnknown, Release)
				mov m_bVtblMode, FALSE
			.else
				mov m_bVtblMode, TRUE
			.endif
		.endif
		.if (m_bVtblMode)
			StatusBar_SetText m_hWndSB, 0, CStr("vtable mode is on")
		.endif

		@mov dwIndex, 0
		.while (m_bScanMode)	;may stop after exception
			mov eax, dwIndex
			.break .if (ax >= [esi].TYPEATTR.cFuncs)
			invoke vf(m_pTypeInfo, ITypeInfo, GetFuncDesc), dwIndex, addr pFuncDesc
			.if (eax == S_OK)
				mov eax, pFuncDesc
				.if (([eax].FUNCDESC.memid == DISPID_VALUE) && ([eax].FUNCDESC.invkind == INVOKE_PROPERTYGET))
					invoke GetProperty, eax, FALSE
					.if (eax == S_OK)
						invoke lstrcpyn, addr szDefItem, m_pszText, sizeof szDefItem
					.endif
				.endif
				invoke CheckFuncDesc, pFuncDesc, dwIndex, NULL
				invoke vf(m_pTypeInfo,ITypeInfo,ReleaseFuncDesc),pFuncDesc
			.endif
			inc dwIndex
		.endw
if ?VARSUPP
		@mov dwIndex, 0
		.while (m_bScanMode)	;may stop after exception
			mov eax, dwIndex
			.break .if (ax >= [esi].TYPEATTR.cVars)
if ?SIMFUNCDESC eq 0
			invoke vf(m_pTypeInfo, ITypeInfo, GetVarDesc), dwIndex, addr pVarDesc
			.if (eax == S_OK)
				invoke CheckVarDesc, pVarDesc, dwIndex
				invoke vf(m_pTypeInfo,ITypeInfo,ReleaseVarDesc),pVarDesc
			.endif
else
			mov edi, dwIndex
			or edi, FLAG_VARDESC
			invoke GetFuncDesc, edi, addr pFuncDesc
			.if (eax == S_OK)
				invoke CheckFuncDesc, pFuncDesc, edi, NULL
				invoke ReleaseFuncDesc, edi, pFuncDesc
			.endif
endif
			inc dwIndex
		.endw
endif
done:
		invoke vf(m_pTypeInfo, ITypeInfo, ReleaseTypeAttr), pTypeAttr
	.else
;---------------------------------------- typeinfo may have become invalid
		invoke wsprintf, addr szText, CStr("ITypeInfo::GetTypeAttr failed[%X]"), eax
		StatusBar_SetText m_hWndSB, 0, addr szText
		invoke MessageBeep, MB_OK
	.endif

	invoke vf(m_pTypeInfo,ITypeInfo,GetDocumentation), MEMBERID_NIL, addr bstr, NULL, NULL, NULL
	.if (eax == S_OK)
		invoke WideCharToMultiByte, CP_ACP, 0, bstr, -1, addr szDocText, sizeof szDocText, 0, 0
		invoke SysFreeString, bstr
	.else
		mov szDocText, 0
	.endif

	sub esp, 256
	mov edx, esp
	invoke wsprintf, edx, CStr("%.40s %.40s Properties"), addr szDocText, addr szDefItem
	invoke SetWindowText, m_hWnd, esp
	add esp, 256

	mov m_bScanMode, FALSE

	.if (m_iSortCol != -1)
		invoke LVSort, m_hWndLV, m_iSortCol, m_iSortDir, 0
	.endif

	invoke SetWindowRedraw( m_hWndLV, TRUE)
	invoke SetCursor, hCsrOld

	mov esi, FALSE
	invoke vf(m_pDispatch, IUnknown, QueryInterface), addr IID_IOleObject, addr pOleObject
	.if (eax == S_OK)
		mov esi, TRUE
		invoke vf(pOleObject, IUnknown, Release)
	.endif
	invoke GetDlgItem, m_hWnd, IDM_VIEWOBJECT
	invoke EnableWindow, eax, esi
exit:
	return dwRC
	align 4

RefreshView endp

SetChildDlgPos proc hWndTab:HWND, hWndDlg:HWND, hWndLV:HWND
 
local	rect:RECT

	invoke GetClientRect, hWndTab, addr rect
	invoke TabCtrl_AdjustRect( hWndTab, FALSE, addr rect)

	mov edx, rect.right
	sub edx, rect.left
	mov ecx, rect.bottom
	sub ecx, rect.top
	push edx
	push ecx
	invoke ClientToScreen, hWndTab, addr rect
	invoke ScreenToClient, hWndDlg, addr rect
	pop ecx
	pop edx
	invoke SetWindowPos, hWndLV, HWND_TOP, rect.left, rect.top, edx, ecx, SWP_SHOWWINDOW
	ret
	align 4

SetChildDlgPos endp


OnInitDialog proc uses esi

local	dwWidth[1]:DWORD
local	rect:RECT
local	clsid:CLSID
local	tci:TC_ITEM

	invoke GetGUID@CObjectItem, m_pObjectItem, addr clsid
	invoke SetWindowIcon, m_hWnd, addr clsid
	mov m_hIcon, eax
if ?PROPICON
	.if (!eax)
		mov eax, g_hIconProp
		.if (!eax)
			invoke LoadIcon,g_hInstance,IDI_PROPERTIES
			mov g_hIconProp, eax
		.endif
		.if (eax)
			invoke SendMessage, m_hWnd, WM_SETICON, ICON_SMALL, g_hIconProp
			invoke SendMessage, m_hWnd, WM_SETICON, ICON_BIG, g_hIconProp
		.endif
	.endif
endif
	invoke GetDlgItem, m_hWnd, IDC_TAB1
	mov m_hWndTab, eax
	invoke GetWindowRect, m_hWndTab, addr rect
	invoke ScreenToClient, m_hWnd, addr rect.right
	push rect.bottom
	invoke GetClientRect, m_hWnd, addr rect
	pop ecx
	mov eax, rect.bottom
	sub eax, ecx
	mov m_dwRim, eax

	invoke GetDlgItem, m_hWnd, IDC_LIST1
	mov m_hWndLV, eax
	invoke GetDlgItem, m_hWnd, IDC_STATUSBAR
	mov m_hWndSB, eax
	mov dwWidth[0*sizeof DWORD], -1
	StatusBar_SetParts m_hWndSB, 1, addr dwWidth

	invoke CreateEditListView, m_hWndLV

	mov tci.mask_,TCIF_TEXT or TCIF_PARAM
	mov esi,offset TabDlgPages
	mov ecx,0
	.while (ecx < NUMDLGS)
		push ecx
		mov tci.lParam,esi
		mov eax,[esi].CTabDlgPage.pTabName
		mov tci.pszText,eax
		invoke TabCtrl_InsertItem( m_hWndTab,ecx,addr tci)
		add esi,sizeof CTabDlgPage
		pop ecx
		inc ecx
	.endw

	invoke SetChildDlgPos, m_hWndTab, m_hWnd, m_hWndLV

	invoke ListView_SetExtendedListViewStyle( m_hWndLV, LVS_EX_GRIDLINES or LVS_EX_INFOTIP)
	ret
	align 4

OnInitDialog endp

;--- typelib callback

tlcb proc uses __this thisarg, pTypeLib:LPTYPELIB, dwIndex:DWORD

local	pTypeAttr:ptr TYPEATTR
local	pTypeInfo:LPTYPEINFO

	mov __this,this@
	invoke vf(pTypeLib, ITypeLib, GetTypeInfo), dwIndex, addr pTypeInfo
	.if (eax == S_OK)
		invoke vf(pTypeInfo, ITypeInfo, GetTypeAttr), addr pTypeAttr
		.if (eax == S_OK)
			mov ecx, pTypeAttr
			.if (([ecx].TYPEATTR.typekind == TKIND_DISPATCH) || \
				([ecx].TYPEATTR.typekind == TKIND_INTERFACE))
				invoke ComPtrAssign, addr m_pTypeInfo, pTypeInfo
;----------------------- this has to be done to get a real refresh
				mov ecx, m_pTypeLibDlg
				invoke PostMessage, [ecx].CDlg.hWnd, WM_CLOSE, 0, 0
				invoke ListView_DeleteAllItems( m_hWndLV)
				invoke PostMessage, m_hWnd, WM_COMMAND, IDM_REFRESH, 0
			.else
				mov ecx, m_pTypeLibDlg
				invoke MessageBox, [ecx].CDlg.hWnd, CStr("Only types of kind DISPATCH or INTERFACE may be selected"), 0, MB_OK
			.endif
			invoke vf(pTypeInfo, ITypeInfo, ReleaseTypeAttr), pTypeAttr
		.endif
		invoke vf(pTypeInfo, IUnknown, Release)
	.endif
	ret
	align 4

tlcb endp


;--- WM_COMMAND


OnCommand proc wParam:WPARAM, lParam:LPARAM

local	pFuncDesc:ptr FUNCDESC
local	dwContext:DWORD
local	bstr:BSTR
local	lvi:LVITEM
local	szText[128]:byte

	movzx eax,word ptr wParam
	.if (eax == IDCANCEL)

		invoke PostMessage, m_hWnd, WM_CLOSE, 0, 0

	.elseif ((eax == IDOK) || (eax == IDM_EDITITEM))

		invoke GetFocus
		.if (eax == m_hWndLV)
			.if (word ptr wParam == IDOK)
				mov ecx, 1
			.else
				mov ecx, 2
			.endif
			invoke OnEdit, ecx
		.endif

	.elseif (eax == IDM_COPYVALUE)

		invoke OnCopyValue

	.elseif (eax == IDM_TYPEINFO)

		xor g_bSyncTypeInfoAndProp, 1
		invoke UpdateTypeInfoDlg, -1

	.elseif (eax == IDM_SHOWALL)

		xor g_bShowAllMembers, 1
		mov eax, TAB_UNDEFINED
		xchg eax, m_iTab
		invoke RefreshView, eax

	.elseif (eax == IDM_USETIINVOKE)

		xor m_bUseTIInvoke, 1

	.elseif (eax == IDM_HEXADECIMAL)

		xor m_bHexadecimal, 1
		mov al, m_bHexadecimal
		mov g_bHexadecimal, al
		invoke PostMessage, m_hWnd, WM_COMMAND, IDM_REFRESH, 0

	.elseif (eax == IDC_TYPELIBDLG)

		.if (m_pTypeLibDlg)
			mov ecx, m_pTypeLibDlg
			invoke RestoreAndActivateWindow, [ecx].CDlg.hWnd
		.else
			invoke Create4@CTypeLibDlg, m_pTypeInfo
			.if (eax)
				mov m_pTypeLibDlg, eax
				invoke GetWindow, m_hWnd, GW_OWNER
				invoke Show@CTypeLibDlg, m_pTypeLibDlg, eax, FALSE
			.else
				invoke MessageBox, m_hWnd, CStr("no type library found"), 0, MB_OK
			.endif
		.endif

	.elseif (eax == IDM_FORCETYPEINFO)

		invoke SendMessage, m_hWnd, WM_COMMAND, IDC_TYPELIBDLG, 0
		.if (m_pTypeLibDlg)
			invoke SetCallBack@CTypeLibDlg, m_pTypeLibDlg, offset tlcb, __this
			invoke EnableWindow, m_hWnd, FALSE
		.endif

	.elseif (eax == IDM_REFRESH)

		invoke RefreshView, m_iTab

	.elseif (eax == IDM_OBJECTDLG)

		invoke vf(m_pObjectItem, IObjectItem, GetFlags)
		or eax, OBJITEMF_IGNOREOV
		invoke vf(m_pObjectItem, IObjectItem, SetFlags), eax
		mov ecx, g_pMainDlg
		invoke vf(m_pObjectItem, IObjectItem, ShowObjectDlg), [ecx].CDlg.hWnd

	.elseif (eax == IDM_VIEWOBJECT)

		mov ecx, g_pMainDlg
		invoke vf(m_pObjectItem, IObjectItem, ShowViewObjectDlg), [ecx].CDlg.hWnd, NULL

	.elseif (eax == IDM_CONTEXTHELP)

		invoke ListView_GetNextItem( m_hWndLV, -1, LVNI_SELECTED)
		.if (eax != -1)
			mov lvi.iItem, eax
			mov lvi.iSubItem, 0
			mov lvi.mask_, LVIF_PARAM
			invoke ListView_GetItem( m_hWndLV, addr lvi)
			invoke GetFuncDesc, lvi.lParam, addr pFuncDesc
			.if (eax == S_OK)
				mov ecx, pFuncDesc
				invoke vf(m_pTypeInfo, ITypeInfo, GetDocumentation), [ecx].FUNCDESC.memid, NULL, NULL, addr dwContext, addr bstr
				.if ((eax == S_OK) && bstr)
					invoke WideCharToMultiByte,CP_ACP,0, bstr, -1, addr szText, sizeof szText,0,0
					invoke SysFreeString, bstr
					.if (dwContext)
						invoke ShowHtmlHelp, addr szText, HH_HELP_CONTEXT, dwContext
						.if (!eax)
							push esi
							sub esp, MAX_PATH+32
							mov esi, esp
							invoke wsprintf, esi, CStr("HtmlHelp('%s', %u) failed"), addr szText, dwContext
							StatusBar_SetText m_hWndSB, 0, esi
							StatusBar_SetTipText m_hWndSB, 0, esi
							invoke MessageBeep, MB_OK
							add esp, MAX_PATH+32
							pop esi
						.endif
					.endif
				.endif
				invoke ReleaseFuncDesc, lvi.lParam, pFuncDesc
			.endif
		.endif
	.endif
	ret
	align 4

OnCommand endp

;--- WM_SIZE

	.const

BtnTab label dword
	dd IDM_OBJECTDLG
	dd IDM_VIEWOBJECT
	dd IDC_TYPELIBDLG
	dd IDM_REFRESH
	dd IDCANCEL
NUMBUTTONS textequ %($ - BtnTab) / sizeof DWORD

	.code

OnSize proc uses edi esi hWnd:HWND, dwType:dword, dwWidth:dword, dwHeight:dword

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
	lea edi, hWndBtn
	mov ecx, NUMBUTTONS
	.while (ecx)
		push ecx
		lodsd
		invoke GetDlgItem, hWnd, eax
		mov [edi], eax
		add edi, sizeof HWND
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
	mov edi, eax

	mov eax, dwHeight
	sub eax, dwRim
	sub eax, m_dwRim
	mov dwHeightLV, eax
	test eax, eax
	.if (SIGN?)
		@mov dwHeightLV, 0
	.endif
	invoke DeferWindowPos, edi, m_hWndTab, NULL, 0, 0, dwWidth, dwHeightLV, SWP_NOMOVE or SWP_NOZORDER or SWP_NOACTIVATE

	mov eax, m_dwRim
	sub eax, dwHeightSB
	sub eax, dwHeightBtn
	shr eax, 1
	add eax, dwHeightLV
	add eax, dwRim
	mov dwYPos, eax

	mov eax, dwWidth
	sub eax, dwWidthBtn
	.if (!CARRY?)
		xor edx, edx
		mov ecx, NUMBUTTONS + 1
		div ecx
		mov dwAddX, eax
		mov dwXPos, eax
	.else
		mov dwAddX, 1
		mov dwXPos, 0
	.endif

	lea esi, hWndBtn
	mov ecx, NUMBUTTONS
	.while (ecx)
		push ecx
		lodsd
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

	invoke DeferWindowPos, edi, m_hWndSB, NULL, 0, 0, 0, 0, SWP_NOZORDER or SWP_NOACTIVATE

	invoke EndDeferWindowPos, edi

	invoke SetChildDlgPos, m_hWndTab, m_hWnd, m_hWndLV

	ret
	align 4

OnSize endp


CPropertiesDialog proc uses __this thisarg, message:dword, wParam:WPARAM, lParam:LPARAM

local	rect:RECT
local	dwRC:DWORD

	mov __this,this@

	mov eax,message
	.if (eax == WM_INITDIALOG)

		.if (!m_ptPos.x)
			invoke CenterWindow, m_hWnd
		.else
		.endif
		invoke OnInitDialog
		invoke RefreshView, TAB_PROPERTIES
		.if (!eax)
			invoke PostMessage, m_hWnd, WM_CLOSE, 0, 0
		.else
			.if (m_ptPos.x)
				invoke SetWindowPos, m_hWnd, NULL, m_ptPos.x, m_ptPos.y,
					0, 0, SWP_NOSIZE or SWP_NOZORDER or SWP_SHOWWINDOW 
			.else
				invoke ShowWindow, m_hWnd,SW_SHOWNORMAL
			.endif
		.endif
		mov eax,1

	.elseif (eax == WM_CLOSE)
if ?MODELESS
		invoke DestroyWindow, m_hWnd
else
		invoke EndDialog, m_hWnd, 0
endif
		mov eax,1

	.elseif (eax == WM_DESTROY)

		invoke Destroy@CPropertiesDlg, __this
if ?MODELESS
	.elseif (eax == WM_ACTIVATE)

		movzx eax,word ptr wParam
		.if (eax == WA_INACTIVE)
			mov g_hWndDlg, NULL
		.else
			mov eax, m_hWnd
			mov g_hWndDlg, eax
			.if (g_pTypeInfoDlg && m_pTypeInfo)
				invoke SetTypeInfo@CTypeInfoDlg, g_pTypeInfoDlg, m_pTypeInfo
				.if (eax)
					invoke UpdateTypeInfoDlg, -1
				.endif
			.endif
		.endif
endif
	.elseif (eax == WM_SIZE)

		.if (wParam != SIZE_MINIMIZED)
			movzx eax, word ptr lParam+0
			movzx ecx, word ptr lParam+2
			invoke OnSize, m_hWnd, wParam, eax, ecx
		.endif

	.elseif (eax == WM_NOTIFY)

		invoke OnNotify, lParam

	.elseif (eax == WM_COMMAND)

		invoke OnCommand, wParam, lParam

	.elseif (eax == WM_ENTERMENULOOP)

		StatusBar_SetSimpleMode m_hWndSB, TRUE

	.elseif (eax == WM_EXITMENULOOP)

		StatusBar_SetSimpleMode m_hWndSB, FALSE

	.elseif (eax == WM_MENUSELECT)

		movzx ecx, word ptr wParam+0
		invoke DisplayStatusBarString, m_hWndSB, ecx
if ?HTMLHELP
	.elseif (eax == WM_HELP)

		invoke DoHtmlHelp, HH_DISPLAY_TOPIC, CStr("PropertiesDialog.htm")
endif
	.elseif (eax == WM_WNDDESTROYED)

		push esi
		mov eax, lParam
		mov edx, g_pTypeInfoDlg
		mov esi, m_pTypeLibDlg
		.if (edx && (eax == [edx].CDlg.hWnd))
			mov g_pTypeInfoDlg, NULL
		.elseif (esi && (eax == [esi].CDlg.hWnd))
			mov m_pTypeLibDlg, NULL
			invoke EnableWindow, m_hWnd, TRUE
			invoke RestoreAndActivateWindow, m_hWnd
		.endif
		pop esi

if ?POSTSETDISP
	.elseif (eax == WM_SETDISPATCH)

		invoke SetDispatch@CPropertiesDlg, __this, lParam, TRUE
endif
	.else
		xor eax,eax ;indicates "no processing"
	.endif
	ret
	align 4

CPropertiesDialog endp

	end
