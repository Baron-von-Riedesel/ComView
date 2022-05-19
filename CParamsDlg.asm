
;*** definition of class CParamsDlg
;*** CParamsDlg implements a dialog to enter method parameters
;*** fills a PARAMRETURN structure as return

	.386
	.model flat,stdcall
	option casemap :none   ; case sensitive
	option proc:private

	include COMView.inc
	include statusbar.inc
INSIDE_CPARAMSDLG equ 1
	include classes.inc
	include rsrc.inc
	include debugout.inc

?USEGETIDSOFNAMES	equ 0		;seems not to work in all cases

protoExecCB	typedef proto :ptr CPropertiesDlg, :HWND, :ptr FUNCDESC, :DWORD, :ptr VARIANT, hWndSB:HWND
LPEXECCB	typedef ptr protoExecCB


BEGIN_CLASS CParamsDlg, CDlg
pTypeInfo		LPTYPEINFO	?
pFuncDesc		LPFUNCDESC	?
dwRows			DWORD		?	;number of parameters
hWndCurEdit		HWND		?
dwIDCurEdit		DWORD		?
hWndSB			HWND		?
dwRC			DWORD		?
ParamReturn		PARAMRETURN <>
pPropertiesDlg	pCPropertiesDlg ?
execcb			LPEXECCB	?
bOk				BOOLEAN		?
END_CLASS

CBEDITID	equ 3e9h

	.data

g_rect		RECT {0,0,0,0}

	.code

;--------------------------------------------------------------
;--- class CParamsDlg
;--------------------------------------------------------------

__this	textequ <ebx>
_this	textequ <[__this].CParamsDlg>
thisarg	textequ <this@:ptr CParamsDlg>


	MEMBER hWnd, pDlgProc, pTypeInfo, pFuncDesc, hWndCurEdit,
	MEMBER dwIDCurEdit, hWndSB, bOk, dwRows, pPropertiesDlg, execcb, ParamReturn, dwRC


Create@CParamsDlg proc public uses esi __this pTypeInfo:LPTYPEINFO, pFuncDesc:ptr FUNCDESC,
				pPropertiesDlg:ptr CPropertiesDlg, pVoid:LPVOID

	invoke malloc, sizeof CParamsDlg
	.if (!eax)
		ret
	.endif

	mov __this,eax
	mov m_pDlgProc, CParamsDialog

	mov eax, pTypeInfo
	mov m_pTypeInfo, eax
	invoke vf(m_pTypeInfo, IUnknown, AddRef)
	mov eax, pFuncDesc
	mov m_pFuncDesc, eax
	mov eax, pPropertiesDlg
	mov ecx, pVoid
	mov m_pPropertiesDlg, eax
	mov m_execcb, ecx
	mov m_dwRC, -1
	return __this
	align 4

Create@CParamsDlg endp


Destroy@CParamsDlg proc uses __this thisarg

	mov __this,this@
	.if (m_pTypeInfo)
		invoke vf(m_pTypeInfo, IUnknown, Release)
	.endif
	invoke ParamReturnClear, addr m_ParamReturn
	invoke free, __this
	ret
	align 4

Destroy@CParamsDlg endp


;*** create a "open file" dialog to browse for a filename
;*** which is returned in pszFileName


OnBrowse proc uses esi pszFileName:LPSTR, dwSize:DWORD

local	ofn:OPENFILENAME
local	szFilter[128]:byte

;------------------------------- prepare GetOpenFileName dialog
	mov esi,pszFileName
	mov byte ptr [esi],0

	invoke ZeroMemory,addr szFilter, sizeof szFilter
	invoke lstrcpy,addr szFilter,CStr("All files (*.*)")
	invoke lstrlen,addr szFilter
	inc eax
	lea ecx,szFilter
	add ecx,eax
	invoke lstrcpy,ecx,CStr("*.*")

	invoke ZeroMemory, addr ofn, sizeof OPENFILENAME
	mov ofn.lStructSize,sizeof OPENFILENAME
	mov eax,m_hWnd
	mov ofn.hwndOwner,eax
	lea eax,szFilter
	mov ofn.lpstrFilter,eax

	mov ofn.lpstrCustomFilter,NULL
	mov ofn.nMaxCustFilter,0

	mov ofn.nFilterIndex,0
	mov ofn.lpstrFile,esi
	mov eax, dwSize
	mov ofn.nMaxFile, eax
	mov ofn.Flags,OFN_EXPLORER or OFN_NOVALIDATE

	invoke GetOpenFileName,addr ofn

	ret
	align 4

OnBrowse endp


;--- esi -> ELEMDESC
;--- edi -> VARIANT

TranslateUDT proc

local	pTypeInfoRef:LPTYPEINFO
local	pTypeAttr:ptr TYPEATTR
local	pVarDesc:ptr VARDESC
local	dwIndex:DWORD
local	dwTmp:DWORD
local	bstrVar:BSTR

	assume esi:ptr ELEMDESC
	assume edi:ptr VARIANT

	DebugOut "CParamsDlg, TranslateUDT"
	invoke vf(m_pTypeInfo, ITypeInfo, GetRefTypeInfo), [esi].tdesc.hreftype, addr pTypeInfoRef
	.if (eax == S_OK)
if ?USEGETIDSOFNAMES
;------------------------------- get the DISPID of the var
		invoke vf(pTypeInfoRef, ITypeInfo, GetIDsOfNames), addr [edi].bstrVal, 1, addr dwMemId
		.if (eax == S_OK)
endif
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
if ?USEGETIDSOFNAMES
						mov eax, dwMemId
;------------------------------- if found, get the value of the constant
						.if (eax == [ecx].VARDESC.memid)
else
						invoke vf(pTypeInfoRef, ITypeInfo, GetNames), [ecx].VARDESC.memid, addr bstrVar, 1, addr dwTmp
						.if (eax == S_OK)
							invoke _strcmpW, [edi].bstrVal, bstrVar
							push eax
							invoke SysFreeString, bstrVar
							pop eax
						.else
							mov eax, 1
						.endif
						.if (eax == 0)
endif
							invoke VariantClear, edi
							mov ecx, pVarDesc
							invoke VariantCopy, edi, [ecx].VARDESC.lpvarValue
							invoke vf(pTypeInfoRef, ITypeInfo, ReleaseVarDesc), pVarDesc
							.break
						.endif
						invoke vf(pTypeInfoRef, ITypeInfo, ReleaseVarDesc), pVarDesc
					.endif
					inc dwIndex
				.endw
				invoke vf(pTypeInfoRef, ITypeInfo, ReleaseTypeAttr), pTypeAttr
			.endif
if ?USEGETIDSOFNAMES
		.endif
endif
		invoke vf(pTypeInfoRef, ITypeInfo, Release)
	.endif
	ret
	assume esi:nothing
	assume edi:nothing
	align 4

TranslateUDT endp


;*** process arguments typed in by user
;*** returns TRUE if everything ok, else FALSE


ParseArguments proc uses esi edi

local	dwSize:DWORD
local	dwMemId:DWORD
local	dwNumParams:DWORD
local	bTranslated:BOOL
local   pbstr:ptr BSTR
local   dwCtrlIDEdit: dword
local   dwCtrlIDCB: dword
local	hWndCB:HWND
local	hWndEdit:HWND
local	dwCBData:DWORD
local	bstrVar:BSTR
local	szText[MAX_PATH]:byte

		mov dwCtrlIDEdit,IDC_EDIT1
		mov dwCtrlIDCB,IDC_COMBO1

		mov ecx, m_ParamReturn.iNumVariants
		mov edi, m_ParamReturn.pVariants
		assume edi:ptr VARIANT

		mov eax,sizeof VARIANT
		mul ecx
		add edi,eax

		mov esi, m_pFuncDesc
		movzx ecx,[esi].FUNCDESC.cParams
		mov esi,[esi].FUNCDESC.lprgelemdescParam
		assume esi:ptr ELEMDESC

;------------------------------- now do in a loop:
;------------------------------- 1. get text from controls IDC_EDITx
;------------------------------- 2. convert it to a BSTR
;------------------------------- 3. change type to parameter type
;------------------------------- registers hold:
;------------------------------- ecx=number of parameters for function
;------------------------------- esi=ptr to ELEMDESC array
;------------------------------- edi=ptr to variant array (reverse order!)
		.while (ecx)
			.if ([esi].paramdesc.wParamFlags & (PARAMFLAG_FLCID or PARAMFLAG_FRETVAL))
				dec ecx
				add esi,sizeof ELEMDESC
				.continue
			.endif
			push ecx
			sub edi,sizeof VARIANT
			invoke VariantInit, edi
			mov [edi].vt,VT_ERROR
			mov [edi].scode,DISP_E_PARAMNOTFOUND
			mov szText, 0
			invoke GetDlgItem, m_hWnd, dwCtrlIDEdit
			mov hWndEdit, eax
			invoke GetDlgItem, m_hWnd, dwCtrlIDCB
			mov hWndCB, eax
			invoke ComboBox_GetCurSel( hWndCB)
			.if (eax == -1)
				invoke SetFocus, hWndCB
				jmp error
			.endif
			invoke ComboBox_GetItemData( hWndCB, eax)
			mov dwCBData, eax
			invoke GetWindowText, hWndEdit, addr szText, sizeof szText
			DebugOut "CParamsDlg, ParseArguments: %s", addr szText
			.if (eax || (dwCBData == VT_BSTR))
				inc eax
				mov dwSize, eax
				.if (dwCBData == VT_UNKNOWN)
					invoke GetDlgItemInt, m_hWnd, dwCtrlIDEdit, addr bTranslated, FALSE
					.if (bTranslated)
						mov [edi].vt, VT_UNKNOWN
						mov [edi].punkVal, eax
					.else
						invoke wsprintf, addr szText, CStr("No object"), eax
						StatusBar_SetText m_hWndSB, 0, addr szText
						invoke SetFocus, hWndEdit
						jmp error
					.endif
				.else
					invoke SysStringFromLPSTR, addr szText, dwSize
					mov [edi].vt,VT_BSTR
					mov [edi].bstrVal,eax
				.endif
				movzx ecx, [edi].vt
				.if (ecx != dwCBData)
					invoke VariantChangeType, edi, edi, 0, dwCBData
					.if (eax != S_OK)
						invoke wsprintf, addr szText, CStr("VariantChangeType failed [%X]"), eax
						StatusBar_SetText m_hWndSB, 0, addr szText
						invoke SetFocus, hWndEdit
						jmp error
					.endif
				.endif
;------------------------------- if it's a userdefined type try to "translate"
;------------------------------- the string we've got
				.if ([esi].tdesc.vt == VT_USERDEFINED)
					invoke TranslateUDT
				.else
					.if ([esi].tdesc.vt == VT_PTR)
						mov eax, [esi].tdesc.lptdesc
						movzx eax, [eax].TYPEDESC.vt
					.else
						movzx eax,[esi].tdesc.vt
					.endif
					.if (eax != VT_VARIANT)
						push eax
						invoke VariantChangeType, edi, edi, 0, eax
						pop ecx
						.if (eax != S_OK)
;------------------------- VariantChangeType has problems converting VT_BSTR to VT_LPWSTR!
;------------------------- so do it here (will cause a memory leak)
							.if (ecx == VT_LPWSTR)
								mov [edi].vt, cx
							.else
								invoke wsprintf, addr szText, CStr("VariantChangeType failed [%X]"), eax
								StatusBar_SetText m_hWndSB, 0, addr szText
								invoke SetFocus, hWndEdit
								jmp error
							.endif
						.endif
					.endif
if 0
					.if (eax == DISP_E_TYPEMISMATCH)
						.if ([esi].tdesc.vt == VT_PTR)
							mov eax, [esi].tdesc.lptdesc
							movzx ecx, [eax].TYPEDESC.vt
						.else
							movzx ecx,[esi].tdesc.vt
						.endif
;-------------------------------------- if parameter type is VT_VARIANT
;-------------------------------------- try to convert to number (VT_I4)
						.if (ecx == VT_VARIANT)
							invoke VariantChangeType, edi, edi, 0, VT_I4
						.endif
					.endif
endif
				.endif
			.elseif (dwCBData == VT_PTR)
;-------------------------------------- what to do here?
				mov ecx, [esi].tdesc.lptdesc
				movzx ecx, [ecx].TYPEDESC.vt
				.if (cx == VT_PTR)
					mov cx, VT_DISPATCH
				.endif
				or cx, VT_BYREF
				mov [edi].vt,cx
				invoke malloc, sizeof VARIANT
				mov [edi].byref, eax
			.else
				.if ((dwCBData != VT_EMPTY) && (dwCBData != VT_ERROR))
					StatusBar_SetText m_hWndSB, 0, CStr("no valid value for this type")
					invoke SetFocus, hWndCB
					jmp error
				.endif
			.endif

			inc dwCtrlIDEdit
			inc dwCtrlIDCB
			add esi,sizeof ELEMDESC
			pop ecx
			dec ecx
		.endw
		mov eax,1
		ret
error:
ifdef @StackBase
@StackBase = @StackBase + 4
endif
		pop ecx
		invoke MessageBeep, MB_OK
		xor eax, eax
		ret
		align 4
		assume esi:nothing, edi:nothing

ParseArguments endp


	.const

ctrltab dd IDC_BROWSE, IDC_PASTE, IDOK, IDCANCEL
SIZECTRLTAB equ ($ - ctrltab) / sizeof DWORD

unktypes	dw VT_UNKNOWN, -1
disptypes	dw VT_DISPATCH, -1
vartypes	dw VT_NULL, VT_I2, VT_I4, VT_R4, VT_R8, VT_CY, VT_DATE
			dw VT_BSTR, VT_DISPATCH, VT_BOOL, VT_UNKNOWN
			dw VT_I1, VT_UI1, VT_UI2, VT_UI4, -1
stdtypes	dw VT_BSTR, -1

	.code

;------ 

OnInitDialog proc uses esi edi __this

local	dwNumNames:dword
local	dwTmp:DWORD
local	pTypeInfoRef:LPTYPEINFO
local	pTypeAttr:ptr TYPEATTR
local	pVarDesc:ptr VARDESC
local	pVariant:ptr VARIANT
local	bstrVar:BSTR
local	hWndCB:HWND
local	dwSize:DWORD
local	dwParmIndex:DWORD
local	dwIndex:DWORD
local	rect:RECT
local	rect2:RECT
local	dwDiff:DWORD
local	dwDiffSum:DWORD
local	dwReturn:DWORD
local	pbstr:ptr BSTR
local	dwCtrlID: dword
local	dwCtrlIDEdit: dword
local	dwCtrlIDCB: dword
local	hWndChild:HWND
local	szText[256]:byte
local	wszText[256]:word
local	szStr[128]:byte
local	szName[64]:byte

		mov esi, m_pFuncDesc
		assume esi:ptr FUNCDESC

		invoke GetDlgItem, m_hWnd, IDC_STATUSBAR
		mov m_hWndSB, eax

		movzx eax,[esi].cParams
		mov edx,[esi].lprgelemdescParam
		mov ecx, eax
;---------------------------------------- filter retval + lcid
		mov edi, m_ParamReturn.pVariants
		.if (!edi)
			.while (ecx)
				.if ([edx].ELEMDESC.paramdesc.wParamFlags & (PARAMFLAG_FLCID or PARAMFLAG_FRETVAL))
					dec eax
				.endif
				add edx, sizeof ELEMDESC
				dec ecx
			.endw
			push eax
			mov ecx,sizeof VARIANT
			mul ecx
			invoke malloc, eax
			pop edx
			mov m_ParamReturn.pVariants, eax
			mov m_ParamReturn.iNumVariants, edx
			mov edi, eax
			.while (edx)
				push eax
				push edx
				invoke VariantInit, eax
				pop edx
				pop eax
				add eax,sizeof VARIANT
				dec edx
			.endw
		.endif
		assume edi:ptr VARIANT

		movzx ecx,[esi].cParams
		inc ecx
		mov dwNumNames, ecx
		mov eax,sizeof BSTR
		mul ecx
		invoke malloc, eax
		mov pbstr, eax
		invoke vf(m_pTypeInfo,ITypeInfo,GetNames),[esi].memid, pbstr, dwNumNames, addr dwReturn

		mov ecx, m_ParamReturn.iNumVariants

		mov eax,sizeof VARIANT
		mul ecx
		add edi,eax
;--------------------------------------- change controls IDC_STATICx in dlg
;--------------------------------------- to display names of parameters;
;--------------------------------------- Enable IDC_EDITx controls; init variant
		mov dwCtrlID, IDC_STATIC1
		mov dwCtrlIDCB, IDC_COMBO1
		mov dwCtrlIDEdit, IDC_EDIT1

		mov dwParmIndex, 0
		movzx ecx,[esi].cParams
		mov esi,[esi].lprgelemdescParam
		assume esi:ptr ELEMDESC
		.while (ecx)
;--------------------------------------- ignore LCID or retval parameters
			.if ([esi].paramdesc.wParamFlags & (PARAMFLAG_FLCID or PARAMFLAG_FRETVAL))
				add esi,sizeof ELEMDESC
				inc dwParmIndex
				dec ecx
				.continue
			.endif
			push ecx
			sub edi,sizeof VARIANT
			invoke GetDlgItem, m_hWnd, dwCtrlIDEdit
			.if (eax)
				mov hWndChild, eax

				invoke GetDlgItem, m_hWnd, dwCtrlIDCB
				mov hWndCB, eax
				invoke GetVarType, VT_EMPTY
				invoke ComboBox_AddString( hWndCB, eax)
				invoke ComboBox_SetItemData( hWndCB, eax, VT_EMPTY)
				movzx ecx, [esi].tdesc.vt
				.if (ecx == VT_PTR)
					mov eax, [esi].tdesc.lptdesc
					movzx ecx, [eax].TYPEDESC.vt
				.endif
				.if (ecx == VT_UNKNOWN)
					mov eax, offset unktypes
				.elseif (ecx == VT_DISPATCH)
					mov eax, offset disptypes
				.elseif (ecx == VT_VARIANT)
					mov eax, offset vartypes
				.else
					mov eax, offset stdtypes
				.endif
				push esi
				mov esi, eax
				.while (word ptr [esi] != -1)
					lodsw
					movzx eax, ax
					push eax
					invoke GetVarType, eax
					invoke ComboBox_AddString( hWndCB, eax)
					pop ecx
					invoke ComboBox_SetItemData( hWndCB, eax, ecx)
				.endw
				pop esi
				movzx eax, [edi].vt
				.if (eax == VT_ERROR)
					mov eax, VT_EMPTY
				.endif
				invoke GetVarType, eax
				invoke ComboBox_SelectString( hWndCB, -1, eax)
				.if (eax == CB_ERR)
					invoke GetVarType, VT_BSTR
					invoke ComboBox_SelectString( hWndCB, -1, eax)
				.endif

				.if ([edi].vt != VT_EMPTY)
					invoke VariantChangeType, edi, edi, 0, VT_BSTR
					.if (eax == S_OK)
						invoke SysStringLen, [edi].bstrVal
						add eax,4
						and al, 0FCh
						sub esp, eax
						mov dwSize, eax
						mov ecx, esp
						invoke WideCharToMultiByte, CP_ACP, 0, [edi].bstrVal, -1, ecx, dwSize, 0, 0
						mov ecx, esp
						invoke SetWindowText, hWndChild, ecx
						add esp, dwSize
					.endif
					invoke VariantClear, edi
				.endif

				movzx ecx, [esi].tdesc.vt
				invoke SetWindowLong, hWndChild, GWL_USERDATA, ecx

;--------------------------------------- userdefined type: make a combobox

				.if ([esi].tdesc.vt == VT_USERDEFINED)
					invoke GetWindowRect, hWndChild, addr rect
					mov eax, rect.right
					sub eax, rect.left
					mov rect.right, eax
					mov eax, rect.bottom
					sub eax, rect.top
					shl eax, 2
					mov rect.bottom, eax
					invoke ScreenToClient, m_hWnd, addr rect
					invoke SendMessage, hWndChild, WM_GETFONT, 0, 0
					push eax
					invoke DestroyWindow, hWndChild
					invoke CreateWindowEx, WS_EX_CLIENTEDGE, CStr("combobox"), NULL,
						WS_CHILD or WS_VISIBLE or WS_VSCROLL or WS_TABSTOP or CBS_DROPDOWN or CBS_AUTOHSCROLL,
						rect.left, rect.top, rect.right, rect.bottom, m_hWnd, dwCtrlIDEdit, g_hInstance, NULL
					mov hWndCB, eax
					mov hWndChild, eax
					pop ecx
					invoke SendMessage, eax, WM_SETFONT, ecx, 0
					invoke vf(m_pTypeInfo, ITypeInfo, GetRefTypeInfo), [esi].tdesc.hreftype, addr pTypeInfoRef
					.if (eax == S_OK)
						invoke vf(pTypeInfoRef, ITypeInfo, GetTypeAttr), addr pTypeAttr
						.if (eax == S_OK)
							mov edx, pTypeAttr
							movzx ecx, [edx].TYPEATTR.cVars
							mov dwIndex, 0
							.while (ecx)
								push ecx
								invoke vf(pTypeInfoRef, ITypeInfo, GetVarDesc), dwIndex, addr pVarDesc
								.if (eax == S_OK)
									mov ecx, pVarDesc
									invoke vf(pTypeInfoRef, ITypeInfo, GetNames), [ecx].VARDESC.memid, addr bstrVar, 1, addr dwTmp
									.if (eax == S_OK)
										invoke SysStringLen, bstrVar
										add eax, 4
										and al,0FCh
										mov dwSize, eax
										sub esp, eax
										mov ecx, esp
										mov edx, pVariant
									    invoke WideCharToMultiByte,CP_ACP,0, bstrVar,-1,ecx,dwSize,0,0 
										mov ecx, esp
										invoke ComboBox_AddString( hWndCB, ecx)
										add esp, dwSize
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
					.endif
				.endif		;end UDT
;--------------------------------------- transform name to ASCIIZ
				mov szText,0
				mov edx, dwParmIndex
				inc edx					;that's a 1-based index
				shl edx, 2				;* 4 (= sizeof BSTR)
				add edx, pbstr
				.if (dword ptr [edx] == NULL)
					mov edx, pbstr
				.endif
				invoke WideCharToMultiByte,CP_ACP,0,[edx],-1,addr szText,sizeof szText,0,0 

;--------------------------------------- get parameter type

				invoke GetParameterType, m_pTypeInfo, esi, addr szStr, sizeof szStr
				invoke lstrcat, addr szText, CStr(": ")
				invoke lstrcat, addr szText, addr szStr
				.if ([esi].paramdesc.wParamFlags & PARAMFLAG_FOPT)
					invoke lstrcat, addr szText, CStr(" (opt)")
				.endif

;--------------------------------------- set prompt

				invoke SetDlgItemText, m_hWnd, dwCtrlID, addr szText

				movzx eax, [esi].paramdesc.wParamFlags
				and eax, PARAMFLAG_FIN or PARAMFLAG_FOUT
				.if (eax == PARAMFLAG_FOUT)
					invoke EnableWindow, hWndChild, FALSE
					invoke EnableWindow, hWndCB, FALSE
					movzx ecx, [esi].tdesc.vt
					.if (ecx == VT_PTR)
						invoke GetVarType, VT_PTR
						push eax
						invoke ComboBox_AddString( hWndCB, eax)
						invoke ComboBox_SetItemData( hWndCB, eax, VT_PTR)
						pop eax
						invoke ComboBox_SelectString( hWndCB, -1, eax)
					.endif
				.endif

				inc dwCtrlID
				inc dwCtrlIDEdit
				inc dwCtrlIDCB
			.endif
			pop ecx
			add esi,sizeof ELEMDESC
			inc dwParmIndex
			dec ecx
		.endw

		mov ecx, dwCtrlID
		sub ecx, IDC_STATIC1-1
		mov m_dwRows, ecx

		assume esi:nothing
		assume edi:nothing


		invoke GetDlgItem, m_hWnd, dwCtrlID
		.if (eax)
			mov hWndChild, eax
			mov dwDiffSum, 0
			invoke GetDlgItem, m_hWnd, IDC_STATIC1
			lea ecx, rect
			invoke GetWindowRect, eax, ecx
			invoke GetDlgItem, m_hWnd, IDC_STATIC2
			lea ecx, rect2
			invoke GetWindowRect, eax, ecx
			mov eax, rect2.top
			sub eax, rect.top
			mov dwDiff, eax
			invoke GetWindowRect, hWndChild, addr rect
			invoke ScreenToClient, m_hWnd, addr rect
			.while (1)
				invoke GetDlgItem, m_hWnd, dwCtrlID
				.break .if (!eax)
				invoke DestroyWindow, eax
				invoke GetDlgItem, m_hWnd, dwCtrlIDEdit
				invoke DestroyWindow, eax
				invoke GetDlgItem, m_hWnd, dwCtrlIDCB
				invoke DestroyWindow, eax
				mov eax, dwDiff
				add dwDiffSum, eax
				inc dwCtrlID
				inc dwCtrlIDEdit
				inc dwCtrlIDCB
			.endw
;---------------------------- move buttons (Browse, Paste, Ok, Cancel) just
;---------------------------- below last control
			mov ecx, SIZECTRLTAB
			mov esi, offset ctrltab
			.while (ecx)
				push ecx
				lodsd
				invoke GetDlgItem, m_hWnd, eax
				mov hWndChild, eax
				invoke GetWindowRect, hWndChild, addr rect2
				invoke ScreenToClient, m_hWnd, addr rect2
				mov eax, rect.top
				mov rect2.top,eax
				invoke SetWindowPos, hWndChild, NULL, rect2.left, rect2.top, 0, 0, SWP_NOSIZE or SWP_NOZORDER or SWP_NOACTIVATE
				pop ecx
				dec ecx
			.endw
;---------------------------- now adjust height of main dialog wnd
			invoke GetWindowRect, m_hWnd, addr rect
			mov eax, rect.bottom
			sub	eax, dwDiffSum
;---------------------------- 4. compute dx (doesnt change) and dy
			sub eax, rect.top
			mov rect.bottom, eax	;dy of dialog wnd
			mov eax, rect.right
			sub eax, rect.left
			mov rect.right, eax		;dx of dialog wnd
			invoke SetWindowPos, m_hWnd, NULL, 0, 0, rect.right, rect.bottom, SWP_NOMOVE or SWP_NOZORDER or SWP_NOACTIVATE
			invoke SetWindowPos, m_hWndSB, NULL, 0, 0, 0, 0, SWP_NOZORDER or SWP_NOACTIVATE
		.endif

;---------------------------- free array of bstrs
		.if (pbstr)
			mov ecx, dwNumNames
			mov esi, pbstr
			.while (ecx)
				push ecx
				lodsd
				invoke SysFreeString, eax
				pop ecx
				dec ecx
			.endw
			invoke free, pbstr
		.endif

		ret
		align 4

OnInitDialog endp

CParamsDialog proc uses __this thisarg, message:dword, wParam:WPARAM, lParam:LPARAM

local	hWndCB:HWND
local	hWndEdit:HWND
local	dwData:DWORD
local	bTranslated:BOOL
local	pDataObject:LPDATAOBJECT
local	fe:FORMATETC
local	stm:STGMEDIUM
local	szText[32]:byte
local	wp:WINDOWPLACEMENT

		mov __this,this@

		mov eax,message
		.if (eax == WM_INITDIALOG)

			.if (g_rect.left)
				invoke SetWindowPos, m_hWnd, NULL, g_rect.left, g_rect.top,
					0, 0, SWP_NOZORDER or SWP_NOACTIVATE or SWP_NOSIZE
			.else
				invoke CenterWindow, m_hWnd
			.endif
			invoke OnInitDialog
			invoke GetDlgItem, m_hWnd, IDC_EDIT1
			invoke SetFocus, eax			
			mov eax,FALSE

		.elseif (eax == WM_CLOSE)

			mov wp.length_, sizeof WINDOWPLACEMENT
			invoke GetWindowPlacement, m_hWnd, addr wp
			invoke CopyRect, addr g_rect, addr wp.rcNormalPosition

			invoke EndDialog, m_hWnd, m_dwRC
			mov eax,1

		.elseif (eax == WM_DESTROY)

			invoke Destroy@CParamsDlg, __this

		.elseif (eax == WM_COMMAND)

			mov ecx, m_dwRows
			add ecx, IDC_EDIT1
			mov edx, m_dwRows
			add edx, IDC_COMBO1

			movzx eax,word ptr wParam
			.if (eax == IDCANCEL)

				invoke PostMessage,m_hWnd,WM_CLOSE,0,0

			.elseif (eax == IDOK)

				invoke ParseArguments
				.if (eax)
					invoke m_execcb, m_pPropertiesDlg, m_hWnd, m_pFuncDesc,
						m_ParamReturn.iNumVariants, m_ParamReturn.pVariants, m_hWndSB
					.if (eax == S_OK)
						mov m_dwRC, eax
						invoke PostMessage,m_hWnd,WM_CLOSE,0,0
					.endif
				.endif

			.elseif (eax == IDC_BROWSE)

				sub esp, MAX_PATH
				mov edx, esp
				invoke OnBrowse, edx, MAX_PATH
				mov edx, esp
				.if (byte ptr [edx])
					invoke SetWindowText, m_hWndCurEdit, edx
				.endif
				add esp, MAX_PATH

			.elseif (eax == IDC_PASTE)

				invoke OleGetClipboard, addr pDataObject
				.if (eax == S_OK)
					mov eax, g_dwMyCBFormat
					mov fe.cfFormat, ax
					mov fe.ptd, NULL
					mov fe.dwAspect, DVASPECT_CONTENT
					mov fe.lindex, -1
					mov fe.tymed, TYMED_HGLOBAL
					invoke vf(pDataObject, IDataObject, GetData), addr fe, addr stm
					.if (eax == S_OK)
						.if (stm.tymed == TYMED_HGLOBAL)
							invoke GlobalLock, stm.hGlobal
							mov eax, [eax].VARIANT.punkVal
							push eax
							invoke wsprintf, addr szText, CStr("%u"), eax
							invoke GlobalUnlock, stm.hGlobal
							invoke EnableWindow, m_hWndCurEdit, FALSE
							invoke SetWindowText, m_hWndCurEdit, addr szText
							mov eax, m_dwIDCurEdit
							sub eax, IDC_EDIT1
							add eax, IDC_COMBO1
							invoke GetDlgItem, m_hWnd, eax
							push eax
							invoke GetVarType, VT_UNKNOWN
							pop ecx
							invoke ComboBox_SelectString( ecx, -1, eax)
							pop eax
							invoke vf(eax, IUnknown, AddRef)
						.endif
						invoke ReleaseStgMedium, addr stm
					.endif
					invoke vf(pDataObject, IUnknown, Release)
				.endif

			.elseif ((eax >= IDC_EDIT1) && (eax < ecx))

				mov m_dwIDCurEdit, eax
				mov eax, lParam
				mov m_hWndCurEdit, eax
				invoke IsWindowEnabled, eax
				.if (eax)
					movzx ecx,word ptr wParam+2
					.if ((ecx == EN_UPDATE) || (ecx == CBN_SELENDOK))
						movzx eax,word ptr wParam
						sub eax, IDC_EDIT1
						add eax, IDC_COMBO1
						invoke GetDlgItem, m_hWnd, eax
						mov hWndCB, eax
						invoke ComboBox_GetCurSel( eax)
						.if (eax != CB_ERR)
							invoke ComboBox_GetItemData( hWndCB, eax)
							mov dwData, eax
						.endif
						movzx eax, word ptr wParam+2
						.if (eax == EN_UPDATE)
							invoke GetWindowTextLength, lParam
						.endif
						.if (eax)
							.if (dwData == VT_EMPTY)
								movzx ecx, word ptr wParam
								invoke GetDlgItemInt, m_hWnd, ecx, addr bTranslated, FALSE
								.if (bTranslated)
									invoke GetVarType, VT_I4
									invoke ComboBox_SelectString( hWndCB, -1, eax)
								.endif
								.if ((!bTranslated) || (eax == CB_ERR))
									invoke GetVarType, VT_BSTR
									invoke ComboBox_SelectString( hWndCB, -1, eax)
								.endif
								.if (eax == CB_ERR)
									invoke MessageBeep, MB_OK
									StatusBar_SetText m_hWndSB, 0, CStr("no keyboard input allowed")
									invoke SetFocus, lParam
									invoke SetWindowText, lParam, addr g_szNull
								.endif
							.elseif ((dwData == VT_UNKNOWN) || (dwData == VT_DISPATCH))
								invoke MessageBeep, MB_OK
								StatusBar_SetText m_hWndSB, 0, CStr("no keyboard input allowed")
								invoke SetFocus, lParam
								invoke SetWindowText, lParam, addr g_szNull
							.endif
						.endif

					.elseif (ecx == EN_SETFOCUS)

						invoke GetDlgItem, m_hWnd, IDC_PASTE
						push eax
						invoke GetWindowLong, lParam, GWL_USERDATA
						.if ((eax == VT_UNKNOWN) || (eax == VT_DISPATCH))
							invoke IsClipboardFormatAvailable, g_dwMyCBFormat
						.else
							xor eax, eax
						.endif
						pop ecx
						invoke EnableWindow, ecx, eax
					.endif		;EN_SETFOCUS
				.endif			;IsWindowEnabled

			.elseif ((eax >= IDC_COMBO1) && (eax < edx))

				movzx ecx,word ptr wParam+2
				.if (ecx == CBN_SELENDOK)
					invoke ComboBox_GetCurSel( lParam)
					invoke ComboBox_GetItemData( lParam, eax)
					.if (eax == VT_UNKNOWN)
						mov ecx, ecx
					.endif
				.endif

			.endif

if ?HTMLHELP
		.elseif (eax == WM_HELP)

			invoke DoHtmlHelp, HH_DISPLAY_TOPIC, CStr("ParamDialog.htm")
endif
		.else
			xor eax,eax ;indicates "no processing"
		.endif
		ret
		align 4

CParamsDialog endp

	end
