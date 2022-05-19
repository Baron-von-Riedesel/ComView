

;*** definition of class CCreateInclude
;*** CCreateInclude creates ASM includes
;*** has 2 constructors:
;*** 1. will require GUID, lcid, verMajor + verMinor
;*** 2. will require LPSTR (filename)

	.386
	.model flat,stdcall
	option casemap :none   ; case sensitive
	option proc:private

	include COMView.inc
INSIDE_CCREATEINCLUDE equ 1
	include classes.inc
	include rsrc.inc
	include debugout.inc


;---- we scan the typeinfos 2 times to check for dependencies
CType struct
dwIdx	DWORD ?
dwPrty	DWORD ?
pszCur	LPSTR ?
pszDep	LPSTR ?
CType ends


CCreateInclude struct
hWnd		HWND	?
dwMode		DWORD	?
pszTypeLib	LPSTR	?
if 0
pszGUID		LPSTR	?
lcid		LCID	?
dwMajor		DWORD	?
dwMinor		DWORD	?
endif
pTypeLib	LPTYPELIB ?
pTypeInfo	LPTYPEINFO ?
hFile		DWORD	?
pszOut		LPSTR	?
CCreateInclude ends

__this	textequ <ebx>
_this	textequ <[__this].CCreateInclude>

	MEMBER hWnd, dwMode, pszTypeLib
if 0
	MEMBER pszGUID, lcid, dwMajor, dwMinor
endif
	MEMBER pTypeLib, pTypeInfo, hFile, pszOut

	.data

g_pDepArray		LPSTR 0
g_pCurDep		LPSTR 0
g_iPass			DWORD 0
g_bWrite		BOOL TRUE

	.code

WriteString proc pStr:ptr byte

local	dwWritten:DWORD

	.if (g_bWrite)
		invoke lstrlen, pStr
		mov ecx,eax
		invoke WriteFile, m_hFile, pStr, ecx, addr dwWritten, NULL
	.endif
	ret
	align 4

WriteString endp


WriteLine proc pStr:ptr byte

	invoke WriteString, pStr
	invoke WriteString, CStr(0dh,0ah)
	ret
	align 4
WriteLine endp

IsRecord proc uses edi pTypeDesc:ptr TYPEDESC

local	pTypeInfoRef:LPTYPEINFO
local	pTypeAttrRef:ptr TYPEATTR
local	bIsRecord:BOOL

	mov edi, pTypeDesc
	mov bIsRecord, FALSE
	.if ([edi].TYPEDESC.vt == VT_USERDEFINED)
		invoke vf(m_pTypeInfo,ITypeInfo,GetRefTypeInfo), [edi].TYPEDESC.hreftype, addr pTypeInfoRef
		.if (eax == S_OK)
			invoke vf(pTypeInfoRef, ITypeInfo, GetTypeAttr), addr pTypeAttrRef
			.if (eax == S_OK)
				mov ecx, pTypeAttrRef
				.if ([ecx].TYPEATTR.typekind == TKIND_RECORD)
					mov bIsRecord, TRUE
				.endif
				invoke vf(pTypeInfoRef, ITypeInfo, ReleaseTypeAttr), pTypeAttrRef
			.endif
			invoke vf(pTypeInfoRef, ITypeInfo, Release)
		.endif
	.endif
	return bIsRecord
	align 4

IsRecord endp

;--- scan variable table
;--- that's RECORD, UNION, ENUM, ALIAS, (MODULE), (DISPATCH) types 


ScanTypeInfoVariables proc uses esi edi pszName:ptr byte, pszFirst:LPSTR, dwMode:DWORD

local	pTypeAttr:ptr TYPEATTR
local	pVarDesc:ptr VARDESC
local	rect:RECT
local	dwReturn:dword
local	bstr:BSTR
local	pArray:ptr
local	dwIndex:DWORD
local	pVariant:ptr VARIANT
local	pStrTmp:LPSTR
local	szHex[32]:byte
local	szStr[512]:byte
local	szVariant[260]:byte
local	szText[128]:byte


		invoke vf(m_pTypeInfo,ITypeInfo,GetTypeAttr),addr pTypeAttr
		.if (eax == S_OK)

			mov esi,pTypeAttr
			assume esi:ptr TYPEATTR

			mov szStr, 0
			mov dwIndex,0

			mov eax, [esi].typekind
;------------------------------------------- for dispatch helpers only watch typekind DISPATCH
			.if ((eax != TKIND_DISPATCH) && (dwMode == INCMODE_DISPHLP))
				jmp done
			.endif

			.if (eax == TKIND_ENUM)

				.if (g_iPass != 1)
					mov g_bWrite, FALSE
				.endif
				invoke wsprintf,addr szStr,CStr("%s typedef SDWORD"),pszName

			.elseif (eax == TKIND_ALIAS)

				.if (g_iPass != 2)
					mov g_bWrite, FALSE
				.endif
				lea ecx,[esi].tdescAlias
				mov szVariant,0
				invoke GetParameterTypeAsm, m_pTypeInfo, ecx,\
					addr szVariant, sizeof szVariant
				invoke wsprintf,addr szStr,CStr("%s typedef %s"),pszName,addr szVariant

			.elseif (eax == TKIND_RECORD)

				.if (g_iPass != 2)
					mov g_bWrite, FALSE
				.endif

				.if ([esi].cbAlignment > 1)
					movzx ecx, [esi].cbAlignment
					invoke wsprintf, addr szStr, CStr("%s struct %u"),pszName, ecx
				.else
					invoke wsprintf, addr szStr, CStr("%s struct"),pszName
				.endif

			.elseif (eax == TKIND_UNION)

				.if (g_iPass != 2)
					mov g_bWrite, FALSE
				.endif
				.if ([esi].cbAlignment > 1)
					movzx ecx, [esi].cbAlignment
					invoke wsprintf, addr szStr, CStr("%s union %u"), pszName, ecx
				.else
					invoke wsprintf, addr szStr, CStr("%s union"), pszName
				.endif

			.elseif (eax == TKIND_DISPATCH)

				.if (g_iPass != 2)
					mov g_bWrite, FALSE
				.endif

			.else

				.if (g_iPass != 1)
					mov g_bWrite, FALSE
				.endif

			.endif
			
			.if (szStr || [esi].cVars)
				invoke WriteString, pszFirst
				mov ecx, pszFirst
				mov byte ptr [ecx],0
				.if (szStr)
					invoke WriteLine, addr szStr
				.endif
			.endif

			.while (1)
				mov eax, dwIndex
				.break .if (ax >= [esi].cVars)

				invoke vf(m_pTypeInfo, ITypeInfo, GetVarDesc), dwIndex, addr pVarDesc
				.if (eax == S_OK)

					mov edi,pVarDesc

					mov bstr, NULL

					invoke vf(m_pTypeInfo,ITypeInfo,GetNames),[edi].VARDESC.memid, addr bstr, 1, addr dwReturn

					.if (bstr)
		    			invoke WideCharToMultiByte,CP_ACP,0,bstr,-1, addr szStr, sizeof szStr, 0, 0
					.else
						invoke lstrcpy,addr szStr,CStr("?")
					.endif

;------------------------------------ if var is a VAR_DISPATCH, define DISPID only
;------------------------------------ define it as comment only!

					.if ([edi].VARDESC.varkind == VAR_DISPATCH)
						.if (dwMode == INCMODE_BASIC)
							invoke WriteString, CStr(";DISPID_")
						.else
							invoke lstrcpy, addr szVariant, addr szStr
							mov szText,0
							invoke GetParameterTypeStub, m_pTypeInfo,
								addr [edi].VARDESC.elemdescVar, addr szText, sizeof szText
							invoke wsprintf, addr szStr,
								CStr(9,"DEFINE_DISPMETHOD",9,"%s , get_%s, 0%Xh, PROPERTYGET, %s"),
								pszName, addr szVariant, [edi].VARDESC.memid, addr szText
							invoke WriteLine,  addr szStr
							invoke wsprintf, addr szStr,
								CStr(9,"DEFINE_DISPMETHOD",9,"%s , put_%s, 0%Xh, PROPERTYPUT, , %s"),
								pszName, addr szVariant, [edi].VARDESC.memid, addr szText
							jmp renderdone
						.endif
					.endif


					.if ([edi].VARDESC.varkind == VAR_PERINSTANCE)
						invoke CheckReservedNames, addr szStr
					.endif
					invoke WriteString, addr szStr

					.if ([edi].VARDESC.varkind == VAR_CONST)
if 0;def _DEBUG
						.if ([esi].typekind == TKIND_MODULE)
							invoke DebugBreak
						.endif
endif
						mov eax,[edi].VARDESC.lpvarValue
						.if ([eax].VARIANT.vt == VT_BSTR)
							invoke WriteString, CStr(9,"TEXTEQU",9,3Ch)
						.else
							invoke WriteString, CStr(9,"EQU",9)
						.endif
						invoke GetVariant, [edi].VARDESC.lpvarValue,
							addr szVariant,sizeof szVariant, NULL
						mov eax,[edi].VARDESC.lpvarValue
						mov ecx,1
						.if ([eax].VARIANT.vt == VT_I4)
							mov eax,[eax].VARIANT.lVal
						.elseif ([eax].VARIANT.vt == VT_UI4)
							mov eax,[eax].VARIANT.ulVal
						.elseif ([eax].VARIANT.vt == VT_I2)
							movsx eax,[eax].VARIANT.iVal
						.elseif ([eax].VARIANT.vt == VT_UI2)
							movzx eax,[eax].VARIANT.uiVal
						.elseif ([eax].VARIANT.vt == VT_I1)
							movsx eax,[eax].VARIANT.cVal
						.elseif ([eax].VARIANT.vt == VT_UI1)
							movzx eax,[eax].VARIANT.bVal
						.elseif ([eax].VARIANT.vt == VT_BSTR)
							push esi
							mov esi,[eax].VARIANT.byref
							mov szVariant,0
							.if (esi)
								invoke SysStringLen, esi
								mov ecx,eax
								.if (ecx > 32)
									mov ecx,32
								.endif
								.while (ecx)
									lodsw
									movzx eax,ax
									push ecx
									invoke wsprintf, addr szHex, CStr("0%Xh,"),eax
									invoke lstrcat, addr szVariant, addr szHex
									pop ecx
									dec ecx
								.endw
								invoke lstrlen, addr szVariant
								.if (eax)
									lea ecx,szVariant
									mov byte ptr [ecx+eax-1],0
								.endif
							.else
								invoke lstrcat, addr szVariant, CStr("NULL")
							.endif
							invoke lstrcat, addr szVariant, CStr(3Eh)
							mov ecx,0
							pop esi
						.else
							mov ecx,0
						.endif
						.if (ecx && eax > 15 && eax != -1)
							invoke wsprintf,addr szHex,CStr(9,9,9,";=0%Xh"),eax
						.else
							mov szHex,0
						.endif
						invoke wsprintf, addr szStr, CStr("%s%s"), addr szVariant, addr szHex

					.elseif ([edi].VARDESC.varkind == VAR_PERINSTANCE)

						invoke WriteString, CStr(9)
						mov szStr,0
						invoke GetParameterTypeAsm, m_pTypeInfo,
							addr [edi].VARDESC.elemdescVar, addr szStr, sizeof szStr
						mov eax,dword ptr szStr
						or eax, 202020h
						.if (eax == " rtp" || eax == "rtp")
							invoke lstrcpy, addr szStr, CStr("LPVOID")
						.endif

						.if ([edi].VARDESC.elemdescVar.tdesc.vt == VT_CARRAY)

							mov eax, [edi].VARDESC.elemdescVar.tdesc.lpadesc
							invoke IsRecord, addr [eax].ARRAYDESC.tdescElem
							.if (eax)
								mov ecx, g_pCurDep
								or [ecx].CType.dwPrty, 1
								mov pStrTmp, CStr(9,"(",3Ch,3Eh,")")
							.else
								mov pStrTmp, CStr(9,"(?)")
							.endif

						.elseif ([edi].VARDESC.elemdescVar.tdesc.vt == VT_SAFEARRAY)

							mov pStrTmp, CStr(9,3Ch,3Eh)

						.elseif ([edi].VARDESC.elemdescVar.tdesc.vt == VT_USERDEFINED)

							invoke IsRecord, addr [edi].VARDESC.elemdescVar.tdesc
							.if (eax)
								mov ecx, g_pCurDep
								or [ecx].CType.dwPrty, 1
								mov pStrTmp, CStr(9,3Ch,3Eh)
							.else
								mov pStrTmp, CStr(9,"?")
							.endif

						.else
							mov pStrTmp, CStr(9,"?")
						.endif
						invoke lstrcat, addr szStr, pStrTmp

					.elseif ([edi].VARDESC.varkind == VAR_DISPATCH)

						invoke WriteString, CStr(9,"EQU",9)
						invoke wsprintf, addr szStr, CStr("0%Xh"),[edi].VARDESC.memid

					.else
						mov szStr,0
					.endif
renderdone:
					invoke WriteLine,  addr szStr

					.if (bstr)
						invoke SysFreeString,bstr
					.endif

					invoke vf(m_pTypeInfo,ITypeInfo,ReleaseVarDesc), edi
				.endif
				inc dwIndex
			.endw
			.if (([esi].typekind == TKIND_RECORD) || ([esi].typekind == TKIND_UNION))
				invoke wsprintf,addr szStr,CStr("%s ends"),pszName
				invoke WriteLine, addr szStr
			.endif
			.if (szStr)
				invoke WriteLine, CStr("")
			.endif
done:
			invoke vf(m_pTypeInfo,ITypeInfo,ReleaseTypeAttr), esi
		.endif
		ret
		align 4
		assume esi:nothing

ScanTypeInfoVariables endp

;------------- the worker proc: scan all functions of a typeinfo
;------------- this will write a BEGIN_INTERFACE ... END_INTERFACE block

ScanTypeInfoFunctions proc uses esi edi pszName:ptr byte, pszFirst:LPSTR, dwMode:DWORD

local	pTypeAttr:ptr TYPEATTR
local	pTypeInfoRef:LPTYPEINFO
local	pTypeAttrRef:ptr TYPEATTR
local	pFuncDesc:ptr FUNCDESC
local	typekind:TYPEKIND
local	szStr[512]:byte
local	szName[128]:byte
local	szNameRef[128]:byte
local	szMethod[64]:byte
local	szPrefix[128]:byte
local	rect:RECT
local	dwReturn:dword
local	bstr:BSTR
local	pbstr:ptr BSTR
local	pArray:ptr
local   dwNumNames:dword
local	reftype:DWORD
local	bVtbl:BOOL
local	bStdMethod:BOOL
local	dwIndex:DWORD
local	dwInhVftSize:DWORD


		invoke vf(m_pTypeInfo, ITypeInfo, GetTypeAttr),addr pTypeAttr
		.if (eax != S_OK)
			xor eax,eax
			ret
		.endif
		.if (g_iPass != 2)
			mov g_bWrite, FALSE
		.endif

;------------- only check types TKIND_DISPATCH + TKIND_INTERFACE + TKIND_MODULE

		mov esi,pTypeAttr
		mov eax, [esi].TYPEATTR.typekind

;;		.if (eax != TKIND_DISPATCH && eax != TKIND_INTERFACE)
		.if ((eax != TKIND_DISPATCH) && (eax != TKIND_INTERFACE) && (eax != TKIND_MODULE))
			invoke vf(m_pTypeInfo,ITypeInfo,ReleaseTypeAttr),pTypeAttr
			ret
		.endif
		.if ((eax == TKIND_MODULE) && ([esi].TYPEATTR.cFuncs == 0))
			invoke vf(m_pTypeInfo,ITypeInfo,ReleaseTypeAttr),pTypeAttr
			ret
		.endif

;------------- if we are generating dispatch helpers, only check TKIND_DISPATCH
		.if (dwMode && (eax != TKIND_DISPATCH))
			invoke vf(m_pTypeInfo,ITypeInfo,ReleaseTypeAttr),pTypeAttr
			ret
		.endif

		mov typekind, eax				;save it so we can access it anywhere

		.if (g_iPass > 1)
			invoke WriteString, pszFirst
			mov ecx, pszFirst
			mov byte ptr [ecx],0
		.endif

		mov bVtbl,FALSE
		.if (typekind == TKIND_DISPATCH)
;------------- does a vtable exist?
			.if (!([esi].TYPEATTR.wTypeFlags & TYPEFLAG_FDUAL))
				.if (dwMode == 0)
					invoke wsprintf, addr szStr, CStr(";--- dispinterface only"), pszName
					invoke WriteLine, addr szStr
				.endif
			.else
;;				invoke MessageBox, 0, CStr("This shouldn't happen any more"), 0, MB_OK
				mov bVtbl,TRUE
			.endif
		.elseif (typekind == TKIND_INTERFACE)
			mov bVtbl,TRUE
		.endif

		mov szNameRef,0
		mov dwInhVftSize,0
		.if ([esi].TYPEATTR.cImplTypes)
			invoke vf(m_pTypeInfo, ITypeInfo, GetRefTypeOfImplType), 0, addr reftype
			.if (eax == S_OK)
				invoke vf(m_pTypeInfo,ITypeInfo,GetRefTypeInfo), reftype, addr pTypeInfoRef
				.if (eax == S_OK)
					invoke vf(pTypeInfoRef,ITypeInfo,GetTypeAttr), addr pTypeAttrRef
					.if (eax == S_OK)
						mov eax,pTypeAttrRef
						movzx eax,[eax].TYPEATTR.cbSizeVft
						mov dwInhVftSize, eax
						invoke vf(pTypeInfoRef,ITypeInfo,ReleaseTypeAttr), pTypeAttrRef
					.endif
					invoke vf(pTypeInfoRef,ITypeInfo,GetDocumentation),
							MEMBERID_NIL,addr bstr,NULL,NULL,NULL
					.if (bstr)
						invoke WideCharToMultiByte,CP_ACP,0,bstr,-1,addr szNameRef,sizeof szNameRef,0,0
						invoke SysFreeString, bstr
					.endif
					invoke vf(pTypeInfoRef,ITypeInfo,Release)
				.endif
			.endif
		.endif

		.if (typekind == TKIND_MODULE)
			;
		.elseif (dwMode == 0)
;-------------------------------------------- avoid multiple interface definitions
			invoke wsprintf, addr szStr, CStr("ifndef %s%s"), pszName, CStr("Vtbl")
			invoke WriteLine, addr szStr
			.if (szNameRef)
;------------------------------------- in first pass save name of type + parent
				.if (g_iPass == 1)
					invoke lstrlen, pszName
					inc eax
					invoke malloc, eax
					mov ecx, g_pCurDep
					mov [ecx].CType.pszCur,eax
					invoke lstrcpy, eax, pszName

					invoke lstrlen, addr szNameRef
					inc eax
					invoke malloc, eax
					mov ecx, g_pCurDep
					mov [ecx].CType.pszDep,eax
					mov ecx, eax
					invoke lstrcpy, ecx, addr szNameRef
				.endif
				invoke wsprintf, addr szStr, CStr("BEGIN_INTERFACE %s, %s"), pszName, addr szNameRef
			.else
				invoke wsprintf, addr szStr, CStr("BEGIN_INTERFACE %s"), pszName
			.endif
			invoke WriteLine, addr szStr
		.endif

		mov dwIndex,0
		.while (1)
			mov eax, dwIndex
			.break .if (ax >= [esi].TYPEATTR.cFuncs)
			push esi
			invoke vf(m_pTypeInfo,ITypeInfo,GetFuncDesc), dwIndex, addr pFuncDesc
			.if (eax == S_OK)
				mov esi,pFuncDesc
				movzx eax,[esi].FUNCDESC.oVft
				movzx ecx,[esi].FUNCDESC.wFuncFlags
				.if ((bVtbl == FALSE && (!(cx & FUNCFLAG_FRESTRICTED))) || (eax >= dwInhVftSize))
					movzx ecx,[esi].FUNCDESC.cParams
					inc ecx
					mov dwNumNames, ecx
					mov eax,sizeof BSTR
					mul ecx
					invoke malloc, eax
					mov pbstr, eax

					invoke vf(m_pTypeInfo,ITypeInfo,GetNames),[esi].FUNCDESC.memid,pbstr,dwNumNames,addr dwReturn

					mov szPrefix, 0
					mov edx,pbstr
					.if (eax == S_OK && dword ptr [edx])
						invoke WideCharToMultiByte,CP_ACP,0,[edx],-1,addr szMethod,sizeof szMethod,0,0
						mov edx,pbstr
						invoke SysFreeString,[edx]
					.else
						invoke lstrcpy, addr szMethod, CStr("?")
					.endif

					mov bStdMethod,FALSE
					mov eax,[esi].FUNCDESC.funckind
					.if (eax == FUNC_STATIC)
					.elseif	(eax == FUNC_NONVIRTUAL)
					.else
						mov bStdMethod,TRUE
						.if (typekind == TKIND_MODULE)
							;
						.elseif (dwMode == INCMODE_DISPHLP)
							invoke lstrcat, addr szPrefix, CStr(9,"DEFINE_DISPMETHOD",9)
							invoke lstrcat, addr szPrefix, pszName
							invoke lstrcat, addr szPrefix, CStr(" , ")
						.elseif (bVtbl)
							invoke lstrcat, addr szPrefix, CStr(9,"STDMETHOD",9)
						.else
							invoke lstrcat, addr szPrefix, CStr(9,"DISPMETHOD",9)
						.endif
					.endif
					mov eax,[esi].FUNCDESC.invkind
					.if (eax == INVOKE_PROPERTYGET)
						invoke lstrcat, addr szPrefix, CStr("get_")
					.elseif (eax == INVOKE_PROPERTYPUT)
						invoke lstrcat, addr szPrefix, CStr("put_")
					.elseif (eax == INVOKE_PROPERTYPUTREF)
						invoke lstrcat, addr szPrefix, CStr("putref_")
					.else
						invoke CheckReservedNames, addr szMethod
					.endif
;----------------------------------- copy function name
					invoke lstrcat, addr szPrefix, addr szMethod

					.if (bStdMethod == FALSE)
						invoke lstrcat, addr szPrefix, CStr(" proto")
					.endif

;					invoke wsprintf,addr szStr,CStr("0x%X"),[esi].FUNCDESC.memid
;					invoke GetParameterType, m_pTypeInfo,
;						addr [esi].FUNCDESC.elemdescFunc, addr szStr, sizeof szStr


					.if (dwMode)
						mov eax,[esi].FUNCDESC.invkind
						.if (eax == INVOKE_PROPERTYGET)
							mov ecx, CStr("PROPERTYGET")
						.elseif (eax == INVOKE_PROPERTYPUT)
							mov ecx, CStr("PROPERTYPUT")
						.elseif (eax == INVOKE_PROPERTYPUTREF)
							mov ecx, CStr("PROPERTYPUTREF")
						.else
							mov ecx, CStr("METHOD")
						.endif
						invoke wsprintf, addr szStr, CStr(", 0%Xh, %s, "),[esi].FUNCDESC.memid, ecx
						invoke lstrcat, addr szPrefix, addr szStr
						movzx eax, [esi].FUNCDESC.elemdescFunc.tdesc.vt
						.if (eax != VT_VOID && eax != VT_HRESULT)
							invoke GetParameterTypeStub, m_pTypeInfo,
								addr [esi].FUNCDESC.elemdescFunc, addr szStr, sizeof szStr
							invoke lstrcat, addr szPrefix, addr szStr
						.endif
					.else
						invoke lstrcat, addr szPrefix, CStr(9)
					.endif

					mov eax,[esi].FUNCDESC.invkind
					movzx ecx,[esi].FUNCDESC.cParams
					.if (ecx && bStdMethod)
						invoke lstrcat, addr szPrefix, CStr(", ")
					.endif

					pushad
					mov szStr,0
					mov edi,pbstr
					add edi, sizeof BSTR
					movzx ecx,[esi].FUNCDESC.cParams		;load into registers to
					mov esi,[esi].FUNCDESC.lprgelemdescParam
					.while (ecx)
						push ecx
						mov szName,0
						.if (dword ptr [edi])
							invoke WideCharToMultiByte,CP_ACP,0,[edi],-1,addr szName,sizeof szName,0,0 
							invoke SysFreeString,[edi]
						.endif
if 0	;; no parameter names, just types
						invoke lstrcat, addr szStr, addr szName
endif
						.if (dwMode == 0)
							invoke lstrcat, addr szStr, CStr(":")
						.endif
						invoke lstrlen, addr szStr
						lea ecx, szStr
						mov edx, sizeof szStr
						sub edx,eax
						add eax,ecx
						.if (dwMode == 0)
							invoke GetParameterTypeAsm, m_pTypeInfo, esi, eax, edx
						.else
							invoke GetParameterTypeStub, m_pTypeInfo, esi, eax, edx
						.endif
						pop ecx
						push ecx
						.if (ecx > 1)
							invoke lstrcat,addr szStr, CStr(",")
						.endif
						pop ecx
						dec ecx
						add esi,sizeof ELEMDESC
						add edi, sizeof BSTR
					.endw
					invoke free, pbstr
					popad

					movzx eax, [esi].FUNCDESC.elemdescFunc.tdesc.vt
					.if (typekind == TKIND_MODULE)
						;
					.elseif ((dwMode == 0) && (!bVtbl) && (eax != VT_VOID))
						invoke lstrcat, addr szStr, CStr(", :ptr ")
						invoke lstrlen, addr szStr
						lea ecx, szStr
						add ecx, eax
						mov edx, sizeof szStr
						sub edx, eax
						invoke GetParameterTypeAsm, m_pTypeInfo,
							addr [esi].FUNCDESC.elemdescFunc, ecx, edx
					.endif


					invoke lstrlen, addr szStr
					movzx ecx,[esi].FUNCDESC.cParams
;-------------------------------------------- if parameterstring too large
;-------------------------------------------- or too many parameters (29)
;-------------------------------------------- write prototype as comment
					.if ((eax > 255 || ecx > 29))
						invoke WriteString, CStr(";+++ ")
;-------------------------------------------- write a dummy prototype
						.if (bVtbl)
							invoke WriteString, addr szPrefix
							invoke WriteLine, addr szStr
							invoke lstrcpy, addr szStr, CStr(":DWORD")
						.endif
					.endif
;--------------------------------------------- now write to file
					invoke WriteString, addr szPrefix
					invoke WriteLine, addr szStr

					invoke vf(m_pTypeInfo,ITypeInfo,ReleaseFuncDesc),pFuncDesc
				.endif
			.endif
			inc dwIndex
			pop esi
		.endw
		.if (typekind == TKIND_MODULE)
			;
		.elseif (dwMode == 0)
			invoke WriteLine, CStr("END_INTERFACE")
			invoke WriteLine, CStr("endif")
		.endif
		invoke WriteLine, CStr("")
		invoke vf(m_pTypeInfo,ITypeInfo,ReleaseTypeAttr),pTypeAttr
		ret
		align 4

ScanTypeInfoFunctions endp

;------------- transform a GUID in ASM compatible syntax

GetAsmGUID proc pStrOut:ptr byte, pGUID: ptr GUID

local	dwTemp[11]:DWORD

		mov edx,pGUID
		mov eax,[edx].GUID.Data1
		mov dwTemp[0*sizeof DWORD],eax
		movzx eax,[edx].GUID.Data2
		mov dwTemp[1*sizeof DWORD],eax
		movzx eax,[edx].GUID.Data3
		mov dwTemp[2*sizeof DWORD],eax
		xor ecx,ecx
		.while (ecx < 8)
			movzx eax,byte ptr [ecx+edx].GUID.Data4
			mov dwTemp[3*sizeof DWORD+ecx*4],eax
			inc ecx
		.endw

		invoke wvsprintf,pStrOut,
				CStr("{0%08Xh,0%04Xh,0%04Xh,{0%02Xh,0%02Xh,0%02Xh,0%02Xh,0%02Xh,0%02Xh,0%02Xh,0%02Xh}}"),
				addr dwTemp
		ret
		align 4
GetAsmGUID endp

;------------- actually no interface scan is needed
;------------- some GUID TEXTEQUs are written though

ScanTypeInfoInterfaces proc uses esi edi pszName:ptr byte, pszFirst:LPSTR

local	pTypeAttr:ptr TYPEATTR
local	pFuncDesc:ptr FUNCDESC
local	szStr[512]:byte
local	szGUID[80]:byte
local	szName[128]:byte
local	rect:RECT
local	dwReturn:dword
local	pbstr:ptr BSTR

		.if (g_iPass != 2)		;write in second pass only
			mov g_bWrite, FALSE
		.endif

		invoke vf(m_pTypeInfo, ITypeInfo, GetTypeAttr),addr pTypeAttr
		.if (eax != S_OK)
			xor eax,eax
			ret
		.endif

		mov szStr, 0
		mov esi,pTypeAttr
		.if ([esi].TYPEATTR.typekind == TKIND_COCLASS)
			invoke GetAsmGUID, addr szGUID, addr [esi].TYPEATTR.guid
			invoke wsprintf, addr szStr, CStr("sCLSID_%s textequ ",3Ch,"GUID %s",3Eh),pszName, addr szGUID
		.elseif ([esi].TYPEATTR.typekind == TKIND_DISPATCH)
			invoke GetAsmGUID, addr szGUID, addr [esi].TYPEATTR.guid
			invoke wsprintf, addr szStr, CStr("sIID_%s textequ ",3Ch,"IID %s",3Eh),pszName, addr szGUID
		.elseif ([esi].TYPEATTR.typekind == TKIND_INTERFACE)
			invoke GetAsmGUID, addr szGUID, addr [esi].TYPEATTR.guid
			invoke wsprintf, addr szStr, CStr("sIID_%s textequ ",3Ch,"IID %s",3Eh),pszName, addr szGUID
		.endif

		.if (szStr)
			invoke WriteString, pszFirst
			invoke WriteLine, addr szStr
			invoke WriteLine, CStr("")
		.endif

		invoke vf(m_pTypeInfo,ITypeInfo,ReleaseTypeAttr),pTypeAttr
		ret
		align 4

ScanTypeInfoInterfaces endp

;--- get filename to save ASM include file for typelib

GetSourceFileName proc hWnd:HWND, pszName:ptr byte, pszExt:ptr byte, pszDesc:ptr byte

local	ofn:OPENFILENAME
local	szStr1[MAX_PATH]:byte
local	szStr2[128]:byte
local	szStr3[128]:byte


;------------------------------- prepare GetSaveFileName dialog

		invoke wsprintf, addr szStr1, CStr("%s%s"), pszName, pszExt
;;		invoke lstrcpy,addr szStr1, pszName
;;		invoke lstrcat,addr szStr1, CStr(".inc")

		invoke ZeroMemory,addr szStr3,sizeof szStr3
		invoke wsprintf, addr szStr3, CStr("%s (*%s)"), pszDesc, pszExt
;;		invoke lstrcpy,addr szStr3,CStr("Includes (*.inc)")
;;		invoke lstrlen,addr szStr3
		inc eax
		lea ecx,szStr3
		add ecx,eax
		invoke wsprintf, ecx, CStr("*%s"), pszExt
;;		invoke lstrcpy, ecx, CStr("*.inc")

		invoke ZeroMemory,addr szStr2,sizeof szStr2
		invoke lstrcpy,addr szStr2,CStr("All files (*.*)")
		invoke lstrlen,addr szStr2
		inc eax
		lea ecx,szStr2
		add ecx,eax
		invoke lstrcpy,ecx,CStr("*.*")

		invoke ZeroMemory,addr ofn,sizeof OPENFILENAME
		mov ofn.lStructSize,sizeof OPENFILENAME
		mov eax,hWnd
		mov ofn.hwndOwner,eax
		lea eax,szStr2
		mov ofn.lpstrFilter,eax

		lea eax,szStr3
		mov ofn.lpstrCustomFilter,eax
		mov ofn.nMaxCustFilter,sizeof szStr3

		mov ofn.nFilterIndex,0
		lea eax,szStr1
		mov ofn.lpstrFile,eax
		mov ofn.nMaxFile,sizeof szStr1
		mov ofn.Flags,OFN_EXPLORER or OFN_OVERWRITEPROMPT 

		invoke GetSaveFileName,addr ofn
		.if (eax != 0)
			invoke CreateFile, addr szStr1, GENERIC_WRITE, 0,
				NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL
			.if (eax == INVALID_HANDLE_VALUE)
				push eax
				invoke MessageBox, hWnd, CStr("cannot create file"), 0, MB_OK
				pop eax
			.endif
		.endif

		ret
		align 4

GetSourceFileName endp

ifdef @StackBase
	option stackbase:ebp
endif
	option prologue:@sehprologue
	option epilogue:@sehepilogue

;--------------- scan a type library to write ASM include

ScanTypeLib proc uses esi edi

local	guid:GUID
local	pTypeLib:LPTYPELIB
local	pTypeInfoRef:LPTYPEINFO
local	dwCount:dword
local	pTypeAttr:ptr TYPEATTR
local	pTLibAttr:ptr TLIBATTR
local	reftype:HREFTYPE
local	dwSize:dword
local	iType:dword
local	hKey:HANDLE
local	bstr1:BSTR
local	bstr2:BSTR
local	dwIndex:DWORD
local	fSwap:BOOL
local	dwExc:dword
local	dwExcAddr:dword
local	szStr[MAX_PATH]:byte
local	szTmpFile[MAX_PATH]:byte
local	szName[80]:byte
local	szDoc[128]:byte
local	szType[32]:byte
local	wszTypeLib[256]:word
local	wszGUID[40]:word
local	szGUID[80]:byte

;		invoke OutputDebugString, CStr("ScanTypeLib MS 1",13,10)
		.if (m_pTypeLib)
			mov eax, m_pTypeLib
			mov pTypeLib, eax
			invoke vf(pTypeLib, IUnknown, AddRef)
if 0
		.elseif (m_pszTypeLib == NULL)
			invoke MultiByteToWideChar,CP_ACP,MB_PRECOMPOSED, m_pszGUID, 40, addr wszGUID, 40
			invoke CLSIDFromString,addr wszGUID,addr guid
			invoke LoadRegTypeLib, addr guid, m_dwMajor, m_dwMinor, m_lcid, addr pTypeLib
endif
		.else
			invoke MultiByteToWideChar,CP_ACP,MB_PRECOMPOSED,
				m_pszTypeLib, -1, addr wszTypeLib, 256 
			.try
				invoke LoadTypeLibEx, addr wszTypeLib, REGKIND_NONE, addr pTypeLib
			.exceptfilter
				mov eax,_exception_info()
				mov eax, [eax].EXCEPTION_POINTERS.ExceptionRecord
				mov ecx, [eax].EXCEPTION_RECORD.ExceptionCode
				mov edx, [eax].EXCEPTION_RECORD.ExceptionAddress
				mov dwExc, ecx
				mov dwExcAddr,edx
				mov eax,EXCEPTION_EXECUTE_HANDLER
			.except
				mov eax, E_UNEXPECTED
			.endtry
			.if (eax != S_OK)
				.if (eax == E_UNEXPECTED)
					invoke printf, CStr("Exception 0x%08X occured at 0x%08X (LoadTypeLibEx)",13,10), dwExc, dwExcAddr
				.else
					invoke printf, CStr("LoadTypeLibEx() failed [%X]",13,10), eax
				.endif
				xor eax,eax
				ret
			.endif
		.endif
;		invoke OutputDebugString, CStr("ScanTypeLib MS 2",13,10)

;---------------------------------------------------- get name of TypeLib
		mov szName,0
		mov szDoc,0
		invoke vf(pTypeLib,ITypeLib,GetDocumentation),MEMBERID_NIL,
			addr bstr1, addr bstr2, NULL, NULL
		.if (eax == S_OK)
			mov eax,1
			.if (bstr1)
				invoke WideCharToMultiByte,CP_ACP,0,bstr1, -1 ,addr szName,sizeof szName,0,0 
				invoke SysFreeString,bstr1
			.endif
			.if (bstr2)
				invoke WideCharToMultiByte,CP_ACP,0,bstr2,-1,addr szDoc,sizeof szDoc,0,0
				invoke SysFreeString,bstr2
			.endif
		.endif
;-------------------------------------------- get name of file to write to
		.if (m_dwMode == INCMODE_DISPHLP)
			invoke lstrcat, addr szName, CStr("c")
		.endif
		mov szTmpFile, 0
		.if (m_hWnd)
			.if (g_bWriteClipBoard)
				mov szStr, 0
				invoke GetTempPath, MAX_PATH, addr szStr
				invoke GetTempFileName, addr szStr, CStr("~CV"), NULL, addr szTmpFile
				invoke CreateFile, addr szTmpFile, GENERIC_WRITE, 0,
					NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL
			.else
				invoke GetSourceFileName, m_hWnd, addr szName, CStr(".inc"), CStr("Includes")
				.if (eax == 0)
					ret
				.endif
			.endif
		.else
			.if (m_pszOut)
				invoke lstrlen, m_pszOut
				mov ecx, CStr("%s\%s.inc")
				mov edx, m_pszOut
				.if (byte ptr [edx+eax-1] == '\')
					mov ecx, CStr("%s%s.inc")
				.endif
				invoke wsprintf, addr szStr, ecx, m_pszOut, addr szName
			.else
				invoke wsprintf, addr szStr, CStr("%s.inc"), addr szName
			.endif
			invoke CreateFile, addr szStr, GENERIC_WRITE, 0,
				NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL
		.endif
		.if (eax == INVALID_HANDLE_VALUE)
			invoke MessageBeep, MB_OK
			invoke MessageBox, m_hWnd, addr szStr, CStr("File creation error"), MB_OK
			xor eax, eax
			ret
		.endif
		mov m_hFile,eax

;-------------------------------------------- write header (typelib infos)
		invoke WriteLine, CStr("")
;-------------------------------------------- abort if we cannot write
		.if (eax == 0)
			invoke MessageBox, m_hWnd, CStr("couldn't write to file"), 0, MB_OK
			xor eax,eax
			ret
		.endif
		invoke WriteLine, CStr(";--- Generated .INC file by COMView")

		invoke wsprintf,addr szStr,CStr(";--- typelib name=%s"),addr szName
		invoke WriteLine, addr szStr
		invoke wsprintf,addr szStr,CStr(";--- documentation=%s"),addr szDoc
		invoke WriteLine, addr szStr
		.if (m_pszTypeLib)
			invoke wsprintf,addr szStr,CStr(";--- DLL=%s"),m_pszTypeLib
			invoke WriteLine, addr szStr
		.endif

		invoke vf(pTypeLib,ITypeLib,GetTypeInfoCount)
		mov dwCount,eax

		invoke vf(pTypeLib,ITypeLib,GetLibAttr),addr pTLibAttr
		.if (eax == S_OK)
			mov edx,pTLibAttr
			invoke StringFromGUID2, addr [edx].TLIBATTR.guid, addr wszGUID,40
;;			invoke WideCharToMultiByte,CP_ACP,0, addr wszGUID, -1, addr szGUID, 40, 0, 0
			invoke wsprintf,addr szStr,CStr(";--- GUID=%S"),addr wszGUID
			invoke WriteLine, addr szStr
			mov edx,pTLibAttr
			movzx eax,[edx].TLIBATTR.wMajorVerNum
			movzx ecx,[edx].TLIBATTR.wMinorVerNum
			invoke wsprintf,addr szStr,CStr(";--- Version %u.%u, LCID %X"),eax,ecx,[edx].TLIBATTR.lcid
			invoke WriteLine, addr szStr
			invoke vf(pTypeLib,ITypeLib,ReleaseTLibAttr),pTLibAttr
		.endif

		invoke WriteLine, CStr("")
		invoke WriteLine, CStr(";--- prototypes too complex for MASM may exist")
		invoke WriteLine, CStr(";--- if such ones reside in a vtable they will be replaced by dummies")
		invoke WriteLine, CStr(";--- search for occurances of ';+++' to have a closer look at these")
		invoke WriteLine, CStr("")
		.if (m_dwMode == INCMODE_DISPHLP)
			.if (g_bCreateMaxDispHlp)
				invoke WriteLine, CStr(";--- this is a maximized dispatch helper include.")
				invoke WriteLine, CStr(";--- that is, definitions for all dispatch interfaces are included.")
				invoke WriteLine, CStr(";--- Methods described here may be called with macro dm() in any case.")
				invoke WriteLine, CStr(";--- if interface is dual, macro vf() may be used as well.")
			.else
				invoke WriteLine, CStr(";--- here dispatchonly interfaces are described")
				invoke WriteLine, CStr(";--- whose methods/properties can be accessed thru IDispatch::Invoke only")
				invoke WriteLine, CStr(";--- use macro dm() to call a method/property of such an interface.")
			.endif
			invoke WriteLine, CStr(";--- macro dm(), which is described in file objbase.inc,")
			invoke WriteLine, CStr(";--- requires placing macro DEFINE_INVOKEHELPER somewhere in your source code")
			invoke WriteLine, CStr("")
		.else
			mov edx,pTLibAttr
			invoke GetAsmGUID, addr szGUID, addr [edx].TLIBATTR.guid
			invoke wsprintf, addr szStr, CStr("sTLBID_%s textequ ",3Ch,"GUID %s",3Eh), addr szName, addr szGUID
			invoke WriteLine, addr szStr
			mov edx,pTLibAttr
			movzx ecx, [edx].TLIBATTR.wMajorVerNum
			invoke wsprintf, addr szStr, CStr("_MajorVer_%s equ %u"), addr szName, ecx
			invoke WriteLine, addr szStr
			mov edx,pTLibAttr
			movzx ecx, [edx].TLIBATTR.wMinorVerNum
			invoke wsprintf, addr szStr, CStr("_MinorVer_%s equ %u"), addr szName, ecx
			invoke WriteLine, addr szStr
			invoke WriteLine, CStr("")
		.endif

;-------------------------------------------- header of include is written now

		mov eax, dwCount
		mov ecx, sizeof CType
		mul ecx
		invoke malloc, eax
		mov g_pDepArray, eax
;-------------------------------------------- we scan twice now to do a
;-------------------------------------------- bit resorting
		mov g_iPass, 1
NextScan:
;-------------------------------------------- scan all typeinfos in typelib
		mov dwIndex, 0
		.while (1)
			mov eax, dwIndex
			.break .if (eax >= dwCount)
			mov esi, g_pDepArray
			mov eax, sizeof CType
			mul dwIndex
			add esi, eax
			mov g_pCurDep, esi
			.if (g_iPass == 1)
				mov [esi].CType.pszCur, NULL
				mov [esi].CType.pszDep, NULL
				mov eax, dwIndex
				mov [esi].CType.dwIdx, eax
			.endif
			invoke vf(pTypeLib,ITypeLib,GetTypeInfo), [esi].CType.dwIdx ,addr m_pTypeInfo
			.if (eax == S_OK)
				invoke vf(m_pTypeInfo, ITypeInfo, GetTypeAttr),addr pTypeAttr
				.if (eax == S_OK)
					mov esi,pTypeAttr
					assume esi:ptr TYPEATTR

;-------------------------------------------- generate maximum dispatch helpers?
					.if ((m_dwMode == INCMODE_DISPHLP) && g_bCreateMaxDispHlp)
						;
;-------------------------------------------- if its a dual dispatch
;-------------------------------------------- get referenced TKIND_INTERFACE typeinfo
					.elseif (([esi].typekind == TKIND_DISPATCH) &&	([esi].wTypeFlags & TYPEFLAG_FDUAL))
						invoke vf(m_pTypeInfo, ITypeInfo, GetRefTypeOfImplType), -1, addr reftype
						.if (eax == S_OK)
							invoke vf(m_pTypeInfo, ITypeInfo, GetRefTypeInfo), reftype, addr pTypeInfoRef
							.if (eax == S_OK)
								invoke vf(m_pTypeInfo,ITypeInfo,ReleaseTypeAttr),pTypeAttr
								invoke vf(m_pTypeInfo, ITypeInfo, Release)
								mov eax, pTypeInfoRef
								mov m_pTypeInfo, eax
								invoke vf(m_pTypeInfo,ITypeInfo,GetTypeAttr),addr pTypeAttr
								mov esi,pTypeAttr
							.endif
						.endif
					.endif

					mov bstr1,NULL
					invoke vf(m_pTypeInfo, ITypeInfo, GetDocumentation),MEMBERID_NIL,addr bstr1,NULL,NULL,NULL
					.if (eax == S_OK)
						invoke WideCharToMultiByte,CP_ACP,0,bstr1,-1,addr szDoc,sizeof szDoc,0,0 
						invoke SysFreeString, bstr1
					.else
						mov szDoc,0
					.endif
					invoke StringFromGUID2,addr [esi].guid,addr wszGUID,40
					invoke WideCharToMultiByte,CP_ACP,0,addr wszGUID,40,addr szName, sizeof szName,0,0


					invoke GetTypekindStr, [esi].typekind
					mov edx,eax
					invoke wsprintf, addr szStr,
						CStr(";--- %s, Type=%s, GUID=%s",13,10,13,10),
						addr szDoc, edx, addr szName

					invoke CheckReservedNames, addr szDoc

					.if (g_iPass == 1)
						mov ecx, g_pCurDep
						.if ([esi].typekind == TKIND_ENUM)
							mov [ecx].CType.dwPrty,0
						.elseif ([esi].typekind == TKIND_RECORD)
							mov [ecx].CType.dwPrty,10h
						.elseif ([esi].typekind == TKIND_UNION)
							mov [ecx].CType.dwPrty,20h
						.elseif ([esi].typekind == TKIND_ALIAS)
							mov [ecx].CType.dwPrty,30h
						.elseif ([esi].typekind == TKIND_MODULE)
							mov [ecx].CType.dwPrty,40h
						.elseif ([esi].typekind == TKIND_INTERFACE)
							mov [ecx].CType.dwPrty,50h
						.elseif ([esi].typekind == TKIND_DISPATCH)
							mov [ecx].CType.dwPrty,60h
						.else
							mov [ecx].CType.dwPrty,70h
						.endif
					.endif

					mov g_bWrite, TRUE
					invoke ScanTypeInfoVariables, addr szDoc, addr szStr, m_dwMode
					mov g_bWrite, TRUE
					invoke ScanTypeInfoFunctions, addr szDoc, addr szStr, m_dwMode
					mov g_bWrite, TRUE
					.if (m_dwMode == INCMODE_BASIC)
						invoke ScanTypeInfoInterfaces, addr szDoc, addr szStr
					.endif
					mov g_bWrite, TRUE

					invoke vf(m_pTypeInfo,ITypeInfo,ReleaseTypeAttr),pTypeAttr
				.endif
				invoke vf(m_pTypeInfo, ITypeInfo, Release)
			.endif
			inc dwIndex
		.endw
		assume esi:ptr CType

if 0							;for debugging: display array of types
		mov esi, g_pDepArray
		mov ecx, dwCount
		.while (ecx)
			push ecx
			mov eax, [esi].CType.pszCur
			mov ecx, [esi].CType.pszDep
			.if (!eax)
				mov eax, CStr("");
			.endif
			.if (!ecx)
				mov ecx, CStr("");
			.endif
			invoke wsprintf, addr szStr, CStr("%u, %s, %s",13,10),[esi].CType.dwIdx, eax, ecx
			invoke OutputDebugString, addr szStr
			add esi, sizeof CType
			pop ecx
			dec ecx
		.endw
endif

		.if (g_iPass == 1)
			push edi
			push ebx
			mov esi, g_pDepArray
			mov edi, dwCount
			.while (edi)
				mov ebx, esi
				mov ecx, edi
				.while (ecx > 1)
					add ebx, sizeof CType
					mov fSwap, FALSE
					mov eax, [esi].CType.dwPrty
					.if (eax > [ebx].CType.dwPrty)
						mov fSwap, TRUE
					.elseif ([esi].CType.pszDep && [ebx].CType.pszCur)
						push ecx
						invoke lstrcmp, [esi].CType.pszDep, [ebx].CType.pszCur
						pop ecx
						.if (eax == 0)
							mov fSwap, TRUE
						.endif
					.endif
					.if (fSwap)
						push edi
						mov ecx, [esi].CType.dwIdx
						mov edx, [esi].CType.pszCur
						mov eax, [esi].CType.pszDep
						mov edi, [esi].CType.dwPrty
						xchg ecx, [ebx].CType.dwIdx
						xchg edx, [ebx].CType.pszCur
						xchg eax, [ebx].CType.pszDep
						xchg edi, [ebx].CType.dwPrty
						mov [esi].CType.dwIdx, ecx
						mov [esi].CType.pszCur, edx
						mov [esi].CType.pszDep, eax
						mov [esi].CType.dwPrty, edi
						pop edi
						mov ebx, esi
						mov ecx, edi
						.continue
					.endif
					dec ecx
				.endw	
				invoke free, [esi].CType.pszCur	;no longer needed
				invoke free, [esi].CType.pszDep	;no longer needed
				add esi, sizeof CType
				dec edi
			.endw
			pop ebx
			pop edi
			inc g_iPass
			jmp NextScan
		.endif

		invoke free, g_pDepArray

		invoke WriteLine, CStr(";--- end of type library ---")

		invoke CloseHandle, m_hFile
		.if (szTmpFile)
			invoke CopyFileToClipboard, m_hWnd, addr szTmpFile
			invoke DeleteFile, addr szTmpFile
			invoke printf, CStr("ASM declarations written to the clipboard",13,10)
		.endif

		mov eax,1
		ret
		align 4
		assume esi:nothing

ScanTypeLib endp

	option prologue: prologuedef
	option epilogue: epiloguedef
ifdef @StackBase
	option stackbase:esp
endif

;--- constructor 1

if 0
Create@CCreateInclude proc public uses __this pszGUID:ptr byte, lcid:LCID, dwMajor:DWORD, dwMinor:DWORD

	invoke malloc, sizeof CCreateInclude
	.if (!eax)
		ret
	.endif
	mov __this, eax

;------------------------- use these parms onyl for loadtypelib!

	mov eax, pszGUID
	mov m_pszGUID, eax
	mov eax, lcid
	mov m_lcid,eax
	mov eax, dwMajor
	mov m_dwMajor, eax
	mov eax, dwMinor
	mov m_dwMinor, eax
	return __this
	align 4

Create@CCreateInclude endp
endif

;--- constructor 2

Create2@CCreateInclude proc public uses __this pszTypeLib:LPSTR

	invoke malloc, sizeof CCreateInclude
	.if (!eax)
		ret
	.endif
	mov __this, eax
	mov eax, pszTypeLib
	mov m_pszTypeLib, eax
	return __this
	align 4

Create2@CCreateInclude endp

Create3@CCreateInclude proc public uses __this pTypeLib:LPTYPELIB

	invoke malloc, sizeof CCreateInclude
	.if (!eax)
		ret
	.endif
	mov __this, eax
	mov eax, pTypeLib
	mov m_pTypeLib, eax
	invoke vf(m_pTypeLib, IUnknown, AddRef)
	return __this
	align 4

Create3@CCreateInclude endp

;--- run the creation pass

Run@CCreateInclude proc public uses __this this_: ptr CCreateInclude, hWnd:HWND, dwMode:DWORD
	mov __this, this_
	mov eax, hWnd
	mov m_hWnd, eax
	mov eax, dwMode
	mov m_dwMode, eax
	invoke ScanTypeLib
	ret
Run@CCreateInclude endp

SetOutputFile@CCreateInclude proc public uses __this this_: ptr CCreateInclude, pszOut:LPSTR
	mov __this, this_
	invoke lstrlen, pszOut
	inc eax
	invoke malloc, eax
	xchg eax, m_pszOut
	invoke free, eax
	invoke lstrcpy, m_pszOut, pszOut
	ret
	align 4
SetOutputFile@CCreateInclude endp

;--- destructor

Destroy@CCreateInclude proc public uses __this this_: ptr CCreateInclude

	mov __this, this_
	.if (m_pszOut)
		invoke free, m_pszOut
	.endif
	.if (m_pTypeLib)
		invoke vf(m_pTypeLib, IUnknown, Release)
	.endif
	invoke free, __this
	ret
	align 4
Destroy@CCreateInclude endp


	end

