

;*** definition of class CContainer
;*** CContainer implements a simple OLE container, so being able
;*** to host an OLE control

	.386
	.model flat,stdcall
	option casemap :none   ; case sensitive
	option proc:private

	include COMView.inc
	.nolist
	.nocref
INSIDE_CCONTAINER equ 1
	include classes.inc
	include rsrc.inc
	include debugout.inc
	include servprov.inc
	.list
	.cref

?NEWMETHOD		equ 1	;select routine for size calculation
?WINDOWLESS		equ 1	;support IOleInPlaceSiteWindowless
?DISPATCH		equ 1	;support IDispatch
?DOCUMENT		equ 1	;support IOleDocumentSite (act as document site)
?COMMANDTARGET	equ 1	;support IOleCommandTarget
?OLECONTAINER	equ 1	;support IOleContainer
?OLELINK		equ 0	;support IOleLink
?CALLFACTORY	equ 0	;support ICallFactory
?SERVICEPROVIDER equ 1	;support IServiceProvider
?DOCHOSTSHOWUI	equ 0	;support IDocHostShowUI (is said to be used by MSHTML)
?USEMYPROPBAG	equ 1	;create CPropertyBag object
?POINTERINACTIVE equ 1	;

if ?DOCUMENT
	include DocObj.inc
endif
if ?DOCHOSTSHOWUI
	include MsHtmHst.inc
endif

if 0
protostrlenW typedef proto :ptr WORD
externdef _imp__lstrlenW@4:PTR protostrlenW
lstrlenW equ <_imp__lstrlenW@4>
endif

@MakeStub macro name, ofs, suffix
ifb <suffix>
name&_:
else
name&suffix:
endif
	sub dword ptr [esp+4], ofs
	jmp name
	endm


BEGIN_CLASS CContainer
OleClientSite			IOleClientSite <>
OleInPlaceSite			IOleInPlaceSite <>
OleInPlaceFrame			IOleInPlaceFrame <>
OleControlSite			IOleControlSite <>
if ?DISPATCH
Dispatch				IDispatch <>
endif
if ?DOCUMENT
OleDocumentSite			IOleDocumentSite <>
endif
if ?COMMANDTARGET
OleCommandTarget		IOleCommandTarget <>
endif
if ?OLECONTAINER
OleContainer			IOleContainer <>
endif
if ?CALLFACTORY
CallFactory				ICallFactory <>
endif
if ?SERVICEPROVIDER
ServiceProvider			IServiceProvider <>
endif
if ?DOCHOSTSHOWUI
DocHostShowUI			IDocHostShowUI <>
endif
dwRefCount				dd ?
pViewObjectDlg			pCViewObjectDlg ?
hWndSite				HWND ?
pOleObject				LPOLEOBJECT ?
pObjectItem				LPOBJECTITEM ?
pObjectWithSite			LPOBJECTWITHSITE ?
pOleInPlaceObject		LPOLEINPLACEOBJECT ?
pOleInPlaceActiveObject	LPOLEINPLACEACTIVEOBJECT ?
if ?WINDOWLESS
pOleInPlaceObjectWindowless	LPOLEINPLACEOBJECTWINDOWLESS ?
endif
if ?DOCUMENT
pOleDocumentView		LPOLEDOCUMENTVIEW ?
endif
if ?OLELINK
pOleLink				LPOLELINK ?
endif
if ?POINTERINACTIVE
pPointerInactive		LPPOINTERINACTIVE ?
endif
pTypeInfo				LPTYPEINFO ?
pAdviseSink				LPADVISESINK ?
pMonikerCon				LPMONIKER ?
pMonikerRel				LPMONIKER ?
pMonikerFull			LPMONIKER ?
dwConnection			DWORD ?
dwDataConnection		DWORD ?
rect					RECT <>		;comment missing!
rectBorderSpace			RECT <>
hMenuView				HMENU ?
bUIActivated			BOOLEAN ?
bShownInWindow			BOOLEAN ?
bClientSiteSet			BOOLEAN ?
bViewConnected			BOOLEAN ?
bLoadPropertyBag		BOOLEAN ?	;initialized with IPersistPropertyBag::Load
END_CLASS

__this	textequ <ebx>
_this	textequ <[__this].CContainer>

	MEMBER OleClientSite, OleInPlaceSite, OleInPlaceFrame, OleControlSite
if ?DISPATCH
	MEMBER Dispatch
endif
	MEMBER dwRefCount, pViewObjectDlg, hWndSite, rect, pObjectItem
	MEMBER pMonikerCon, pMonikerRel, pMonikerFull
	MEMBER pOleObject, pObjectWithSite, pOleInPlaceObject, pOleInPlaceActiveObject
if ?WINDOWLESS
	MEMBER pOleInPlaceObjectWindowless
endif
if ?DOCUMENT
	MEMBER OleDocumentSite, pOleDocumentView
endif
if ?COMMANDTARGET
	MEMBER OleCommandTarget
endif
if ?OLELINK
	MEMBER pOleLink
endif
if ?OLECONTAINER
	MEMBER OleContainer
endif
if ?CALLFACTORY
	MEMBER CallFactory
endif
if ?SERVICEPROVIDER
	MEMBER ServiceProvider
endif
if ?DOCHOSTSHOWUI
	MEMBER DocHostShowUI
endif
if ?POINTERINACTIVE
	MEMBER pPointerInactive
endif
	MEMBER pTypeInfo, pAdviseSink, dwConnection, bUIActivated, bShownInWindow
	MEMBER bClientSiteSet, bViewConnected, hMenuView, rectBorderSpace
	MEMBER dwDataConnection, bLoadPropertyBag

;--- private methods

Destroy@CContainer proto :ptr CContainer
ActivateObject proto

	.data

externdef IID_IViewObject:IID 
externdef IID_IViewObject2:IID 
externdef IID_IOleObject:IID 

g_pStream	LPSTREAM NULL
g_pStorage	LPSTORAGE NULL
g_dqNull	LARGE_INTEGER <<0,0>>
g_bUIDead	BOOLEAN FALSE
;;g_bLoadFile	BOOLEAN FALSE
g_bScribbleMode	BOOLEAN TRUE

externdef g_dwBorder:DWORD
;;public g_bLoadFile

;DISPID_AMBIENT_USERMODE	equ -709
;DISPID_AMBIENT_UIDEAD		equ -710
;DISPID_AMBIENT_LOCALEID	equ -705

	.const

;*** vtbl interface IOleClientSite + IUnknown

COleClientSiteVtbl label IOleClientSiteVtbl
	IUnknownVtbl {QueryInterface, AddRef, Release}
	dd SaveObject
	dd GetMoniker
	dd GetContainer
	dd ShowObject
	dd OnShowWindow
	dd RequestNewObjectLayout

;*** vtbl of interface IOleWindow, IOleInPlaceSite, IOleInPlaceSiteEx, IOleInPlaceSiteWindowless

COleInPlaceSiteVtbl label IOleInPlaceSiteVtbl
	IUnknownVtbl {QueryInterface1, AddRef1, Release1}
;--- IOleWindow
	dd offset GetWindow__
	dd offset ContextSensitiveHelp_
;--- IOleInPlaceSite
	dd offset CanInPlaceActivate
	dd offset OnInPlaceActivate_
	dd offset OnUIActivate_
	dd offset GetWindowContext_
	dd offset Scroll_
	dd offset OnUIDeactivate_
	dd offset OnInPlaceDeactivate_
	dd offset DiscardUndoState
	dd offset DeactivateAndUndo_
	dd offset OnPosRectChange_
;--- IOleInPlaceSiteEx
	dd offset OnInPlaceActivateEx_
	dd offset OnInPlaceDeactivateEx_
	dd offset RequestUIActivate
if ?WINDOWLESS
;--- IOleInPlaceSiteWindowless
	dd offset CanWindowlessActivate
	dd offset GetCapture__
	dd offset SetCapture__
	dd offset GetFocus__
	dd offset SetFocus__
	dd offset GetDC__
	dd offset ReleaseDC__
	dd offset InvalidateRect__
	dd offset InvalidateRgn__
	dd offset ScrollRect
	dd offset AdjustRect
	dd offset OnDefWindowMessage_
endif

;*** vtbl of interface IOleInPlaceFrame + IOleInPlaceUIWindow

COleInPlaceFrameVtbl label IOleInPlaceFrameVtbl
	IUnknownVtbl {QueryInterface2, AddRef2, Release2}
	dd offset GetWindow_2
	dd offset ContextSensitiveHelp2
	dd offset GetBorder_
	dd offset RequestBorderSpace_
	dd offset SetBorderSpace_
	dd offset SetActiveObject_
	dd offset InsertMenus_
	dd offset SetMenu__
	dd offset RemoveMenus_
	dd offset SetStatusText_
	dd offset EnableModeless
	dd offset TranslateAccelerator_

;*** vtbl of interface IOleControlSite

COleControlSiteVtbl label IOleControlSiteVtbl
	IUnknownVtbl {QueryInterface4, AddRef4, Release4}
	dd offset OnControlInfoChanged
	dd offset LockInPlaceActive
	dd offset GetExtendedControl
	dd offset TransformCoords
	dd offset TranslateAccelerator__
	dd offset OnFocus
	dd offset ShowPropertyFrame

;*** vtbl of interface IDispatch

if ?DISPATCH
CDispatchVtbl label IDispatchVtbl
	IUnknownVtbl {QueryInterface3, AddRef3, Release3}
	dd offset GetTypeInfoCount
	dd offset GetTypeInfo
	dd offset GetIDsOfNames
	dd offset Invoke_
endif

if ?DOCUMENT
COleDocumentSiteVtbl label IOleDocumentSiteVtbl
	IUnknownVtbl {QueryInterface5, AddRef5, Release5}
	dd offset ActivateMe_
endif

if ?COMMANDTARGET
COleCommandTargetVtbl label dword
	IUnknownVtbl {QueryInterface6, AddRef6, Release6}
	dd QueryStatus_, Exec_
endif

if ?OLECONTAINER
COleContainerVtbl label dword
	IUnknownVtbl {QueryInterface7, AddRef7, Release7}
	dd ParseDisplayName_, EnumObjects__, LockContainer_
endif

if ?CALLFACTORY
CCallFactoryVtbl label dword
	IUnknownVtbl {QueryInterface8, AddRef8, Release8}
	dd CreateCall_
endif

if ?SERVICEPROVIDER
CServiceProviderVtbl label dword
	IUnknownVtbl {QueryInterface9, AddRef9, Release9}
	dd QueryService_
endif

;--- seems a MASM bug: lines have to be commented out in debug version,
;--- although ?DOCHOSTSHOWUI is zero

;if ?DOCHOSTSHOWUI
;CDocHostShowUIVtbl label dword
;	IUnknownVtbl {QueryInterface10, AddRef10, Release10}
;	dd ShowMessage_, ShowHelp_
;endif

	public g_szContainer

g_szContainer			db "Container",0
g_szIOleClientSite		db "IOleClientSite",0
g_szIOleInPlaceSite		db "IOleInPlaceSite",0
g_szIOleInPlaceSiteEx	db "IOleInPlaceSiteEx",0
g_szIOleInPlaceSiteWindowless	db "IOleInPlaceSiteWindowless",0
g_szIOleInPlaceFrame	db "IOleInPlaceFrame",0
g_szIOleControlSite		db "IOleControlSite",0

	align 4

wszExcl	dw '!',0

;szFmtIdBag	GUID {20001801h, 5De6h, 11d1h, {8eh, 38h, 00h, 0c0h, 4fh, 0b9h, 38h, 6dh}}

	.data

;*** table of supported interfaces (in .data since it will be changed

iftab label dword
	dd IID_IUnknown				, CContainer.OleClientSite
	dd IID_IOleClientSite		, CContainer.OleClientSite
	dd IID_IOleWindow			, CContainer.OleInPlaceSite
	dd IID_IOleInPlaceSite		, CContainer.OleInPlaceSite
	dd IID_IOleInPlaceUIWindow	, CContainer.OleInPlaceFrame
	dd IID_IOleInPlaceFrame		, CContainer.OleInPlaceFrame
	dd IID_IOleControlSite		, CContainer.OleControlSite
if ?OLECONTAINER
	dd IID_IOleContainer		, CContainer.OleContainer
endif
if ?CALLFACTORY
	dd IID_ICallFactory			, CContainer.CallFactory
endif
if ?DOCHOSTSHOWUI
	dd IID_IDocHostShowUI		, CContainer.DocHostShowUI
endif
NUMIFENTRIES equ ($ - offset iftab) / (4 * 2)

freeentries label dword

if ?DISPATCH
	dd 0, 0
	.const
DispatchEntry label dword
	dd IID_IDispatch, CContainer.Dispatch
	.data
endif

	dd 0, 0
	.const
OleInPlaceSiteExEntry label dword
	dd offset IID_IOleInPlaceSiteEx, CContainer.OleInPlaceSite
	.data

if ?WINDOWLESS
	dd 0, 0
	.const
OleInPlaceSiteWindowlessEntry label dword
	dd offset IID_IOleInPlaceSiteWindowless, CContainer.OleInPlaceSite
	.data
endif
	
if ?DOCUMENT
	dd 0, 0
	.const
OleDocumentSiteEntry label dword
	dd offset IID_IOleDocumentSite, CContainer.OleDocumentSite
	.data
endif

if ?COMMANDTARGET
	dd 0, 0
	.const
OleCommandTargetEntry label dword
	dd offset IID_IOleCommandTarget, CContainer.OleCommandTarget
	.data
endif

if ?SERVICEPROVIDER
	dd 0, 0
	.const
ServiceProviderEntry label dword
	dd offset IID_IServiceProvider, CContainer.ServiceProvider
	.data
endif

	.code

DisplayHResult proc uses __this this_:ptr CContainer, pszText:LPSTR, HResult:DWORD
	
local szText[128]:byte

	mov __this, this_
	invoke wsprintf, addr szText, pszText, HResult
	invoke SetStatusText@CViewObjectDlg, m_pViewObjectDlg, 0, addr szText
	ret

DisplayHResult endp

SetStatus proc 
	.if (m_bUIActivated)
		mov eax, CStr("UIActivated")
	.elseif (m_pOleInPlaceObject)
		mov eax, CStr("Activated")
	.elseif (m_bShownInWindow)
		mov eax, CStr("Shown in window")
	.else
		mov eax, CStr("Loaded")
	.endif
	invoke SetStatusText@CViewObjectDlg, m_pViewObjectDlg, 1, eax
if ?WINDOWLESS
	.if (m_pOleInPlaceObjectWindowless)
		mov eax, CStr("Windowless")
	.elseif (m_pOleInPlaceObject)
		mov eax, CStr("Windowed")
	.else
		mov eax, CStr("")
	.endif
	invoke SetStatusText@CViewObjectDlg, m_pViewObjectDlg, 2, eax
endif
	ret
SetStatus endp

;--------------------------------------------------------------
;--- interface IUnknown
;--------------------------------------------------------------

AddRef proto :ptr CContainer


QueryInterface proc uses esi edi __this this_:ptr CContainer, riid:ptr IID, ppReturn:ptr ptr

local	wszIID[40]:word
local	szKey[128]:byte
local	dwSize:DWORD
local	hKey:HANDLE

	mov __this,this_

	mov edx, NUMIFENTRIES
	mov edi, offset freeentries
if ?DISPATCH
	.if (g_bDispatchSupp)
		mov esi, offset DispatchEntry
		movsd
		movsd
		inc edx
	.endif
endif
	.if (g_bInPlaceSiteExSupp)
		mov esi, offset OleInPlaceSiteExEntry
		movsd
		movsd
		inc edx
	.endif
if ?WINDOWLESS
	.if (g_bAllowWindowless)
		mov esi, offset OleInPlaceSiteWindowlessEntry
		movsd
		movsd
		inc edx
	.endif
endif
if ?DOCUMENT
	.if (g_bDocumentSiteSupp)
		mov esi, offset OleDocumentSiteEntry
		movsd
		movsd
		inc edx
	.endif
endif
if ?COMMANDTARGET
	.if (g_bCommandTargetSupp)
		mov esi, offset OleCommandTargetEntry
		movsd
		movsd
		inc edx
	.endif
endif
if ?SERVICEPROVIDER
	.if (g_bServiceProviderSupp)
		mov esi, offset ServiceProviderEntry
		movsd
		movsd
		inc edx
	.endif
endif

	invoke IsInterfaceSupported, riid, offset iftab, edx,  this_, ppReturn

	.if (g_bLogActive && g_bDispQueryIFCalls)
;--------------------- print the name of the interface we have just been queried
		push eax
		invoke StringFromGUID2,riid, addr wszIID,40
		invoke wsprintf, addr szKey, CStr("%s\%S"),addr g_szInterface, addr wszIID
		invoke RegOpenKeyEx, HKEY_CLASSES_ROOT, addr szKey, 0, KEY_READ, addr hKey 
		.if (eax == ERROR_SUCCESS)
			mov dwSize, sizeof szKey
			invoke RegQueryValueEx,hKey,addr g_szNull,NULL,NULL,addr szKey,addr dwSize
			invoke RegCloseKey, hKey
		.else
			mov szKey,0
		.endif
		pop eax
		push eax
;;		DebugOut "%s_IUnknown::QueryInterface(%S[%s])=%X", addr g_szContainer, addr wszIID, addr szKey, eax
		invoke printf@CLogWindow, CStr("%s_IUnknown::QueryInterface(%S[%s])=%X",10),
					addr g_szContainer, addr wszIID, addr szKey, eax
		pop eax
	.endif
	ret

QueryInterface endp


AddRef proc uses __this this_:ptr CContainer

	mov __this,this_
	inc m_dwRefCount
ifdef _DEBUG
	.if (g_bDispQueryIFCalls)
		DebugOut "CContainer::AddRef = %u", m_dwRefCount
	.endif
endif
	mov eax, m_dwRefCount
	ret

AddRef endp

Release proc uses __this this_:ptr CContainer

	mov __this,this_
	dec m_dwRefCount
ifdef _DEBUG
	.if (g_bDispQueryIFCalls)
		DebugOut "CContainer::Release = %u", m_dwRefCount
	.endif
endif
	mov eax, m_dwRefCount
	.if (eax == 0)
		invoke Destroy@CContainer, __this
		xor eax,eax
	.endif
	ret

Release endp


;--------------------------------------------------------------
;--- interface IOleClientSite
;--------------------------------------------------------------


SaveObject proc uses __this esi this_:ptr CContainer

local	hr:DWORD
local	pStorage:LPSTORAGE
local	pPersistStreamInit:LPPERSISTSTREAMINIT
local	pPersistPropertyBag:LPPERSISTPROPERTYBAG

	DebugOut "IOleClientSite::SaveObject"
	mov __this, this_
	mov hr, S_OK

	.repeat
		.if (g_bUseIPersistPropBag && m_bLoadPropertyBag)
			invoke vf(m_pOleObject, IUnknown, QueryInterface), addr IID_IPersistPropertyBag, addr pPersistPropertyBag
			.if (eax == S_OK)
				invoke vf(pPersistPropertyBag, IUnknown, Release)
				mov esi, SAVE_PROPBAG
				.break
			.endif
		.endif
		.if (g_bUseIPersistStream)
			invoke vf(m_pOleObject, IUnknown, QueryInterface), addr IID_IPersistStreamInit, addr pPersistStreamInit
			.if (eax == S_OK)
				invoke vf(pPersistStreamInit, IUnknown, Release)
				mov esi, SAVE_STREAM
				.break
			.endif
		.endif
		mov esi, SAVE_STORAGE
	.until (1)

	mov eax, IDYES
	.if (g_bConfirmSaveReq)
		invoke MessageBox, m_hWndSite, CStr("Object requests to be saved. Save now?"), CStr("IOleClientSite::SaveObject"), MB_YESNO
	.endif
	.if (eax == IDYES)
		invoke Save@CContainer, __this, esi, FALSE
		.if (eax != S_OK)
			mov hr, E_FAIL
		.endif
	.endif
	.if (g_bDispContainerCalls)
		invoke printf@CLogWindow, CStr("%s_%s::SaveObject called, returns %X",10),
			addr g_szContainer, addr g_szIOleClientSite, hr
	.endif
	return hr

SaveObject endp

GetMoniker proc uses __this esi this_:ptr CContainer, dwAssign:DWORD, dwWhichMoniker:DWORD, ppmk:ptr LPMONIKER

local clsid:CLSID
local wszCLSID[40]:word

	DebugOut "IOleClientSite::GetMoniker(%X, %X, %X)", dwAssign, dwWhichMoniker, ppmk
	mov __this, this_
	mov esi, dwWhichMoniker
	.if  ((esi == OLEWHICHMK_CONTAINER) || (esi == OLEWHICHMK_OBJFULL))
		.if ((!m_pMonikerCon) && (dwAssign == OLEGETMONIKER_FORCEASSIGN))
			invoke CreateFileMoniker, CStrW(L("COMView.exe")), addr m_pMonikerCon
		.endif
		.if (esi == OLEWHICHMK_CONTAINER)
			mov eax, m_pMonikerCon
			jmp done
		.endif
	.endif
	.if  ((esi == OLEWHICHMK_OBJREL) || (esi == OLEWHICHMK_OBJFULL))
		.if ((!m_pMonikerRel) && (dwAssign == OLEGETMONIKER_FORCEASSIGN))
			invoke GetGUID@CObjectItem, m_pObjectItem, addr clsid
			invoke StringFromGUID2, addr clsid, addr wszCLSID, 40
			invoke CreateItemMoniker, addr wszExcl, addr wszCLSID, addr m_pMonikerRel
		.endif
		.if (esi == OLEWHICHMK_OBJREL)
			mov eax, m_pMonikerRel
			jmp done
		.endif
	.endif
	.if ((!m_pMonikerFull) && (dwAssign == OLEGETMONIKER_FORCEASSIGN))
		invoke CreateGenericComposite, m_pMonikerCon, m_pMonikerRel, addr m_pMonikerFull
	.endif
	mov eax, m_pMonikerFull
done:
	mov ecx, ppmk
	mov [ecx], eax
	.if (eax)
		invoke vf(eax, IUnknown, AddRef)
		mov eax, S_OK
	.else
		mov eax, E_FAIL
	.endif
	ret

GetMoniker endp

GetContainer proc this_:ptr CContainer, ppContainer:ptr LPOLECONTAINER

	DebugOut "IOleClientSite::GetContainer(%X)", ppContainer
	mov eax,ppContainer
if ?OLECONTAINER
	mov edx, this_
	lea ecx, [edx].CContainer.OleContainer
	mov [eax], ecx
	inc [edx].CContainer.dwRefCount
	return S_OK
else
	mov dword ptr [eax], NULL
	return E_NOINTERFACE
endif

GetContainer endp

ShowObject proc this_:ptr CContainer

	DebugOut "IOleClientSite::ShowObject"
	return S_OK

ShowObject endp

;--- if true, object is shown in separate window

OnShowWindow proc uses __this this_:ptr CContainer, fShow:dword

	mov __this, this_
	mov eax, fShow
	mov m_bShownInWindow, al
;;	DebugOut "IOleClientSite::OnShowWindow(%X)", fShow
	.if (g_bDispContainerCalls)
		invoke printf@CLogWindow, CStr("%s_%s::OnShowWindow(%u) called",10),
			addr g_szContainer, addr g_szIOleClientSite,
			fShow
	.endif
	invoke SetStatus
	return S_OK

OnShowWindow endp

RequestNewObjectLayout proc this_:ptr CContainer

	DebugOut "IOleClientSite::RequestNewObjectLayout"
	return E_NOTIMPL

RequestNewObjectLayout endp

;--------------------------------------------------------------

;--- CContainer uses multiple inheritance, so implementing
;--- more than 1 vtable. methods called from vtables which are not
;--- located at offset 0 in the object need "this" adjustment
;--- thats done here by "sub [esp+4],X" or otherwise "sub this_, X"

;--------------------------------------------------------------
;--- interface IOleInPlaceSite
;--------------------------------------------------------------

;--- IUnknown
	@MakeStub QueryInterface,	CContainer.OleInPlaceSite, 1
	@MakeStub AddRef,			CContainer.OleInPlaceSite, 1
	@MakeStub Release,			CContainer.OleInPlaceSite, 1
;--- IOleWindow
	@MakeStub GetWindow_,			CContainer.OleInPlaceSite
	@MakeStub ContextSensitiveHelp,	CContainer.OleInPlaceSite
;--- IOleInPlaceSite
	@MakeStub OnUIActivate,			CContainer.OleInPlaceSite
	@MakeStub GetWindowContext,		CContainer.OleInPlaceSite
	@MakeStub OnUIDeactivate,		CContainer.OleInPlaceSite
	@MakeStub OnInPlaceActivate,	CContainer.OleInPlaceSite
	@MakeStub OnInPlaceDeactivate,	CContainer.OleInPlaceSite
	@MakeStub DeactivateAndUndo,	CContainer.OleInPlaceSite
	@MakeStub OnPosRectChange,		CContainer.OleInPlaceSite
;--- IOleInPlaceSiteEx
	@MakeStub OnInPlaceActivateEx,	CContainer.OleInPlaceSite
	@MakeStub OnInPlaceDeactivateEx,CContainer.OleInPlaceSite
;--- IOleInPlaceSiteWindowless
	@MakeStub GetCapture_,			CContainer.OleInPlaceSite
	@MakeStub SetCapture_,			CContainer.OleInPlaceSite
	@MakeStub GetFocus_,			CContainer.OleInPlaceSite
	@MakeStub SetFocus_,			CContainer.OleInPlaceSite
	@MakeStub GetDC_,				CContainer.OleInPlaceSite
	@MakeStub ReleaseDC_,			CContainer.OleInPlaceSite
	@MakeStub InvalidateRect_,		CContainer.OleInPlaceSite
	@MakeStub InvalidateRgn_,		CContainer.OleInPlaceSite
	@MakeStub OnDefWindowMessage,	CContainer.OleInPlaceSite


GetWindow_ proc uses __this this_:ptr CContainer, phwnd:ptr HWND

	DebugOut "IOleInPlaceSite::GetWindow"
	mov __this, this_
	mov eax, phwnd
	mov ecx, m_hWndSite
	mov [eax], ecx

	return S_OK

GetWindow_ endp


ContextSensitiveHelp proc this_:ptr CContainer, fEnterMode:BYTE
	DebugOut "IOleInPlaceSite::ContextSensitiveHelp"
	return E_NOTIMPL
ContextSensitiveHelp endp


CanInPlaceActivate proc this_:ptr CContainer
	DebugOut "IOleInPlaceSite::CanInPlaceActivate"
	return S_OK
CanInPlaceActivate endp


OnInPlaceActivate proc uses __this this_:ptr CContainer

	mov __this, this_

;;	DebugOut "IOleInPlaceSite::OnInPlaceActivate"

	.if (g_bDispContainerCalls)
		invoke printf@CLogWindow, CStr("%s_%s::OnInPlaceActivate",10),
			addr g_szContainer, addr g_szIOleInPlaceSite
	.endif

	.if (m_pOleObject)
		invoke vf(m_pOleObject, IUnknown, QueryInterface), addr IID_IOleInPlaceObject, addr m_pOleInPlaceObject
	.endif
	invoke SetStatus
	invoke SetMenu@CViewObjectDlg, m_pViewObjectDlg, IDM_INPLACEDEACTIVATE, MF_ENABLED
	return S_OK

OnInPlaceActivate endp

OnUIActivate proc uses __this this_:ptr CContainer
	mov __this,this_
;;	DebugOut "IOleInPlaceSite::OnUIActivate"
	.if (g_bDispContainerCalls)
		invoke printf@CLogWindow, CStr("%s_%s::OnUIActivate",10),
			addr g_szContainer, addr g_szIOleInPlaceSite
	.endif
	mov m_bUIActivated, TRUE
	invoke SetMenu@CViewObjectDlg, m_pViewObjectDlg, IDM_UIDEACTIVATE, MF_ENABLED
	invoke SetStatus
	return S_OK

OnUIActivate endp

;*** IOleInPlaceSite::GetWindowContext

GetWindowContext proc uses __this this_:ptr CContainer, ppFrame:ptr LPOLEINPLACEFRAME,
								ppDoc:ptr LPOLEINPLACEUIWINDOW, lprcPosRect:ptr RECT,
								lprcClipRect:ptr RECT, lpFrameInfo:ptr OLEINPLACEFRAMEINFO

local	rect:RECT

	DebugOut "IOleInPlaceSite::GetWindowContext"

	mov __this,this_

	invoke CopyRect, lprcClipRect, addr m_rect
	invoke CopyRect, lprcPosRect, addr m_rect

	mov ecx,ppFrame
	lea eax, m_OleInPlaceFrame
	mov dword ptr [ecx],eax
	invoke vf(eax, IUnknown, AddRef)

	mov eax,ppDoc
	mov dword ptr [eax],NULL

	mov eax,lpFrameInfo
	mov [eax].OLEINPLACEFRAMEINFO.fMDIApp,FALSE
	mov ecx, m_hWndSite
	mov [eax].OLEINPLACEFRAMEINFO.hwndFrame,ecx
	mov [eax].OLEINPLACEFRAMEINFO.haccel,NULL
	mov [eax].OLEINPLACEFRAMEINFO.cAccelEntries,0
	return S_OK

GetWindowContext endp

Scroll_ proc this_:ptr CContainer, scrollExtent:SIZEL
	DebugOut "IOleInPlaceSite::Scroll"
	return S_FALSE
	align 4

Scroll_ endp

OnUIDeactivate proc uses __this this_:ptr CContainer, fUndoable:DWORD

	mov __this,this_
;;	DebugOut "IOleInPlaceSite::OnUIDeactivate(%X)", fUndoable
	.if (g_bDispContainerCalls)
		invoke printf@CLogWindow, CStr("%s_%s::OnUIDeactivate",10),
			addr g_szContainer, addr g_szIOleInPlaceSite
	.endif
	mov m_bUIActivated, FALSE
	.if (m_hMenuView)
		invoke SetMenu, m_hWndSite, m_hMenuView
	.endif
	invoke SetStatus
	invoke SetMenu@CViewObjectDlg, m_pViewObjectDlg, IDM_UIDEACTIVATE, MF_GRAYED
	return S_OK
	align 4

OnUIDeactivate endp

OnInPlaceDeactivate proc uses __this this_:ptr CContainer

;;	DebugOut "IOleInPlaceSite::OnInPlaceDeactivate"

	.if (g_bDispContainerCalls)
		invoke printf@CLogWindow, CStr("%s_%s::OnInPlaceDeactivate",10),
			addr g_szContainer, addr g_szIOleInPlaceSite
	.endif

	mov __this,this_

	.if (m_pOleInPlaceActiveObject)
		invoke vf(m_pOleInPlaceActiveObject, IUnknown, Release)
		mov m_pOleInPlaceActiveObject, NULL
	.endif
if ?WINDOWLESS
	.if (m_pOleInPlaceObjectWindowless)
		invoke vf(m_pOleInPlaceObjectWindowless, IUnknown, Release)
		mov m_pOleInPlaceObjectWindowless, NULL
	.endif
endif
	.if (m_pOleInPlaceObject)
		invoke vf(m_pOleInPlaceObject, IUnknown, Release)
		mov m_pOleInPlaceObject, NULL
	.endif
	invoke SetStatus
	invoke SetMenu@CViewObjectDlg, m_pViewObjectDlg, IDM_INPLACEDEACTIVATE, MF_GRAYED

	return S_OK
	align 4

OnInPlaceDeactivate endp

DiscardUndoState proc this_:ptr CContainer
	DebugOut "IOleInPlaceSite::DiscardUndoState"
	return S_OK
DiscardUndoState endp

DeactivateAndUndo proc uses __this this_:ptr CContainer
	DebugOut "IOleInPlaceSite::DeactivateAndUndo"
	mov __this, this_
	.if (m_bUIActivated)
		invoke vf(m_pOleInPlaceObject, IOleInPlaceObject, UIDeactivate)
	.endif
	return S_OK
DeactivateAndUndo endp

OnPosRectChange proc uses __this this_:ptr CContainer, lprcPosRect:ptr RECT

local	rect:RECT

ifdef _DEBUG
	mov ecx, lprcPosRect
	DebugOut "IOleInPlaceSite::OnPosRectChange(%X,%X,%X,%X)",[ecx].RECT.left,\
		[ecx].RECT.top, [ecx].RECT.right, [ecx].RECT.bottom
endif
	mov __this,this_

	.if (m_pOleInPlaceObject)
		invoke CopyRect, addr m_rect, lprcPosRect
		invoke vf(m_pOleInPlaceObject, IOleInPlaceObject, SetObjectRects), lprcPosRect, lprcPosRect
	.endif

	return S_OK

OnPosRectChange endp

;-------------------------------------------------
;IOleInPlaceSiteEx methods
;-------------------------------------------------

OnInPlaceActivateEx proc uses __this this_:ptr CContainer, pfNoRedraw:ptr BOOL, dwFlags:DWORD

	mov __this,this_
;;	DebugOut "IOleInPlaceSiteEx::OnInPlaceActivateEx"
	.if (g_bDispContainerCalls)
		invoke printf@CLogWindow, CStr("%s_%s::OnInPlaceActivateEx(%u,%X)",10),
			addr g_szContainer, addr g_szIOleInPlaceSiteEx,\
			pfNoRedraw, dwFlags
	.endif

	.if (dwFlags & ACTIVATE_WINDOWLESS)
		invoke vf(m_pOleObject, IUnknown, QueryInterface), addr IID_IOleInPlaceObjectWindowless,\
			addr m_pOleInPlaceObjectWindowless
	.endif
	mov eax,pfNoRedraw
	mov dword ptr [eax],FALSE
	invoke OnInPlaceActivate, this_
	ret

OnInPlaceActivateEx endp

OnInPlaceDeactivateEx proc uses ebx this_:ptr CContainer, fNoRedraw:BOOL

;;	DebugOut "IOleInPlaceSiteEx::OnInPlaceDeactivateEx"

	.if (g_bDispContainerCalls)
		invoke printf@CLogWindow, CStr("%s_%s::OnInPlaceDeactivateEx(%u) called",10),
			addr g_szContainer, addr g_szIOleInPlaceSiteEx,\
			fNoRedraw
	.endif
	invoke OnInPlaceDeactivate, this_
	ret

OnInPlaceDeactivateEx endp

RequestUIActivate proc this_:ptr CContainer

	DebugOut "IOleInPlaceSiteEx::RequestUIActivate"
	return S_OK
RequestUIActivate endp


if ?WINDOWLESS
;-------------------------------------------------
;IOleInPlaceSiteWindowless methods
;-------------------------------------------------

CanWindowlessActivate proc this_:ptr CContainer
;;	DebugOut "IOleInPlaceSiteWindowless::CanWindowlessActivate"
	.if (g_bAllowWindowless)
		mov eax, S_OK
	.else
		mov eax, S_FALSE
	.endif
	.if (g_bDispContainerCalls)
		invoke printf@CLogWindow, CStr("%s_%s::CanWindowlessActivate will return %u",10),
			addr g_szContainer, addr g_szIOleInPlaceSiteWindowless, eax
	.endif
	ret
CanWindowlessActivate endp

GetCapture_ proc uses __this this_:ptr CContainer
	mov __this, this_
	DebugOut "IOleInPlaceSiteWindowless::GetCapture"
	invoke GetCapture
	.if (eax == m_hWndSite)
		mov eax, S_OK
	.else
		mov eax, S_FALSE
	.endif
	ret
GetCapture_ endp

SetCapture_ proc uses __this this_:ptr CContainer, fCapture:BOOL
	mov __this, this_
	DebugOut "IOleInPlaceSiteWindowless::SetCapture"
	.if (fCapture)
		invoke SetCapture, m_hWndSite
	.else
		invoke GetCapture
		.if (eax == m_hWndSite)
			invoke ReleaseCapture
		.endif
	.endif
	return S_OK
SetCapture_ endp

GetFocus_ proc uses __this this_:ptr CContainer
	mov __this, this_
	DebugOut "IOleInPlaceSiteWindowless::GetFocus"
	invoke GetFocus
	.if (eax == m_hWndSite)
		mov eax, S_OK
	.else
		mov eax, S_FALSE
	.endif
	ret
GetFocus_ endp

SetFocus_ proc uses __this this_:ptr CContainer, fFocus:BOOL
	mov __this, this_
	DebugOut "IOleInPlaceSiteWindowless::SetFocus"
	.if (fFocus)
		invoke SetFocus, m_hWndSite
	.else
		invoke GetFocus
		.if (eax == m_hWndSite)
			invoke SetFocus, NULL
		.endif
	.endif
	return S_OK

SetFocus_ endp

GetDC_ proc uses __this this_:ptr CContainer, pRect:ptr RECT, grfFlags:DWORD, phDC:ptr HDC
	mov __this, this_
	DebugOut "IOleInPlaceSiteWindowless::GetDC"
	invoke GetDC, m_hWndSite
	mov ecx, phDC
	mov [ecx],eax
	return S_OK
GetDC_ endp

ReleaseDC_ proc uses __this this_:ptr CContainer, hDC:HDC
	mov __this, this_
	DebugOut "IOleInPlaceSiteWindowless::ReleaseDC"
	invoke ReleaseDC, m_hWndSite, hDC
	return S_OK
ReleaseDC_ endp

InvalidateRect_ proc uses __this this_:ptr CContainer, pRect:ptr RECT, fErase:BOOL
	mov __this, this_
	DebugOut "IOleInPlaceSiteWindowless::InvalidateRect"
	invoke InvalidateRect, m_hWndSite, pRect, fErase
	return S_OK
InvalidateRect_ endp

InvalidateRgn_ proc uses __this this_:ptr CContainer, hRGN:HRGN, fErase:BOOL
	mov __this, this_
	DebugOut "IOleInPlaceSiteWindowless::InvalidateRgn"
	invoke InvalidateRgn, m_hWndSite, hRGN, fErase
	return S_OK
InvalidateRgn_ endp

ScrollRect proc this_:ptr CContainer, dx_:DWORD, dy:DWORD, pRectScroll:ptr RECT, pRectClip:ptr RECT
	DebugOut "IOleInPlaceSiteWindowless::ScrollRect"
	return S_FALSE
ScrollRect endp

AdjustRect proc this_:ptr CContainer, prc:ptr RECT
	DebugOut "IOleInPlaceSiteWindowless::AdjustRect"
	return S_FALSE
AdjustRect endp

OnDefWindowMessage proc uses __this this_:ptr CContainer, msg:DWORD, wParam:WPARAM, lParam:LPARAM, plResult:ptr DWORD
	mov __this, this_
	DebugOut "IOleInPlaceSiteWindowless::OnDefWindowMessage"
	invoke DefWindowProc, m_hWndSite, msg, wParam, lParam
	mov ecx, plResult
	mov [ecx], eax
	return S_OK
OnDefWindowMessage endp

endif

;--------------------------------------------------------------
;--- interface IOleInPlaceFrame (includes IOleInPlaceUIWindow)
;--------------------------------------------------------------

E_NOTOOLSPACE equ 800401A1h

	@MakeStub QueryInterface,	CContainer.OleInPlaceFrame, 2
	@MakeStub AddRef,			CContainer.OleInPlaceFrame, 2
	@MakeStub Release,			CContainer.OleInPlaceFrame, 2
	@MakeStub GetWindow_,		CContainer.OleInPlaceFrame, 2
	@MakeStub ContextSensitiveHelp,	CContainer.OleInPlaceFrame, 2
	@MakeStub GetBorder,		CContainer.OleInPlaceFrame
	@MakeStub RequestBorderSpace,CContainer.OleInPlaceFrame
	@MakeStub SetBorderSpace,	CContainer.OleInPlaceFrame
	@MakeStub SetActiveObject,	CContainer.OleInPlaceFrame
	@MakeStub InsertMenus,		CContainer.OleInPlaceFrame
	@MakeStub SetMenu_,			CContainer.OleInPlaceFrame
	@MakeStub RemoveMenus,		CContainer.OleInPlaceFrame
	@MakeStub SetStatusText,	CContainer.OleInPlaceFrame


GetBorder proc uses __this this_:ptr CContainer, lprectBorder:ptr RECT

	DebugOut "IOleInPlaceUIWindow::GetBorder(%X)", lprectBorder
	mov __this,this_
	invoke CopyRect, lprectBorder, addr m_rect
	return S_OK

GetBorder endp

;--- BORDERWIDTHS is same as RECT

RequestBorderSpace proc uses __this this_:ptr CContainer, pborderwidths:ptr BORDERWIDTHS

	DebugOut "IOleInPlaceUIWindow::RequestBorderSpace(%X)", pborderwidths
	mov __this,this_
	return S_OK 

RequestBorderSpace endp


SetBorderSpace proc uses __this this_:ptr CContainer, pborderwidths:ptr BORDERWIDTHS

local rect:RECT

	mov __this,this_
	mov eax,pborderwidths
	.if (eax)
		DebugOut "IOleInPlaceUIWindow::SetBorderSpace(%X,%X,%X,%X)",\
			[eax].BORDERWIDTHS.left, [eax].BORDERWIDTHS.top,\
			[eax].BORDERWIDTHS.right, [eax].BORDERWIDTHS.bottom
		invoke CopyRect, addr m_rectBorderSpace, eax
	.endif
	return S_OK

SetBorderSpace endp


SetActiveObject proc uses __this this_:ptr CContainer, pActiveObject:LPOLEINPLACEACTIVEOBJECT, pszObjName:ptr WORD

	DebugOut "IOleInPlaceUIWindow::SetActiveObject(%X)", pActiveObject
	mov __this,this_
	.if (m_pOleInPlaceActiveObject)
		invoke vf(m_pOleInPlaceActiveObject, IOleInPlaceActiveObject, Release)
	.endif
	mov eax, pActiveObject
	mov m_pOleInPlaceActiveObject, eax
	.if (m_pOleInPlaceActiveObject)
		invoke vf(m_pOleInPlaceActiveObject, IOleInPlaceActiveObject, AddRef)
	.endif
	return S_OK

SetActiveObject endp


InsertMenus proc uses esi __this this_:ptr CContainer, hmenuShared:HMENU, lpMenuWidths:ptr OLEMENUGROUPWIDTHS

	DebugOut "IOleInPlaceFrame::InsertMenus"
	mov __this,this_
	.if (!m_hMenuView)
		invoke GetMenu, m_hWndSite
		mov m_hMenuView, eax
	.endif
	mov ecx, lpMenuWidths
	mov [ecx].OLEMENUGROUPWIDTHS.width_[0*sizeof DWORD], 1
	mov [ecx].OLEMENUGROUPWIDTHS.width_[2*sizeof DWORD], 1
	mov [ecx].OLEMENUGROUPWIDTHS.width_[4*sizeof DWORD], 1
	invoke GetSubMenu, m_hMenuView, 0
	mov ecx, eax
	invoke InsertMenu, hmenuShared, 0, MF_BYPOSITION or MF_POPUP, ecx, CStr("&File")
	invoke GetSubMenu, m_hMenuView, 1
	mov ecx, eax
	invoke InsertMenu, hmenuShared, 1, MF_BYPOSITION or MF_POPUP, ecx, CStr("&Actions")
	invoke GetSubMenu, m_hMenuView, 2
	mov ecx, eax
	invoke InsertMenu, hmenuShared, 2, MF_BYPOSITION or MF_POPUP, ecx, CStr("&Options")
	return S_OK

InsertMenus endp

SetMenu_ proc uses __this this_:ptr CContainer, hmenuShared:HMENU, holemenu:HANDLE, hwndActiveObject:HWND

	DebugOut "IOleInPlaceFrame::SetMenu(%X, %X, %X)", hmenuShared, holemenu, hwndActiveObject
	mov __this,this_
	.if (hmenuShared)
		invoke SetMenu, m_hWndSite, hmenuShared
		invoke OleSetMenuDescriptor, holemenu, m_hWndSite, hwndActiveObject,
			addr m_OleInPlaceFrame, m_pOleInPlaceActiveObject
	.endif
	return S_OK

SetMenu_ endp

RemoveMenus proc uses __this this_:ptr CContainer, hmenuShared:HMENU

	DebugOut "IOleInPlaceFrame::RemoveMenus(%X)", hmenuShared
	mov __this,this_
	invoke GetSubMenu, m_hMenuView, 0
	invoke RemoveMenu, hmenuShared, eax, MF_BYCOMMAND
	invoke GetSubMenu, m_hMenuView, 1
	invoke RemoveMenu, hmenuShared, eax, MF_BYCOMMAND
	invoke GetSubMenu, m_hMenuView, 2
	invoke RemoveMenu, hmenuShared, eax, MF_BYCOMMAND
	invoke OleSetMenuDescriptor, NULL, m_hWndSite, NULL,
		addr m_OleInPlaceFrame, m_pOleInPlaceActiveObject
	return S_OK

RemoveMenus endp

SetStatusText proc uses __this this_:ptr CContainer, pszStatusText:ptr WORD

local dwSize:DWORD

	mov __this,this_
	.if (pszStatusText)
		invoke lstrlenW, pszStatusText
		add eax, 4
		and al, 0FCh
		sub esp, eax
		mov dwSize, eax
		mov edx, esp
		invoke WideCharToMultiByte, CP_ACP, 0, pszStatusText, -1, edx, dwSize, NULL, NULL
		.if (g_bDispContainerCalls)
			invoke printf@CLogWindow, CStr("%s_%s::SetStatusText",28h,22h),
				addr g_szContainer, addr g_szIOleInPlaceFrame
			invoke printf@CLogWindow, CStr("%s"), esp
			invoke printf@CLogWindow, CStr(22h,29h,10)
		.endif
		invoke SetStatusText@CViewObjectDlg, m_pViewObjectDlg, 0, esp
		add esp, dwSize
	.endif

	return S_OK

SetStatusText endp

EnableModeless proc this_:ptr CContainer, fEnable:DWORD

	DebugOut "IOleInPlaceFrame::EnableModeless(%u)", fEnable
	return S_OK

EnableModeless endp

TranslateAccelerator_ proc this_:ptr CContainer, lpmsg:ptr MSG, wID:WORD

	DebugOut "IOleInPlaceFrame::TranslateAccelerator"
	return S_FALSE

TranslateAccelerator_ endp

;-------------------------------------------------
;--- IOleControlSite interface
;-------------------------------------------------

	@MakeStub QueryInterface,	CContainer.OleControlSite, 4
	@MakeStub AddRef,			CContainer.OleControlSite, 4
	@MakeStub Release,			CContainer.OleControlSite, 4

;	@MakeStub OnControlInfoChanged,	CContainer.OleControlSite
;	@MakeStub LockInPlaceActive,	CContainer.OleControlSite
;	@MakeStub GetExtendedControl,	CContainer.OleControlSite
;	@MakeStub TransformCoords,		CContainer.OleControlSite
;	@MakeStub TranslateAccelerator,	CContainer.OleControlSite
;	@MakeStub OnFocus,				CContainer.OleControlSite
;	@MakeStub ShowPropertyFrame,	CContainer.OleControlSite

;*** IOleControlSite::OnControlInfoChanged

OnControlInfoChanged proc uses __this this_:ptr CContainer

;;	DebugOut "IOleInPlaceFrame::OnControlInfoChanged"
	.if (g_bDispContainerCalls)
		invoke printf@CLogWindow, CStr("%s_%s::OnControlInfoChanged called",10),
			addr g_szContainer, addr g_szIOleControlSite
	.endif
	return S_OK

OnControlInfoChanged endp


;*** IOleControlSite::LockInPlaceActive


LockInPlaceActive proc this_:ptr CContainer, fLock:BOOL

;;	DebugOut "IOleInPlaceFrame::LockInPlaceActive(%u)", fLock
	.if (g_bDispContainerCalls)
		invoke printf@CLogWindow, CStr("%s_%s::LockInPlaceActive(%u) called",10),
			addr g_szContainer, addr g_szIOleControlSite, fLock
	.endif
	return E_NOTIMPL

LockInPlaceActive endp


;*** IOleControlSite::GetExtendedControl


GetExtendedControl proc this_:ptr CContainer, ppDisp:ptr LPDISPATCH

	DebugOut "IOleControlSite::GetExtendedControl"
	mov eax,ppDisp
	mov dword ptr [eax],NULL
	return E_NOTIMPL

GetExtendedControl endp


;*** IOleControlSite::TransformCoords


TransformCoords proc this_:ptr CContainer, pPtlHimetric:ptr POINTL, pPtfContainer:ptr POINTF, dwFlags:DWORD

	DebugOut "IOleControlSite::TransformCoords"
	return E_NOTIMPL

TransformCoords endp


;*** IOleControlSite::TranslateAccelerator


TranslateAccelerator__ proc this_:ptr CContainer, pMsg:ptr MSG, grfModifiers:DWORD

	DebugOut "IOleControlSite::TranslateAccelerator"
	return S_FALSE

TranslateAccelerator__ endp


;*** IOleControlSite::OnFocus


OnFocus proc this_:ptr CContainer, fGotFocus:BOOL

	DebugOut "IOleControlSite::OnFocus(%u)", fGotFocus
	return S_OK

OnFocus endp


;*** IOleControlSite::ShowPropertyFrame
;*** return NOT_IMPL to let control show properties itself


ShowPropertyFrame proc this_:ptr CContainer

	DebugOut "IOleControlSite::ShowPropertyFrame"
	return E_NOTIMPL

ShowPropertyFrame endp


if ?DISPATCH
;--------------------------------------------------------------
;--- interface IDispatch
;--------------------------------------------------------------

	@MakeStub QueryInterface,	CContainer.Dispatch, 3
	@MakeStub AddRef,			CContainer.Dispatch, 3
	@MakeStub Release,			CContainer.Dispatch, 3

GetTypeInfoCount proc this_:ptr CContainer, pctinfo:ptr DWORD
	DebugOut "IDispatch::GetTypeInfoCount"
	mov eax, pctinfo
	mov dword ptr [eax],0
	return S_OK
GetTypeInfoCount endp

GetTypeInfo proc this_:ptr CContainer, iTInfo:DWORD, lcid:LCID, ppTInfo:ptr LPTYPEINFO
	DebugOut "IDispatch::GetTypeInfo"
	mov eax, ppTInfo
	mov dword ptr [eax], NULL
	return E_NOTIMPL
GetTypeInfo endp

GetIDsOfNames proc this_:ptr CContainer, riid:ptr IID, rgszNames:DWORD, cNames:DWORD, lcid:LCID, rgDispId:ptr DISPID
	DebugOut "IDispatch::GetIDsOfNames"
	return DISP_E_UNKNOWNINTERFACE
GetIDsOfNames endp


Invoke_:
	sub	dword ptr [esp+4], CContainer.Dispatch

Invoke$ proc uses __this esi this_:ptr CContainer, dispIdMember:DISPID, riid:ptr IID,
			lcid:LCID, wFlags:DWORD, pDispParams:ptr DISPPARAMS,
			pVarResult:ptr VARIANT, pExcepInfo:DWORD, puArgErr:ptr DWORD

local	dwNumNames:dword
local	valRet:DWORD
local	bstr:BSTR
local	pwszName:ptr WORD
local	szText[256]:byte
local	szName[128]:byte
local	szType[32]:byte
local	szParams[128]:byte
local	szIID[40]:byte
local	wszIID[40]:word


	mov __this,this_

	mov esi, S_OK
	@mov valRet, 0
	.if (wFlags & DISPATCH_PROPERTYGET)
		invoke GetAmbientProp, dispIdMember, pVarResult
;--------------------------- this function returns value in edx
		mov esi, eax
		.if (eax == S_OK)
			.if (dispIdMember == DISPID_AMBIENT_DISPLAYNAME && (edx == NULL))
				invoke VariantInit, pVarResult
				invoke vf(m_pOleObject, IOleObject, GetUserType), USERCLASSTYPE_SHORT, addr pwszName
				.if (eax == S_OK)
					invoke SysAllocString, pwszName
					push eax
					invoke CoTaskMemFree, pwszName
					pop eax
				.else
					invoke SysAllocString, CStrW(L("[undefined]"))
				.endif
				mov edx, eax
				mov ecx,pVarResult
				mov [ecx].VARIANT.vt, VT_BSTR
				mov [ecx].VARIANT.bstrVal, edx
			.endif
		.endif
		.if (esi == S_OK)
			mov valRet, edx
		.endif
	.else
		mov esi, DISP_E_MEMBERNOTFOUND
	.endif


	.if (g_bLogActive && g_bDispContainerCalls)

		mov szName,0
		.if (dispIdMember >= 0)
			.if (m_pTypeInfo != 0)
				invoke vf(m_pTypeInfo, ITypeInfo, GetNames), dispIdMember, addr bstr, 1, addr dwNumNames
				.if ((eax == S_OK) && (dwNumNames > 0))
					invoke WideCharToMultiByte,CP_ACP,0,bstr,-1,addr szName,sizeof szName,0,0 
					invoke SysFreeString, bstr
				.endif
			.endif
		.else
			invoke GetStdDispIdStr, dispIdMember
			.if (eax)
				push eax
				invoke lstrcpy, addr szName, CStr("DISPID_")
				pop	eax
				invoke lstrcat, addr szName, eax
			.endif
		.endif
		.if (wFlags & DISPATCH_PROPERTYGET)
			invoke lstrcpy,addr szType,CStr("PropertyGet")
		.elseif (wFlags & DISPATCH_PROPERTYPUT)
			invoke lstrcpy,addr szType,CStr("PropertyPut")
		.elseif (wFlags & DISPATCH_METHOD)
			invoke lstrcpy,addr szType,CStr("Method")
		.elseif (wFlags & DISPATCH_PROPERTYPUTREF)
			invoke lstrcpy,addr szType,CStr("PropertyPutRef")
		.else
			mov szType,0
		.endif

		mov szIID,0
		.if (riid)
			push edi
			mov edi,riid
			mov ecx,4
			xor eax,eax
			repz scasd
			pop edi
			.if (!ZERO?)
				invoke StringFromGUID2, riid, addr wszIID, 40
				invoke WideCharToMultiByte, CP_ACP, 0, addr wszIID, -1, addr szIID, sizeof szIID,0,0 
			.endif
		.endif

		mov szParams,0
		.if (pDispParams)
			mov word ptr szParams,'('
			mov eax,pDispParams
			mov ecx,[eax].DISPPARAMS.cArgs
			mov edx,[eax].DISPPARAMS.rgvarg
			.while (ecx)
				push ecx
				push edx
				invoke GetArgument, edx, addr szParams
				pop edx
				pop ecx
				add edx,sizeof VARIANT
				dec ecx
				.if (ecx)
					pushad
					invoke lstrcat, addr szParams, CStr(",")
					popad
				.endif
			.endw
			invoke lstrcat, addr szParams, CStr(29h)
		.endif

		invoke printf@CLogWindow, CStr("%s_IDispatch::Invoke %s, %s, ID=%d, %s%s, valRet=%X, HResult=%X",10),
			addr g_szContainer, addr szIID, addr szType, dispIdMember, addr szName, addr szParams, valRet, esi

	.endif

	return esi
    align 4
Invoke$ endp

endif

if ?DOCUMENT
;--------------------------------------------------------------
;--- interface IOleDocumentSite
;--------------------------------------------------------------

;--- IUnknown
	@MakeStub QueryInterface,	CContainer.OleDocumentSite, 5
	@MakeStub AddRef,			CContainer.OleDocumentSite, 5
	@MakeStub Release,			CContainer.OleDocumentSite, 5
;--- IOleDocumentSite
	@MakeStub ActivateMe,		CContainer.OleDocumentSite

ActivateMe proc uses __this this_:ptr CContainer, pViewToActivate:LPOLEDOCUMENTVIEW

local pOleDocument:LPOLEDOCUMENT

	mov __this, this_
	DebugOut "IOleDocumentSite::ActivateMe(%X)", pViewToActivate
;------------------------------------ release old view
	.if (m_pOleDocumentView)
		invoke vf(m_pOleDocumentView, IUnknown, Release)
		mov m_pOleDocumentView, NULL
	.endif
	mov eax, pViewToActivate
	.if (eax)
		mov m_pOleDocumentView, eax
if 1
		invoke vf(m_pOleDocumentView, IOleDocumentView, SetInPlaceSite), addr m_OleInPlaceSite
endif
		invoke vf(m_pOleDocumentView, IUnknown, AddRef)
	.else
		invoke vf(m_pOleObject, IUnknown, QueryInterface), addr IID_IOleDocument, addr pOleDocument
		.if (eax == S_OK)
			invoke vf(pOleDocument, IOleDocument, CreateView),
				addr m_OleInPlaceSite, NULL, NULL, addr m_pOleDocumentView
			invoke vf(pOleDocument, IUnknown, Release)
		.endif
	.endif
	.if (m_pOleDocumentView)
		invoke vf(m_pOleDocumentView, IOleDocumentView, UIActivate), TRUE
		.if (!m_bShownInWindow)
			invoke vf(m_pOleDocumentView, IOleDocumentView, SetRect), addr m_rect
		.endif
		invoke vf(m_pOleDocumentView, IOleDocumentView, Show), TRUE
	.endif
	return S_OK
ActivateMe endp

endif

if ?COMMANDTARGET
;--------------------------------------------------------------
;--- interface IOleCommandTarget
;--------------------------------------------------------------

;--- IUnknown
	@MakeStub QueryInterface,	CContainer.OleCommandTarget, 6
	@MakeStub AddRef,			CContainer.OleCommandTarget, 6
	@MakeStub Release,			CContainer.OleCommandTarget, 6
;--- IOleCommandTarget
	@MakeStub QueryStatus,		CContainer.OleCommandTarget
	@MakeStub Exec,				CContainer.OleCommandTarget

QueryStatus proc uses __this this_:ptr CContainer, pguidCmdGroup:REFGUID, cCmds:DWORD, prgCmds:ptr OLECMD, pCmdText:ptr OLECMDTEXT

local szGUID[40]:byte	
local wszGUID[40]:word	

	.if (g_bDispContainerCalls)
		.if (pguidCmdGroup)
			invoke StringFromGUID2, pguidCmdGroup, addr wszGUID, 40
			invoke WideCharToMultiByte, CP_ACP, 0, addr wszGUID, -1, addr szGUID, 40, NULL, NULL
		.else
			invoke lstrcpy, addr szGUID, CStr("NULL")
		.endif
		mov ecx, prgCmds
		invoke printf@CLogWindow, CStr("%s_IOleCommandTarget::QueryStatus(%s, %u, [%X,%X])",10),
			addr g_szContainer, addr szGUID, cCmds, [ecx].OLECMD.cmdID, [ecx].OLECMD.cmdf
	.endif
	return OLECMDERR_E_UNKNOWNGROUP

QueryStatus endp

Exec proc uses __this this_:ptr CContainer, pguidCmdGroup:REFGUID, nCmdID:DWORD, nCmdExecOpt:DWORD, pvaIn:ptr VARIANT, pvaOut:ptr VARIANT

local szGUID[40]:byte	
local wszGUID[40]:word	

	.if (g_bDispContainerCalls)
		.if (pguidCmdGroup)
			invoke StringFromGUID2, pguidCmdGroup, addr wszGUID, 40
			invoke WideCharToMultiByte, CP_ACP, 0, addr wszGUID, -1, addr szGUID, 40, NULL, NULL
		.else
			invoke lstrcpy, addr szGUID, CStr("NULL")
		.endif
		invoke printf@CLogWindow, CStr("%s_IOleCommandTarget::Exec(%s, %X, %X)",10),
			addr g_szContainer, addr szGUID, nCmdID, nCmdExecOpt
	.endif
	return OLECMDERR_E_DISABLED

Exec endp

endif

if ?OLECONTAINER
;--------------------------------------------------------------
;--- interface IOleContainer
;--------------------------------------------------------------

;--- IUnknown
	@MakeStub QueryInterface,	CContainer.OleContainer, 7
	@MakeStub AddRef,			CContainer.OleContainer, 7
	@MakeStub Release,			CContainer.OleContainer, 7
;--- IOleContainer
	@MakeStub ParseDisplayName,	CContainer.OleContainer
	@MakeStub EnumObjects_,		CContainer.OleContainer
	@MakeStub LockContainer,	CContainer.OleContainer

ParseDisplayName proc uses __this this_:ptr CContainer,
		pbc:ptr IBindCtx, pszDisplayName:LPOLESTR, pchEaten:ptr DWORD, ppmkOut:ptr LPMONIKER

	DebugOut "IOleContainer::ParseDisplayName(%X, %X, %X, %X)", pbc, pszDisplayName, pchEaten, ppmkOut

	mov ecx, pchEaten
	mov dword ptr [ecx], 0
	mov ecx, ppmkOut
	mov dword ptr [ecx], 0
	return MK_E_NOOBJECT

ParseDisplayName endp

EnumObjects_ proc uses __this this_:ptr CContainer, grfFlags:DWORD, ppenum:ptr ptr IEnumUnknown

	DebugOut "IOleContainer::EnumObjects(%X, %X)", grfFlags, ppenum
	mov ecx, ppenum
	mov dword ptr [ecx], 0
	return E_NOTIMPL

EnumObjects_ endp

LockContainer proc uses __this this_:ptr CContainer, fLock:BOOL

	DebugOut "IOleContainer::LockContainer(%u)", fLock
	return S_OK

LockContainer endp

endif

if ?CALLFACTORY

;--- IUnknown
	@MakeStub QueryInterface,	CContainer.CallFactory, 8
	@MakeStub AddRef,			CContainer.CallFactory, 8
	@MakeStub Release,			CContainer.CallFactory, 8
;--- IOleContainer
	@MakeStub CreateCall,		CContainer.CallFactory

CreateCall proc uses __this this_:ptr CContainer, riid1:REFIID, pUnk:LPUNKNOWN, riid2:REFIID, ppObj:ptr LPUNKNOWN

local	wszIID1[40]:word
local	wszIID2[40]:word
local	szKey1[128]:byte
local	szKey2[128]:byte
local	dwSize:DWORD
local	hKey:HANDLE

	mov ecx, ppObj
	mov dword ptr [ecx], NULL
	mov eax, E_NOINTERFACE
	.if (g_bLogActive && g_bDispQueryIFCalls)
;--------------------- print the name of the interface we have just been queried
	    push eax
		invoke StringFromGUID2, riid1, addr wszIID1,40
		invoke wsprintf, addr szKey1, CStr("%s\%S"),addr g_szInterface, addr wszIID1
		invoke RegOpenKeyEx, HKEY_CLASSES_ROOT, addr szKey1, 0, KEY_READ, addr hKey
		.if (eax == ERROR_SUCCESS)
			mov dwSize, sizeof szKey1
			invoke RegQueryValueEx,hKey,addr g_szNull,NULL,NULL,addr szKey1,addr dwSize
			invoke RegCloseKey, hKey
		.else
			mov szKey1,0
		.endif
		invoke StringFromGUID2, riid2, addr wszIID2,40
		invoke wsprintf, addr szKey2, CStr("%s\%S"),addr g_szInterface, addr wszIID2
		invoke RegOpenKeyEx, HKEY_CLASSES_ROOT, addr szKey2, 0, KEY_READ, addr hKey
		.if (eax == ERROR_SUCCESS)
			mov dwSize, sizeof szKey2
			invoke RegQueryValueEx,hKey,addr g_szNull,NULL,NULL,addr szKey2,addr dwSize
			invoke RegCloseKey, hKey
		.else
			mov szKey2,0
		.endif
		pop eax
		push eax
;;		DebugOut "%s_ICallFactory::CreateCall(%S[%s], %S[%s])=%X", addr g_szContainer, addr wszIID1, addr szKey1, addr wszIID2, addr szKey2, eax
		invoke printf@CLogWindow, CStr("%s_ICallFactory::CreateCall(%S[%s], %S[%s])=%X",10),
					addr g_szContainer, addr wszIID1, addr szKey1, addr wszIID2, addr szKey2, eax
		pop eax
	.endif
	ret

CreateCall endp

endif

if ?SERVICEPROVIDER
;--- IUnknown
	@MakeStub QueryInterface,	CContainer.ServiceProvider, 9
	@MakeStub AddRef,			CContainer.ServiceProvider, 9
	@MakeStub Release,			CContainer.ServiceProvider, 9
;--- IServiceProvider
	@MakeStub QueryService,		CContainer.ServiceProvider

QueryService proc uses __this this_:ptr CContainer, guidService:REFGUID, riid:REFIID, ppv:ptr LPVOID

local hr:DWORD
local hKey:HANDLE
local dwSize:DWORD
local wszGUID[40]:word	
local wszIID[40]:word
local szKey[128]:byte
local szKey2[128]:byte

	mov hr, E_NOINTERFACE
	.if (g_bLogActive && g_bDispContainerCalls)

		invoke StringFromGUID2, guidService, addr wszGUID, LENGTHOF wszGUID
		invoke wsprintf, addr szKey, CStr("%s\%S"), addr g_szInterface, addr wszGUID
		invoke RegOpenKeyEx, HKEY_CLASSES_ROOT, addr szKey, 0, KEY_READ, addr hKey
		.if (eax == ERROR_SUCCESS)
			mov dwSize, sizeof szKey
			invoke RegQueryValueEx,hKey,addr g_szNull,NULL,NULL,addr szKey,addr dwSize
			invoke RegCloseKey, hKey
		.else
			mov szKey,0
		.endif

		invoke StringFromGUID2, riid, addr wszIID, LENGTHOF wszIID
		invoke wsprintf, addr szKey2, CStr("%s\%S"), addr g_szInterface, addr wszIID
		invoke RegOpenKeyEx, HKEY_CLASSES_ROOT, addr szKey2, 0, KEY_READ, addr hKey
		.if (eax == ERROR_SUCCESS)
			mov dwSize, sizeof szKey
			invoke RegQueryValueEx,hKey,addr g_szNull,NULL,NULL,addr szKey2,addr dwSize
			invoke RegCloseKey, hKey
		.else
			mov szKey2,0
		.endif
		invoke printf@CLogWindow, CStr("%s_IServiceProvider::QueryService(%S[%s], %S[%s])=%X",10),
					addr g_szContainer, addr wszGUID, addr szKey, addr wszIID, addr szKey2, hr
	.endif
	mov ecx, ppv
	mov dword ptr [ecx], NULL
	return hr

QueryService endp

endif

if ?DOCHOSTSHOWUI
;--- IUnknown
	@MakeStub QueryInterface,	CContainer.DocHostShowUI, 10
	@MakeStub AddRef,			CContainer.DocHostShowUI, 10
	@MakeStub Release,			CContainer.DocHostShowUI, 10
;--- IDocHostShowUI
	@MakeStub ShowMessage,		CContainer.DocHostShowUI
	@MakeStub ShowHelp,			CContainer.DocHostShowUI

ShowMessage proc uses __this this_:ptr CContainer, hwnd:HWND, lpstrText:LPOLESTR, lpstrCaption:LPOLESTR, dwType:DWORD, lpstrHelpFile:LPOLESTR, dwHelpContext:DWORD, plResult:ptr DWORD

	.if (g_bLogActive && g_bDispContainerCalls)
		invoke printf@CLogWindow, CStr("%s_IDocHostShowUI::ShowMessage(%.60S, %.60S)",10),
				addr g_szContainer, lpstrText, lpstrCaption
	.endif
	return S_FALSE

ShowMessage endp

ShowHelp proc uses __this this_:ptr CContainer, hwnd:HWND, pszHelpFile:LPOLESTR, uCommand:DWORD, dwData:DWORD, ptMouse:POINT, pDispatchObjectHit:LPDISPATCH

	.if (g_bLogActive && g_bDispContainerCalls)
		invoke printf@CLogWindow, CStr("%s_IDocHostShowUI::ShowHelp(%.60S, %X, %X)",10),
				addr g_szContainer, pszHelpFile, uCommand, dwData
	.endif
	return S_FALSE

ShowHelp endp

endif

;-----------------------------------------------------------
;--- public methods
;-----------------------------------------------------------


Close@CContainer proc public uses __this this_:ptr CContainer

	DebugOut "Close@CContainer enter"
	mov __this,this_
if 0				;this makes problems with excel at IPersistStorage::Save
	.if (m_pOleInPlaceActiveObject)
		invoke vf(m_pOleInPlaceObject, IOleInPlaceObject, InPlaceDeactivate)
	.endif
else
	.if (m_bUIActivated)
		invoke vf(m_pOleInPlaceObject, IOleInPlaceObject, UIDeactivate)
	.endif
endif
if 0
	.if (m_pOleInPlaceObject)
		invoke vf(m_pOleInPlaceObject, IOleInPlaceObject, Release)
		mov m_pOleInPlaceObject, NULL
	.endif
endif
	.if (m_pOleObject)
		.if (m_bClientSiteSet)
			invoke vf(m_pObjectItem, IObjectItem, Close), OLECLOSE_SAVEIFDIRTY
			invoke vf(m_pOleObject, IOleObject, SetClientSite), NULL
		.endif
if 0
		invoke vf(m_pOleObject, IOleObject, Release)
endif
	.endif
	.if (m_pObjectWithSite)
		invoke vf(m_pObjectWithSite, IObjectWithSite, SetSite), NULL
	.endif
	ret

Close@CContainer endp


OnAmbientPropertyChange@CContainer proc public uses __this this_:ptr CContainer, DispId:DWORD

local pOleControl:LPOLECONTROL

	mov __this,this_
	.if (m_pOleObject)
		invoke vf(m_pOleObject, IUnknown, QueryInterface), addr IID_IOleControl, addr pOleControl
		.if (eax == S_OK)
			invoke vf(pOleControl, IOleControl, OnAmbientPropertyChange), DispId
			invoke vf(pOleControl, IUnknown, Release)
		.endif
	.endif
	ret

OnAmbientPropertyChange@CContainer endp


Close2@CContainer proc public uses __this this_:ptr CContainer

	mov __this,this_
	.if (m_pOleObject)
;;		invoke vf(m_pOleObject, IOleObject, Close), OLECLOSE_NOSAVE
		invoke vf(m_pOleObject, IOleObject, Close), OLECLOSE_PROMPTSAVE
		push eax
		invoke DisplayHResult, __this, CStr("IOleObject::Close returned %X"), eax
		pop eax
	.endif
	ret

Close2@CContainer endp


InPlaceDeactivate@CContainer proc public uses __this this_:ptr CContainer

	mov __this,this_
	.if (m_pOleInPlaceObject)
		invoke vf(m_pOleInPlaceObject, IOleInPlaceObject, InPlaceDeactivate)
		push eax
		invoke DisplayHResult, __this, CStr("IOleInPlaceObject::InPlaceDeactivate returned %X"), eax
		pop eax
	.endif
	ret

InPlaceDeactivate@CContainer endp


UIDeactivate@CContainer proc public uses __this this_:ptr CContainer

	mov __this,this_
if ?DOCUMENT
	.if (m_pOleDocumentView)
		invoke vf(m_pOleDocumentView, IOleDocumentView, Show), FALSE
		invoke vf(m_pOleDocumentView, IOleDocumentView, UIActivate), FALSE
		invoke vf(m_pOleDocumentView, IUnknown, Release)
		mov m_pOleDocumentView, NULL
	.endif
endif
	.if (m_pOleInPlaceObject)
		invoke vf(m_pOleInPlaceObject, IOleInPlaceObject, UIDeactivate)
		push eax
		invoke DisplayHResult, __this, CStr("IOleInPlaceObject::UIDeactivate returned %X"), eax
		pop eax
	.endif
	ret

UIDeactivate@CContainer endp


Update@CContainer proc public uses __this this_:ptr CContainer
	mov __this,this_
	.if (m_pOleObject)
		invoke vf(m_pOleObject, IOleObject, Update)
		invoke DisplayHResult, __this, CStr("IOleObject::Update returned %X"), eax
	.endif
	ret
Update@CContainer endp

;--- iType:
;--- 0 = IOleObject::Advise/Unadvise
;--- 1 = IViewObject::SetAdvise
;--- 2 = IDataObject::DAdvise/DUnadvise

Advise@CContainer proc public uses __this this_:ptr CContainer, iType:DWORD

local hr:DWORD
local pViewObject:LPVIEWOBJECT
local pDataObject:LPDATAOBJECT
local formatetc:FORMATETC

	mov hr, E_FAIL
    mov __this,this_

	.if (iType == 1)
		invoke vf(m_pOleObject, IUnknown, QueryInterface), addr IID_IViewObject, addr pViewObject
		.if (eax != S_OK)
			mov hr, eax
			invoke DisplayHResult, __this, CStr("QueryInterface(IViewObject) returned %X"), hr
			jmp done
		.endif
	.elseif (iType == 2)
		invoke vf(m_pOleObject, IUnknown, QueryInterface), addr IID_IDataObject, addr pDataObject
		.if (eax != S_OK)
			mov hr, eax
			invoke DisplayHResult, __this, CStr("QueryInterface(IDataObject) returned %X"), hr
			jmp done
		.endif
	.endif

	.if ((iType == 0) && m_dwConnection)
		invoke vf(m_pOleObject, IOleObject, Unadvise), m_dwConnection
		mov hr, eax
		.if (eax == S_OK)
			mov m_dwConnection, 0
		.endif
		invoke DisplayHResult, __this, CStr("IOleObject::Unadvise returned %X"), hr
	.elseif ((iType == 1) && m_bViewConnected)
		invoke vf(pViewObject, IViewObject, SetAdvise), DVASPECT_CONTENT, 0, NULL
		mov hr, eax
		mov m_bViewConnected, FALSE
		invoke DisplayHResult, __this, CStr("IViewObject::SetAdvise(0) returned %X"), hr
	.elseif ((iType == 2) && m_dwDataConnection)
		invoke vf(pDataObject, IDataObject, DUnadvise), m_dwDataConnection
		mov hr, eax
		.if (eax == S_OK)
			mov m_dwDataConnection, 0
		.endif
		invoke DisplayHResult, __this, CStr("IDataObject::DUnadvise(0) returned %X"), hr
	.else
		mov eax, m_pAdviseSink
		.if (!eax)
			invoke Create@CAdviseSink
			mov m_pAdviseSink, eax
		.endif
		.if (eax)
			.if (iType ==  0)
				lea ecx, m_dwConnection
				invoke vf(m_pOleObject, IOleObject, Advise), eax, ecx
				mov hr, eax
				invoke DisplayHResult, __this, CStr("IOleObject::Advise returned %X"), hr
			.elseif (iType == 1)
				invoke vf(pViewObject, IViewObject, SetAdvise), DVASPECT_CONTENT, 0, eax
				mov hr, eax
				.if (eax == S_OK)
					mov m_bViewConnected, TRUE
				.endif
				invoke DisplayHResult, __this, CStr("IViewObject::SetAdvise returned %X"), hr
			.else
				mov formatetc.cfFormat, 0
				mov formatetc.ptd, NULL
				mov formatetc.dwAspect, -1
				mov formatetc.lindex, -1
				mov formatetc.tymed, -1
				lea ecx, m_dwDataConnection
				invoke vf(pDataObject, IDataObject, DAdvise), addr formatetc, ADVF_NODATA, eax, ecx
				mov hr, eax
				invoke DisplayHResult, __this, CStr("IDataObject::DAdvise returned %X"), hr
			.endif
		.endif
	.endif

	.if (iType == 1)
		invoke vf(pViewObject, IUnknown, Release)
	.elseif (iType == 2)
		invoke vf(pDataObject, IUnknown, Release)
	.endif
done:
	return hr

Advise@CContainer endp


if ?NEWMETHOD

HIMETRIC_PER_INCH   equ 2540

MAP_PIX_TO_LOGHIM macro x, ppli

	mov edx, HIMETRIC_PER_INCH
	mov eax, x
	mul edx
	mov ecx, ppli
	shr ecx, 1
	add eax, ecx
	cdq
	mov ecx, ppli
	div ecx
	endm

AtlPixelToHiMetric proc  lpSizeInPix:ptr SIZEL, lpSizeInHiMetric: ptr SIZEL

local	nPixelsPerInchX:DWORD
local	nPixelsPerInchY:DWORD
local	hDCScreen:HDC

	invoke GetDC, NULL
	mov hDCScreen, eax
	invoke GetDeviceCaps, hDCScreen, LOGPIXELSX
	mov nPixelsPerInchX, eax
	invoke GetDeviceCaps, hDCScreen, LOGPIXELSY
	mov nPixelsPerInchY, eax
	invoke ReleaseDC, NULL, hDCScreen

	mov ecx, lpSizeInPix
	MAP_PIX_TO_LOGHIM [ecx].SIZEL.cx_, nPixelsPerInchX
	mov ecx, lpSizeInHiMetric
	mov [ecx].SIZEL.cx_, eax

	mov ecx, lpSizeInPix
	MAP_PIX_TO_LOGHIM [ecx].SIZEL.cy, nPixelsPerInchY
	mov ecx, lpSizeInHiMetric
	mov [ecx].SIZEL.cy, eax
	ret
AtlPixelToHiMetric endp

SetRect@CContainer proc public uses __this this_:ptr CContainer, prect:ptr RECT

local sizel:SIZEL
local sizel2:SIZEL

	mov __this,this_

	invoke CopyRect, addr m_rect, prect
	.if (m_pOleInPlaceActiveObject)
		invoke vf(m_pOleInPlaceActiveObject, IOleInPlaceActiveObject, ResizeBorder),
			addr m_rect, addr m_OleInPlaceFrame, TRUE
	.endif

	mov eax, prect
;;	DebugOut "SetRect@CContainer(%d, %d, %d, %d)", [eax].RECT.left, [eax].RECT.top, [eax].RECT.right, [eax].RECT.bottom

	mov ecx, m_rectBorderSpace.left
	add [eax].RECT.left, ecx
	mov ecx, m_rectBorderSpace.top
	add [eax].RECT.top, ecx
	mov ecx, m_rectBorderSpace.right
	sub [eax].RECT.right, ecx
	mov ecx, m_rectBorderSpace.bottom
	sub [eax].RECT.bottom, ecx

	mov ecx, [eax].RECT.right
	sub ecx, [eax].RECT.left
	jnc @F
	xor ecx, ecx
@@:
	mov sizel.cx_, ecx
	mov ecx, [eax].RECT.bottom
	sub ecx, [eax].RECT.top
	jnc @F
	xor ecx, ecx
@@:
	mov sizel.cy, ecx
	invoke AtlPixelToHiMetric, addr sizel, addr sizel2
	.if (m_pOleObject)
		invoke vf(m_pOleObject, IOleObject, SetExtent), DVASPECT_CONTENT, addr sizel2
		DebugOut "IOleObject::SetExtent(%d,%d)=%X", sizel2.cx_, sizel2.cy, eax
	.endif
	.if (m_pOleInPlaceObject)
		invoke vf(m_pOleInPlaceObject, IOleInPlaceObject, SetObjectRects), prect, prect
ifdef _DEBUG
		mov ecx, prect
		DebugOut "IOleInPlaceObject::SetObjectRects(%d,%d,%d,%d)=%X", [ecx].RECT.left, [ecx].RECT.top, [ecx].RECT.right, [ecx].RECT.bottom, eax
endif
	.endif
;---------------------------------------- refresh window (this is a dialog
;---------------------------------------- window with CS_SAVEBITS set)
	.if (m_pOleInPlaceObjectWindowless)
		invoke InvalidateRect, m_hWndSite, 0, 1
	.endif
	ret
	align 4

SetRect@CContainer endp

else

SetRect@CContainer proc public this_:ptr CContainer, prect:ptr RECT

	invoke OnPosRectChange, this_, prect
	ret

SetRect@CContainer endp

endif

OnMouseMove@CContainer proc public uses __this this_:ptr CContainer, xPos:DWORD, yPos:DWORD, dwKeyState:DWORD
if ?POINTERINACTIVE
local dwPolicy:DWORD

	mov __this,this_
	.if (m_pPointerInactive)
		invoke IsActive@CContainer, __this
		.if (eax)
			jmp done
		.endif
		invoke vf(m_pPointerInactive, IPointerInactive, GetActivationPolicy), addr dwPolicy
		.if (eax == S_OK)
			DebugOut "OnMouseMove@CContainer, GetActivationPolicy=%X", dwPolicy
			.if (dwPolicy & POINTERINACTIVE_ACTIVATEONENTRY)
				invoke ActivateObject
				jmp done
			.endif
		.endif
		invoke vf(m_pPointerInactive, IPointerInactive, OnInactiveMouseMove), addr m_rect, xPos, yPos, dwKeyState
	.endif
endif
done:
	ret

OnMouseMove@CContainer endp

OnMouseClick@CContainer proc public uses __this this_:ptr CContainer 

if ?POINTERINACTIVE
local dwPolicy:DWORD

	mov __this,this_
	invoke IsActive@CContainer, __this
	.if (!eax)
		invoke ActivateObject
	.endif
endif
done:
	ret

OnMouseClick@CContainer endp

OnSetCursor@CContainer proc public uses __this this_:ptr CContainer, xPos:DWORD, yPos:DWORD, dwMessageId:DWORD
if ?POINTERINACTIVE
local dwPolicy:DWORD
local hr:DWORD

	mov __this,this_
	mov hr, S_FALSE
	.if (m_pPointerInactive)
		invoke IsActive@CContainer, __this
		.if (eax)
			jmp done
		.endif
		invoke vf(m_pPointerInactive, IPointerInactive, GetActivationPolicy), addr dwPolicy
		.if (eax == S_OK)
			DebugOut "OnSetCursor@CContainer, GetActivationPolicy=%X", dwPolicy
			.if (dwPolicy & POINTERINACTIVE_ACTIVATEONENTRY)
				invoke ActivateObject
				jmp done
			.endif
		.endif
		invoke vf(m_pPointerInactive, IPointerInactive, OnInactiveSetCursor), addr m_rect, xPos, yPos, dwMessageId, FALSE
		mov hr, eax
	.endif
endif
done:
	return hr
    
OnSetCursor@CContainer endp

ifdef @StackBase
	option stackbase:ebp
endif
	option prologue:@sehprologue
	option epilogue:@sehepilogue

;--- called on WM_ACTIVATE

OnActivate@CContainer proc public uses esi edi __this this_:ptr CContainer, fBool:BOOL

    mov __this,this_
	.if (m_pOleInPlaceActiveObject)
		.try
		invoke vf(m_pOleInPlaceActiveObject, IOleInPlaceActiveObject, OnFrameWindowActivate), fBool
		.exceptfilter
			mov __this,this_	;reload this register
			mov eax, _exception_info()
			mov eax, [eax].EXCEPTION_POINTERS.ExceptionRecord
			sub esp, 256
			mov edx, esp
			invoke wsprintf, edx, CStr("Exception %X at %X in IOleInPlaceActiveObject::OnFrameWindowActivate"),
				[eax].EXCEPTION_RECORD.ExceptionCode, [eax].EXCEPTION_RECORD.ExceptionAddress
			invoke SetStatusText@CViewObjectDlg, m_pViewObjectDlg, 0, esp
			add esp, 256
		.except
			mov __this,this_	;reload this register
		.endtry
	.endif
	ret
	align 4

OnActivate@CContainer endp

	option prologue: prologuedef
	option epilogue: epiloguedef
ifdef @StackBase
	option stackbase:esp
endif

if ?WINDOWLESS

MsgTab label dword
	dd WM_LBUTTONDOWN,	WM_RBUTTONDOWN,		WM_MBUTTONDOWN
	dd WM_LBUTTONUP,	WM_RBUTTONUP,		WM_MBUTTONUP
	dd WM_LBUTTONDBLCLK,WM_RBUTTONDBLCLK,	WM_MBUTTONDBLCLK
	dd WM_SETCURSOR,	WM_MOUSEMOVE
KeyMsgs label dword
	dd WM_KEYDOWN,		WM_KEYUP,		WM_CHAR,	WM_DEADCHAR
	dd WM_SYSKEYDOWN,	WM_SYSKEYUP,	WM_SYSDEADCHAR
	dd WM_CANCELMODE,	WM_HELP
SIZEMSGTAB equ ($ - offset MsgTab) / sizeof DWORD


OnMessage@CContainer proc public uses __this this_:ptr CContainer, uMsg:DWORD, wParam:WPARAM, lParam:LPARAM, plResult:ptr DWORD
	
local pt:POINT
local lResult:DWORD

	mov __this, this_
	mov eax, uMsg
	mov edx, edi
	mov edi, offset MsgTab
	mov ecx, SIZEMSGTAB
	repnz scasd
	xchg edi, edx
	.if (ZERO?)
		.if (edx <= offset KeyMsgs)
			invoke GetMessagePos
			movzx ecx, ax
			shr eax, 16
			mov pt.y, eax
			mov pt.x, ecx
			invoke ScreenToClient, m_hWndSite, addr pt
			mov eax, pt.x
			mov ecx, pt.y
			.if ((eax >= m_rect.left) && (eax < m_rect.right) && (ecx >= m_rect.top) && (ecx < m_rect.bottom))
				mov eax, 1
			.else
				xor eax, eax
			.endif
		.endif
		.if (eax)
			invoke vf(m_pOleInPlaceObjectWindowless, IOleInPlaceObjectWindowless, OnWindowMessage),
				uMsg, wParam, lParam, plResult
			ret
		.endif
	.endif
	return S_FALSE

OnMessage@CContainer endp

IsWindowless@CContainer proc public uses __this this_:ptr CContainer

	mov __this, this_
	.if (m_pOleInPlaceObjectWindowless)
		mov eax, TRUE
	.else
		mov eax, FALSE
	.endif
	ret

IsWindowless@CContainer endp
endif

IsActive@CContainer proc public uses __this this_:ptr CContainer

	mov __this, this_
	.if (m_pOleInPlaceObject)
		mov eax, TRUE
	.else
		mov eax, FALSE
	.endif
	ret

IsActive@CContainer endp

MyGetStorage proc

local pUnknown:LPUNKNOWN
local pStorage:LPSTORAGE

	invoke vf(m_pObjectItem, IObjectItem, GetStorage)
	.if (eax)
		mov pUnknown, eax
		invoke vf(pUnknown, IUnknown, QueryInterface), addr IID_IStorage, addr pStorage
		.if (eax == S_OK)
			invoke vf(pStorage, IStorage, Release)
			mov eax, pUnknown
			jmp done
		.endif
	.endif
	.if (!g_pStorage)
		invoke StgCreateDocfile,
			NULL, STGM_READWRITE or STGM_SHARE_EXCLUSIVE or STGM_TRANSACTED or STGM_DELETEONRELEASE,
			NULL, addr g_pStorage
	.endif
	mov eax, g_pStorage
done:
	ret
MyGetStorage endp


MyGetStream proc

local pUnknown:LPUNKNOWN
local pStream:LPSTREAM

	invoke vf(m_pObjectItem, IObjectItem, GetStorage)
	.if (eax)
		mov pUnknown, eax
		invoke vf(pUnknown, IUnknown, QueryInterface), addr IID_IStream, addr pStream
		.if (eax == S_OK)
			invoke vf(pStream, IStream, Release)
			mov eax, pUnknown
			jmp done
		.endif
	.endif
	.if (!g_pStream)
		invoke CreateStreamOnHGlobal, NULL, TRUE, addr g_pStream
	.endif
	mov eax, g_pStream
done:
	ret
MyGetStream endp


LogInit proc pszText:LPSTR, hr:DWORD
	invoke printf@CLogWindow, CStr("%s: %s returned %X",10), addr g_szContainer, pszText, hr
	ret
LogInit endp


TranslateAccelerator@CContainer proc public uses __this this_:ptr CContainer, pMsg:ptr MSG

	mov __this, this_
	.if (m_pOleInPlaceActiveObject)
		invoke vf(m_pOleInPlaceActiveObject, IOleInPlaceActiveObject, TranslateAccelerator_), pMsg
		.if (eax == S_OK)
			return TRUE
		.endif
	.endif
	xor eax, eax
	ret

TranslateAccelerator@CContainer endp

;--- load properties from property bag

LoadProperties proc pPersistPropertyBag:LPPERSISTPROPERTYBAG, pStorage:LPSTORAGE

local	hr:DWORD
local	pPropertyBag:LPPROPERTYBAG
local	clsid:CLSID

		mov pPropertyBag, NULL
		.if (pStorage)
if ?USEMYPROPBAG
			invoke vf(pPersistPropertyBag, IPersistPropertyBag, GetClassID), addr clsid
			invoke Create@CPropertyBag, pStorage, addr clsid, FALSE, addr pPropertyBag
			.if (eax != S_OK)
				invoke printf@CLogWindow, CStr("--- Create@PropertyBag failed [%X]",10), eax
			.endif
else
			invoke vf(pStorage, IUnknown, QueryInterface), addr IID_IPropertyBag, addr pPropertyBag
			.if (eax != S_OK)
				invoke printf@CLogWindow, CStr("--- IStorage::QueryInterface(IPropertyBag) failed [%X]",10), eax
			.endif
endif
			mov hr, eax
			.if (eax != S_OK)
				jmp done
			.endif
		.endif
		.if (pPropertyBag)
			invoke vf(pPersistPropertyBag, IPersistPropertyBag, Load), pPropertyBag, NULL
			mov hr, eax
;;			DebugOut "IPersistPropertyBag::Load(%X) returned %X", pPropertyBag, eax
			invoke LogInit, CStr("IPersistPropertyBag::Load"), eax
			invoke vf(pPropertyBag, IUnknown, Release)
			.if (hr == S_OK)
				mov m_bLoadPropertyBag, TRUE
			.else
				invoke OutputMessage, m_hWndSite, hr, CStr("IPersistPropertyBag::Load"), 0
			.endif
		.else
			invoke vf(pPersistPropertyBag, IPersistPropertyBag, InitNew)
			mov hr, eax
;;			DebugOut "IPersistPropertyBag::InitNew returned %X", eax
			invoke LogInit, CStr("IPersistPropertyBag::InitNew"), eax
		.endif
done:
		return hr

LoadProperties endp

ifdef @StackBase
	option stackbase:ebp
endif
	option prologue:@sehprologue
	option epilogue:@sehepilogue

DoVerb@CContainer proc public uses esi edi __this this_:ptr CContainer, verb:DWORD

local	szText[128]:BYTE

	mov __this,this_
	.try
	.if (m_pOleObject)
		invoke vf(m_pObjectItem, IObjectItem, GetFlags)
		.if (eax & OBJITEMF_ROT)
			invoke SetStatusText@CViewObjectDlg, m_pViewObjectDlg, 0, CStr("No activation for ROT items possible")
			jmp done
		.endif
		lea ecx, m_OleClientSite
		invoke vf(m_pOleObject, IOleObject, DoVerb),
				verb, NULL, ecx, NULL, m_hWndSite, addr m_rect
		push eax
		invoke wsprintf, addr szText, CStr("IOleObject::DoVerb(%d) returned %X"), verb, eax
		invoke SetStatusText@CViewObjectDlg, m_pViewObjectDlg, 0, addr szText
		pop eax
	.endif
	.exceptfilter
		mov __this,this_	;reload "this" register
		mov eax, _exception_info()
		invoke DisplayExceptionInfo, m_hWndSite, eax, CStr("DoVerb@CContainer"), EXCEPTION_EXECUTE_HANDLER
	.except
		mov eax, E_FAIL
	.endtry
done:
	ret

DoVerb@CContainer endp

SafeRelease proc public uses ebx esi edi pUnknown:LPUNKNOWN

;--------------------- use a silent try block for that
	nop
	.try
	.if (pUnknown)
		invoke vf(pUnknown, IUnknown, Release)
	.endif
	.exceptfilter
		mov eax, _exception_info()
		mov eax, [eax].EXCEPTION_POINTERS.ExceptionRecord
		invoke printf@CLogWindow, CStr("%s: Exception 0x%08X occured at 0x%08X",10),
			CStr("Calling IUnknown::Release"), [eax].EXCEPTION_RECORD.ExceptionCode, [eax].EXCEPTION_RECORD.ExceptionAddress
		invoke MessageBeep, MB_OK
		mov eax, EXCEPTION_EXECUTE_HANDLER
	.except
		mov eax, E_FAIL
	.endtry
	ret
SafeRelease endp

Save@CContainer proc public uses esi edi __this this_:ptr CContainer, iType:DWORD, bForceSave:BOOL

local pPersistStorage:LPPERSISTSTORAGE
local pPersistStreamInit:LPPERSISTSTREAMINIT
local pPersistFile:LPPERSISTFILE
local pPersistPropertyBag:LPPERSISTPROPERTYBAG
local pStorage:LPSTORAGE
local pStream:LPSTREAM
local pPropertyBag:LPPROPERTYBAG
local pwszFile:ptr WORD
local clsid:CLSID
local szFile[MAX_PATH]:byte
local wszFile[MAX_PATH]:word

	mov __this,this_

	.if (!m_pOleObject)
		mov eax, E_FAIL
		jmp done
	.endif

	mov pPersistStorage, NULL
	mov pPersistStreamInit, NULL
	mov pPersistPropertyBag, NULL
	mov pPersistFile, NULL
	mov pPropertyBag, NULL
	mov pwszFile, NULL

	.try
	.if (iType == SAVE_STORAGE)

		invoke MyGetStorage
		mov pStorage, eax
		.if (eax)
			invoke vf(m_pOleObject, IUnknown, QueryInterface), addr IID_IPersistStorage, addr pPersistStorage
			.if (eax == S_OK)
				.if (!bForceSave)
					invoke vf(pPersistStorage, IPersistStorage, IsDirty)
					.if (g_bDispContainerCalls)
						invoke printf@CLogWindow, CStr("%s: IPersistStorage::IsDirty returned %X (S_FALSE[=1] will NOT save)",10),
							addr g_szContainer, eax
					.endif
				.endif
				.if (eax != S_FALSE)
					invoke vf(pPersistStorage, IPersistStorage, GetClassID), addr clsid
					invoke WriteClassStg, pStorage, addr clsid
					mov eax, pStorage
					mov ecx, TRUE
					.if (eax == g_pStorage)
						mov ecx, FALSE
					.endif
					invoke vf(pPersistStorage, IPersistStorage, Save), pStorage, ecx
;;					DebugOut "IPersistStorage::Save(%X) returned %X", pStorage, eax
					push eax
					invoke vf(pPersistStorage, IPersistStorage, SaveCompleted), NULL
					pop eax
					mov ecx, CStr("IPersistStorage::Save")
				.else
					xor ecx, ecx
					mov eax, S_OK
				.endif
			.else
				mov ecx, CStr("QueryInterface(IPersistStorage)")
			.endif
		.else
			mov ecx, CStr("StgCreateDocfile")
		.endif

	.elseif (iType == SAVE_STREAM)

		invoke MyGetStream
		mov pStream, eax
		.if (eax)
			invoke vf(m_pOleObject, IUnknown, QueryInterface), addr IID_IPersistStreamInit, addr pPersistStreamInit
			.if (eax == S_OK)
				.if (!bForceSave)
					invoke vf(pPersistStreamInit, IPersistStreamInit, IsDirty)
					.if (g_bDispContainerCalls)
						invoke printf@CLogWindow, CStr("%s: IPersistStreamInit::IsDirty returned %X (S_FALSE[=1] will NOT save)",10),
							addr g_szContainer, eax
					.endif
				.endif
				.if (eax != S_FALSE)
					invoke vf(pStream, IStream, Seek), g_dqNull, STREAM_SEEK_SET, NULL
					invoke vf(pPersistStreamInit, IPersistStreamInit, GetClassID), addr clsid
					invoke WriteClassStm, pStream, addr clsid
					mov eax, pStream
					mov ecx, TRUE
					.if (eax == g_pStream)
						mov ecx, FALSE
					.endif
					invoke vf(pPersistStreamInit, IPersistStreamInit, Save), pStream, ecx
;;					DebugOut "IPersistStreamInit::Save(%X) returned %X", pStream, eax
					mov ecx, CStr("IPersistStreamInit::Save")
				.else
					xor ecx, ecx
					mov eax, S_OK
				.endif
			.else
				mov ecx, CStr("QueryInterface(IPersistStreamInit)")
			.endif
		.else
			mov ecx, CStr("CreateStreamOnHGlobal")
		.endif

	.elseif (iType ==  SAVE_PROPBAG)

		invoke MyGetStorage
		.if (eax)
			mov pStorage, eax
			invoke vf(m_pOleObject, IUnknown, QueryInterface), addr IID_IPersistPropertyBag, addr pPersistPropertyBag
			.if (eax == S_OK)
if ?USEMYPROPBAG
				invoke vf(pPersistPropertyBag, IPersistPropertyBag, GetClassID), addr clsid
				invoke Create@CPropertyBag, pStorage, addr clsid, TRUE, addr pPropertyBag
else
				invoke vf(pStorage, IUnknown, QueryInterface), addr IID_IPropertyBag, addr pPropertyBag
endif
				.if (eax == S_OK)
					invoke vf(pPersistPropertyBag, IPersistPropertyBag, GetClassID), addr clsid
					invoke WriteClassStg, pStorage, addr clsid
					invoke vf(pPersistPropertyBag, IPersistPropertyBag, Save), pPropertyBag, TRUE, FALSE
;;					DebugOut "IPersistPropertyBag::Save(%X) returned %X", pPropertyBag, eax
					mov ecx, CStr("IPersistPropertyBag::Save")
				.else
if ?USEMYPROPBAG
					mov ecx, CStr("Create@CPropertyBag")
else
					mov ecx, CStr("QueryInterface(IPropertyBag)")
endif
				.endif
			.else
				mov ecx, CStr("QueryInterface(IPersistPropertyBag)")
			.endif
		.else
			mov ecx, CStr("StgCreateDocfile")
		.endif

	.elseif (iType ==  SAVE_FILE)

		invoke vf(m_pOleObject, IUnknown, QueryInterface), addr IID_IPersistFile, addr pPersistFile
		.if (eax == S_OK)
			mov szFile, 0
			invoke vf(pPersistFile, IPersistFile, GetCurFile), addr pwszFile
			.if ((eax == S_OK) && pwszFile)
				invoke WideCharToMultiByte,CP_ACP,0,pwszFile,-1,addr szFile,sizeof szFile,0,0
			.endif

			invoke MyGetFileName, m_hWndSite, addr szFile, sizeof szFile, NULL, 0, 1, NULL
			.if (eax)
				invoke MultiByteToWideChar,CP_ACP,MB_PRECOMPOSED, addr szFile, -1, addr wszFile, MAX_PATH
				invoke vf(pPersistFile, IPersistFile, Save), addr wszFile, FALSE
;;				DebugOut "IPersistFile::Save returned %X", eax
				push eax
				invoke vf(pPersistFile, IPersistFile, SaveCompleted), addr wszFile
				pop eax
			.endif
			mov ecx, CStr("IPersistFile::Save")
		.else
			mov ecx, CStr("QueryInterface(IPersistFile)")
		.endif

	.else
		mov eax, E_FAIL
		mov ecx, CStr("NI")
	.endif
	.if (ecx)
		invoke LogInit, ecx, eax
		.if (eax != S_OK)
			push eax
			invoke OutputMessage, m_hWndSite, eax, ecx, 0
			pop eax
		.endif
	.endif

	.exceptfilter
		mov __this,this_	;reload "this" register
		mov eax, _exception_info()
		invoke DisplayExceptionInfo, m_hWndSite, eax, CStr("Save@CContainer"), EXCEPTION_EXECUTE_HANDLER
	.except
		mov __this,this_	;reload this register
		mov eax, E_FAIL
	.endtry

	push eax
	.if (pwszFile)
		invoke CoTaskMemFree, pwszFile
	.endif
	invoke SafeRelease, pPersistStorage
	invoke SafeRelease, pPersistStreamInit
	invoke SafeRelease, pPersistPropertyBag
	invoke SafeRelease, pPersistFile
	invoke SafeRelease, pPropertyBag
	pop eax

done:
	ret
	align 4

Save@CContainer endp

Load@CContainer proc public uses __this this_:ptr CContainer

local	pPersistFile:LPPERSISTFILE
local	szFile[MAX_PATH]:byte
local	wszFile[MAX_PATH]:word

	mov __this, this_

	invoke vf(m_pOleObject, IUnknown, QueryInterface), addr IID_IPersistFile, addr pPersistFile
	.if (eax != S_OK)
		jmp done
	.endif
	mov szFile, 0
	invoke MyGetFileName, m_hWndSite, addr szFile, sizeof szFile, NULL, 0, 0, CStr("Select File for IPersistFile::Load")
	.if (eax)
		invoke MultiByteToWideChar,CP_ACP,MB_PRECOMPOSED, addr szFile, -1, addr wszFile, MAX_PATH
		invoke vf(pPersistFile, IPersistFile, Load), addr wszFile, NULL
;;		DebugOut "IPersistFile::Load returned %X", eax
		invoke LogInit, CStr("IPersistFile::Load"), eax
	.else
		mov eax, E_FAIL
	.endif
	push eax
	invoke SafeRelease, pPersistFile
	pop eax
done:
	ret
Load@CContainer endp

;--- initialize/load object

InitObject proc uses esi edi __this pUnknown:LPUNKNOWN

local	pStorage:LPSTORAGE
local	pStream:LPSTREAM
local	pPersistStreamInit:LPPERSISTSTREAMINIT
local	pPersistStorage:LPPERSISTSTORAGE
local	pPersistPropertyBag:LPPERSISTPROPERTYBAG
;local	pPersistFile:LPPERSISTFILE
local	clsid:CLSID
local	dwObjFlags:DWORD
local	this_:ptr CContainer
;local	szFile[MAX_PATH]:byte

	mov pStorage, NULL
	mov pStream, NULL
	invoke vf(m_pObjectItem, IObjectItem, GetFlags)
	.if (eax & OBJITEMF_INIT)
		jmp done
	.endif
	or eax, OBJITEMF_INIT
	mov dwObjFlags, eax

	.if (pUnknown)
		invoke vf(pUnknown, IUnknown, QueryInterface), addr IID_IStorage, addr pStorage
		.if (eax != S_OK)
			invoke vf(pUnknown, IUnknown, QueryInterface), addr IID_IStream, addr pStream
		.endif
	.endif

	mov this_, __this

	.try

	.if (g_bUseIPersistFile && (!pStorage) && (!pStream))
		invoke Load@CContainer, __this
		.if (eax == S_OK)
			jmp done2
		.endif
	.endif

	.if (g_bUseIPersistPropBag)
		invoke vf(m_pOleObject, IUnknown, QueryInterface), addr IID_IPersistPropertyBag, addr pPersistPropertyBag
		.if (eax == S_OK)
			invoke LoadProperties, pPersistPropertyBag, pUnknown
			push eax
			invoke vf(pPersistPropertyBag, IUnknown, Release)
			pop eax
			.if (eax == S_OK)
				jmp done2
			.endif
		.endif
	.endif

	.if (g_bUseIPersistStream)
		invoke vf(m_pOleObject, IUnknown, QueryInterface), addr IID_IPersistStreamInit, addr pPersistStreamInit
		.if (eax == S_OK)
			.if (pStorage)
				invoke vf(pStorage, IStorage, OpenStream), CStrW(L("contents")),\
					NULL, STGM_READ or STGM_SHARE_EXCLUSIVE, NULL, addr pStream
			.endif
			.if (pStream)
				invoke vf(pStream, IStream, Seek), g_dqNull, STREAM_SEEK_SET, NULL
				.if (!pStorage)
					invoke ReadClassStm, pStream, addr clsid
				.endif
				invoke vf(pPersistStreamInit, IPersistStreamInit, Load), pStream
;;				DebugOut "IPersistStreamInit::Load(%X) returned %X", pStream, eax
				invoke LogInit, CStr("IPersistStreamInit::Load"), eax
				.if (eax != S_OK)
					push eax
					invoke OutputMessage, m_hWndSite, eax, CStr("IPersistStreamInit::Load"), 0
					pop eax
				.endif
			.else
				invoke vf(pPersistStreamInit, IPersistStreamInit, InitNew)
;;				DebugOut "IPersistStreamInit::InitNew() returned %X", eax
				invoke LogInit, CStr("IPersistStreamInit::InitNew"), eax
			.endif
			push eax
			invoke vf(pPersistStreamInit, IUnknown, Release)
			pop eax
			.if (eax == S_OK)
				jmp done2
			.endif
		.endif
	.endif

	invoke vf(m_pOleObject, IUnknown, QueryInterface), addr IID_IPersistStorage, addr pPersistStorage
	.if (eax == S_OK)
		.if (!pStorage)
			invoke StgCreateDocfile, NULL, STGM_READWRITE or STGM_SHARE_EXCLUSIVE or STGM_DELETEONRELEASE,\
				NULL, addr pStorage
			DebugOut "StgCreateDocfile(NULL) returned %X", eax
			.if (eax == S_OK)
				invoke vf(pPersistStorage, IPersistStorage, GetClassID), addr clsid
				invoke WriteClassStg, pStorage, addr clsid
				invoke vf(pPersistStorage, IPersistStorage, InitNew), pStorage
;;				DebugOut "IPersistStorage::InitNew(%X) returned %X", pStorage, eax
				invoke LogInit, CStr("IPersistStorage::InitNew"), eax
			.else
				push eax
				invoke OutputMessage, m_hWndSite, eax, CStr("StgCreateDocfile"), 0
				pop eax
			.endif
		.else
			invoke vf(pPersistStorage, IPersistStorage, Load), pStorage
;;			DebugOut "IPersistStorage::Load(%X) returned %X", pStorage, eax
			invoke LogInit, CStr("IPersistStorage::Load"), eax
			.if (eax != S_OK)
				push eax
				invoke OutputMessage, m_hWndSite, eax, CStr("IPersistStorage::Load"), 0
				pop eax
			.endif
		.endif
		push eax
		invoke vf(pPersistStorage, IUnknown, Release)
		pop eax
	.endif

done2:
	.if (eax == S_OK)
		invoke vf(m_pObjectItem, IObjectItem, SetFlags), dwObjFlags
	.endif

	.exceptfilter
		mov __this,this_	;reload "this" register
		mov eax, _exception_info()
		invoke DisplayExceptionInfo, m_hWndSite, eax, CStr("Object initialization"), EXCEPTION_EXECUTE_HANDLER
	.except
		mov __this,this_	;reload this register
	.endtry

	invoke SafeRelease, pStream
	invoke SafeRelease, pStorage
done:
	ret

InitObject endp

;--- destructor

Destroy@CContainer proc uses esi edi __this this_:ptr CContainer

local	pViewObject:LPVIEWOBJECT
local	pDataObject:LPDATAOBJECT

	DebugOut "Destroy@CContainer enter"
    mov __this,this_
	.try
if ?OLELINK
	.if (m_pOleLink)
		invoke vf(m_pOleLink, IOleLink, UnbindSource)
		invoke vf(m_pOleLink, IOleLink, Release)
	.endif
endif
if ?DOCUMENT
	.if (m_pOleDocumentView)
		invoke vf(m_pOleDocumentView, IUnknown, Release)
	.endif
endif
if ?POINTERINACTIVE
	.if (m_pPointerInactive)
		invoke vf(m_pPointerInactive, IUnknown, Release)
	.endif
endif
if 1
	.if (m_pOleInPlaceActiveObject)
		invoke vf(m_pOleInPlaceActiveObject, IUnknown, Release)
	.endif
if ?WINDOWLESS
	.if (m_pOleInPlaceObjectWindowless)
		invoke vf(m_pOleInPlaceObjectWindowless, IUnknown, Release)
	.endif
endif
	.if (m_pOleInPlaceObject)
		invoke vf(m_pOleInPlaceObject, IUnknown, Release)
	.endif
endif
	.if (m_pAdviseSink)
		.if (m_bViewConnected)
			invoke vf(m_pOleObject, IUnknown, QueryInterface), addr IID_IViewObject, addr pViewObject
			.if (eax == S_OK)
				invoke vf(pViewObject, IViewObject, SetAdvise), DVASPECT_CONTENT, 0, NULL
				invoke vf(pViewObject, IUnknown, Release)
			.endif
		.endif
		.if (m_dwDataConnection)
			invoke vf(m_pOleObject, IUnknown, QueryInterface), addr IID_IDataObject, addr pDataObject
			.if (eax == S_OK)
				invoke vf(pDataObject, IDataObject, DUnadvise), m_dwDataConnection
				invoke vf(pDataObject, IUnknown, Release)
			.endif
		.endif
		.if (m_dwConnection)
			invoke vf(m_pOleObject, IOleObject, Unadvise), m_dwConnection
		.endif
		invoke vf(m_pAdviseSink, IUnknown, Release)
	.endif
	.if (m_pOleObject)
		invoke vf(m_pOleObject, IUnknown, Release)
	.endif

	.if (m_pMonikerFull)
		invoke vf(m_pMonikerFull, IUnknown, Release)
	.endif
	.if (m_pMonikerCon)
		invoke vf(m_pMonikerCon, IUnknown, Release)
	.endif
	.if (m_pMonikerRel)
		invoke vf(m_pMonikerRel, IUnknown, Release)
	.endif

	.if (m_pObjectWithSite)
		invoke vf(m_pObjectWithSite, IUnknown, Release)
	.endif
	.if (m_hMenuView)
		invoke DestroyMenu, m_hMenuView
	.endif
	
	.exceptfilter
		mov __this,this_	;reload this register
		mov eax, _exception_info()
		invoke DisplayExceptionInfo, m_hWndSite, eax, CStr("Destroy@CContainer"), EXCEPTION_EXECUTE_HANDLER
	.except
		mov __this,this_	;reload this register
	.endtry

	invoke free, __this

	invoke printf@CLogWindow, CStr("--- container %X destroyed",10), __this
	DebugOut "Destroy@CContainer exit"
	ret
Destroy@CContainer endp

;--- in place activate object

ActivateObject proc uses esi 

if ?DOCUMENT
local	pOleDocument:LPOLEDOCUMENT
endif

		mov esi, OLEIVERB_INPLACEACTIVATE
if ?DOCUMENT
		mov pOleDocument, NULL
;---------------------------------- avoid to call INPLACEACTIVATE
		.if (g_bDocumentSiteSupp)
			invoke vf(m_pOleObject, IUnknown, QueryInterface), addr IID_IOleDocument, addr pOleDocument
			.if (eax == S_OK)
				mov esi, OLEIVERB_UIACTIVATE
				invoke vf(pOleDocument, IUnknown, Release)
				DebugOut "Object supports IOleDocument"
				jmp step2
			.endif
		.endif
endif
		DebugOut "Starting activation, first verb is InPlaceActivate/UIActivate, then Show"

		invoke vf(m_pOleObject, IOleObject, DoVerb),\
				esi, NULL, __this, NULL, m_hWndSite, addr m_rect
		DebugOut "IOleObject::DoVerb(%d) returned %X", esi, eax
		.if (eax != S_OK)
step2:
			invoke vf(m_pOleObject, IOleObject, DoVerb),\
				OLEIVERB_SHOW, NULL, __this, NULL, m_hWndSite, addr m_rect
			DebugOut "IOleObject::DoVerb(SHOW) returned %X", eax
		.endif
		.if (eax != S_OK)
			invoke DisplayHResult, __this, CStr("Last IOleObject:DoVerb returned %X"), eax
		.endif
		ret
ActivateObject endp

Create@CContainer proc public uses esi edi __this pObjectItem:LPOBJECTITEM, pViewObjectDlg:ptr CViewObjectDlg, prect:ptr RECT

local	dwESP:DWORD
local	dwMiscStatus:DWORD
local	pRunnableObject:LPRUNNABLEOBJECT
local	pQuickActivate:LPQUICKACTIVATE
local	qacontainer:QACONTAINER
local	qacontrol:QACONTROL
local	pwszName:LPOLESTR
local	pUnknown:LPUNKNOWN
local	bOleDocument:BOOL
local	bDontActivate:BOOLEAN
local	sizel:SIZEL
local	rect:RECT
local	this_:ptr CContainer
;;local	szText[128]:byte

	DebugOut "Create@CContainer enter"
	invoke malloc, sizeof CContainer
	.if (!eax)
		jmp exit
	.endif

	mov dwESP, esp
	mov __this, eax
	mov this_, eax
	mov m_OleClientSite.lpVtbl,		offset COleClientSiteVtbl
	mov m_OleInPlaceSite.lpVtbl,	offset COleInPlaceSiteVtbl
	mov m_OleInPlaceFrame.lpVtbl,	offset COleInPlaceFrameVtbl
	mov m_OleControlSite.lpVtbl,	offset COleControlSiteVtbl
if ?DISPATCH
	mov m_Dispatch.lpVtbl,			offset CDispatchVtbl
endif
if ?COMMANDTARGET
	mov m_OleCommandTarget.lpVtbl,	offset COleCommandTargetVtbl
endif
if ?DOCUMENT
	mov m_OleDocumentSite.lpVtbl,	offset COleDocumentSiteVtbl
endif
if ?OLECONTAINER
	mov m_OleContainer.lpVtbl,		offset COleContainerVtbl
endif
if ?CALLFACTORY
	mov m_CallFactory.lpVtbl,		offset CCallFactoryVtbl
endif
if ?SERVICEPROVIDER
	mov m_ServiceProvider.lpVtbl,	offset CServiceProviderVtbl
endif
if ?DOCHOSTSHOWUI
	mov m_DocHostShowUI.lpVtbl,		offset CDocHostShowUIVtbl
endif
	mov eax, pViewObjectDlg
	mov m_pViewObjectDlg, eax
	mov eax, [eax].CDlg.hWnd
	mov m_hWndSite, eax
	.if (prect)
		invoke CopyRect, addr m_rect, prect
	.else
		lea eax, m_rect
		mov prect, eax
		invoke GetClientRect, m_hWndSite, addr m_rect
		mov eax, g_dwBorder
		add m_rect.left, eax
		add m_rect.top, eax
		sub m_rect.right, eax
		sub m_rect.bottom, eax
	.endif
	mov m_pOleInPlaceObject, NULL
	mov m_pOleInPlaceActiveObject, NULL

	mov eax, pObjectItem
	mov m_pObjectItem, eax
	invoke GetUnknown@CObjectItem, eax
	mov pUnknown, eax

if ?OLELINK
	invoke vf(pUnknown, IUnknown, QueryInterface), addr IID_IOleLink, addr m_pOleLink
	.if (eax == S_OK)
		invoke vf(m_pOleLink, IOleLink, BindToSource), 0, NULL
	.endif
endif

	invoke printf@CLogWindow, CStr("--- container %X created",10), __this

	mov m_bClientSiteSet, FALSE
	.try

		invoke vf(pUnknown, IUnknown, QueryInterface), addr IID_IOleObject, addr m_pOleObject
		.if (eax != S_OK)
			invoke vf(pUnknown, IUnknown, QueryInterface), addr IID_IObjectWithSite, addr m_pObjectWithSite
			.if (eax != S_OK)
				invoke Destroy@CContainer, __this
				return 0
			.endif
		.endif
		mov m_dwRefCount, 1

		mov dwMiscStatus, 0
		.if (m_pOleObject)

			invoke vf(m_pOleObject, IOleObject, GetMiscStatus), DVASPECT_CONTENT, addr dwMiscStatus

			mov bDontActivate, FALSE
if ?POINTERINACTIVE
			.if (g_bUseIPointerInactive)
				invoke vf(pUnknown, IUnknown, QueryInterface), addr IID_IPointerInactive,\
					addr m_pPointerInactive
				.if ((eax == S_OK)) ; && (dwMiscStatus & OLEMISC_IGNOREACTIVATEWHENVISIBLE))
					mov bDontActivate, TRUE
					jmp step1
				.endif
			.endif
endif
			.if (g_bUseIQuickActivate)
				invoke vf(pUnknown, IUnknown, QueryInterface), addr IID_IQuickActivate, addr pQuickActivate
				.if (eax == S_OK)
					invoke ZeroMemory, addr qacontainer, sizeof QACONTAINER
					mov qacontainer.cbSize, sizeof QACONTAINER
					mov qacontainer.pClientSite, __this
					xor ecx, ecx
					.if (g_bUserMode)
						or ecx, QACONTAINER_USERMODE
					.endif
					.if (g_bUIDead)
						or ecx, QACONTAINER_UIDEAD
					.endif
					mov qacontainer.dwAmbientFlags, ecx
					mov qacontainer.colorFore, COLOR_WINDOWTEXT
					mov qacontainer.colorBack, COLOR_APPWORKSPACE 
					mov eax, g_LCID
					mov qacontainer.lcid, eax
					mov qacontrol.cbSize, sizeof QACONTROL
					invoke vf(pQuickActivate, IQuickActivate, QuickActivate), addr qacontainer, addr qacontrol
					.if (eax == S_OK)
						mov eax, qacontrol.dwMiscStatus
						mov dwMiscStatus, eax
						mov m_bClientSiteSet, TRUE
					.endif
					invoke vf(pQuickActivate, IUnknown, Release)
				.endif
			.endif
step1:
			.if (!m_bClientSiteSet)
				.if (dwMiscStatus & OLEMISC_SETCLIENTSITEFIRST) 
					invoke vf(m_pOleObject, IOleObject, SetClientSite), __this
					DebugOut "IOleObject::SetClientSite() returned %X", eax
					.if (eax == S_OK)
						mov m_bClientSiteSet, TRUE
					.endif
				.endif
			.endif

			invoke vf(m_pObjectItem, IObjectItem, GetStorage)
			invoke InitObject, eax

			.if (m_bClientSiteSet == FALSE) 
				invoke vf(m_pOleObject, IOleObject, SetClientSite), __this
				DebugOut "IOleObject::SetClientSite() returned %X", eax
				.if (eax == S_OK)
					mov m_bClientSiteSet, TRUE
				.else
					invoke DisplayHResult, __this, CStr("IOleObject:SetClientSite returned %X"), eax
				.endif
			.endif

			.if (dwMiscStatus & OLEMISC_STATIC)
				jmp label1
			.endif
			.if (m_bClientSiteSet)
				invoke vf(m_pObjectItem, IObjectItem, GetDisplayName), addr pwszName
				mov ecx, pwszName
				.if (!ecx)
					mov ecx, CStrW(L("unnamed"))
				.endif
				invoke vf(m_pOleObject, IOleObject, SetHostNames),\
					CStrW(L("COMView")), ecx
				DebugOut "IOleObject::SetHostNames() returned %X", eax
				invoke CoTaskMemFree, pwszName		;pwszName may be NULL

				invoke vf(m_pObjectItem, IObjectItem, GetFlags) 
				.if (eax & OBJITEMF_ROT)
					jmp label1
				.endif

				.if (!bDontActivate)
					invoke ActivateObject
				.endif
			.endif
label1:
if 0
			invoke vf(m_pOleObject, IOleObject, GetExtent),\
					DVASPECT_CONTENT, addr sizel
endif
		.else
			invoke vf(m_pObjectWithSite, IObjectWithSite, SetSite), __this
		.endif
		invoke SetStatus

	.exceptfilter
		mov __this,this_	;reload this register
		mov eax, _exception_info()
		invoke DisplayExceptionInfo, m_hWndSite, eax, CStr("OLE Container"), EXCEPTION_EXECUTE_HANDLER
	.except
		mov __this,this_	;reload this register
		.if (m_bClientSiteSet)
			mov m_bClientSiteSet, FALSE
			invoke vf(m_pOleObject, IOleObject, SetClientSite), NULL
		.endif
		invoke Destroy@CContainer, __this
		xor eax, eax
		jmp exit
	.endtry
done:
	.if (esp != dwESP)
		invoke MessageBox, m_hWndSite, CStr("Object has modified ESP during initialization",10,"some possible reasons are:",10,"wrong assumptions about number/size of parameters", 10, "wrong calling convention"), 0, MB_OK
		mov esp, dwESP
	.endif
	mov eax, __this
exit:
	DebugOut "Create@CContainer exit"
	ret

Create@CContainer endp

	option prologue: prologuedef
	option epilogue: epiloguedef
ifdef @StackBase
	option stackbase:esp
endif

	end
