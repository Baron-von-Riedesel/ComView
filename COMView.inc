
;----------------------- includes ---------------------

	.nolist
	.nocref
WIN32_LEAN_AND_MEAN	equ 1
INCL_OLE2			equ 1
	include windows.inc
	include commctrl.inc
	include commdlg.inc
	include windowsx.inc
	include olectl.inc
	include propidl.inc

	include macros.inc
	include try.inc

	.list
	.cref

;----------------------- global constants ---------------------

MAXINPUTTEXT		equ 80		;max length of edit text for std. input dialog
?POSTSETDISP		equ 1		;use PostMessage to set a new IDispatch/ITypeInfo
WM_WNDDESTROYED		equ WM_USER+99	;special msg to inform all popup windows
								;about a window being destroyed
?REMOVEITEM			equ 1		;support REMOVEITEM
?HTMLHELP			equ 1		;use HtmlHelp instead of simple richedit text
?UPDATECHK			equ 1
?HDRBMPS			equ 1		;show listview header bitmaps

if ?HTMLHELP
protoHtmlHelp	typedef proto :HWND, :LPSTR, :DWORD, :DWORD
LPHTMLHELP		typedef ptr protoHtmlHelp
HH_DISPLAY_TOPIC        equ 0000h
HH_DISPLAY_TOC          equ 0001h
HH_DISPLAY_INDEX        equ 0002h
HH_DISPLAY_SEARCH       equ 0003h
HH_HELP_CONTEXT         equ 000Fh
HH_INITIALIZE           equ 001Ch
HH_UNINITIALIZE         equ 001Dh
ShowHtmlHelp proto pszFile:LPSTR, uCommand:DWORD, dwData:DWORD
DoHtmlHelp proto uCommand:DWORD, dwData:DWORD
endif

;--- funcions from IProp.dll (ole32.dll)

protoPropStgNameToFmtId	typedef proto :LPOLESTR, :ptr FMTID
protoStgCreatePropStg	typedef proto :LPUNKNOWN, :ptr FMTID, :ptr CLSID, grfFlags:DWORD, dwReserved:DWORD, ppPropStg:ptr LPPROPERTYSTORAGE
protoStgOpenPropStg		typedef proto :LPUNKNOWN, :ptr FMTID, grfFlags:DWORD, dwReserved:DWORD, ppPropStg:ptr LPPROPERTYSTORAGE

LPPROPSTGNAMETOFMT	typedef ptr protoPropStgNameToFmtId
LPSTGCREATEPROPSTG	typedef ptr protoStgCreatePropStg
LPSTGOPENPROPSTG	typedef ptr protoStgOpenPropStg

;----------------------- publics/externals ---------------------

externdef pRegKeys:LPSTR

externdef g_szCLSID			:byte
externdef g_szTypeLib		:byte
externdef g_szInterface		:byte

externdef g_szRootCLSID		:byte
externdef g_szRootTypeLib	:byte
externdef g_szRootInterface	:byte
externdef g_szRootAppID		:byte
externdef g_szRootCompCat	:byte

externdef g_hInstance:HINSTANCE
externdef c _onexitbegin:LPVOID
externdef g_hCurAccel:HACCEL		;current accelerator
externdef g_hAccel2:HACCEL			;accelerator for dialogs
externdef g_hWndDlg:HWND			;current dialog window
externdef g_hMenu:HMENU
externdef g_hWndOption:HWND
externdef g_hIconApp:HICON
externdef g_pcdi:ptr CInterfaceList
externdef g_argc:dword
externdef g_argv:LPVOID
externdef g_himlLV:HANDLE
externdef g_heap:HANDLE
externdef g_szUserColCLSID:BYTE
externdef g_szUserColHKCR:BYTE
externdef g_szUserColInterface:BYTE
externdef g_hCsrWait:HCURSOR
externdef g_bLogActive:BOOLEAN
externdef g_LCID:LCID
externdef g_rectMain:RECT
externdef g_MaxCollItems:DWORD
externdef g_pMainDlg:ptr CDlg
externdef g_pViewObjectDlg:ptr CViewObjectDlg
externdef g_dwMyCBFormat:DWORD
externdef g_bAcceptDrop:BOOLEAN
externdef g_szHint:BYTE
externdef g_szWarning:BYTE
externdef g_szNull:BYTE
externdef g_bConfirmDelete:BOOLEAN
externdef g_pStream:LPSTREAM
externdef g_pStorage:LPSTORAGE
externdef g_dwFontWidth:DWORD
externdef g_pObjects:ptr CList
externdef g_dqNull:LARGE_INTEGER
externdef g_szCOMViewHttp:BYTE
externdef g_pDataObject:LPDATAOBJECT
externdef g_pfnStgCreatePropStg:LPSTGCREATEPROPSTG
externdef g_pfnStgOpenPropStg:LPSTGOPENPROPSTG
externdef g_aViewControlClass:WORD
externdef g_pszMenuHelp:LPSTR

externdef g_bCtxInProcHandler:BOOLEAN
externdef g_bCtxInProcServer:BOOLEAN
externdef g_bCtxLocalServer:BOOLEAN
externdef g_bCtxRemoteServer:BOOLEAN
externdef g_bUseQueryPath:BOOLEAN
externdef g_bCreateMaxDispHlp:BOOLEAN
externdef g_bOwnWndForPropDlg:BOOLEAN
externdef g_bDepTypeLibDlg:BOOLEAN
externdef g_bSyncTypeInfoAndProp:BOOLEAN
externdef g_bNewDlgForMethods:BOOLEAN
externdef g_bTypeFlagsAsNumber:BOOLEAN
externdef g_bObjDlgsAsTopLevelWnd:BOOLEAN
externdef g_bMemIdInDecimal:BOOLEAN
externdef g_bViewDlgAsTopLevelWnd:BOOLEAN
externdef g_bPropDlgAsTopLevelWnd:BOOLEAN
externdef g_bDispContainerCalls:BOOLEAN
externdef g_bDispQueryIFCalls:BOOLEAN
externdef g_bUserMode:BOOLEAN
externdef g_bUIDead:BOOLEAN
externdef g_bShowAllMembers:BOOLEAN
externdef g_bFreeLibs:BOOLEAN
externdef g_bExcludeProxy:BOOLEAN
externdef g_bExcludeTypeLib:BOOLEAN
externdef g_bQueryMI:BOOLEAN
externdef g_bDispUserCalls:BOOLEAN
externdef g_bTranslateUDTs:BOOLEAN
externdef g_bShowForceTypeInfo:BOOLEAN
externdef g_bNoDispatchPropScan:BOOLEAN
externdef g_bAllowWindowless:BOOLEAN
externdef g_bDrawIfNotActive:BOOLEAN
externdef g_bWriteClipBoard:BOOLEAN
externdef g_bUseClassFactory2:BOOLEAN
externdef g_bAddAutoTreatEntries:BOOLEAN
externdef g_bCollDlgAsTopLevelWnd:BOOLEAN
externdef g_bCloseCollDlgOnDlbClk:BOOLEAN
externdef g_bBindIsDefault:BOOLEAN
externdef g_bDocumentSiteSupp:BOOLEAN
externdef g_bOneInstance:BOOLEAN
externdef g_bCommandTargetSupp:BOOLEAN
externdef g_bInPlaceSiteExSupp:BOOLEAN
externdef g_bConfirmSaveReq:BOOLEAN
externdef g_bUseTypeInfoInvoke:BOOLEAN
externdef g_bDispatchSupp:BOOLEAN
externdef g_bMultiDoc:BOOLEAN
externdef g_bUseIPersistFile:BOOLEAN
externdef g_bTLibDlgAsTopLevelWnd:BOOLEAN
externdef g_bValueInDecimal:BOOLEAN
externdef g_bUseIPersistPropBag:BOOLEAN
externdef g_bUseIPersistStream:BOOLEAN
externdef g_bUseIQuickActivate:BOOLEAN
externdef g_bLogToDebugWnd:BOOLEAN
externdef g_bServiceProviderSupp:BOOLEAN
externdef g_bUseEnumCPs:BOOLEAN
externdef g_bGrayClrforIF:BOOLEAN
externdef g_bUseIPointerInactive:BOOLEAN

;----------------------- typedefs and prototypes ---------------------

PVOID		typedef ptr
PBYTE		typedef ptr byte
PSTR		typedef ptr sbyte
PPSTR		typedef ptr PSTR
PGUID		typedef ptr GUID
LCID		typedef dword		;is needed here
;HTREEITEM	typedef HANDLE
;DLGPROC		typedef proto :ptr, message:dword,wParam:dword,lParam:dword
;PDLGPROC	typedef ptr DLGPROC

ifndef NMLVDISPINFO
NMLVDISPINFO typedef LV_DISPINFO
endif

;--- structure for return VARIANT parameters 

PARAMRETURN struct
iCurItem		dd ?
iNumVariants	dd ?
pVariants		LPVARIANT ?
PARAMRETURN ends

LPPARAMRETURN typedef ptr PARAMRETURN

;--- general prototypes (functions in utility.asm)

;strchr				proto pStr1:LPSTR,dwByte:dword
_strcmpW			proto pszwStr1:ptr WORD, pszwStr2:ptr WORD

malloc				proto dwBytes:DWORD
free				proto pMem:ptr BYTE
printf				proto c :ptr BYTE, :VARARG
SetArguments		proto
String2Number		proto pStr:LPSTR, pDW:ptr dword, radix:dword
String2DWord		proto pStr:LPSTR, pDW:ptr dword
String22DWords		proto pStr:LPSTR, pDW1:ptr dword, pDW2:ptr dword
CenterWindow		proto hWnd:HWND
SetLVColumns		proto :HWND, numCols:dword,pColHdr:ptr CColHdr
GetChildPos			proto hWnd:HWND
ReadAllRegSubKeys	proto hWnd:HWND, pszKey:LPSTR, pNumEntries:ptr dword
DeleteKeyWithSubKeys proto hKey:HANDLE,pszKey:LPSTR
ResetListViewSelection proto hWndLV:HWND
GetTextFromCLSID    proto pGUID:ptr GUID,pStr:LPSTR,dwSize:dword
GUIDFromLPSTR		proto pszGUID:LPSTR, pGUID: ptr GUID
OutputMessage		proto hWnd:HWND,dwErr:dword,pszCap:LPSTR,pszPrefix:LPSTR
LVSort				proto hWndLV:HWND,iSortCol:dword,iSortDir:dword,dwFlags:dword
comparelvproc		proto lParam1:dword, lParam2:dword, plvc:ptr
GetTypekindStr		proto tk:dword
GetVarType			proto dwVarType:DWORD
GetVarTypeAsm		proto dwVarType:DWORD
GetParameterType	proto pTypeInfo:LPTYPEINFO, pTypeDesc:ptr TYPEDESC, pStr:LPSTR, iMax:dword
GetParameterTypeAsm	proto pTypeInfo:LPTYPEINFO, pTypeDesc:ptr TYPEDESC, pStr:LPSTR, iMax:dword
GetParameterTypeStub	proto pTypeInfo:LPTYPEINFO, pTypeDesc:ptr TYPEDESC, pStr:LPSTR, iMax:dword
GetParameterSize	proto pTypeInfo:LPTYPEINFO, pTypeDesc:ptr TYPEDESC
inputdlgproc		proto :HWND, :DWORD, :WPARAM, :LPARAM
CheckReservedNames	proto pStr:ptr byte
BroadCastMessage	proto uMsg:DWORD, :WPARAM, :LPARAM
RestoreAndActivateWindow proto hWnd:HWND
ParamReturnClear	proto :ptr PARAMRETURN
GetStdDispIdStr		proto dispid:sdword
GetArgument			proto pVariant:ptr VARIANT, pStrOut:LPSTR
CopyStringToClipboard proto hWnd:HWND, pszGUID:LPSTR
CopyFileToClipboard proto hWnd:HWND, pszFile:LPSTR
String2DWords		proto pStr:LPSTR, iNum:DWORD, pArray:ptr DWORD
GetItemPosition		proto hWnd:HWND, bMouse:BOOL, pPoint:ptr POINT
SetWindowIcon		proto :HWND, :ptr GUID
MakeUDColumnList	proto handle:HANDLE, mode:DWORD, dwIDStart:DWORD
SetSBParts			proto hWndSB:HWND, pParts:ptr DWORD, iNumParts:DWORD
GetAmbientDispId	proto dwIndex:DWORD, pDispId:ptr DWORD, pDispIdStr:ptr LPSTR
MyGetFileName		proto hWnd:HWND, pszFileName:LPSTR, iMax:DWORD, pszCustom:LPSTR, iMaxCustom:DWORD, iMode:DWORD, pszTitle:LPSTR
SysStringFromLPSTR	proto pszString:LPSTR, dwSize:DWORD
ComPtrAssign		proto pp:ptr LPUNKNOWN, lp:LPUNKNOWN
CheckUpdate			proto :HWND
InitIProp			proto
_cexit				proto c
atexit				proto c pExitProc:LPVOID
IsInterfaceSupported proto :REFIID,:ptr REFIID, :DWORD, :LPUNKNOWN, :ptr LPUNKNOWN
GetVarKind			proto :DWORD
GetFuncKind			proto :DWORD
GetInvokeKind		proto :DWORD
GetCallConv			proto callconv:DWORD
GetImplTypeFlags_	proto impltypeflags:DWORD, pStr:LPSTR, iMax:dword
GetTypeFlags		proto dwFlags:DWORD, pszOut:LPSTR
GetFuncFlags		proto dwFlags:DWORD, pszOut:LPSTR
GetVarFlags			proto dwFlags:DWORD, pszOut:LPSTR
GetParamFlags		proto dwFlags:DWORD, pszOut:LPSTR
SetHeaderBitmap		proto hWndLV:HWND, iColumn:DWORD, bSortDesc:BOOL
ResetHeaderBitmap	proto hWndLV:HWND
DisplayExceptionInfo proto hWnd:HWND, pEIP:ptr EXCEPTION_INFO_PTRS, pszCaption:LPSTR, :DWORD
SetImageListHdr		proto :HWND
DisplayStatusBarString proto hWnd:HWND, dwID:DWORD
RefreshObjectView	proto
SaveNormalWindowPos	proto hWnd:HWND, pRect:ptr RECT
MySetWindowPos		proto hWnd:HWND, pRect:ptr RECT
GetVariant			proto :ptr VARIANT, pStr:LPSTR, dwMax:dword, pszHexPrefix:LPSTR
SafeRelease			proto :LPUNKNOWN

GetDefaultInterfaceFromCoClass		proto :LPTYPEINFO, :BOOL
GetTypeInfoFromIProvideClassInfo	proto pUnknown:LPUNKNOWN, :BOOL
GetTypeInfoFromIDispatch			proto pUnknown:LPUNKNOWN
TypeInfoDlgFromIProvideClassInfo	proto :HWND, :LPUNKNOWN
TypeInfoDlgFromIDispatch			proto :HWND, :LPUNKNOWN

classdialogproc		proto hWnd:HWND,message:dword,wParam:WPARAM,lParam:LPARAM
helpdialogproc		proto hWnd:HWND,message:dword,wParam:WPARAM,lParam:LPARAM
aboutdialogproc		proto hWnd:HWND,message:dword,wParam:WPARAM,lParam:LPARAM

;--- interface viewer procs

externdef InterfaceViewerTab:DWORD

protoViewProc typedef proto :HWND, :LPUNKNOWN, :ptr CLSID
LPVIEWPROC	typedef ptr protoViewProc

;---------------------------- macros/equates -----------------------------

;--- submenu index of g_hMenu

ID_SUBMENU_OBJECTDLG		equ 0
ID_SUBMENU_TYPEINFODLG		equ 1
ID_SUBMENU_TYPELIBDLG		equ 2
ID_SUBMENU_NOTIFICATIONDLG	equ 3
ID_SUBMENU_PROPERTIESDLG	equ 4
ID_SUBMENU_OUT_OBJECTDLG	equ 5
ID_SUBMENU_COLLECTIONDLG	equ 6

FILEOP_PROPERTIES	equ 0
FILEOP_UNREGISTER	equ 1

FillMemory equ <RtlFillMemory>
ZeroMemory equ <RtlZeroMemory>
CopyMemory equ <RtlMoveMemory>

;--- @MakeIUnknownStubs: used for multiple inheritance
;--- in COM objects. adjust this_ ptr, then call standard IUnknown method

@MakeIUnknownStubs macro objofs, suffix
QueryInterface_&suffix::
if objofs
	sub dword ptr [esp+4],objofs
endif
	jmp QueryInterface
AddRef_&suffix::
if objofs
	sub dword ptr [esp+4],objofs
endif
	jmp AddRef
Release_&suffix::
if objofs
	sub dword ptr [esp+4],objofs
endif
	jmp Release
	endm

;------------- Debug Support

@SetLocalThis macro class
ifdef _DEBUG
local	this_:ptr class
		mov this_, __this
endif
		endm

ifdef _DEBUG
externdef DEBUGPREFIX:LPSTR
endif

;*** end of file
