
;*** definition of COptions methods

	.386
	.model flat,stdcall
	option casemap :none   ; case sensitive
	option proc:private

	include COMView.inc
INSIDE_COPTIONS equ 1
	include classes.inc
	include rsrc.inc
	include debugout.inc


	.data

g_LCID					LCID LOCALE_SYSTEM_DEFAULT
g_bConfirmDelete		BOOLEAN TRUE	;confirm delete requests required (general) 

;--- the following options are saved to a profile file
;--- place new BOOLEAN options ALWAYS just before BoolOptionEnd label!

OptionTableStart label byte
BoolOptionStart label byte
g_bUseQueryPath			BOOLEAN FALSE	;use QueryPathOfRegLib in typelib dlgs
g_bCreateMaxDispHlp		BOOLEAN FALSE	;create max. dispatch helper includes
g_bOwnWndForPropDlg		BOOLEAN TRUE	;open each properties in its own window
g_bDepTypeLibDlg		BOOLEAN FALSE	;close typelib dlg with properties
g_bSyncTypeInfoAndProp	BOOLEAN FALSE	;show sel. member from "properties" in "typeinfo" dlg
g_bNewDlgForMethods		BOOLEAN FALSE	;open "properties" for methods returning IDispatch
g_bTypeFlagsAsNumber	BOOLEAN FALSE	;display type flags as numbers
g_bObjDlgsAsTopLevelWnd	BOOLEAN FALSE	;"objects" as top level window
g_bCtxInProcHandler		BOOLEAN TRUE	;flag INPROCHANDLER for CoCreateInstance
g_bCtxInProcServer		BOOLEAN TRUE	;flag INPROCSERVER for CoCreateInstance
g_bCtxLocalServer		BOOLEAN TRUE	;flag LOCALSERVER for CoCreateInstance
g_bCtxRemoteServer		BOOLEAN FALSE	;flag REMOTE for CoCreateInstance
g_bMemIdInDecimal		BOOLEAN FALSE	;show MEMBERID in decimal
g_bViewDlgAsTopLevelWnd	BOOLEAN FALSE	;"view control" as top level
g_bPropDlgAsTopLevelWnd	BOOLEAN FALSE	;"properties" as top level
g_bDispContainerCalls	BOOLEAN FALSE	;display container methods being called
g_bDispQueryIFCalls		BOOLEAN FALSE	;display QueryInterface being called in "log window"
g_bUserMode				BOOLEAN FALSE	;ambient property user mode/design mode
g_bShowAllMembers		BOOLEAN FALSE	;show restricted members as well
g_bFreeLibs				BOOLEAN FALSE	;call CoFreeUnusedLibraries
g_bExcludeProxy			BOOLEAN FALSE	;exclude proxystubclsid32 from edit (interface)
g_bExcludeTypeLib		BOOLEAN FALSE	;exclude typelib GUID from edit (interface)
g_bQueryMI				BOOLEAN TRUE	;use IQueryMI if available (usually for remote objects)
g_bDispUserCalls		BOOLEAN FALSE	;display user initiated calls in "properties" dialog
g_bTranslateUDTs		BOOLEAN TRUE	;translate UDTs in scan mode
g_bShowForceTypeInfo	BOOLEAN FALSE	;show menu item "Set TypeInfo"
g_bNoDispatchPropScan	BOOLEAN TRUE	;dont scan "IDispatch" properties
g_bAllowWindowless		BOOLEAN TRUE	;alloc windowless activation
g_bDrawIfNotActive		BOOLEAN FALSE	;use ViewObject to draw object if inactive
g_bWriteClipBoard		BOOLEAN FALSE	;write ASM includes in clipboard
g_bUseClassFactory2		BOOLEAN FALSE	;use CoGetClassObject + IClassFactory2
g_bAddAutoTreatEntries	BOOLEAN FALSE	;add AutoConvertTo + TreatAs CLSID entries
g_bCollDlgAsTopLevelWnd	BOOLEAN FALSE	;collection dialog as top level
g_bCloseCollDlgOnDlbClk	BOOLEAN TRUE	;close collection dialog after dblclk
g_bBindIsDefault		BOOLEAN TRUE	;Bind to Object is default
g_bDocumentSiteSupp		BOOLEAN TRUE	;support IOleDocumentSite
g_bOneInstance			BOOLEAN TRUE	;allow 1 instance of COMView only
g_bCommandTargetSupp	BOOLEAN TRUE	;support IOleCommandTarget
g_bInPlaceSiteExSupp	BOOLEAN TRUE	;support IOleInPlaceSiteEx
g_bConfirmSaveReq		BOOLEAN FALSE	;confirm control's save requests
g_bUseTypeInfoInvoke	BOOLEAN FALSE	;use ITypeInfo::Invoke for TKIND_DISPATCH
g_bDispatchSupp			BOOLEAN TRUE	;support IDispatch
g_bMultiDoc				BOOLEAN TRUE	;hold multiple documents
g_bUseIPersistFile		BOOLEAN FALSE	;use IPersistFile::Load for initialization
g_bTLibDlgAsTopLevelWnd	BOOLEAN FALSE	;typelib dialog as top level
g_bValueInDecimal		BOOLEAN TRUE	;show value in decimal (typeinfo/variables)
g_bUseIPersistPropBag	BOOLEAN TRUE	;use IPersistPropertyBag interface
g_bUseIPersistStream	BOOLEAN TRUE	;use IPersistStreamInit interface
g_bUseIQuickActivate	BOOLEAN FALSE	;use IQuickActivate interface
g_bLogToDebugWnd		BOOLEAN FALSE	;redirect log to debug terminal
g_bServiceProviderSupp	BOOLEAN TRUE	;support IServiceProvider
g_bUseEnumCPs			BOOLEAN TRUE	;use IEnumConnectionPoints
g_bGrayClrforIF			BOOLEAN TRUE	;display interfaces in gray with no typeinfo
g_bUseIPointerInactive	BOOLEAN FALSE	;use IPointerInactive interface
BoolOptionEnd label byte

g_szUserColCLSID db 64 dup (0)		;1 user defined column in CLSID view
g_szUserColHKCR  db 64 dup (0)		;1 user defined column in HKCR view
g_szUserColInterface db 64 dup (0)	;1 user defined column in Interface view
g_rectMain	RECT {}
g_MaxCollItems	DWORD 2000
g_dwFontWidth	DWORD 8

OptionTableEnd label byte
OptionTableSize$ textequ %( OptionTableEnd - OptionTableStart )

SavedOptions db OptionTableEnd - OptionTableStart dup (?)

	.const

;--- sections

szOptions			db "Options",0
szInterfaces		db "Interfaces", 0

;--- key names

szBooleans			db "Booleans", 0
szUserColClsid		db "UserColClsid", 0
szUserColHKCR		db "UserColHKCR", 0
szUserColInterface	db "UserColInterface", 0
szMainWnd			db "MainWnd", 0
szMaxCollItems		db "MaxCollItems", 0
szFixedFontWidth	db "FixedFontWidth", 0

	.code


;--- static method Load()


Load@COptions proc public uses esi edi pszFile:LPSTR

local	szBuffer[256]:byte

		mov szBuffer, 0
		invoke GetPrivateProfileString, addr szOptions, addr szBooleans,
			CStr(""), addr szBuffer, sizeof szBuffer, pszFile
		lea esi, szBuffer
		mov edi, offset BoolOptionStart
		.while ((byte ptr [esi]) && (edi < offset BoolOptionEnd))
			lodsb
			.if ((al >= '0') && (al <= '1'))
				sub al, '0'
				mov [edi],al
			.elseif (al == ',')
				inc edi
			.endif
		.endw
		invoke GetPrivateProfileString, addr szOptions, addr szUserColClsid,
			CStr(""), addr g_szUserColCLSID, sizeof g_szUserColCLSID, pszFile
		invoke GetPrivateProfileString, addr szOptions, addr szUserColHKCR,
			CStr(""), addr g_szUserColHKCR, sizeof g_szUserColHKCR, pszFile
		invoke GetPrivateProfileString, addr szOptions, addr szUserColInterface,
			CStr(""), addr g_szUserColInterface, sizeof g_szUserColInterface, pszFile
		invoke GetPrivateProfileString, addr szOptions, addr szMainWnd,
			CStr(""), addr szBuffer, sizeof szBuffer, pszFile
		.if (szBuffer)
			invoke String2DWords, addr szBuffer, 4, addr g_rectMain
		.endif
		invoke GetPrivateProfileInt, addr szOptions, addr szMaxCollItems,
			g_MaxCollItems, pszFile
		mov g_MaxCollItems, eax
		invoke GetPrivateProfileInt, addr szOptions, addr szFixedFontWidth,
			g_dwFontWidth, pszFile
		mov g_dwFontWidth, eax
		invoke CopyMemory, addr SavedOptions, addr OptionTableStart, OptionTableSize$
		invoke UpdateLogSwitch@CLogWindow
		ret
		align 4
Load@COptions endp

;--- static method Load()

Save@COptions proc public uses esi edi pszFile:LPSTR

local	szBuffer[128]:byte

		mov esi, offset OptionTableStart
		mov edi, offset SavedOptions
		mov ecx, OptionTableEnd - OptionTableStart
		repz cmpsb
		.if (!ZERO?)
			lea edi, szBuffer
			mov esi, offset BoolOptionStart
			.while (esi < offset BoolOptionEnd)
				lodsb
				movzx eax, al
				invoke wsprintf, edi, CStr("%u,"), eax
				add edi, eax
			.endw
			mov byte ptr [edi-1],0

			invoke WritePrivateProfileString, addr szOptions, addr szBooleans,
				addr szBuffer, pszFile
			invoke WritePrivateProfileString, addr szOptions, addr szUserColClsid,
				addr g_szUserColCLSID, pszFile
			invoke WritePrivateProfileString, addr szOptions, addr szUserColHKCR,
				addr g_szUserColHKCR, pszFile
			invoke WritePrivateProfileString, addr szOptions, addr szUserColInterface,
				addr g_szUserColInterface, pszFile
			invoke wsprintf, addr szBuffer, CStr("%u,%u,%u,%u"), g_rectMain.left, g_rectMain.top, g_rectMain.right, g_rectMain.bottom
			invoke WritePrivateProfileString, addr szOptions, addr szMainWnd,
				addr szBuffer, pszFile
			invoke wsprintf, addr szBuffer, CStr("%u"), g_MaxCollItems
			invoke WritePrivateProfileString, addr szOptions, addr szMaxCollItems,
				addr szBuffer, pszFile
			invoke wsprintf, addr szBuffer, CStr("%u"), g_dwFontWidth
			invoke WritePrivateProfileString, addr szOptions, addr szFixedFontWidth,
				addr szBuffer, pszFile
		.endif

		ret
		align 4
Save@COptions endp

GetCoCreateFlags@COptions proc public
		xor eax, eax
		.if (g_bCtxInProcHandler)
			or eax, CLSCTX_INPROC_HANDLER
		.endif
		.if (g_bCtxInProcServer)
			or eax, CLSCTX_INPROC_SERVER
		.endif
		.if (g_bCtxLocalServer)
			or eax, CLSCTX_LOCAL_SERVER
		.endif
		.if (g_bCtxRemoteServer)
			or eax, CLSCTX_REMOTE_SERVER
		.endif
		ret
		align 4
GetCoCreateFlags@COptions endp

ifdef @StackBase
	option stackbase:ebp
endif

GetInterfaces@COptions proc public uses esi edi pIL:ptr CInterfaceList

local pBuffer:LPSTR
local szModule[MAX_PATH]:byte
local dwRest:DWORD
local dwCount:DWORD
local pszNames:LPSTR
local iid:IID
local wszIID[40]:WORD
local szKey[40]:byte
local szValue[128]:byte

		invoke malloc, 8000h
		.if (eax)
			mov pBuffer, eax
			mov esi, eax
			mov dwCount, 0
			mov pszNames, NULL
			invoke GetModuleFileName, NULL, addr szModule,  sizeof szModule
			lea ecx, szModule
			mov dword ptr [ecx+eax-3],"ini"
			invoke GetPrivateProfileSection, addr szInterfaces, esi, 7FFFh, addr szModule
			.if (eax)
				inc eax
				mov dwRest, eax
				invoke malloc, eax
				mov pszNames, eax
			.endif
			.while (byte ptr [esi])
				lea edi, szKey
				mov byte ptr [edi],0
				mov ecx, sizeof szKey - 1
				lodsb
				.while (ecx && al && (al != '='))
					stosb
					dec ecx
					lodsb
				.endw
				.break .if (!al)
				xor al, al
				stosb
				lea edi, szValue
				mov ecx, sizeof szValue - 1
				lodsb
				.while (ecx && al)
					stosb
					dec ecx
					lodsb
				.endw
				xor al, al
				stosb
				invoke MultiByteToWideChar, CP_ACP, MB_PRECOMPOSED,
						addr szKey, -1, addr wszIID, 40 
				sub esp, sizeof iid
				invoke IIDFromString,addr wszIID, esp
				.if (eax == S_OK)
					DebugOut "GetInterfaces: added IID=%s", addr szKey
					invoke Find@CInterfaceList, pIL, addr iid
					.if (!eax)
						invoke lstrlen, addr szValue
						inc eax
						sub dwRest, eax
						mov ecx, pszNames
						add ecx, dwRest
						invoke lstrcpy, ecx, addr szValue
						inc dwCount
						sub esp, sizeof iid
					.endif
				.endif
				add esp, sizeof iid
			.endw
			.if (dwCount)
				mov ecx, esp
				mov eax, pszNames
				add eax, dwRest
				invoke AddIIDs@CInterfaceList, pIL, dwCount, ecx, eax
				mov eax, dwCount
				shl eax, 4
				add esp, eax
			.endif
			invoke free, pBuffer
			invoke free, pszNames
		.endif
		ret
		align 4

GetInterfaces@COptions endp

ifdef @StackBase
	option stackbase:esp
endif

	end
