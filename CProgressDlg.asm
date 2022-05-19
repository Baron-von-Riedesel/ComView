
;*** definition of CProgressDlg methods

	.386
	.model flat,stdcall
	option casemap :none   ; case sensitive
	option proc:private

INSIDE_CPROGRESSDLG equ 1
	include COMView.inc
	include classes.inc
	include rsrc.inc

BEGIN_CLASS CProgressDlg, CDlg
hWndLV		HWND ?
hWndProgress HWND ?
iNumCols	DWORD ?
hThread		HANDLE ?
dwThreadId	DWORD ?
pszFileName	LPSTR ?
bCancel		BOOLEAN ?
iSaveMode	BYTE ?
END_CLASS

__this	textequ <edi>
_this	textequ <[__this].CProgressDlg>
thisarg textequ <this@:ptr CProgressDlg>

Destroy@CProgressDlg proto thisarg

	MEMBER hWnd, pDlgProc
	MEMBER hWndLV, hWndProgress, hThread, dwThreadId, iNumCols
	MEMBER pszFileName, bCancel, iSaveMode

BLOCKLEN equ 8000h

	.data

	.code

SaveListView proc uses ebx __this thisarg

	local iNextItem:dword
	local hFile:HANDLE
	local dwSize:dword
	local hGlobal:HGLOBAL
	local hNewGlobal:HGLOBAL
	local iLines:dword
	local iLinesAll:dword
	local iNewPos:dword
	local iOldPos:dword
	local j:dword
	local iSize:dword
	local lvi:LVITEM
	local hList:HANDLE
	local pGlobal:ptr byte
	local iReadType:dword
	local pszLine:ptr byte
	local dwWritten:dword
	local lvc:LVCOLUMN
	local szLine[1024]:byte
	local szStr[260]:byte


		xor eax,eax
		mov iOldPos,eax
		mov j,eax
		mov iSize,eax
		mov hList,eax
		mov pszLine,eax

		mov __this,this@

		.if (m_iSaveMode == SAVE_CLIPBOARD)
			invoke ListView_GetSelectedCount( m_hWndLV)
			mov iLinesAll,eax

			mov iReadType,LVNI_SELECTED
			mov dwSize,BLOCKLEN
			invoke GlobalAlloc,GMEM_MOVEABLE or GMEM_DDESHARE,dwSize
			mov hGlobal,eax
			.if (eax != 0)
				invoke GlobalLock,hGlobal
				mov pGlobal,eax
			.else
				invoke MessageBox,m_hWnd,CStr("memory error"),0,MB_OK
				mov m_bCancel,TRUE
			.endif
			invoke lstrcpy,addr szStr,CStr("copy to clipboard")
		.else
			invoke ListView_GetItemCount( m_hWndLV)
			mov iLinesAll,eax
			mov iReadType,LVNI_ALL
			invoke CreateFile, m_pszFileName, GENERIC_WRITE, 0, NULL,
					CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL 
			mov hFile,eax
			.if (eax == INVALID_HANDLE_VALUE)
				invoke wsprintf,addr szLine,CStr("file",0dh,0ah,"%s",0dh,0ah,"cannot be opened"),m_pszFileName
				invoke MessageBox,m_hWnd,addr szLine,0,MB_OK
				mov m_bCancel,TRUE
			.endif	
			invoke malloc,BLOCKLEN
			mov pGlobal,eax
			.if (eax == 0)
				invoke MessageBox,m_hWnd,CStr("memory error"),0,MB_OK
				mov m_bCancel,TRUE
			.endif
			invoke lstrcpy,addr szStr,CStr("copy to disk")
		.endif

		invoke SetWindowText,m_hWnd,addr szStr

		lea eax, szLine
		mov lvc.pszText, eax
		mov lvc.cchTextMax, sizeof szLine
		mov lvc.mask_, LVCF_TEXT
		@mov ebx, 0
		.while (ebx < m_iNumCols)
			invoke ListView_GetColumn( m_hWndLV, ebx, addr lvc)
			invoke lstrlen, lvc.pszText
			mov ecx, lvc.pszText
			mov byte ptr [ecx+eax],9
			inc eax
			add lvc.pszText, eax
			sub lvc.cchTextMax, eax
			inc ebx
		.endw
		mov ecx, lvc.pszText
		mov eax,000A0Dh
		mov [ecx-1], eax
		inc ecx
		lea eax, szLine
		sub ecx, eax
		mov iSize, ecx
		invoke CopyMemory, pGlobal, eax, ecx

		mov iLines, 1

		invoke ListView_GetNextItem( m_hWndLV,-1,iReadType)
		.while ((eax != -1) && (m_bCancel == FALSE))
			mov iNextItem,eax
			mov lvi.iItem,eax
			lea eax,szLine
			mov byte ptr [eax],0
			mov pszLine,eax
			@mov ebx,0
			.while (ebx < m_iNumCols)
				mov lvi.iSubItem,ebx
				mov lvi.mask_,LVIF_TEXT
				mov eax,pszLine
				mov lvi.pszText,eax
				lea ecx,szLine
				sub eax,ecx
				mov ecx,sizeof szLine
				sub ecx,eax
				mov lvi.cchTextMax,ecx
				invoke ListView_GetItem( m_hWndLV,addr lvi)
				invoke lstrlen,pszLine
				mov j,eax
				add pszLine,eax
				mov ecx, pszLine
				mov byte ptr [ecx],9
				inc pszLine
				inc ebx
			.endw
			mov eax,000A0Dh
			mov ecx,pszLine
			mov dword ptr [ecx-1],eax
			inc ecx
			mov pszLine,ecx

			lea eax,szLine
			sub ecx,eax
			mov j,ecx	
			mov ecx,pGlobal
			add ecx,iSize
			invoke CopyMemory,ecx,addr szLine,j
			mov eax,j
			add iSize,eax
			.if (m_iSaveMode == SAVE_DISK)
				mov eax,BLOCKLEN
				sub eax,sizeof szLine
				.if (iSize > eax)
					invoke WriteFile,hFile,pGlobal,iSize,addr dwWritten,NULL
					mov ecx,dwWritten
					.if ((eax == 0) || (ecx != iSize))
						invoke MessageBeep,MB_OK
						invoke MessageBox,m_hWnd,CStr("error writing file"),0,MB_OK
						mov m_bCancel,TRUE
					.endif
					mov iSize,0
				.endif
			.else
				mov eax,dwSize
				sub eax,sizeof szLine
				.if (iSize > eax)
					invoke GlobalUnlock,hGlobal
					add dwSize,BLOCKLEN
					invoke GlobalReAlloc,hGlobal,dwSize,GMEM_MOVEABLE
					mov hNewGlobal,eax
					.if (!eax)
						invoke MessageBeep,MB_OK
						invoke MessageBox,m_hWnd,CStr("memory error"),0,MB_OK
						mov m_bCancel,TRUE
					.endif
					mov eax,hNewGlobal
					mov hGlobal,eax
					invoke GlobalLock,hGlobal
					mov pGlobal,eax
				.endif
			.endif

			inc iLines
			.if (m_hWndProgress != 0)
				mov eax,iLines
				mov ecx,100
				mul ecx
				mov ecx,iLinesAll
				div ecx
				mov iNewPos,eax
				.if (iNewPos > 100)
					mov iNewPos,100
				.endif
				mov eax,iOldPos
				add eax,3
				.if (iNewPos >= eax)
					 invoke SendMessage,m_hWndProgress, PBM_SETPOS, iNewPos, 0
					 mov iOldPos,eax
				.endif
			.endif
			invoke ListView_GetNextItem( m_hWndLV,iNextItem,iReadType)
		.endw

		.if (m_iSaveMode == SAVE_DISK)
			.if (m_bCancel == FALSE)
				.if (iSize > 0)
					invoke WriteFile,hFile,pGlobal,iSize,addr dwWritten,NULL
				.endif
				invoke free,pGlobal
				invoke CloseHandle,hFile
				.if (m_hWndProgress != 0)
					invoke SendMessage,m_hWndProgress,PBM_SETPOS, 100, 0
				.endif
				invoke wsprintf,addr szLine,CStr("file %s written"),m_pszFileName
				invoke MessageBox,m_hWnd, addr szLine, addr g_szHint, MB_OK
			.else
				invoke free,pGlobal
				invoke CloseHandle,hFile
				invoke DeleteFile,m_pszFileName
			.endif
		.elseif (m_iSaveMode == SAVE_CLIPBOARD)
			invoke GlobalUnlock,hGlobal
			.if (m_bCancel == FALSE)
				mov eax,pGlobal
				mov ecx,iSize
				mov byte ptr [eax+ecx],0
				.if (m_hWndProgress)
					invoke SendMessage,m_hWndProgress, PBM_SETPOS, 100, 0
				.endif
				invoke OpenClipboard,m_hWnd
				invoke EmptyClipboard
				invoke SetClipboardData,CF_TEXT,hGlobal
				invoke CloseClipboard
			.else
				invoke GlobalFree,hGlobal
			.endif
		.endif
		invoke PostMessage,m_hWnd,WM_CLOSE,0,0

		ret
		align 4

SaveListView endp


;*** Dialog Proc 


ProgressDialog proc uses __this thisarg, message:dword,wParam:dword,lParam:dword

		mov __this,this@

		mov eax,message
		.if (eax == WM_INITDIALOG)
			invoke GetDlgItem,m_hWnd,IDC_PROGRESS1
			mov m_hWndProgress,eax
			invoke ListView_GetSelectedCount( m_hWndLV)
			.if (eax > 24 || m_iSaveMode == SAVE_DISK)
				invoke CreateThread,0,1000h,SaveListView,__this,0,addr m_dwThreadId
				mov m_hThread,eax
				invoke CenterWindow,m_hWnd
			.else
				invoke SaveListView, __this
			.endif
			mov eax,1
		.elseif (eax == WM_CLOSE)

			invoke EndDialog, m_hWnd, 0
			mov eax,1

		.elseif (eax == WM_DESTROY)

			invoke Destroy@CProgressDlg, __this

		.elseif (eax == WM_COMMAND)
			mov eax,wParam
			.if (ax == IDCANCEL)
				mov m_bCancel,TRUE
			.endif
		.else
			xor eax,eax ;indicates "no processing"
		.endif
		ret
		align 4

ProgressDialog endp


Destroy@CProgressDlg proc uses __this thisarg

		mov __this,this@
		invoke free, __this
		ret
		align 4

Destroy@CProgressDlg endp


;*** constructor


Create@CProgressDlg proc public uses __this hWndLV:HWND, pszFileName:LPSTR, iSaveMode:dword, iNumCols:DWORD

		invoke malloc, sizeof CProgressDlg
		.if (!eax)
			ret
		.endif
		mov __this,eax

		mov eax,hWndLV
		mov m_hWndLV,eax
		mov eax,pszFileName
		mov m_pszFileName,eax
		mov	eax,iSaveMode
		mov m_iSaveMode,al
		mov eax,iNumCols
		.if (eax == -1)
			invoke ListView_GetHeader( hWndLV)
			invoke Header_GetItemCount( eax)
		.endif
		mov m_iNumCols,eax
		mov m_pDlgProc,ProgressDialog
		mov m_bCancel,FALSE
		return __this
		align 4

Create@CProgressDlg endp

	end
