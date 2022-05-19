
;DEFBCKCOLOR				equ 00C0F0F0h
DEFADRTXTCOLOR			equ 00800000h	;address area text color
;DEFDTATXTCOLOR			equ 00000000h	;data area text color
DEFASCTXTCOLOR			equ 00008000h	;ascii area text color
;DEFSELBCKCOLOR			equ 00800000h	;selection area bgcolor
;DEFSELTXTCOLOR			equ 00FFFFFFh	;selection area text color
DEFSELBARCOLOR			equ 00C0C0C0h
DEFSELBARPEN			equ 00808080h
DEFLNRCOLOR				equ 00800000h

HEM_BASE				equ WM_USER+1000

;Private messages
HEM_SETFONT				equ HEM_BASE+1		;wParam=nLineSpacing, lParam=lpRAFONT
HEM_GETFONT				equ HEM_BASE+2		;wParam=0, lParam=lpRAFONT
HEM_SETCOLOR			equ HEM_BASE+3		;wParam=0, lParam=lpRACOLOR
HEM_GETCOLOR			equ HEM_BASE+4		;wParam=0, lParam=lpRACOLOR
HEM_VCENTER				equ HEM_BASE+5		;wParam=0, lParam=0
HEM_REPAINT				equ HEM_BASE+6		;wParam=0, lParam=0
HEM_ANYBOOKMARKS		equ HEM_BASE+7		;wParam=0, lParam=0
HEM_TOGGLEBOOKMARK		equ HEM_BASE+8		;wParam=nLine, lParam=0
HEM_CLEARBOOKMARKS		equ HEM_BASE+9		;wParam=0, lParam=0
HEM_NEXTBOOKMARK		equ HEM_BASE+10		;wParam=0, lParam=0
HEM_PREVIOUSBOOKMARK	equ HEM_BASE+11		;wParam=0, lParam=0
HEM_SELBARWIDTH			equ HEM_BASE+12		;wParam=nWidth, lParam=0
HEM_LINENUMBERWIDTH		equ HEM_BASE+13		;wParam=nWidth, lParam=0
HEM_GETSIZE				equ HEM_BASE+14		;wParam=0, lParam=0
HEM_RESET				equ HEM_BASE+15		;wParam=0, lParam=0

FR_HEX					equ 2

STYLE_NOSPLITT			equ 001h			;No splitt button
STYLE_NOLINENUMBER		equ 002h			;No linenumber button
STYLE_NOHSCROLL			equ 004h			;No horizontal scrollbar
STYLE_NOVSCROLL			equ 008h			;No vertical scrollbar
STYLE_NOSIZEGRIP		equ 010h			;No size grip
STYLE_NOSTATE			equ 020h			;No state indicator
STYLE_NOADDRESS			equ 040h			;No adress field
STYLE_NOASCII			equ 080h			;No ascii field
STYLE_NOUPPERCASE		equ 100h			;Hex numbers is lowercase letters
STYLE_READONLY			equ 200h			;Text is locked

HEFONT struct
	hFont		dd ?		;Code edit normal
	hLnrFont	dd ?		;Line numbers
HEFONT ends

HECOLOR struct
	bckcol		dd ?		;Back color
	adrtxtcol	dd ?		;Text color
	dtatxtcol	dd ?		;Text color
	asctxtcol	dd ?		;Text color
	selbckcol	dd ?		;Sel back color
	seltxtcol	dd ?		;Sel text color
	selbarbck	dd ?		;Selection bar
	selbarpen	dd ?		;Selection bar pen
	lnrcol		dd ?		;Line numbers color
HECOLOR ends

HESELCHANGE struct
	nmhdr		NMHDR <?>
	chrg		CHARRANGE <?>
	seltyp		dw ?		;SEL_TEXT or SEL_OBJECT
	line		dd ?		;Line number
	nlines		dd ?		;Total number of lines
	fchanged	dd ?		;TRUE if changed since last
HESELCHANGE ends

HEBMK struct
	hWin		dd ?		;Handle of window having the bookmark
	nLine		dd ?		;Bokkmarked line
HEBMK ends

Init@CHexEdit	proto :HINSTANCE
Deinit@CHexEdit	proto

