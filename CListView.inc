
;--- new messages

LVM_ENDEDITMODE		equ WM_USER+100		;message for listview to end edit mode
LVM_COMBOBOXMODE	equ WM_USER+101		;start combobox mode
LVM_GETCOMBOBOXCONTROL	equ WM_USER+102	;get combobox control

;--- adds edit features to a listview window

CreateEditListView	proto hWnd:HWND
