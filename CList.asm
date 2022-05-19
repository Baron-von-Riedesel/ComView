
;*** classes CList ***

	.386
	.model flat,stdcall
	option casemap :none   ; case sensitive
	option proc:private

	include COMView.inc
INSIDE_CLIST equ 1
	include classes.inc
	include rsrc.inc

;-------------------------------------------------------

BEGIN_CLASS CListItem
pNext	pCListItem ?
pItem	LPVOID ?
lParam	LPARAM ?
END_CLASS

BEGIN_CLASS CList
pFirst	pCListItem ?
dwFlags	DWORD ?
END_CLASS

	.code

Create@CListItem proc pItem:LPVOID
		invoke malloc, sizeof CListItem
		.if (eax)
			mov [eax].CListItem.pNext, NULL
			mov ecx, pItem
			mov [eax].CListItem.pItem, ecx
		.endif
		ret
Create@CListItem endp

if 0
Destroy@CListItem proc this_:ptr CListItem
		invoke free, this_
		ret
Destroy@CListItem endp
else
Destroy@CListItem textequ <free>
endif


;-----------------------------------------------


AddItem@CList proc public uses esi this_:ptr CList, pItem:LPVOID

		mov ecx, this_
		.if ([ecx].CList.dwFlags & LISTF_STRINGS)
			invoke lstrlen, pItem
			inc eax
			invoke malloc, eax
			.if (eax)
				mov ecx, pItem
				mov pItem, eax
				invoke lstrcpy, eax, ecx
			.else
				jmp error
			.endif
		.endif
		invoke Create@CListItem, pItem
		.if (eax)
			xor esi, esi
			mov ecx, this_
			mov edx,[ecx].CList.pFirst
			.while (edx)
				mov ecx,edx
				mov edx,[edx].CListItem.pNext
				inc esi
			.endw
			mov [ecx].CListItem.pNext, eax
			mov eax, esi
		.else
error:
			mov eax, -1
		.endif
		ret

AddItem@CList endp



DeleteItem@CList proc public uses esi this_:ptr CList, dwIndex:DWORD

		mov eax, dwIndex
		mov ecx, this_
		mov esi, [ecx].CList.dwFlags
		mov edx,[ecx].CList.pFirst
		.while (edx && eax)
			mov ecx, edx
			mov edx,[edx].CListItem.pNext
			dec eax
		.endw
		xor eax, eax
		.if (edx)
			mov eax, [edx].CListItem.pNext
			mov [ecx].CListItem.pNext, eax
			.if (esi & LISTF_STRINGS)
				push edx
				invoke free, [edx].CListItem.pItem
				pop edx
			.endif
			invoke Destroy@CListItem, edx
			mov eax, 1
		.endif
		ret
DeleteItem@CList endp



GetItem@CList proc public this_:ptr CList, dwIndex:DWORD

		mov edx, this_
		mov ecx, dwIndex
		mov eax, [edx].CList.pFirst
		.while (ecx && eax)
			mov eax, [eax].CListItem.pNext
			dec ecx
		.endw
		.if (eax)
			mov eax, [eax].CListItem.pItem
		.endif
		ret

GetItem@CList endp



FindItem@CList proc public this_:ptr CList, pItem:LPVOID

		mov eax, pItem
		mov edx, this_
		mov edx, [edx].CList.pFirst
		xor ecx, ecx
		.while (edx && (eax != [edx].CListItem.pItem))
			mov edx,[edx].CListItem.pNext
			inc ecx
		.endw
		mov eax, -1
		.if (edx)
			mov eax, ecx
		.endif
		ret

FindItem@CList endp


;--- GetItemCount


GetItemCount@CList proc public this_:ptr CList
		xor eax,eax
		mov ecx,this_
		mov ecx, [ecx].CList.pFirst
		.while (ecx)
			inc eax
			mov ecx,[ecx].CListItem.pNext
		.endw
		ret
GetItemCount@CList endp

GetItemData@CList proc public this_:ptr CList, dwIndex:DWORD

		mov edx, this_
		mov edx, [edx].CList.pFirst
		mov ecx, dwIndex
		xor eax, eax
		.while (edx && ecx)
			mov edx,[edx].CListItem.pNext
			dec ecx
		.endw
		.if (edx)
			mov eax, [edx].CListItem.lParam
		.endif
		ret
GetItemData@CList endp

SetItemData@CList proc public this_:ptr CList, dwIndex:DWORD, lParam:LPARAM

		mov edx, this_
		mov edx, [edx].CList.pFirst
		mov ecx, dwIndex
		xor eax, eax
		.while (edx && ecx)
			mov edx,[edx].CListItem.pNext
			dec ecx
		.endw
		.if (edx)
			mov eax, lParam
			xchg eax, [edx].CListItem.lParam
		.endif
		ret

SetItemData@CList endp

FindItemData@CList proc public this_:ptr CList, lParam:LPARAM

		mov eax, lParam
		mov edx, this_
		mov edx, [edx].CList.pFirst
		xor ecx, ecx
		.while (edx && (eax != [edx].CListItem.lParam))
			mov edx,[edx].CListItem.pNext
			inc ecx
		.endw
		mov eax,-1
		.if (edx)
			mov eax, ecx
		.endif
		ret
FindItemData@CList endp

FindAllItemData@CList proc public this_:ptr CList, lParam:LPARAM

		mov ecx, lParam
		mov edx, this_
		mov edx, [edx].CList.pFirst
		xor eax, eax
		.while (edx)
			.if (ecx == [edx].CListItem.lParam)
				inc eax
			.endif
			mov edx,[edx].CListItem.pNext
		.endw
		ret

FindAllItemData@CList endp

;--- Destroy


Destroy@CList proc public uses esi this_:ptr CList

		.repeat
			invoke DeleteItem@CList, this_, 0
		.until (!eax)
		invoke free, this_
		ret

Destroy@CList endp


;--- Create a linked list


Create@CList proc public dwFlags:DWORD

		invoke malloc, sizeof CList
		.if (eax)
			mov [eax].CList.pFirst, NULL
			mov ecx, dwFlags
			mov [eax].CList.dwFlags, ecx
		.endif
		ret

Create@CList endp

	end
