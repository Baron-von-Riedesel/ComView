
# this will create COMView.exe
# to create a debug version use "nmake debug=1"
# output will be in subdir RELEASE or DEBUG, format
# of object modules is COFF format, true flat

!ifndef DEBUG
DEBUG=0
!endif

NAME=COMView

!ifndef MASM
MASM=0
MSLINK=0
!else
MSLINK=1
!endif


!if $(DEBUG)
AOPTD=-D_DEBUG -Zi
OUTDIR=DEBUG
LOPTDM=/DEBUG
LOPTDW=debug watcom
!else
AOPTD=
OUTDIR=RELEASE
LOPTDM=
LOPTDW=
!endif

WININC=\WinInc

SRCMODS = \
!include modules.inc

OBJMODS = $(SRCMODS:.asm=.obj)
!if $(DEBUG)
OBJMODS = $(OBJMODS:.\=DEBUG\)
RMODSX = $(RMODS:.\=DEBUG\)
!else
OBJMODS = $(OBJMODS:.\=RELEASE\)
RMODSX = $(RMODS:.\=RELEASE\)
!endif

AOPT=-c -coff -nologo -Sg -Fl$* -Fo$* $(AOPTD) -I$(WININC)\Include

!if $(MASM)
ASM=ml.exe $(AOPT) 
!else
ASM=jwasm.exe $(AOPT)
!endif

LIBS=kernel32.lib advapi32.lib gdi32.lib user32.lib ole32.lib oleaut32.lib uuid.lib shell32.lib comctl32.lib comdlg32.lib
LIBPATH=$(WININC)\lib

!if $(MSLINK)
LOPTS= /NOLOGO /MAP:$*.map /SUBSYSTEM:WINDOWS /OUT:$*.exe /LIBPATH:$(LIBPATH) /LIBPATH:.\Lib
LINK=link.exe 
LINKPARMS=$(OBJMODS) $*.res $(LOPTS) $(LOPTDM) $(LIBS)
#LINK=polink.exe /FORCE:MULTIPLE
!else
LOPTS= op MAP=$*.map
LINK=jwlink.exe 
LINKPARMS=$(LOPTDW) format win nt ru win file {$(OBJMODS)} name $*.exe op res=$*.res $(LOPTS) LibPath $(LIBPATH) LibPath .\Lib Lib { $(LIBS) }
!endif
HHC="\HTMLHelp\hhc.exe"

.SUFFIXES: .asm .obj .rc

.asm{$(OUTDIR)}.obj:
	@$(ASM) $<

ALL: $(OUTDIR) $(OUTDIR)\$(NAME).exe $(OUTDIR)\$(NAME).chm

$(OUTDIR):
	@mkdir $(OUTDIR)

$(OUTDIR)\$(NAME).exe: $(OBJMODS) Makefile $*.res
	$(LINK) @<<
$(LINKPARMS)
<<

$(OBJMODS): COMView.inc Classes.inc

$(OUTDIR)\$(NAME).res: $(NAME).rc
	rc /fo $*.res /i$(WININC)\include $(NAME).rc

$(OUTDIR)\$(NAME).chm: Makefile HELP\*.htm HELP\*.gif
	cd HELP
	$(HHC) $(NAME).hhp
	copy $(NAME).chm ..\$(OUTDIR)\*.*
	cd ..

clean:
	erase $(OUTDIR)\*.obj
	erase $(OUTDIR)\*.lst
	erase $(OUTDIR)\*.map
	erase $(OUTDIR)\*.res
	erase $(OUTDIR)\*.exe
