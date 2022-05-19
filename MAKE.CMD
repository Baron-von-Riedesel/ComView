@echo off
rem jwasm.exe -c -coff -nologo -Fl=Release\ -I\WinInc\Include -Fo Release\ -Fi=fpo.inc *.asm
jwasm.exe -c -coff -nologo -Sg -I\WinInc\Include -Fo Release\ *.asm
rem ml.exe -c -coff -nologo -I\WinInc\Include -Fo Release\ *.asm
rc -i \WinInc\Include -fo Release\COMView.res COMView.rc
cd Release
rem polink *.obj ComView.res /FORCE:MULTIPLE /OUT:COMView.exe /SUBSYSTEM:WINDOWS /MAP /LIBPATH:\WinInc\Lib kernel32.lib advapi32.lib user32.lib gdi32.lib ole32.lib oleaut32.lib comctl32.lib comdlg32.lib shell32.lib uuid.lib ..\Lib\except.lib
rem link *.obj ComView.res /OUT:COMView.exe /SUBSYSTEM:WINDOWS /MAP /LIBPATH:\WinInc\Lib kernel32.lib advapi32.lib user32.lib gdi32.lib ole32.lib oleaut32.lib comctl32.lib comdlg32.lib shell32.lib uuid.lib ..\Lib\except.lib
jwlink format windows pe runtime windows file *.obj op res ComView.res name COMView.exe op MAP LIBPATH \WinInc\Lib lib kernel32, advapi32, user32, gdi32, ole32, oleaut32, comctl32, comdlg32, shell32, uuid, ..\Lib\except.lib op noreloc, statics
cd ..
