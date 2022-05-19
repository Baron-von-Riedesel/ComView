@echo off
jwasm.exe -c -Fi=fpo.inc -Zi -coff -nologo -I\WinInc\Include -Sg -Fl=Debug\ -Fo Debug\ *.asm
rem jwasm.exe -c -Zi -coff -nologo -I\WinInc\Include -Sg -Fl=Debug\ -Fo Debug\ *.asm
rem ml.exe -c -Zi -coff -nologo -I\WinInc\Include -Fo Debug\ *.asm
rem rc -i \WinInc\Include -fo Debug\COMView.res COMView.rc
cd Debug
rem jwlink debug c op cvp format windows pe runtime windows file *.obj op res ComView.res name COMView.exe op MAP LIBPATH \WinInc\Lib lib kernel32, advapi32, user32, gdi32, ole32, oleaut32, comctl32, comdlg32, shell32, uuid, ..\Lib\except.lib op noreloc
link /subsystem:windows /debug *.obj ComView.res /out:COMView.exe /MAP /LIBPATH:\WinInc\Lib kernel32.lib advapi32.lib user32.lib gdi32.lib ole32.lib oleaut32.lib comctl32.lib comdlg32.lib shell32.lib uuid.lib ..\Lib\except.lib
cd ..
