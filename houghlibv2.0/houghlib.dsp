# Microsoft Developer Studio Project File - Name="houghlib" - Package Owner=<4>
# Microsoft Developer Studio Generated Build File, Format Version 6.00
# ** DO NOT EDIT **

# TARGTYPE "Win32 (x86) Dynamic-Link Library" 0x0102

CFG=houghlib - Win32 Release
!MESSAGE This is not a valid makefile. To build this project using NMAKE,
!MESSAGE use the Export Makefile command and run
!MESSAGE 
!MESSAGE NMAKE /f "houghlib.mak".
!MESSAGE 
!MESSAGE You can specify a configuration when running NMAKE
!MESSAGE by defining the macro CFG on the command line. For example:
!MESSAGE 
!MESSAGE NMAKE /f "houghlib.mak" CFG="houghlib - Win32 Release"
!MESSAGE 
!MESSAGE Possible choices for configuration are:
!MESSAGE 
!MESSAGE "houghlib - Win32 Release" (based on "Win32 (x86) Dynamic-Link Library")
!MESSAGE "houghlib - Win32 Debug" (based on "Win32 (x86) Dynamic-Link Library")
!MESSAGE "houghlib - Win32 Houghlib Win 2000 debug" (based on "Win32 (x86) Dynamic-Link Library")
!MESSAGE "houghlib - Win32 Houglib Windows 2000 Realease" (based on "Win32 (x86) Dynamic-Link Library")
!MESSAGE 

# Begin Project
# PROP AllowPerConfigDependencies 0
# PROP Scc_ProjName ""$/houghlib", PFAAAAAA"
# PROP Scc_LocalPath "."
CPP=cl.exe
MTL=midl.exe
RSC=rc.exe

!IF  "$(CFG)" == "houghlib - Win32 Release"

# PROP BASE Use_MFC 0
# PROP BASE Use_Debug_Libraries 0
# PROP BASE Output_Dir "Release"
# PROP BASE Intermediate_Dir "Release"
# PROP BASE Target_Dir ""
# PROP Use_MFC 0
# PROP Use_Debug_Libraries 0
# PROP Output_Dir "Release"
# PROP Intermediate_Dir "Release"
# PROP Ignore_Export_Lib 0
# PROP Target_Dir ""
# ADD BASE CPP /nologo /MT /W3 /GX /O2 /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D "_MBCS" /D "_USRDLL" /D "HOUGHLIB_EXPORTS" /YX /FD /c
# ADD CPP /nologo /MD /W3 /GX /Z7 /O2 /I "c:\DXSDK\include" /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D "_MBCS" /D "_USRDLL" /D "HOUGHLIB_EXPORTS" /FR /YX /FD /c
# ADD BASE MTL /nologo /D "NDEBUG" /mktyplib203 /win32
# ADD MTL /nologo /D "NDEBUG" /mktyplib203 /win32
# ADD BASE RSC /l 0x1409 /d "NDEBUG"
# ADD RSC /l 0x1409 /d "NDEBUG"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LINK32=link.exe
# ADD BASE LINK32 kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib /nologo /dll /machine:I386
# ADD LINK32 strmbase.lib quartz.lib kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib /nologo /dll /machine:I386 /out:".\Release\laserlib.dll"

!ELSEIF  "$(CFG)" == "houghlib - Win32 Debug"

# PROP BASE Use_MFC 0
# PROP BASE Use_Debug_Libraries 0
# PROP BASE Output_Dir "houghlib___Win32_Debug"
# PROP BASE Intermediate_Dir "houghlib___Win32_Debug"
# PROP BASE Ignore_Export_Lib 0
# PROP BASE Target_Dir ""
# PROP Use_MFC 0
# PROP Use_Debug_Libraries 0
# PROP Output_Dir "houghlib___Win32_Debug"
# PROP Intermediate_Dir "houghlib___Win32_Debug"
# PROP Ignore_Export_Lib 0
# PROP Target_Dir ""
# ADD BASE CPP /nologo /MD /W3 /GX /O2 /I "e:\DXSDK\include" /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D "_MBCS" /D "_USRDLL" /D "HOUGHLIB_EXPORTS" /YX /FD /c
# ADD CPP /nologo /MDd /W3 /GX /ZI /Od /I "c:\DXSDK\include" /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D "_MBCS" /D "_USRDLL" /D "HOUGHLIB_EXPORTS" /FR /YX /FD /c
# ADD BASE MTL /nologo /D "NDEBUG" /mktyplib203 /win32
# ADD MTL /nologo /D "NDEBUG" /mktyplib203 /win32
# ADD BASE RSC /l 0x1409 /d "NDEBUG"
# ADD RSC /l 0x1409
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LINK32=link.exe
# ADD BASE LINK32 strmbase.lib quartz.lib kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib /nologo /dll /machine:I386 /out:"C:/WINNT/system32/laserlib.dll"
# ADD LINK32 strmbase.lib quartz.lib kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib /nologo /dll /incremental:yes /debug /machine:I386 /out:"C:\CBS\Profiler 6.2.3/laserlib.dll"

!ELSEIF  "$(CFG)" == "houghlib - Win32 Houghlib Win 2000 debug"

# PROP BASE Use_MFC 0
# PROP BASE Use_Debug_Libraries 0
# PROP BASE Output_Dir "houghlib___Win32_Houghlib_Win_2000_debug"
# PROP BASE Intermediate_Dir "houghlib___Win32_Houghlib_Win_2000_debug"
# PROP BASE Ignore_Export_Lib 0
# PROP BASE Target_Dir ""
# PROP Use_MFC 0
# PROP Use_Debug_Libraries 0
# PROP Output_Dir "houghlib___Win32_Houghlib_Win_2000_debug"
# PROP Intermediate_Dir "houghlib___Win32_Houghlib_Win_2000_debug"
# PROP Ignore_Export_Lib 0
# PROP Target_Dir ""
# ADD BASE CPP /nologo /MDd /W3 /GX /Od /I "e:\DXSDK\include" /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D "_MBCS" /D "_USRDLL" /D "HOUGHLIB_EXPORTS" /YX /FD /c
# ADD CPP /nologo /MDd /W3 /GX /Od /I "e:\DXSDK\include" /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D "_MBCS" /D "_USRDLL" /D "HOUGHLIB_EXPORTS" /YX /FD /c
# ADD BASE MTL /nologo /D "NDEBUG" /mktyplib203 /win32
# ADD MTL /nologo /D "NDEBUG" /mktyplib203 /win32
# ADD BASE RSC /l 0x1409
# ADD RSC /l 0x1409
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LINK32=link.exe
# ADD BASE LINK32 strmbase.lib quartz.lib kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib /nologo /dll /debug /machine:I386 /out:"C:/windows/system32/laserlib.dll"
# ADD LINK32 strmbase.lib quartz.lib kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib /nologo /dll /debug /machine:I386 /out:"C:/winnt/system32/laserlib.dll"

!ELSEIF  "$(CFG)" == "houghlib - Win32 Houglib Windows 2000 Realease"

# PROP BASE Use_MFC 0
# PROP BASE Use_Debug_Libraries 0
# PROP BASE Output_Dir "houghlib___Win32_Houglib_Windows_2000_Realease"
# PROP BASE Intermediate_Dir "houghlib___Win32_Houglib_Windows_2000_Realease"
# PROP BASE Ignore_Export_Lib 0
# PROP BASE Target_Dir ""
# PROP Use_MFC 0
# PROP Use_Debug_Libraries 0
# PROP Output_Dir "houghlib___Win32_Houglib_Windows_2000_Realease"
# PROP Intermediate_Dir "houghlib___Win32_Houglib_Windows_2000_Realease"
# PROP Ignore_Export_Lib 0
# PROP Target_Dir ""
# ADD BASE CPP /nologo /MD /W3 /GX /O2 /I "c:\DXSDK\include" /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D "_MBCS" /D "_USRDLL" /D "HOUGHLIB_EXPORTS" /YX /FD /c
# ADD CPP /nologo /MD /W3 /GX /O2 /I "c:\DXSDK\include" /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D "_MBCS" /D "_USRDLL" /D "HOUGHLIB_EXPORTS" /FR /YX /FD /c
# ADD BASE MTL /nologo /D "NDEBUG" /mktyplib203 /win32
# ADD MTL /nologo /D "NDEBUG" /mktyplib203 /win32
# ADD BASE RSC /l 0x1409 /d "NDEBUG"
# ADD RSC /l 0x1409 /d "NDEBUG"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LINK32=link.exe
# ADD BASE LINK32 strmbase.lib quartz.lib kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib /nologo /dll /machine:I386 /out:"C:/houghlibv2.0/release/laserlib.dll"
# ADD LINK32 strmbase.lib quartz.lib kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib /nologo /dll /debug /machine:I386 /out:"C:/windows/system32/laserlib.dll"

!ENDIF 

# Begin Target

# Name "houghlib - Win32 Release"
# Name "houghlib - Win32 Debug"
# Name "houghlib - Win32 Houghlib Win 2000 debug"
# Name "houghlib - Win32 Houglib Windows 2000 Realease"
# Begin Group "Source Files"

# PROP Default_Filter "cpp;c;cxx;rc;def;r;odl;idl;hpj;bat"
# Begin Source File

SOURCE=.\centre.cpp

!IF  "$(CFG)" == "houghlib - Win32 Release"

# ADD CPP /Od

!ELSEIF  "$(CFG)" == "houghlib - Win32 Debug"

# ADD CPP /ZI

!ELSEIF  "$(CFG)" == "houghlib - Win32 Houghlib Win 2000 debug"

!ELSEIF  "$(CFG)" == "houghlib - Win32 Houglib Windows 2000 Realease"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\Counter.cpp

!IF  "$(CFG)" == "houghlib - Win32 Release"

# ADD CPP /Z7 /Od

!ELSEIF  "$(CFG)" == "houghlib - Win32 Debug"

# ADD CPP /Z7

!ELSEIF  "$(CFG)" == "houghlib - Win32 Houghlib Win 2000 debug"

# ADD BASE CPP /Z7
# ADD CPP /Z7

!ELSEIF  "$(CFG)" == "houghlib - Win32 Houglib Windows 2000 Realease"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\FishEyeTransform.cpp

!IF  "$(CFG)" == "houghlib - Win32 Release"

# ADD CPP /Z7 /Od

!ELSEIF  "$(CFG)" == "houghlib - Win32 Debug"

# ADD CPP /Z7

!ELSEIF  "$(CFG)" == "houghlib - Win32 Houghlib Win 2000 debug"

# ADD CPP /Z7

!ELSEIF  "$(CFG)" == "houghlib - Win32 Houglib Windows 2000 Realease"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\hough.def
# End Source File
# Begin Source File

SOURCE=.\IPDInterface.cpp

!IF  "$(CFG)" == "houghlib - Win32 Release"

!ELSEIF  "$(CFG)" == "houghlib - Win32 Debug"

# ADD CPP /Z7

!ELSEIF  "$(CFG)" == "houghlib - Win32 Houghlib Win 2000 debug"

!ELSEIF  "$(CFG)" == "houghlib - Win32 Houglib Windows 2000 Realease"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\laserprofiler.cpp

!IF  "$(CFG)" == "houghlib - Win32 Release"

# ADD CPP /Z7 /Od

!ELSEIF  "$(CFG)" == "houghlib - Win32 Debug"

# ADD CPP /ZI

!ELSEIF  "$(CFG)" == "houghlib - Win32 Houghlib Win 2000 debug"

# ADD BASE CPP /Z7
# ADD CPP /Z7

!ELSEIF  "$(CFG)" == "houghlib - Win32 Houglib Windows 2000 Realease"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\LaserTracking.cpp

!IF  "$(CFG)" == "houghlib - Win32 Release"

!ELSEIF  "$(CFG)" == "houghlib - Win32 Debug"

# ADD CPP /Z7

!ELSEIF  "$(CFG)" == "houghlib - Win32 Houghlib Win 2000 debug"

!ELSEIF  "$(CFG)" == "houghlib - Win32 Houglib Windows 2000 Realease"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\profile.cpp

!IF  "$(CFG)" == "houghlib - Win32 Release"

# ADD CPP /Z7 /Od

!ELSEIF  "$(CFG)" == "houghlib - Win32 Debug"

# ADD CPP /Z7

!ELSEIF  "$(CFG)" == "houghlib - Win32 Houghlib Win 2000 debug"

# ADD BASE CPP /Z7
# ADD CPP /Z7

!ELSEIF  "$(CFG)" == "houghlib - Win32 Houglib Windows 2000 Realease"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\RadialProcess.cpp

!IF  "$(CFG)" == "houghlib - Win32 Release"

# ADD CPP /Z7 /Od

!ELSEIF  "$(CFG)" == "houghlib - Win32 Debug"

# ADD CPP /ZI

!ELSEIF  "$(CFG)" == "houghlib - Win32 Houghlib Win 2000 debug"

# ADD CPP /Z7

!ELSEIF  "$(CFG)" == "houghlib - Win32 Houglib Windows 2000 Realease"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\VBlaserprofiler.cpp

!IF  "$(CFG)" == "houghlib - Win32 Release"

# ADD CPP /Z7 /Od

!ELSEIF  "$(CFG)" == "houghlib - Win32 Debug"

# ADD CPP /Z7

!ELSEIF  "$(CFG)" == "houghlib - Win32 Houghlib Win 2000 debug"

# ADD BASE CPP /Z7
# ADD CPP /Z7

!ELSEIF  "$(CFG)" == "houghlib - Win32 Houglib Windows 2000 Realease"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\video.cpp

!IF  "$(CFG)" == "houghlib - Win32 Release"

# ADD CPP /Z7 /Od

!ELSEIF  "$(CFG)" == "houghlib - Win32 Debug"

# ADD CPP /Z7

!ELSEIF  "$(CFG)" == "houghlib - Win32 Houghlib Win 2000 debug"

# ADD BASE CPP /Z7
# ADD CPP /Z7

!ELSEIF  "$(CFG)" == "houghlib - Win32 Houglib Windows 2000 Realease"

!ENDIF 

# End Source File
# End Group
# Begin Group "Header Files"

# PROP Default_Filter "h;hpp;hxx;hm;inl"
# Begin Source File

SOURCE=.\CBSAlgebra.h
# End Source File
# Begin Source File

SOURCE=.\centre.h
# End Source File
# Begin Source File

SOURCE=.\Counter.h
# End Source File
# Begin Source File

SOURCE=.\FishEyeTransform.h
# End Source File
# Begin Source File

SOURCE=.\IPDInterface.h
# End Source File
# Begin Source File

SOURCE=.\laserprofiler.h
# End Source File
# Begin Source File

SOURCE=.\LaserTracking.h
# End Source File
# Begin Source File

SOURCE=.\profile.h
# End Source File
# Begin Source File

SOURCE=.\RadialProcess.h
# End Source File
# Begin Source File

SOURCE=.\VBlaserprofiler.h
# End Source File
# Begin Source File

SOURCE=.\video.h
# End Source File
# End Group
# Begin Group "Resource Files"

# PROP Default_Filter "ico;cur;bmp;dlg;rc2;rct;bin;rgs;gif;jpg;jpeg;jpe"
# Begin Source File

SOURCE=.\Dump.txt
# End Source File
# End Group
# Begin Source File

SOURCE=.\changelog.txt
# End Source File
# Begin Source File

SOURCE=C:\ClearLineProfilerV4.0\ErrorCodes.txt
# End Source File
# Begin Source File

SOURCE=.\ScrapBook.txt
# End Source File
# End Target
# End Project
