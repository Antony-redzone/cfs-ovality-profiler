# Microsoft Developer Studio Project File - Name="clearline" - Package Owner=<4>
# Microsoft Developer Studio Generated Build File, Format Version 6.00
# ** DO NOT EDIT **

# TARGTYPE "Win32 (x86) Dynamic-Link Library" 0x0102

CFG=clearline - Win32 Debug
!MESSAGE This is not a valid makefile. To build this project using NMAKE,
!MESSAGE use the Export Makefile command and run
!MESSAGE 
!MESSAGE NMAKE /f "clearline.mak".
!MESSAGE 
!MESSAGE You can specify a configuration when running NMAKE
!MESSAGE by defining the macro CFG on the command line. For example:
!MESSAGE 
!MESSAGE NMAKE /f "clearline.mak" CFG="clearline - Win32 Debug"
!MESSAGE 
!MESSAGE Possible choices for configuration are:
!MESSAGE 
!MESSAGE "clearline - Win32 Release" (based on "Win32 (x86) Dynamic-Link Library")
!MESSAGE "clearline - Win32 Debug" (based on "Win32 (x86) Dynamic-Link Library")
!MESSAGE 

# Begin Project
# PROP AllowPerConfigDependencies 0
# PROP Scc_ProjName ""$/ClearLineProfilerV5.6/clearline", RPAAAAAA"
# PROP Scc_LocalPath "."
CPP=cl.exe
MTL=midl.exe
RSC=rc.exe

!IF  "$(CFG)" == "clearline - Win32 Release"

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
# ADD BASE CPP /nologo /MT /W3 /GX /O2 /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D "_MBCS" /D "_USRDLL" /D "CLEARLINE_EXPORTS" /YX /FD /c
# ADD CPP /nologo /MT /W3 /GX /O2 /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D "_MBCS" /D "_USRDLL" /D "CLEARLINE_EXPORTS" /FR /YX /FD /c
# ADD BASE MTL /nologo /D "NDEBUG" /mktyplib203 /win32
# ADD MTL /nologo /D "NDEBUG" /mktyplib203 /win32
# ADD BASE RSC /l 0x1409 /d "NDEBUG"
# ADD RSC /l 0x1409 /d "NDEBUG"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LINK32=link.exe
# ADD BASE LINK32 kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib /nologo /dll /machine:I386
# ADD LINK32 kernel32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib /nologo /dll /machine:I386 /out:".\Release\clearline.dll"

!ELSEIF  "$(CFG)" == "clearline - Win32 Debug"

# PROP BASE Use_MFC 0
# PROP BASE Use_Debug_Libraries 1
# PROP BASE Output_Dir "Debug"
# PROP BASE Intermediate_Dir "Debug"
# PROP BASE Target_Dir ""
# PROP Use_MFC 0
# PROP Use_Debug_Libraries 1
# PROP Output_Dir "Debug"
# PROP Intermediate_Dir "Debug"
# PROP Ignore_Export_Lib 0
# PROP Target_Dir ""
# ADD BASE CPP /nologo /MTd /W3 /Gm /GX /ZI /Od /D "WIN32" /D "_DEBUG" /D "_WINDOWS" /D "_MBCS" /D "_USRDLL" /D "CLEARLINE_EXPORTS" /YX /FD /GZ /c
# ADD CPP /nologo /MTd /W3 /GX /Z7 /Od /D "WIN32" /D "_DEBUG" /D "_WINDOWS" /D "_MBCS" /D "_USRDLL" /D "CLEARLINE_EXPORTS" /YX /FD /GZ /c
# ADD BASE MTL /nologo /D "_DEBUG" /mktyplib203 /win32
# ADD MTL /nologo /D "_DEBUG" /mktyplib203 /win32
# ADD BASE RSC /l 0x1409 /d "_DEBUG"
# ADD RSC /l 0x1409 /d "_DEBUG"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LINK32=link.exe
# ADD BASE LINK32 kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib /nologo /dll /debug /machine:I386 /pdbtype:sept
# ADD LINK32 kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib /nologo /dll /debug /machine:I386 /out:"c:/windows/system32/clearline.dll" /pdbtype:sept

!ENDIF 

# Begin Target

# Name "clearline - Win32 Release"
# Name "clearline - Win32 Debug"
# Begin Group "Source Files"

# PROP Default_Filter "cpp;c;cxx;rc;def;r;odl;idl;hpj;bat"
# Begin Source File

SOURCE=.\AutoRotate.cpp
# End Source File
# Begin Source File

SOURCE=.\Capacity.cpp
# End Source File
# Begin Source File

SOURCE=.\CentreCalculations.cpp

!IF  "$(CFG)" == "clearline - Win32 Release"

!ELSEIF  "$(CFG)" == "clearline - Win32 Debug"

# ADD CPP /ZI

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\clearline.cpp

!IF  "$(CFG)" == "clearline - Win32 Release"

!ELSEIF  "$(CFG)" == "clearline - Win32 Debug"

# ADD CPP /ZI /FR

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\clearline.def
# End Source File
# Begin Source File

SOURCE=..\ClearLine.ini
# End Source File
# Begin Source File

SOURCE=.\DeltaMaxMin.cpp
# End Source File
# Begin Source File

SOURCE=.\EditProfile.cpp

!IF  "$(CFG)" == "clearline - Win32 Release"

!ELSEIF  "$(CFG)" == "clearline - Win32 Debug"

# ADD CPP /ZI

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\EmbeddedFile.cpp
# End Source File
# Begin Source File

SOURCE=.\FilterGraph.cpp
# End Source File
# Begin Source File

SOURCE=.\Flat3D.cpp

!IF  "$(CFG)" == "clearline - Win32 Release"

!ELSEIF  "$(CFG)" == "clearline - Win32 Debug"

# ADD CPP /FR

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\Fractile.cpp
# End Source File
# Begin Source File

SOURCE=.\LoadPVD.cpp
# End Source File
# Begin Source File

SOURCE=.\Median.cpp
# End Source File
# Begin Source File

SOURCE=.\Ovality.cpp
# End Source File
# Begin Source File

SOURCE=..\Modules\PageFunctions.bas
# End Source File
# Begin Source File

SOURCE=.\Percentile.cpp
# End Source File
# Begin Source File

SOURCE=..\Modules\ScreenDrawing.bas
# End Source File
# Begin Source File

SOURCE=.\Shapes.cpp
# End Source File
# Begin Source File

SOURCE=.\XYDiameter.cpp
# End Source File
# Begin Source File

SOURCE=.\XYDiameterMaxMin.cpp
# End Source File
# Begin Source File

SOURCE=.\ZipProfile.cpp
# End Source File
# End Group
# Begin Group "Header Files"

# PROP Default_Filter "h;hpp;hxx;hm;inl"
# Begin Source File

SOURCE=.\AutoRotate.h
# End Source File
# Begin Source File

SOURCE=.\Capacity.h
# End Source File
# Begin Source File

SOURCE=..\houghlibv2.0\CBSAlgebra.h
# End Source File
# Begin Source File

SOURCE=.\CentreCalculations.h
# End Source File
# Begin Source File

SOURCE=.\Common.h
# End Source File
# Begin Source File

SOURCE=.\DeltaMaxMin.h
# End Source File
# Begin Source File

SOURCE=.\EditProfile.h
# End Source File
# Begin Source File

SOURCE=.\EmbeddedFile.h
# End Source File
# Begin Source File

SOURCE=.\FilterGraph.h
# End Source File
# Begin Source File

SOURCE=.\Flat3d.h
# End Source File
# Begin Source File

SOURCE=.\Fractile.h
# End Source File
# Begin Source File

SOURCE=.\LoadPVD.h
# End Source File
# Begin Source File

SOURCE=.\median.h
# End Source File
# Begin Source File

SOURCE=.\Ovality.h
# End Source File
# Begin Source File

SOURCE=.\Percentile.h
# End Source File
# Begin Source File

SOURCE=.\Shapes.h
# End Source File
# Begin Source File

SOURCE=.\XYDiameter.h
# End Source File
# Begin Source File

SOURCE=.\XYDiameterMaxMin.h
# End Source File
# Begin Source File

SOURCE=.\ZipProfile.h
# End Source File
# End Group
# Begin Group "Resource Files"

# PROP Default_Filter "ico;cur;bmp;dlg;rc2;rct;bin;rgs;gif;jpg;jpeg;jpe"
# End Group
# Begin Source File

SOURCE=..\Forms\PipelineDetails.frm
# End Source File
# End Target
# End Project
