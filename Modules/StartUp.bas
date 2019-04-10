Attribute VB_Name = "StartUp"
Option Explicit

'Contants
Public Const PVFlat As Integer = 0
Public Const PVMedianDiameter As Integer = 1
Public Const PVOvality As Integer = 2
Public Const PVMaxDiameter As Integer = 3
Public Const PVXYDiameter As Integer = 4
Public Const PVCapacity As Integer = 5
Public Const PVYDiameter As Integer = 6
Public Const PVDebris As Integer = 7
Public Const PVMinDiameter As Integer = 8
Public Const PVOvalitySmooth As Integer = 9 'PNC9999
Public Const PVXDiameterSmooth As Integer = 10
Public Const PVYDiameterSmooth As Integer = 11
Public Const PVMedianDiameterSmooth As Integer = 12
Public Const PVMaxDiameterSmooth As Integer = 13
Public Const PVCapacitySmooth As Integer = 14
Public Const PVMinDiameterSmooth As Integer = 15
Public Const PVDeflectionX As Integer = 16 'PCN5186
Public Const PVDeflectionY As Integer = 17 'PCN5196
'PCN6458 Public Const PVInclination As Integer = 18 'PCN6128
'PCN6458 Public Const PVDesignGradient As Integer = 19 'PCN6178
'PCN6458 Public Const PVInclinationSmooth As Integer = 20 'PCN6128


Public SoftwareConfiguration As String 'PCN3809 'Full or Reader

Public D3D_NumberofFrames(10) As Long 'PCN 2453
Public d3d_setlanguage(15) As String 'PCN2473
Public Const InvalidData As Double = -1000000000
Public IPD As Boolean 'PCN3744 ipd flag
Public CapacityDataOffset As Single
Public TrueDiameterOffset As Single
Public LoadVideo As Boolean
Public PVDLoadError As Boolean
Public DebrisOn As Boolean 'PCN4461
Public ShiftOn As Boolean 'PCN4484
Public LockedDonut As Boolean
Public FlatOvality As Boolean
Public SetClearMask As Long 'PCN9999
Public MedianFlat As Boolean 'PCN4974
Public SeaLevelStartHeight As Single 'PCN6128
Public SeaLevelEndHeight As Single 'PCN6128
Public DesignGradient As Single ' PCN6165




'vvvv PCN1970 *************************************************
'Whenever changing the version of a DLL, ensure the number complies with the following guidelines.
'The DLL may be changed more or less often than the VB version.
'Do we want to update a user with a new version of the VB software every time we change the DLL version? Probably not.
'So the VB software will except DLL version with the same major version number. That is if the VB DLL version is 1.0, the VB will accept the DLL version 1.0 to 1.9. The VB will not DLL versions <1.0 or >1.9.
'E.g.: ClearLine Profiler's LaserLib.dll version = 1.0. Then ClearLine will accept LaserLib.dll version from 1.0 to 1.9
'Therefore, for a VB software with a DLL version number of 1.0, ALL DLLs with versions 1.0 to 1.9 MUST work on this VB software.
'If the change in the DLL means it will not work on ALL VB software of the same major DLL version, then the DLL's version MUST increase the major DLL version.
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Const DirectXNeeded As Long = 9 'PCN3140 to keep inline with other version checks will use a constant
                                       'This is the Direct X Version needed. At least 9 in this case

Public Const PVDVersion As String = "V6.4"      'V6.4 Saves water calculation centre
                                                'V6.3 PCN4006 PVData now single, so twice the size in storage per frame, instead of ints
                                                'V6.1 PCN3019 FisheyeOriginalWidth and FisheyeOriginalHeight added
                                                'V5.3 PCN2952 variable introduced

Public Const INIVersion As Double = 7.9         '7.9 Added inclination PCN6128
                                                '7.8 Added DeflectionOrNormal
                                                '7.7 Added delfection graph, PCN4974
                                                '7.5 added titles to the 1k graphs
                                                '7.5 Added min diameter and all its extras
                                                '7.4 added graph subtittles.
                                                '7.3 PCN4799 limit lines and graph setting added to INI
                                                '7.2
                                                '7.1 PCN4294 Added PipeDetails labels to the INI.
                                                '7.0 Added PCN3687 Horizontal calibration and camera version
                                                '6.8 Added scale setting for each of the graphs
                                                '6.7 PCN3197 HASP HL dongle locking field 'HASPLock' added
                                                '6.6 PCN3069 Added report margins
                                                '6.5 PCN2395 Added Video Settings and Video Capture Device
                                                '6.4 Final defult setting for profiling
                                                '6.3 PCN3031 Added Fish_Displayed
                                                '6.2 PCN3024 Added PaperSize
                                                '6.1 PCN3019 Added Fish_OriginalWidth and Fish_OriginalHeight
                                                '6.0 PCN2980 Added PVDiameterMethod="XY"
                                                '5.6 PCN2866 Set defaults for IP, PVGraph YScale and Limit Lines.
                                                '5.5 PCN2829, PCN2703 and PCN2769
                                                '5.4 PCN2443 Verify the ClearLine.ini version number

Public Const LaserLibVersion As Double = 16.5   '16.5 vob (or any video) at resolution 352 x 480 caused overflow, memory leaks fixed, spike introduced in rough point removal fixed
                                                '16.4 Add registry entry (only in administrator mode) for elle card, also set seek reg entry to 3
                                                '16.3 removed baby smoothing, changed im so that its only the size of the video, not xtra. Swaped height with width, added some more memory cleanup. Downsample, was introducing errors, spikes.
                                                '     rewrote with simple averaging of good data, if less than two poings good, then make a hole.
                                                '16.1 release
                                                '16.0 New smoothing from 540 points.
                                                '15.9 Broke water display fixing blue dot overlay.
                                                '15.8 Broke the ability to turn blue dot overlay off when playing back pre recorded video
                                                '15.7 Added 540 and contrast
                                                '15.6 Remove sections
                                                '15.4 PNC9999
                                                '15.3 Ele card demultiplexer and ele card favourtism
                                                '15.1 PCN4596 'When mask is set and turns
                                                '15.0    'Loading video no vidoe sinc, loading pvd video video sync on.
                                                '14.9 Added the ability to lock the donut
                                                '14.8 Added the ability to turn centre calculation off.
                                                '14.7 and more deadlock stuff, final
                                                '14.6 More deadlock stuff
                                                '14.4   video framerewind added
                                                '14.3   'Divide by 0 error on fisheye initialise
                                                '14.2   '14.2
                                                '14.1 'Pre release
                                                '13.9   'Broke the water level, left out the reference, now put back in :(
                                                'Opps, screwed up the time stamp fix, now fixed the fix Antony van Iersel - 3 Aug 2006, 10:44pm
                                                '13.7 Fixed time stamp error, was adding a pause frame to recording. lastTimeReocrded added.
                                                '13.6 vb function calls to cpp changed to sub by ref calls
                                                '13.5 reversh fish eye,
                                                '13.4   'Blurring was adding a laser edge at the edge of the image, this rectified
                                                '13.3   'Added Blur and took out overflow in centre calculation
                                                '13.1   'Pre-testing to see if it does the normal things
                                                '13.0 PCN3561, video crashing, 3567 move reference circle
                                                '11.2   'PCN2990 3401 3402 3373 3441 etc, beta release for demostration 5.6
                                                '10.9  'PCN3??? and PCNant, LaserTracking, inverse video and semieliptical shape
                                                '10.8   'PCN2781 IM is no initialised after the video not before
                                                '10.7   'PCN3122 Bulls Eye for Centre and Centre history. Laser width and Donut now has min limits
                                                '10.6   'PCN???? fatel flaw in laserlib fixed.
                                                'PCN3085 More memory leaks fixed, PCN???? when in inches flat no longer black and white
                                                '10.4 PCN2395 Added two laserlib calls to check capture devices
                                                '10.3   'PCNXXXX new centre
                                                '10.1   'Majour Bug Fixes
                                                '10     'PCN3017 changing over to clearline profiler 5.5
                                                '6.9    'PCN2874 Ticker counter now able to count in feet and meters.
                                                '6.8 PCN2874 Ticker counter now able to count in feet and meters.
                                                '6.7 PCN2865 Call laserlib.dll videoframeadvance added.
                                                '6.6 PCN1970 Created to allow VB to confirm version capatibility
                                                '6.5 PCN2778 Last major memory leak removed, includes memory lead detection, and initHw added (which is used to stop accessing Laserlib)
                                                '6.4 PCN2639 Improvement in Distance counter
                                                '6.3 PCN2639 Distance counter functionality added
                                                '6.2 PCN2405 Finding truer centre. 10/03/04
                                                '6.1 PCN2668 Every frame processed 05/03/04
                                                '6.0 PCN2612 Upgrade to Manual Tuning form interface.
                                                '5.9
                                                '5.8 PCN2575 Contrast Control on Auto Calibration added.
                                                '5.7 PCN2426(Preparation for more circles), PCN2433(Refining auto calibration)
                                                '5.6 PCN2461 (Majour memory leak fix)
                                                '5.5 PCN2400(Adjust center, use variance, Progress bar), PCN2420, PCN2421, PCN2419
                                                'PCN2461 Fix a major memory leak
                                                'PCN2290 FISH-EYE. This makes LaserLibVersion 5.4
                                                'PCN1970 Verify the LaserLib.dll version number
                                                'PCN2488 Ignoring half-circles, laserlib.dll -> 5.9
                                                
                                                '10.6 '  'PCN2990 3401 3402 3373 3441 etc, beta release for demostration 5.6

Public Const ThreeDimVersion As Double = 10.9   'Fixed 3D capture
                                                '10.8 Added back the stl export panel when exporting the 3d data
                                                '10.7 Three dim fix for stuff (textures not being there)
                                                '10.7 Heaps different, eg colour now from vb etc
                                                '10.5 LOD altered and release version
                                                '10.4   'PCN3141 detection of Direct X Version
                                                '10.3   'PCN3111 pas thru units to 3D Pipe, PCN3112 auto select vertexMode
                                                '10.2 PCN3085 was a problem when you loaded multiple videos. But a fault was also
                                                '     found in the three D, a serious one. Now Fixed.
                                                '10.1   'Majour bug fixes
                                                '10     'PCN3017 change over to clearline profiler 5.5
                                                '2.1 PCN2860 Mulitplier passed to 3D so it can display the higher presision, thats for Ovality, capacity and delta
                                                '2.0 PCN2473 Language support added to C++ D3D Pipe
                                                '1.9 PCN2653 Fixed a Crash on Some machines when exiting the program after 3D use.
                                                '1.8 PCN2367 Direction indicator added
                                                '1.7 PCN2510 Coloured Pipe Limits causing crash
                                                '1.6 PCN2465 & PCN2461 & PCN2467, Memory Leaks reduced from 13Meg to 3k, Protection to stop VB accessing
                                                '     Variables and functions when 3D scene is unloaded, Missing Texture files and Directory displayed (Antony van Iersel, 9 Dec 2003)
                                                '1.5 PCN2453 Changed the way VB passed the number of profile points
                                                'PCN2376 Added Export 3D in STL format
                                                'PCN2337 Now Coloured Pipe Limits
                                                'PCN2322 Now unlimited Frames, Textures changed to JPG's
                                                'PCN2240 Verify the ThreeDim.dll version number
                                                
Public Const ClearLineVersion As Double = 5#    'Max and min boundry checks, max and min not initialised
                                                '4.9 smooth graph -100000000 bug
                                                '4.8 PCN6314 bung end frame inclination calc
                                                '4.7Inclination adjustment
                                                '4.6 Richards inclination added
                                                '4.5 added lay of pipe from inlination data
                                                ' 4.4 Added inclination PCN6128
                                                ' 4.3 inverted the delfection graph, so greater than median is possible %, lessthan is negative %
                                                'also out by one on the delfection graph, shoud have been <toFrames, not <= to frames
                                                '4.2 ' in max min calc, max and min were not initialised
                                                '4.1Deflection graph
                                                '4.0PCN4974 median flat graph (Diameter Flat Graph)
                                                '3.9 water level bug in true diameter. Now true dimater, ovality, and capacity, there water level is treated like a hole.
                                                '    abs - was used for smooth, but capacity goes -ve, not marker for capaicyt to not use abs
                                                '3.8 Release
                                                '3.6 Ovality calc now looks for greatest deviance from mean to choose either to use min diameter, or max diamter calculation
                                                '3.3   'Serious bug in smoothing, added almost 2 %, when averageing the graph, total was not initislised to 0. Accumilated error.
                                                '3.2 smoothing egnore bad data.
                                                'Added PCN9999 smoothing
                                                '2.0 Added fill in for xy min max
                                                '2.9 ovality rflection fill now for holes with no water, centre same
                                                '2.8 added a new cetre calc
                                                '2.7 Added refrence shape shift array and put a whole fill on true diameter
                                                '2.6 Added whole filling in techneque for xy graphs
                                                '2.5 Removed all reference to GDI, it stoped the profiler in windows 2000
                                                '2.4       'added selectable centre calcs 'PCN4418
                                                '2.3   'Flat graph filter added to get rid of spikes
                                                '2.1 fixed overflow with fractile calulation
                                                '1.9 Added water whole fill in when finding the median diameter
                                                '1.8 Flat fixed moved left, when it move right and visa versa
                                                '1.7 90% Fractile added ' PCN4296
                                                '1.6 'Imbeded reports were growing the file (which was bad) and not overwriting the file
                                                'more complete auto rotate
                                                '1.4   'Shift embed saving to C++
                                                'Trying to get rid of a overflow error from C++ no quite
                                                '1.2   'divide by 0 error hidden in the fill top whole cacluation for ovality and centre calcs
                                                '1.1 water level was not passed to centre calculation
                                                '1.0     'At last start this ones version number off.
                
Public Const SonarVersion  As Double = 1.1 'Added sonar version


Public Const ArmadilloVersion As String = "6.00"    'PCN4183, incapsulating dll's
                                                    '5.3 PCN2212 Arm Initial creation of variable for Armadillo

Public Const FECVersion As String = "2.0"       '2.0 PCN3019 Created a new FEC file format with version control.

'^^^^ *********************************************************
Public Declare Sub setdeviceinput Lib "laserlib.dll" () 'PCN2289
Public Declare Sub getcounter Lib "laserlib.dll" (ByRef counter As Long) 'PCN2639 Testing distance counter 'PCN3744 moved here.
Public Declare Sub hough_checkforIPD Lib "laserlib.dll" (ByRef AnyIPD As Long) ' returns 1 for ok 0 for not ok
'vvvv PCN2612 **************
'Private Declare Sub setmethod Lib "laserlib.dll" (ByVal i As Long) '        Input:  0 Type 1 'PCN2612
''                                                                                   1 Type 2 'PCN2612
'^^^^ *********************

Public Declare Sub getscalevalue Lib "laserlib.dll" (ByRef Scl As Double)
Public Declare Sub setscalevalue Lib "laserlib.dll" (ByVal S As Double)
Public Declare Sub hough_SetYFishScale Lib "laserlib.dll" (ByVal value As Double) 'PCN3687
Public Declare Sub hough_GetYFishScale Lib "laserlib.dll" (ByRef FishScale As Double) 'PCN3687
Public Declare Sub calculatescale Lib "laserlib.dll" () 'PCN3687
Public Declare Sub createmask Lib "laserlib.dll" () 'PCN3687
Public Declare Sub getimagesize Lib "laserlib.dll" (ByRef height As Long, ByRef width As Long)
Public Declare Sub setoriginalsize Lib "laserlib.dll" (ByVal width As Long, ByVal height As Long)
Public Declare Sub hough_processimageonoff Lib "laserlib.dll" (ByVal IsOn As Boolean)
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, _
                                                                         ByVal lpBuffer As String) As Long

Public Declare Sub d3d_laser_focus Lib "threedim.dll" (ByVal focus As Long)

Public ThreeDActivated As Boolean   'Determines whether the user has paid for the 3d package
Public ThreeDRunning As Boolean

Public LastDataTime As Double
Public RecordTimeIncrement As Integer

Public DrawingFlatGraph As Boolean

'Public CalibrationTypeLength As String
Public CalibrationTypeLength As Double



'Regional Language Settings - Number settings - Richard
'**************************************************************************************************************************
Public Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Public Declare Function SetLocaleInfo Lib "kernel32" Alias "SetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String) As Boolean
Public Declare Function GetUserDefaultLCID% Lib "kernel32" ()


Public Const LOCALE_SDECIMAL              As Long = &HE     'decimal separator
Public Const LOCALE_STHOUSAND             As Long = &HF     'thousand separator
Public Const LOCALE_SGROUPING             As Long = &H10    'digit grouping
Public Const LOCALE_IDIGITS               As Long = &H11    'number of fractional digits
Public Const LOCALE_ILZERO                As Long = &H12    'leading zeros for decimal
Public Const LOCALE_INEGNUMBER            As Long = &H1010  'negative number mode
Public Const LOCALE_SNATIVEDIGITS         As Long = &H13    'native ASCII 0-9
Public Const LOCALE_SPOSITIVESIGN         As Long = &H50    'positive sign
Public Const LOCALE_SNEGATIVESIGN         As Long = &H51    'negative sign

Public RegDecSymbol As String
Public RegThousandSeperator As String
Public RegDigitGrouping As String
Public RegFractionalDigits As String
Public RegLeadingZeros As String
Public RegNegFormat As String
Public RegPosSign As String
Public RegNegSign As String
'**************************************************************************************************************************


Public PVXScaleLimitPerL As Double ' PCN3501 Single 'PCNGL2901032 'PCN2337 Made Global for 3D Pipe Colour Limits
Public PVXScaleLimitPerR As Double ' PCN3501 Single 'PCNGL2901032 'PCN2337 Made Global for 3D Pipe Colour Limits

Public CompanyName As String
Public PhoneNo As String
Public FaxNo As String
Public CompanyLogoPath As String
Public MeasurementUnits As String
Public ThreeDRenderingStyle As Integer 'String 'PCN2266 'PCN4197 changed from a string to a integer, now 0 is auto, 2 is software
'Public CaptureDevice As String 'PCN2289 'PCN2298
'Public ConfigArray(20) As String
Public Config_LineCnt As Integer
'Public CurrentPVPage As Control
Public FontName As String
Public FontType As String
Public FontSize As Long
Public FontColour As String
'Public ProcessMethod As String 'PCN2773
'Public Contrast As Integer 'PCN2773
'Public Enhancement As String 'PCN2773

'Variables to enable detection of user's operating system
Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
 End Type


  Private Declare Function GetVersionEx Lib "kernel32" _
      Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Boolean
      
  ' dwPlatforID Constants
  Public Const VER_PLATFORM_WIN32s = 0
  Public Const VER_PLATFORM_WIN32_WINDOWS = 1
  Public Const VER_PLATFORM_WIN32_NT = 2
'

Public Change As Boolean 'Flag for change of field contents
Public Registered As Boolean 'Flag for product registration
Public PVDSaved As Boolean 'Flag to indicate if .pvd file has been saved

Public MyFile As String
Public ImageDataFile As String
Public PVDFileName As String 'Contains the path and file name of the currently loaded PVD file 'PCNGL140103
Public PVDFileName2nd As String 'PCN4380
Public strMediaFilePath As String
Public strMediaFileName As String
Public AVIInitialised As Boolean 'AVI file status validation before running any video C functions. Required to minimise the likelihood of a C function crash. 'PCNGL150103
'Public MPGInitialised As Boolean 'PCNLS030203
Public AVIFrameTime() As Double  'Time in seconds from the start of the AVI file. Used to put an AVI frame time stamp on each line of the PVData. 'PCNGL150103
Public StartupBypass As Boolean 'Bypass the ClearLineScreen form resize event on startup, PCNGL171202

Public NoOfProfileSegments As Integer 'PCNGL1812022 The number of segments within a single profile (usually 180, may increase to 360)
Public WaterEgnoreList() As Long 'PCN3219 a list of points each matching a profile point, 1 for egnore, 0 to accept point
                                 ' this is used for calculations, normally used for water level

Public PipeInfoArray()


Public NumTimesRecorded As Long

'vvvv PCN2988 *************************************
Public TD_PVDataX() As Single 'PCN4006 'Contains the PVData X loaded from file and passed to the 3D profile modelling program. 'PCNGL090603
Public TD_PVDataY() As Single 'PCN4006 'Contains the PVData Y loaded from file and passed to the 3D profile modelling program. 'PCNGL090603
'^^^^ *********************************************

'vvvv PCN3219 We are going to need adjustable centre if the water level is to be changed after recording
Public TD_PVCentreX() As Single 'PCN4006 Long 'Contains the Centre calculations result for the X coord
Public TD_PVCentreY() As Single 'PCN4006 Contains the Centre calculations result for the Y coord

'vvvv PCN4484 The reference shape and water level values accross the whole of the PVD
Public PVShapeCentreX() As Single
Public PVShapeCentreY() As Single
Public PVWaterLevelStart() As Single
Public PVWaterLevelFinnish() As Single

Public TD_PVDataX2nd() As Single 'PCN4380
Public TD_PVDataY2nd() As Single 'PCN4380
Public TD_PVCentreX2nd() As Single 'PCN4380
Public TD_PVCentreY2nd() As Single 'PCN4380
Public PVDistances2nd() As Double 'PCN4380
Public PVTimes2nd() As Double 'PCN4380

'Public PVGraphType As String 'Determines whether the graph is ovality, capacity, etc replaced with imagegraphstate(0).GraphType PCN???? (9 August 2005, Antony)

'vvvv PCN2240 *************************************
Public PVCapacityFullData() As Single ' PCN3540 was Long, now not saving to file so make it single with no multiplier
'Public PVOvalityFullData() As Single ' PCN3540 was Long, now not saving to file so make it single with no multiplier



'Public PVOvalityOrigFullData() As Single
Public PVDeltaFullMax() As Double 'PCN3450
Public PVDeltaFullMin() As Double 'PCN3450
Public PVDeltaSegFullMax() As Integer 'PCN3450 (Antony, 4 August 2005)
Public PVDeltaSegFullMin() As Integer 'PCN3450 (Antony, 4 August 2005)
'^^^^ *****************************************

'**********************************************
Public PVXDiameterFullData() As Double 'Stores the results of the X Diameter calcs on the PVData -  PCN2703
Public PVYDiameterFullData() As Double 'Stores the results of the Y Diameter calcs on the PVData -  PCN2703
Public PVDiameterFullMax() As Double
Public PVDiameterFullMin() As Double
'Public PVMinMaxSegNosFullData() As Long 'Stores the MinSegNo, MinOppSegNo, MaxSegNo and MaxOppSegNo 'PCN2962 'PCN2966
Public PVDiameterSegFullMax() As Integer 'PCN3450 (Antony, 4 August 2005)
Public PVDiameterSegFullMin() As Integer 'PCN3450 (Antony, 4 August 2005)
Public PVDiameterMedian() As Double 'Stores the calculated median diameter PCN2639 'PCN3489
'Public PVFractile() As Single 'PCN4235


'**********************************************

Public PVFlat3DRed() As Long 'PCNGL060103 'PCNGL140103 Changed to integer 'PCN2970 Changed to Long
Public PVFlat3DBlue() As Long 'PCNGL060103 'PCNGL140103 Changed to integer 'PCN2970 Changed to Long
Public PVFlat3DGreen() As Long 'PCNGL060103 'PCNGL140103 Changed to integer 'PCN2970 Changed to Long

Public PVRecording As Boolean  'PCNLS240103
Public LastRecordedFrame As Long  'PCNLS300103
'Public LastRecordedMainFrame As Long 'PCNLS080203 'Not used and not to be used PCN3289 (3 Feb 2005, Antony)
'Main picture screen mode controls
Public VideoAspectRatio As Double 'Video Screen aspect ratio, Image Height/Width , determined by C code 'PCNGL2401032
Public Const VideoAspectRatio768x576 As Double = 0.75 'MainScreen default video size (768x576 or aspect ratio of 0.75) 'PCNGL2401032
Public CLPScreenMode As String
Public Const PV As String = "PV"
Public Const Video As String = "Video"
Public Const SnapShot As String = "SnapShot"
Public Const ThreeD As String = "ThreeD"
Public Const StillImage As String = "StillImage"

Public CLPScreenActionPrevious As String
Public CLPScreenAction As String 'PCN3569 'What action the CLP screen is up to, like, Move, DrawLine, Calibrate etc
Public CLPScreenDrawState As String 'PCN4046 'What state the drawing routines are in, FirstClick, RubberBand, NextClick
Public CLPScreenDrawAction As String 'What is current happening, like RightClick, LeftClick, WheelScrollUp etc
Public CLPScreenItemSelect As String 'What item is selected of cursor is over
'Public VideoSnapShotMode As String 'PCN4043
Public PicInPicMode As String
'***********************************
'These PV ratios are in mm or inches per pixel


Public PVGraphCapacityXScale As Double  '+/- X scale markers 'PCN2829
Public PVGraphOvalityXScale As Double  '+/- X scale markers 'PCN2829
Public PVGraphDeltaXScale As Double  '+/- X scale markers 'PCN2829
Public PVGraphXYDiaXScale As Double  '+/- X scale markers 'PCN2829
Public PVGraphGeneralXScale As Double 'PCN3373 General X Scale for the switchible graph
Public PVGraphDiaMaxMinXScale As Double 'PCN3540
Public PVGraphDiaMedianXScale As Double 'PCN3540
Public PVGraphDiaMaxXScale As Double
'Public PVGraphFractileXScale As Double



Public PVGraphYRatio As Double 'Contains the scaling ratio for the PVGraph Y -axis
Public GraphStartFrame As Double 'This is the start frame drawn on the graphs. PCN???? (9 August 2005, Antony)
Public GraphEndFrame As Double 'This is the end frame drawn on the graphs. PCN???? (9 August 2005, Antony)
Public YScaleZoomFactor As Integer 'The Y scale zoom factor, the default value is 1 'PCNGL080103


'PCN3402 Graphs offsets for the four line graphs ''''''''''''''''''''''''''''
Public PVGraphCapacityXOffset As Double 'Graph offset from the centre line  '
Public PVGraphOvalityXOffset As Double 'Graph offset from the centre line   '
Public PVGraphDeltaXOffset As Double 'Graph offset from the centre line     '
Public PVGraphXYDiaXOffset As Double 'Graph offset from the centre line     '
Public PVGraphGeneralXOffset As Double 'Graph offset fromt he centre line   '
Public PVGraphDiaMaxMinXOffset As Double 'PCN3540
Public PVGraphDiaMedianXOffset As Double 'PCN3540
Public PVGraphDiaMaxXOffset As Double
'Public PVGraphFractileXOffset As Double



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'****************************************
Public PVFrameNo As Long 'Precision Vision Frame Number
Public PVScaleMarkerStFrame As Long ' PCN3373 Frame start marker for the PV Scale Lines (29 march 2005, Ant)
Public PVScaleMarkerFnFrame As Long ' PCN3373 Frame finnish marker for the PV Scale Lines (29 march 2005, Ant)
Public PVDataNoOfLines As Long 'Total number of Precision Vision Data lines
Public PVDataNoOfLines2nd As Long 'PCN4380
Public Const MaxFrameBufferNo As Integer = 1 'Determines how big the buffer arrays are for PVData, capacity etc. 'PCNGL140103
Public RequestFrameNo As Long 'Video timer requested FrameNo counter 'PCNGL060103
Public MaxDisplayedFrameNo As Long 'PVGraph number of frames that are currently displayed, typically 2 windows on loading PVD file. PCNGL070103
Public MaxCalculatedFrameNo As Long 'PVGraph number of frames that are currently calculated, Flat3D and/or Max/Min 'PCN2970
'Public CurrentPVYScalePageNo As Long 'PCNGL140103
Public CurrentPVYScalePageNo As Integer 'PCNGL140103 'PCN2639
Public CentreLineX As Single 'PV Main Screen grid centre line - X axis
Public CentreLineY As Single 'PV Main Screen grid centre line - Y axis
Public VideoCentreLineX As Single 'PV Main Screen grid centre line - X axis
Public VideoCentreLineY As Single 'PV Main Screen grid centre line - Y axis
Public MainScreenMouseIcon As Long ' Contains the path name under the application dir. 'PCNGL021202

Public CalLen_Global As Double 'Integer PCN3640
Public CalLength_Global As Double
Public CalLengthYScale_Global As Double
Public CalLineExist As Boolean
'Public InternalDiameterExpected_Global As Single 'PCN1836
Public ExpectedDiameter As Double 'PCN???? Need a better sounding global ExpectedDiameter, its used every as temp anyway


Public Const PI As Double = 3.141592

Public CalDrawFlag As Boolean
Public CalCirKind As Integer
Public CalLineExistFlag As Boolean

Public CurrentShape As Integer
Public ArrayCnt As Long
Public PVDrawScreenRatio As Double 'Screen scale ratio, pixels/mm (or pixels/in) 'PCNGL220103
'Delete
Public DelFlag As Boolean 'PCNGL030403-1


'Colours
Public NormalDrawingColor As Long
Public SelectedObjectColor As Long
Public ModiCircleColor As Long
Public ChosenModiCircleColor As Long
Public AreaFillingColor As Long
Public ExtraObjectColor As Long
Public JointCircleColor As Long
Public TempDrawingColor As Long
Public MovingObjectColor As Long
Public ModifyingObjectColor As Long
Public RotatingObjectColor As Long
Public SelectionBoundaryColor As Long
Public TextSizeIndicatorColor As Long 'Loaded from the INI
'^^^^ *******************************************************************
Public Calibration As Double
Public DrawAutoSnap As Boolean 'Replaces autosnap on the Options page 'PCNGL210103
'Drawing Numbers, originally required for the Y scales
Public NumberPic(14, 7, 4) As Long 'PCNGL090103
Public NumberPicX(13, 4, 7) As Long 'PCNLS130103 'PCN2463 'PCN2777
Public FramesPerSec As Integer 'Frames per second for PV play rate, usually 25fpm
'Limit marker settings
Public CapacityLimitL As Double ' PCN3501 Long
Public CapacityLimitR As Double ' PCN3501 Long
Public OvalityLimitL As Double ' PCN3501 Long
Public OvalityLimitR As Double ' PCN3501 Long
Public DeltaLimitL As Double ' PCN3501 Long
Public DeltaLimitR As Double ' PCN3501 Long
Public XYDiameterLimitL As Double ' PCN3501 Long 'PCN2703
Public XYDiameterLimitR As Double ' Long 'PCN2703
Public Flat3dLimitL As Double 'PCNANTONY Now Flat graph has its one limits
Public Flat3dLimitR As Double 'PCNANTONY
Public DiameterMaxMinLimitR As Double 'PCN3540
Public DiameterMaxMinLimitL As Double 'PCN3540
Public DiameterMedianLimitR As Double 'PCN3540
Public DiameterMedianLimitL As Double 'pcn3540
Public DiameterMaxLimitR As Double
Public DiameterMaxLimitL As Double
'Public FractileLimitR As Double 'PCN4235
'Public FractileLimitL As Double 'PCN4235

Public PipeDisplayMoveLastY As Integer  'PCNGL181202           'Last mouse Y position (Pipe indicator movement)
Public IndicatorOffset As Integer

Public VideoFileName As String  'PCNLS201202  'Filename of AVI 'PCNGL140103 Changed to VideoFileName
Public ImageFileName As String 'PCN3194

Public PVPageTop As Single 'PCNGL301202
Public PVPageLeft As Single 'PCNGL301202
Public PVPageHeight As Single 'PCNGL301202
Public PVPageWidth As Single 'PCNGL301202


'PCN3021 removed these global variables, all references to these variables are to
'be replaced with the configinfo equivilant
'Public MediaWidth As Long    'PCNLS290103
'Public MediaHeight As Long    'PCNLS290103
Public MediaOriginalHeight As Long 'The video files original height before it is corrected = MediaHeight. 'PCNGL180503-1
Public MediaOriginalWidth As Long
Public WLDrawStAng As Double
Public WLDrawFhAng As Double


Public SnapshotFilename As String 'The name of the file that the C code produces
                                  'as a snapshot of the avi
Public SystemDir As String 'PCN4337
Public LocToSave As String  'This stores the location (directory)
                            'of the ini file (without the name) so to store any files,
                            'just use this & the filename
Public Const DefaultPVDFileName As String = "PVDRecording.PVD" 'This file is used to store the recording PVD data from the live or AVI video 'PCNGL140103
Public VideoFrame As Long 'This is the global frame number that is established from
                             'C code whent he videotimer is incremented
Public mediatype As String 'This is the type of media playing in the videoscreen
                           'can be either "", "Video" or "Live" or later (not implemented yet) "MPEG"
                           'used for determining how it is uninitialized.
                           'CRUCIAL that this is never anything else
                           'an empty string means no media type has been loaded
Public WindowsTempDirectory As String 'PCN???? this is windows temp directory where we know we can write
                           
                           
'***For Image Processing*****************************************************
'Public XThres As Double 'PCN2612
'Public YThres As Double 'PCN2612
'Public GradThres As Long
Public ShowLaserWidth As Long 'PCN3017
Public ShowProfileCandidates As Long 'PCN3017
Public ShowGreenX As Long
Public ShowGreenY As Long
Public ShowProf As Long
Public ShowColour As Long
Public TextMiddle As Long
Public PercProfPnts As Long
Public TotalPerc As Double
Public BestPerc As Long
Public OvalPipe As Long
'Public StDX As Double
'Public StDY As Double
Public ProportionDone As Double
Public BestProportion As Double
Public XAdjust As Double
Public YAdjust As Double
Public RecordMode As String  'Either NTSC or PAL
Public TuningStyle As String   'Either Automatic or Manual
Public Optimised As Boolean
Public LightInPipe As Boolean  'True if the pipe is light, false otherwise (default dark)
Public UserDefinedLight As Boolean
Public WaterLevelIgnoreCenter As Boolean 'tells the C to use ignorewaterlevel to find center
'Public WaterLevelIgnoreProfile As Boolean 'Allways profile water level 'tells the C to use ignorewaterlevel to find profile
                           
'XThres , YThres, GradThres, StDX, StDY, ShowGreen, ShowProf, ShowColour, PercentProfPnts, TotalPerc, TextMiddle, OvalPipe, XAdjust
                           
'***************************************************************************
'****                 Binary file format variables                     ***** 'PCNGL110103
'Main Header information (the first line of any PVD file
Type PVDFileMainHeaderType
    PVDFileMHAppName As String * 20      'PVD File Main Header Application Name
    PVDFileMHVersionMajor As Integer     'PVD File Main Header Version Number, Major
    PVDFileMHVersionMinor As Integer     'PVD File Main Header Version Number, Minor
    PVDFileMHVersionRev As Integer       'PVD File Main Header Version Number, Revision
    PVDFileMHPointerAddress As Integer       'PVD File Main Header Starting Address of the Pointers
    PVDFileMHNoOfPointers As Integer       'PVD File Main Header NoOfPointers
    'PVDFileMHRecordMode As String        'PVD File Main Header Record Mode of Media 'PCN1914 'Taken out as per GL010403
End Type
Public PVDFileMainHeader As PVDFileMainHeaderType
'File header pointers (must be the second block of data in any PVD file, there are no fix number of pointers)
Type PVDPointerType
    PVDPointerConfigInfo As Long     'PVD File Header Pointer for Pipe Information 'PCNGL130103
    PVDPointerPipeInfo As Long       'PVD File Header Pointer for Pipe Information
    PVDPointerPipeObs As Long        'PVD File Header Pointer for Pipe observations 'PCNGL130103
    PVDPointerFontInfo As Long       'PVD File Header Pointer for Font Information
    PVDPointerDrawInfo As Long       'PVD File Header Pointer for Drawing Information
    PVDPointerPVData As Long         'PVD File Header Pointer for Precision Vision Data
End Type
Public Const PVDFileOutPutNoOfPointers As Integer = 6
Public PVDFilePointers As PVDPointerType
Type PVDHeaderType
    'File header Name (Their must be a Header Descriptor before each block)
    PVDHeaderDescriptor As String * 15      'PVD File Header Descriptor, used for writing header block Descriptor
    'File header CheckSum (These are check numbers used in auditing the length of a block of header data)
    PVDCheck As Long       'PVD File Header CheckSum
End Type
Public PVDHeaderConfigInfo As PVDHeaderType 'PCNGL130103
Public PVDHeaderPipeInfo As PVDHeaderType
Public PVDHeaderPipeObs As PVDHeaderType 'PCNGL130103
Public PVDHeaderFontInfo As PVDHeaderType
Public PVDHeaderDrawInfo As PVDHeaderType
Public PVDHeaderPVData As PVDHeaderType
'PCN3576
Type PVDHeaderEmbeddedType
    'File header Name (Their must be a Header Descriptor before each block)
    Descriptor As String * 15   'PVD File Header Descriptor, used for writing header block Descriptor
    Owner As String * 50        'Who owns this embedded file
    FileLength As Long              'Length of embedded file
    'File header CheckSum (These are check numbers used in auditing the length of a block of header data)
    PVDCheck As Long       'PVD File Header CheckSum
End Type

Type EmbeddedOwners
    FileOffset As Long 'File pointer offset for snapshot data (bitmap) (widht 4 bytes)           '
    FileLength As Long 'File length of snapshot data (bitmap, eg snapshot file, (width 4 bytes)  '
    PVHeaderEmbedded As PVDHeaderEmbeddedType
    EmbeddedType As String
    EmbeddedIndex As Integer
End Type

Public ListEmbeddedOwners() As EmbeddedOwners

'ID4601
'Public Const EmbeddedFileNameAndPath As String = "CBS\Embedded.jpg" 'PCN4233
'Public Const EmbedBMPFileNameAndPath As String = "CBS\EmbedFile.bmp" 'PCN4233
'Public Const EmbedJMPFileNameAndPath As String = "CBS\EmbedFile.jpg" 'PCN4233


Public Const EmbeddedFileNameAndPath As String = "Embedded.jpg" 'ID4601 'PCN4233
Public Const EmbedBMPFileNameAndPath As String = "EmbedFile.bmp" 'ID4601 'PCN4233
Public Const EmbedJMPFileNameAndPath As String = "EmbedFile.jpg" 'ID4601 'PCN4233

'PCN4233 ''''''''''''''''''''''''
Type StoredReportPageIndex_V10  '
    EmbeddedIndex   As Integer  'observations array index where the image filepointer information is stored
    PageNumber As Integer       'Page number, they should match the index after sorting it
End Type                        '
                                '
Type StoredReportType_V10       '''''''''
    Title As String                     'Title of the report that is being printed
    ReportType As Integer               'Type (not yet defined) will be a number representing report type, (Profile, Observation, Single etc)
    NumberOfPages As Integer            'Number of pages that goes with this report.
    Page() As StoredReportPageIndex_V10 'Index information for each report, this points to the place where the image of the report is stored
    ReportNumber As Integer             'Unique ID on each report, each report gets a new mumber when stored
End Type                                '
'''''''''''''''''''''''''''''''''''''''''

Public StoredReportArray() As StoredReportType_V10

'***************************************************************************
'****                 System Configuration Information variables       ***** 'PCNGL130103
'vvvv PCN2392 ************************************
Type ConfigInfoType_V40
    Units As String * 2
    FileCountryCode As Integer
    FileLanguage As Integer
    CalDist As Integer
    CalLineLength As Single
    Ratio As Double
    NoOfProfileSegments As Integer
    LenReal As Double
    LenRealPercent As Double
    WLStartAngle As Single
    WLFinishAngle As Single
    FishEyeHorDistortion As Double 'PCN3687 was area that was never used
    'VideoFileName As String
    VideoFileName As String * 500 'PCN1768
    MediaWidth As Long 'PCN1833
    MediaHeight As Long 'PCN1833
End Type
Public Const ConfigInfoNoOfLines_V40 As Integer = 12
'Public ConfigInfo As ConfigInfoType_V40 'PCN2392
Type ConfigInfoType_V41 'PCN2392
    Units As String * 2
    FileCountryCode As Integer
    FileLanguage As Integer
    CalDist As Integer
    CalLineLength As Single
    Ratio As Double
    NoOfProfileSegments As Integer
    LenReal As Double
    LenRealPercent As Double
    WLStartAngle As Single
    WLFinishAngle As Single
    FishEyeHorDistortion As Double 'PCN3687 was area that was never used
    'VideoFileName As String
    VideoFileName As String * 500
    MediaWidth As Long
    MediaHeight As Long
    'PVD file version 'PCN2392
    PVDFileVersion As String * 4  'PCN2392
    'Fish Eye distortion parameters 'PCN2392
    FishEyeFlag As Boolean 'PCN2392 Fish Eye enable/disable flag
    FishEyeDistortion As Integer 'PCN2392 Fish Eye Distortion factor Fish_Distortion
    FishEyeRatio As Double 'PCN2495
    FishEyeCenterX As Integer 'PCN2497
    FishEyeCenterY As Integer 'PCN2497
End Type
Public Const ConfigInfoNoOfLines_V41 As Integer = 15 'PCN2392 'PCN2639
'Public ConfigInfo As ConfigInfoType_V41 'PCN2392 'PCN2639
Public ConfigInfo_V40 As ConfigInfoType_V40 'PCN2492
'^^^^ *******************************************
'vvvv PCN2639 *******************************************************
Type ConfigInfoType_V50
    Units As String * 2
    FileCountryCode As Integer
    FileLanguage As Integer
    CalDist As Integer
    CalLineLength As Single
    Ratio As Double
    NoOfProfileSegments As Integer
    LenReal As Double
    LenRealPercent As Double
    WLStartAngle As Single
    WLFinishAngle As Single
    FishEyeHorDistortion As Double 'PCN3687 was area that was never used
    'VideoFileName As String
    VideoFileName As String * 500
    MediaWidth As Long
    MediaHeight As Long
    'PVD file version 'PCN2392
    PVDFileVersion As String * 4
    'Fish Eye distortion parameters
    FishEyeFlag As Boolean 'Fish Eye enable/disable flag
    FishEyeDistortion As Integer 'Fish Eye Distortion factor Fish_Distortion
    FishEyeRatio As Double
    FishEyeCenterX As Integer
    FishEyeCenterY As Integer
    DistanceProcessMethod As String * 25 'Contains the distane calculation method if applicable.
    DistanceStart As Double
    DistanceDirection As String * 4
    DistanceFinish As Double
End Type
Public Const ConfigInfoNoOfLines_50 As Integer = 25 'PCN2820
'Public ConfigInfo As ConfigInfoType_V50 'PCN2820
Public ConfigInfo_V41 As ConfigInfoType_V41
'^^^^ ***************************************************************
'vvvv PCN2820 *******************************************************
Type ConfigInfoType_V52 'PCN2829
    Units As String * 2
    FileCountryCode As Integer
    FileLanguage As Integer
    CalDist As Integer
    CalLineLength As Single
    Ratio As Double
    NoOfProfileSegments As Integer
    LenReal As Double
    LenRealPercent As Double
    WLStartAngle As Single
    WLFinishAngle As Single
    FishEyeHorDistortion As Double 'PCN3687 was area that was never used
    'VideoFileName As String
    VideoFileName As String * 500
    MediaWidth As Long
    MediaHeight As Long
    'PVD file version 'PCN2392
    PVDFileVersion As String * 4
    'Fish Eye distortion parameters
    FishEyeFlag As Boolean 'Fish Eye enable/disable flag
    FishEyeDistortion As Integer 'Fish Eye Distortion factor Fish_Distortion
    FishEyeRatio As Double
    FishEyeCenterX As Integer
    FishEyeCenterY As Integer
    DistanceProcessMethod As String * 25 'Contains the distane calculation method if applicable.
    DistanceStart As Double
    DistanceDirection As String * 4
    DistanceFinish As Double
    'Image Processing Settings
    PVShapeCentreX As Double 'Replace IPXThres 'PCN4336
    PVShapeCentreY As Double 'Replace IPYThres 'PCN4336
    IPGradThres As Long
    IPStDX As Double
    IPStDY As Double
    IPProcessMethod As String
    IPZone As Integer
    IPEnhancement As String
    'Limit lines
    LimitCapacityL As Double
    LimitCapacityR As Double
    LimitOvality As Double
    LimitDeltaL As Double
    LimitDeltaR As Double
    LimitXYDiameterL As Double
    LimitXYDiameterR As Double
End Type
Public Const ConfigInfoNoOfLines_V52 As Integer = 40 'PCN2891
'Public ConfigInfo As ConfigInfoType_V51
'Public ConfigInfo As ConfigInfoType_V52 'PCN2829
Public ConfigInfo_V50 As ConfigInfoType_V50
'^^^^ ***************************************************************
'vvvv PCN2891 *******************************************************
Type ConfigInfoType_V60
    Units As String * 2
    FileCountryCode As Integer
    FileLanguage As Integer
    CalDist As Integer
    CalLineLength As Single
    Ratio As Double
    NoOfProfileSegments As Integer
    LenReal As Double
    LenRealPercent As Double
    WLStartAngle As Single
    WLFinishAngle As Single
    FishEyeHorDistortion As Double 'PCN3687 was area that was never used
    'VideoFileName As String
    VideoFileName As String * 500
    MediaWidth As Long
    MediaHeight As Long
    'PVD file version 'PCN2392
    PVDFileVersion As String * 4
    'Fish Eye distortion parameters
    FishEyeFlag As Boolean 'Fish Eye enable/disable flag
    FishEyeDistortion As Integer 'Fish Eye Distortion factor Fish_Distortion
    FishEyeRatio As Double
    FishEyeCenterX As Integer
    FishEyeCenterY As Integer
    DistanceProcessMethod As String * 25 'Contains the distane calculation method if applicable.
    DistanceStart As Double
    DistanceDirection As String * 4
    DistanceFinish As Double
    'Image Processing Settings
    PVShapeCentreX As Double 'PCN4336
    PVShapeCentreY As Double 'PCN4336
    IPGradThres As Long
    IPStDX As Double
    IPStDY As Double
    IPProcessMethod As String
    IPZone As Integer
    IPEnhancement As String
    'Limit lines
    LimitCapacityL As Double
    LimitCapacityR As Double
    LimitOvality As Double
    LimitDeltaL As Double
    LimitDeltaR As Double
    LimitXYDiameterL As Double
    LimitXYDiameterR As Double
    'vvvv PCN2891 *********************************************************
    ProfileRecordingMethod As String
    'Specifies method of determining the PV Profile data.
    'The old method ("Radius") records radius as an integer every 2 degrees.
    'For better accuracy the method "XY" specifies an X and Y co-ordinate,
    'both integers. However this can be translated into a more accurate
    'radius. NOTE: The angle between profile points will NOT be a constant
    'if the FishEye distortion function has been applied to the video.
    'This is because the 180 profile points are taken from the original
    'video and only these profile points are transformed by the FishEye
    'distortion function. Hence the angle between profiles can not be
    'a constant.
    '^^^^ ******************************************************************
End Type
'Public Const ConfigInfoNoOfLines As Integer = 40
'Public ConfigInfo As ConfigInfoType_V60 'PCN3019
Public ConfigInfo_V52 As ConfigInfoType_V52
'^^^^ ***************************************************************
'vvvv PCN3019 *******************************************************
Type ConfigInfoType_V61
    Units As String * 2
    FileCountryCode As Integer
    FileLanguage As Integer
    CalDist As Integer
    CalLineLength As Single
    Ratio As Double
    NoOfProfileSegments As Integer
    LenReal As Double
    LenRealPercent As Double
    WLStartAngle As Single
    WLFinishAngle As Single
    FishEyeHorDistortion As Double 'PCN3687 was area that was never used
    'VideoFileName As String
    VideoFileName As String * 500
    MediaWidth As Long
    MediaHeight As Long
    'PVD file version 'PCN2392
    PVDFileVersion As String * 4
    'Fish Eye distortion parameters
    FishEyeFlag As Boolean 'Fish Eye enable/disable flag
    FishEyeDistortion As Integer 'Fish Eye Distortion factor Fish_Distortion
    FishEyeRatio As Double
    FishEyeCenterX As Integer
    FishEyeCenterY As Integer
    DistanceProcessMethod As String * 25 'Contains the distane calculation method if applicable.
    DistanceStart As Double
    DistanceDirection As String * 4
    DistanceFinish As Double
    'Image Processing Settings
    PVShapeCentreX As Double 'PCN4336
    PVShapeCentreY As Double 'PCN4336
    IPGradThres As Long
    IPStDX As Double
    IPStDY As Double
    IPProcessMethod As String 'ID5395 taken away * 8, set never to change, Added * 8
    IPZone As Integer
    IPEnhancement As String 'ID5395 taken away * 8, set never to change,Added * 8
    'Limit lines
    LimitCapacityL As Double
    LimitCapacityR As Double
    LimitOvality As Double
    LimitDeltaL As Double
    LimitDeltaR As Double
    LimitXYDiameterL As Double
    LimitXYDiameterR As Double
    'vvvv PCN2891 *********************************************************
    ProfileRecordingMethod As String 'ID5395 taken awaw * 6 , never to change 'Added * 6
    'Specifies method of determining the PV Profile data.
    'The old method ("Radius") records radius as an integer every 2 degrees.
    'For better accuracy the method "XY" specifies an X and Y co-ordinate,
    'both integers. However this can be translated into a more accurate
    'radius. NOTE: The angle between profile points will NOT be a constant
    'if the FishEye distortion function has been applied to the video.
    'This is because the 180 profile points are taken from the original
    'video and only these profile points are transformed by the FishEye
    'distortion function. Hence the angle between profiles can not be
    'a constant.
    '^^^^ ******************************************************************
    FishEyeOriginalWidth As Long 'This is the media width of the video used for fisheye calibration
    FishEyeOriginalHeight As Long 'This is the media height of the video used for fisheye calibration
    '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    '!!!!!! Always add new variables to the bottom of this type. !!!!!!!!!!!!
End Type

'ID5495 not needed now, fixed was worse then problem. Now to stop corruption, make sure IPEnhancement IPProcessMethod are not changed in configinfo
'''^^^^ ***************************************************************
'''vvvv PCN3019 *******************************************************
''Type ConfigInfoType_V61_currupting
''    Units As String * 2
''    FileCountryCode As Integer
''    FileLanguage As Integer
''    CalDist As Integer
''    CalLineLength As Single
''    Ratio As Double
''    NoOfProfileSegments As Integer
''    LenReal As Double
''    LenRealPercent As Double
''    WLStartAngle As Single
''    WLFinishAngle As Single
''    FishEyeHorDistortion As Double 'PCN3687 was area that was never used
''    'VideoFileName As String
''    VideoFileName As String * 500
''    MediaWidth As Long
''    MediaHeight As Long
''    'PVD file version 'PCN2392
''    PVDFileVersion As String * 4
''    'Fish Eye distortion parameters
''    FishEyeFlag As Boolean 'Fish Eye enable/disable flag
''    FishEyeDistortion As Integer 'Fish Eye Distortion factor Fish_Distortion
''    FishEyeRatio As Double
''    FishEyeCenterX As Integer
''    FishEyeCenterY As Integer
''    DistanceProcessMethod As String * 25 'Contains the distane calculation method if applicable.
''    DistanceStart As Double
''    DistanceDirection As String * 4
''    DistanceFinish As Double
''    'Image Processing Settings
''    PVShapeCentreX As Double 'PCN4336
''    PVShapeCentreY As Double 'PCN4336
''    IPGradThres As Long
''    IPStDX As Double
''    IPStDY As Double
''    IPProcessMethod As String 'Added * 8
''    IPZone As Integer
''    IPEnhancement As String 'Added * 8
''    'Limit lines
''    LimitCapacityL As Double
''    LimitCapacityR As Double
''    LimitOvality As Double
''    LimitDeltaL As Double
''    LimitDeltaR As Double
''    LimitXYDiameterL As Double
''    LimitXYDiameterR As Double
''    'vvvv PCN2891 *********************************************************
''    ProfileRecordingMethod As String 'Added * 6
''    'Specifies method of determining the PV Profile data.
''    'The old method ("Radius") records radius as an integer every 2 degrees.
''    'For better accuracy the method "XY" specifies an X and Y co-ordinate,
''    'both integers. However this can be translated into a more accurate
''    'radius. NOTE: The angle between profile points will NOT be a constant
''    'if the FishEye distortion function has been applied to the video.
''    'This is because the 180 profile points are taken from the original
''    'video and only these profile points are transformed by the FishEye
''    'distortion function. Hence the angle between profiles can not be
''    'a constant.
''    '^^^^ ******************************************************************
''    FishEyeOriginalWidth As Long 'This is the media width of the video used for fisheye calibration
''    FishEyeOriginalHeight As Long 'This is the media height of the video used for fisheye calibration
''    '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
''    '!!!!!! Always add new variables to the bottom of this type. !!!!!!!!!!!!
''End Type
''Public ConfigInfo_currupting As ConfigInfoType_V61_currupting
Public PossibleConfigInfoCurruption As Boolean
Public Const ConfigInfoNoOfLines As Integer = 42
Public ConfigInfo As ConfigInfoType_V61

Type IPEnhancementAndIPProcessMethod_V1 'ID5395
    IPProcessMethod As String
    IPEnhancement As String
End Type
Public IPEnhancementAndIPProcessMethod As IPEnhancementAndIPProcessMethod_V1


Public ConfigInfo2nd As ConfigInfoType_V61 'PCN4380
Public ConfigInfo_V60 As ConfigInfoType_V60
'^^^^ ***************************************************************
'***************************************************************************
'****                 Pipeline Information variables                   ***** 'PCNGL130103
Type PipeLineInfo_V40
    IntDiameter As Single
    ExtDiameter As Single
    PipeLength As Single
    Comments As String * 250
    Material As String * 50
    AssetNo As String * 100
    SiteID As String * 6
    City As String * 50
    Date As Date
    Time As Date
    StartName As String * 20
    StartLocation As String * 100
    FinishName As String * 20
    FinishLocation As String * 100
End Type


Public Const PipeInfoNoOfLines As Integer = 14
Public PipelineInfo As PipeLineInfo_V40
'***************************************************************************
'****                 Observation  variables                           ***** 'PCNGL130103
Type PipeObservationType_V40
    PipeObsDist As Single
    PipeObs As String * 100 'Observation field width is 100 characters
End Type
'vvvv PCN2928 ************************************************
'Public PipeObservations() As PipeObservationType_V40 'To be redim when observation records are added
'
'Note V50 will still be of a data size that is equivalent to V40.
'Hence V50 WILL be used to read PipeObs in older revision PVD files.
Type PipeObservationType_V50
    PipeObsDist As Single
    PipeObs As String * 96     'Observation field width is 98 characters
    PipeObsFrameNo As Long   'Frame number of Observation (width is 4 Bytes)
End Type

'PCN3576 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Orignally in V40 the  PipeObs was 100 characters, then to make room in the same file space                 '
'the pipeobs was reduced by 4 bytes, 96 characters (V50), now one more time we have to make room            '
'for two more longs witch reduced the pipeobs by another 4 x 2 bytes, 8 bytes so pipe obs for               '
'V60 is now 88 characters.                                                                                  '
Type PipeObservationType_V60                                                                                '
    PipeObsDist As Single                                                                                   '
    PipeObs As String * 88 'Observation field with is 88 characters                                         '
    PipeObsSnapshotOffset As Long 'File pointer offset for snapshot data (bitmap) (widht 4 bytes)           '
    PipeObsSnapshotLength As Long 'File length of snapshot data (bitmap, eg snapshot file, (width 4 bytes)  '
                                  'every byte so it can be restored to a seperate file to load by VB)       '
    PipeObsFrameNo As Long 'Frame number of observation (width is 4 bytes)                                  '
End Type                                                                                                    '
Public PipeObservations() As PipeObservationType_V60 'To be redim when observation records are added        '
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'This type is used to store the Distance Counter known distance points
'for correcting the Estimate distance and Autodistance calculations.
Type DistanceCounterFixedPt_V50
    FixedPtFrameNo As Long
    FixedPtDist As Single
End Type
Public DistanceCounterFixedPt() As DistanceCounterFixedPt_V50 'To be redim when observation records are added
Public NoOfPipeObservations As Integer 'Defines the number of Pipe Observations in the PVD file
Public PipeObsBuffer As Integer 'Defines the file space or buffer assigned to the PVD for Pipe observations. V4 = 10 and V5 = 100
'^^^^ ********************************************************
'***************************************************************************
'****                 Font Information variables                       ***** 'PCNGL130103
Type FontInfo_V40
    FontName As String * 30
    FontType As String * 10
    FontSize As Integer
    FontColour As String * 10
End Type
Public Const FontInfoNoOfLines As Integer = 4
Public FontInfo As FontInfo_V40
'***************************************************************************
'****            PVData and PVGraph calculation variables              ***** 'PCN2639
'Main PV data
Public pvData() As Single  'PCN4006 'Contains the Precision Vision profile data 'PCNGL140103
Public PVDataFrameBlockSize As Integer  'Contains the memory block size for a frame of PVData - PCN2639
Public Const PVDataFrameBlockSize_V40 As Integer = 360 '180x2 Contains the memory block size for a frame of PVData - PCN2639
Public Const PVDataFrameBlockSize_V50 As Integer = 360 '180x2 Contains the memory block size for a frame of PVData - PCN2639
Public Const PVDataFrameBlockSize_V60 As Integer = 720 'PVDataX (180x2) + PVDataY (180x2) Contains the memory block size for a frame of PVData - PCN2891
Public Const PVDataFrameBlockSize_V70 As Integer = 1440 '(PVDataX (180x2) + PVDataY (180x2)) x 2 because its single. Contains the memory block size for a frame of PVData - PCN4006
'PVGraph Calculations
Public pvCapacityData() As Integer 'Stores the results of the capacity calcs on the PVData -  PCNGL101202 'PCNGL140103 Changed to integer
Public PVOvalityData() As Integer 'Stores the results of the out of round calcs on the PVData -PCNGL101202'PCNGL140103 Changed to integer
Public PVDelta() As Integer 'Stores the results of the out of delta calcs on the PVData -PCNGL101202'PCNGL140103  Changed to integer
Public PVCalculationsBlockSize As Integer  'Contains the memory block size for PV Calculations - PCN2639
Public Const PVCalculationsBlockSize_V40 As Integer = 8 '4x2 (2 for Delta), Contains the memory block size for PV Calculations - PCN2639
Public PVDiameterXY() As Integer 'Stores the calculated horizontal and vertical diameters PCN2639

Public PVDiameterMedianForInspection As Double 'PCN3489
Public Const PVCalculationsBlockSize_V50 As Integer = 14 '7x2 (2 each for Delta and PVDiameterXY), Contains the memory block size for PV Calculations - PCN2639
Public Const PVCalculationsBlockSize_V60 As Integer = 14 '7x2 (2 each for Delta and PVDiameterXY), Contains the memory block size for PV Calculations - PCN2891
'vvvv PCN2829***********************************************
Public Const PVCalculationsMultiplier_V52 As Integer = 100 'PVCapacity, PVOvality, PVDelta and PVXY are integers (usually nominal percentages). This multiplier determines how many decimal places these percentages will have.
Public PVCalculationsMultiplier As Long ' PCN2860 (Antony, 2 June 2004)  was Integer no long, so its easier to pass to the 3D
'^^^^ ******************************************************
Public Const PVCalculationsMultiplier_V60 As Integer = 100 'PVCapacity, PVOvality, PVDelta and PVXY are integers (usually nominal percentages). This multiplier determines how many decimal places these percentages will have. 'PCN2891
'Other PV Frame related information
Public PVTimes() As Double  'Contains times (corresponding to frame) for every PV frame currently in PVData

'The related info is Distance and Time, each a double at 8 bytes each, 8 * 2 = 16
Public PVRelatedInfoBlockSize As Integer  'Contains the memory block size for PV frame related information PCN2639
Public Const PVRelatedInfoBlockSize_V40 As Integer = 8 '1x8 Contains the memory block size for PV frame related information PCN2639
Public PVDistances() As Double  'PCN2639 Contains distance information for each PV frame.
Public Const PVRelatedInfoBlockSize_V50 As Integer = 16 '2x8 Contains the memory block size for PV frame related information PCN2639
Public Const PVRelatedInfoBlockSize_V60 As Integer = 16 '2x8 Contains the memory block size for PV frame related information PCN2891
'***************************************************************************

Public D3DLanguageArray() As String * 100 'PCN2473 Antony van Iersel 11 March 2004, to store language array for C++ D3D

'***************************************************************************
Public Const PVGraphReport_GraphWidth As Long = 12000 'Defines the length of the graph in terms of twips 'PCNGL160103
Public MainScaleGrid As Double 'Specifies the MainScreen grid spacing 'PCNGL200103 'PCN1858

'vvvv PCNGL240403-1 **************************************
'The following variables are used to define the area of the
'video image ingnored by the image processing.
Public IgnoreAreaFlag As Boolean
Public IgnoreX1 As Long 'Defines the start X of the area ignored by the IP in the video image 'PCNGL240403-1
Public IgnoreY1 As Long 'Defines the start y of the area ignored by the IP in the video image 'PCNGL240403-1
Public IgnoreX2 As Long 'Defines the finish X of the area ignored by the IP in the video image 'PCNGL240403-1
Public IgnoreY2 As Long 'Defines the finish y of the area ignored by the IP in the video image 'PCNGL240403-1
Public WaterLevelFlag As Boolean 'PCN1939
Public WLStartAngle As Double 'Defines the start angle, relative to the ref circle centre point, for the water level 'PCN1939
Public WLFinishAngle As Double 'Defines the finish angle, relative to the ref circle centre point, for the water level 'PCN1939
'^^^^ ****************************************************
'vvvv PCN2639 ************************************************
'The following variables are used to define the area of the
'video image that difines the distance counter and ingnored by the image processing.
Public IgnoreDistAreaFlag As Boolean
Public IgnoreDistX1 As Long 'Defines the start X of the distance counter area
Public IgnoreDistY1 As Long 'Defines the start y of the distance counter area
Public IgnoreDistX2 As Long 'Defines the finish X of the distance counter area
Public IgnoreDistY2 As Long 'Defines the finish y of the distance counter area
'^^^^ ********************************************************

Public Language As String 'PCN2111
Public EULAFilename As String 'to store the text filename for EULA to load. PCN2111
Public HelpFilename As String 'PCN2111 & PCN2167 7/8/03 by Abe
Public ReaderHelpFile As String
Public ReadOnlyAppPath As String 'PCN2123
' FISH-EYE( PCN2290 )---------v
'Public FishEyeFlag As Boolean 'PCN2392 Now ConfigInfo.FishEyeFlag
'Public iFETransX As Integer 'PCN2392 Now ConfigInfo.FishEyeDistortion
Public iFEScaleX As Integer
Public iFESCaleY As Integer
' FISH-EYE( PCN2290 )---------^

''PCN3513 Background load of flat3d no longer needed (Antony, 12 May 2005)
''
''Public Flat3DCancel As Boolean 'PCN2371

'vvvv PCN2463 **************************************
'The method of determining the distance vs FrameNo
'for the PVGraph and reports
Public DistanceMethod As String
Public CameraSpeedInFrames As Double
Public CameraSpeedInTime As Double
Public DistanceStartTime As Double
Public DistanceStart As Double
Public CountDirection As String 'PCN2639
'^^^^ **********************************************
'vvvv PCN2820 ****************************************
''vvvv PCN2769 ****************************************
'Public LimitCapacityL As Double
'Public LimitCapacityR As Double
'Public LimitOvality As Double
'Public LimitDeltaL As Double
'Public LimitDeltaR As Double
'Public LimitXYDiameterL As Double 'PCN2703
'Public LimitXYDiameterR As Double 'PCN2703
''^^^^ ************************************************
'^^^^ ************************************************

'vvvv PCN2612 ***************************************************
' Precision Vision Tuning and display
Public Declare Sub setmethod Lib "laserlib.dll" (ByVal i As Long) '        Input:  0 Type 1 'PCN2612
'                                                                                   1 Type 2 'PCN2612
Public Declare Sub setprofileoverlay Lib "laserlib.dll" (ByVal i As Long) 'Input:  0 Off
'                                                                                   1 100% of profile size
'                                                                                   2 105% of profile szie
Public Declare Sub setimageanalysis Lib "laserlib.dll" (ByVal i As Long)  'Input:  0 Normal video
'                                                                                   1 Image Enhancement (gray scale)
'                                                                                   2 Black background
Public Declare Sub getimageanalysis Lib "laserlib.dll" (ByRef OnOff As Long) 'Output: 1 or on, 0 for off
'                                                                                   false for off
Public Declare Sub setvideofiltertype Lib "laserlib.dll" (ByVal i As Long) 'Input: 0 Red (Low)
'                                                                                   1 Green (High)
'                                                                                   2 Blue (Standard)
'                                                                                   3 Mixed (eg green + blue)
'Public Declare Sub setselectionfilter Lib "laserlib.dll" (ByVal X As Long, ByVal Y As Long) 'Input:    0 Off
'                                                                                                       1 On
Public Declare Sub setprofilecandidates Lib "laserlib.dll" (ByVal i As Long) 'Input:    0 Off
'                                                                                        1 On
Public Declare Sub hough_showlaserwidth Lib "laserlib.dll" (ByVal i As Long) 'Input 0 off PCN3017
'                                                                                   1 on
Public Declare Sub hough_setinsidezone Lib "laserlib.dll" (ByVal i As Double) 'PCN3017 inside zone circle
Public Declare Sub hough_setoutsidezone Lib "laserlib.dll" (ByVal i As Double) 'PCN3017 ouside zone circle
Public Declare Sub setgradthreshold Lib "laserlib.dll" (ByVal GT As Long)
Public Declare Sub setstandarddeviation Lib "laserlib.dll" (ByVal XT As Double, ByVal YT As Double, ByVal SDX As Double, ByVal SDY As Double) 'PCN2612
Public Declare Sub setprofilecontrast Lib "laserlib.dll" (ByVal i As Long)
Public Declare Sub hough_getcapturedevices Lib "laserlib.dll" (ByVal hwnd As Long) 'PCN2395 Multiple capture device select (21 Sept, Ant)
Public Declare Sub hough_anycapturedevices Lib "laserlib.dll" (ByRef AnyCapture As Long) 'PCN2395 is there any capture devices

Public Declare Sub hough_lockdonut Lib "laserlib.dll" (ByVal diameter As Double) 'PCN4539
Public Declare Sub hough_unlockdonut Lib "laserlib.dll" () 'PCN4539

Public Declare Sub houge_AdjustContrastBright Lib "laserlib.dll" (ByVal rgbScaler As Double, ByVal brightness As Double)

'^^^^ ***********************************************************
Public VideoScreenScale As Double 'PCN2891 Represents the scaling factor between the video resolution and the VideoScreen size 'PCNGL2901032

Public Declare Sub SetWaterLevel Lib "laserlib.dll" Alias "setwaterlevel" (ByRef EgnoreList As Long) 'PCN3219

'vvvv PCN2930 **********************************************************
Type SliderParameters 'Used for the Video frame slider
    Max As Long 'The maximum value for the slider
    Min As Long 'The minimum value for the slider
    value As Double 'The current value for the slider 'PCN2955
    MarkerStart As Long
    MarkerStop As Long
    MarkerPosition As Long
    FrameTop As Integer
    FrameHeight As Integer
    FrameLeft As Integer
    Framewidth As Integer
    FrameRailHeight As Integer
    FrameSpaceMajor As Integer
    FrameMinorSpacing As Integer
End Type
'^^^^ ******************************************************************
Public FishEyeActive As Boolean 'PCN2907 Antony 2 August 2004, so when distance rectangle is set it knows to turn fisheye off firt
Public FishEyeProcessing As Boolean 'PCN2907 Antony 4 August 2004, when the fisheye turning on or off this flag is high to indicate fisheye is busy
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'vvvv PCN2970 *****************************************************************
Public Const NoOfPVGraphs As Integer = 5  'How many PVGraphs (Capacity, Ovality, Delta, XY_Dia and Flat makes 5)
Public Const PVGraphHeightLimit As Long = 240000 'Specifies the height limit of a PVGraph set.
Public PVGraphTotalHeight As Long
Public Const NoOfPVGraphSets As Integer = 4  'How many PVGraph Sets (ie makes up PVGraph length)
'Note: NoOfPVGraphSets can not be increased at this stage since
'      we get the error "Can't create AutoRedraw Image" when AutoRedraw is True.
'      This is because there is not enough memory for the AutoRedraw method.
'^^^^ *************************************************************************
Public PVDiameterMethod As String 'PCN2980

'Public CurrentPVGraphPageNoUPPER As Long 'The current PVGraph Page, Upper Limit 'PCN2970
'Public CurrentPVGraphPageNoLOWER As Long 'The current PVGraph Page, Lower Limit 'PCN2970
'Public BackgroundLoadCancel As Boolean 'PCN2970 Cancel background loading flag PCN3513 (Antony, 12 May 2005) no longer need background load.
'Public Const PVDataXYMultiplier As Long = 10 'PCN2988 Since PVData is an integer, this multiplier transforms PVData into a 1 decimal number.
Public Const PVDataXYMultiplier As Long = 1 ' PCN6004 now PVData is stored is single, dont need multiplier
Public VideoCaptureDevice As Long 'PCN2395 Defines the video capture index
'vvvv PCN3067 ***********************************
'Report margins
Public ReportMarginTop As Integer
Public ReportMarginBottom As Integer
Public ReportMarginLeft As Integer
Public ReportMarginRight As Integer
'^^^^ *******************************************

Public Declare Function HASP_Lock_Validate Lib "HASPLOCK.dll" () As Double 'PCN3197
Public HASPLockActive As Boolean    'PCN3197
'vvvv PCN3*** ********************************
Public DrawShapeType As String ' Draw Shape Type, Circle, Egg, HorseShoe or InvertedHorseShow

'Horse Shoe shape parameters
Public Const HorseShoeAngleSlides As Double = (24.5 / 180)
Public Const HorseShoeAngleBottom As Double = (66 / 180)

' The following types used to define the reference shapes
' Type define for arc shape, Radians with 0rad at East, Counter clockwise
Type ShapeArc_V10
    OriginX As Single
    OriginY As Single
    Radius As Single
    StartAngle As Single
    EndAngle As Single
    Colour As Long
End Type
' Type define for line shape, start x,y coord and end x,y coord
Type ShapeLine_V10
    StartX As Single
    StartY As Single
    EndX As Single
    EndY As Single
    Colour As Long
End Type


' Type define for a shape, reason for Use defined is some pipes have a different shape
' for external and intenal pipe shape
Type ReferenceShape_V10
    Name As String * 256 'This is the name of the shape type eg Egg, Circle, HorseShoe etc
    Use As String * 256 'Type of use, eg "Internal", "External", "All" etc
    Arcs(128) As ShapeArc_V10
    NoArcs As Long
    Lines(128) As ShapeLine_V10
    NoLines As Long
    CentreOffsetX As Single
    CentreOffsetY As Single
    Colour As Long
End Type

Type ShapePolyLine_V10
    Lines(128) As ShapeLine_V10
    NoLines As Long
    Colour As Long
End Type


Public ReferenceShape() As ReferenceShape_V10
Public SemiEllipticalType As Integer 'PCN3055

'(9 March 2005, Antony van Iersel) ''''''
'PCN3373   Setting up the windows functions needed to driect memory draw onto
'          a picture box                '
Public Declare Function VarPtrArray Lib "msvbvm50.dll" Alias "VarPtr" (Ptr() As Any) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

  

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Const SRCCOPY = &HCC0020 ' (DWORD) dest = source

Public Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Public Type SAFEARRAYBOUND             '
    cElements As Long                   '
    lLbound As Long                     '
End Type                                '
                                        '
Public Type SAFEARRAY1D                '
    cDims As Integer                    '
    fFeatures As Integer                '
    cbElements As Long                  '
    cLocks As Long                      '
    pvData As Long                      '
    Bounds(0 To 0) As SAFEARRAYBOUND    '
End Type                                '
                                        
Public Type SAFEARRAY2D                '
    cDims As Integer                    '
    fFeatures As Integer                '
    cbElements As Long                  '
    cLocks As Long                      '
    pvData As Long                      '
    Bounds(0 To 1) As SAFEARRAYBOUND    '
End Type                                '
                                        '
Public Type BITMAP                     '
    bmType As Long                      '
    bmWidth As Long                     '
    bmHeight As Long                    '
    bmWidthBytes As Long                '
    bmPlanes As Integer                 '
    bmBitsPixel As Integer              '
    bmBits As Long                      '
End Type                                '
                                        '
'''''''''''''''''''''''''''''''''''''''''
'PCNSingleInstance
Private Declare Sub SwitchToThisWindow Lib "user32" (ByVal hwnd As Long, ByVal fAltTab As Boolean)
Private Declare Function EnumWindows& Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long)
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function IsWindowVisible& Lib "user32" (ByVal hwnd As Long)
Private Declare Function GetParent& Lib "user32" (ByVal hwnd As Long)
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd&, ByVal wMsg&, wParam As Any, lParam As Any) As Long

Public Const EM_GETLINECOUNT = &HBA
Public Const EM_FMTLINES = &HC8
Public Const WM_USER As Long = &H400
Public Const EM_SETTARGETDEVICE As Long = WM_USER + 72

Private sPattern As String, hFind As Long

Public CalibrationMethodActioned As String 'PCN4194 - Is set when a normal calibration has been performed/actioned. This is a flag that must be set before a vertical calibration is done.

Public UserTitleAnalysis As String      'PCN
Public UserTitleObservations As String
Public UserTitleSummary As String
Public UserTitleProfile As String
Public UserTitle1KFlat As String
Public UserTitle1KOvalityFlat As String

Public SonarIsRegistered As Boolean


Sub Main()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Main Sub  Geoff Logan geofflogan@cbsys.co.nz
'
'Revision history"
'   V0.0    Geoff Logan,    4/11/02     Building initial framework
'
'Description:
'
'
'
'
'Purpose:
'
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim llResult As Long 'PCNSingleInstance
''Dim ProductName As String
Dim ErrorStr As String
''Dim INFFile As String
''Dim INFReadONLY As Boolean
Dim strLoadPVDFile As String
Dim HaspLockFile As String
Const MAX_PATH = 260

Set PipelineDetails.CommonDialog1 = CreateObjectFromFile("COMDLG32.ocx", "CommonDialog1")
If PipelineDetails.CommonDialog1 Is Nothing Then Debug.Print " couldnt create object" Else Debug.Print "Created"

ConfigInfo.IPProcessMethod = "Type2"
ConfigInfo.IPEnhancement = "Standard"


PossibleConfigInfoCurruption = False
DesignGradient = 0 'PCN6165
SeaLevelStartHeight = 0 'PCN6128
SeaLevelEndHeight = 0 'PCN6128
FlatShading = 0 ' PCN4974
FlatOvality = False
ShiftOn = False
LockedDonut = False
LoadVideo = True
MedianFlat = False ' PCN4974
ReadOnlyAppPath = App.Path & "\" 'PCN2123
ReDim GraphInfoContainer(PVOvalitySmooth).DataSingle(0) 'PCN9999
ReDim GraphInfoContainer(PVXDiameterSmooth).DataSingle(0)
ReDim GraphInfoContainer(PVYDiameterSmooth).DataSingle(0)
ReDim GraphInfoContainer(PVMedianDiameterSmooth).DataSingle(0)
ReDim GraphInfoContainer(PVMaxDiameterSmooth).DataSingle(0)
ReDim GraphInfoContainer(PVMinDiameterSmooth).DataSingle(0)
ReDim GraphInfoContainer(PVCapacitySmooth).DataSingle(0)
'PCN6458 ReDim GraphInfoContainer(PVInclinationSmooth).DataSingle(0) 'PCN6128




' get the path of the \TEMP directory
WindowsTempDirectory = Space$(MAX_PATH)
GetTempPath Len(WindowsTempDirectory), WindowsTempDirectory
' trim off characters in excess
WindowsTempDirectory = Left$(WindowsTempDirectory, InStr(WindowsTempDirectory & vbNullChar, vbNullChar) - 1)
'MkDir WindowsTempDirectory & "CBS" & TrimOffStartPathCharacters(App.Path) 'ID4834 every INI needs its own temp directory
WindowsTempDirectory = WindowsTempDirectory & "CBS" & TrimOffStartPathCharacters(App.Path) & "\" 'ID4601
Call CreateDirectory(WindowsTempDirectory) 'ID4601

' Add demultiplexor
Dim Res
Dim RegisterMPGMultiplexer As String
Dim RegisterOCX As String
RegisterMPGMultiplexer = """" & App.Path & "\empgdmx.ax"""
Res = Shell("Regsvr32 /s " & RegisterMPGMultiplexer, vbNormalFocus)

'PCN6027 added stuff below
RegisterOCX = """" & App.Path & "\RICHTX32.OCX"""
Res = Shell("Regsvr32 /s " & RegisterOCX, vbNormalFocus)

RegisterOCX = """" & App.Path & "\COMDLG32.OCX"""
Res = Shell("Regsvr32 /s " & RegisterOCX, vbNormalFocus)

RegisterOCX = """" & App.Path & "\MSCOMCTL.OCX"""
Res = Shell("Regsvr32 /s " & RegisterOCX, vbNormalFocus)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Call GetSystem32Dir 'PCN4332

'Load any translation for the language file
'Call LoadLanguageFromFile(ErrorStr) 'PCN4171




Call LoadLanguageFromFile(ErrorStr) 'PCN4171

If App.PrevInstance Then
'    llResult = FindWindowWild("ClearLine Profiler", False)
'    Call SwitchToThisWindow(llResult, False)
'
'    Load TextToBeSentForm
'
'    TextToBeSentForm.TextToBeSentText.LinkTopic = "ClearLine Profiler|DDE_AddList"
'    TextToBeSentForm.TextToBeSentText.LinkMode = vbLinkManual
'    TextToBeSentForm.TextToBeSentText.LinkExecute Command()
'
'    TextToBeSentForm.TextToBeSentText.LinkMode = vbLinkNone
    
    Call MsgBox(DisplayMessage("This Profiler is already running!")) 'ID4834


    Exit Sub
End If


'vvvv PCN3809 *******************************************
'Check what configuration the software is to be setup in.
'Full or Reader
'This software is defined as a Reader when the hasplock.dll
'does not exist in the application directory
HaspLockFile = Dir(App.Path & "\hasplock.dll")
If LCase(HaspLockFile) <> "hasplock.dll" Then
    SoftwareConfiguration = "Reader"
    Call LoadLanguageFromFile(ErrorStr) 'PCN4171
    Call SetupSoftwareForReaderConfiguration
'    Call SetupSoftwareForFullConfiguration(strLoadPVDFile)

    Call LoadCoreForms

    DoEvents
    Call ControlsScreen.SetupForReaderConfiguration

    Call ReaderLoadPVD
    
Else
    SoftwareConfiguration = "Full"
    Call LoadLanguageFromFile(ErrorStr) 'PCN4171
    Call SetupSoftwareForFullConfiguration(strLoadPVDFile)
    Call LoadCoreForms
    DoEvents
    Call CheckAndSetupInterface(strLoadPVDFile)
    
End If
'^^^^ ***************************************************

'Load any translation for the language file
Call LoadLanguageFromFile(ErrorStr) 'PCN4171



If MeasurementUnits = "mm" Then
    ControlsScreen.ControlsReports(5).ToolTipText = ControlsScreen.Label1kReport.Caption
Else
    ControlsScreen.ControlsReports(5).ToolTipText = ControlsScreen.Label1mlReport.Caption
End If



Exit Sub
Err_Handler:
    Select Case Err
        Case 75: Resume Next
        Case 286: Resume Next 'PCNANT
        Case 52 'Bad file or filename 'PCNGL091202
            Resume Next 'PCNGL091202
        Case Else
            MsgBox Err & "-ST1:" & Error$
    End Select
End Sub


Sub GetINI_Information(MyFile As String)

On Error GoTo Err_Handler

' MGR 17/10/2002
' Read INI File and populate LOGO & Contractor Details

' Get INI file from current directory and load into memory

Dim ConfigLine As String
Dim X As Long
Dim Y As Long
Dim Parameter As String
Dim value As String
Dim SectionHead As String
Dim SectionDetail As String
Dim PathName As String
Dim FileName As String
Dim HaltApp As Boolean
Dim INIRev As Double 'PCN2443

INIRev = 0 'PCN2443
X = 1
SectionDetail = "***"
   
Config_LineCnt = 0

Dim FileNo As Integer
FileNo = FreeFile


Open MyFile For Input As #FileNo
Do While Not EOF(FileNo)
 Line Input #FileNo, ConfigLine
 Config_LineCnt = Config_LineCnt + 1
Loop
Close #FileNo

ReDim ConfigArray(Config_LineCnt)
  
FileNo = FreeFile
  
Open MyFile For Input As #FileNo
Do While Not EOF(FileNo)
 Line Input #FileNo, ConfigLine
 ConfigArray(X) = ConfigLine
 X = X + 1
Loop
Close #FileNo

' Run Through Entire Array and Validate Paths and Files

For Y = 1 To Config_LineCnt
  'vvvv PCN2443 *******************************************
  SectionHead = ConfigArray(Y)
  If SectionHead = "[Revision]" Then
      Y = Y + 1
      SectionDetail = ConfigArray(Y)
      Do Until SectionDetail = "" Or Left(SectionDetail, 1) = "[" Or SectionDetail = "\\" 'End marker of data
        If SectionDetail <> "" Then Call GetParam(SectionDetail, Parameter, value, PathName, FileName)
        Select Case Parameter
            Case "INIRevision="
                INIRev = SafeCDbl(value)
            Case Else
        End Select
        Y = Y + 1
        SectionDetail = ConfigArray(Y)
      Loop
  End If
  '^^^^ **************************************************
  SectionHead = ConfigArray(Y)
  If SectionHead = "[Company Information]" Then
     Y = Y + 1
     SectionDetail = ConfigArray(Y)
     Do Until SectionDetail = "" Or Left(SectionDetail, 1) = "[" Or SectionDetail = "\\" 'End marker of data
      If SectionDetail <> "" Then Call GetParam(SectionDetail, Parameter, value, PathName, FileName)
      Select Case Parameter
        Case "CompanyName="
          CompanyName = value
        Case "PhoneNo="
          PhoneNo = value
        Case "FaxNo="
          FaxNo = value
        Case "MeasurementUnits="
          MeasurementUnits = value
          ConfigInfo.Units = MeasurementUnits 'PCNGL150103
        Case "CompanyLogoPath="
        Call ValidatePath(value, PathName, FileName, HaltApp, Parameter) 'PCN294
          ConfigArray(Y) = Parameter & value
          CompanyLogoPath = value
        Case "CalibrationDistance="
          Calibration = SafeCDbl(value) 'PCN4616
          'ClearLineScreen.CalLen = Calibration
          CalLen_Global = Calibration
        Case "CalibrationLineLength="
          'ClearLineScreen.CalLength_tmp = Val(Value)
          CalLength_Global = SafeCDbl(value) 'PCN4616
          If CalLen_Global <> 0 And CalLength_Global <> 0 Then
'            ClearLineScreen.Ratio = Round(CalLen_Global / CalLength_Global, 3)
            'ConfigInfo.Ratio = Round(CalLen_Global / CalLength_Global, 3) 'PCN3035
            ConfigInfo.Ratio = CalLen_Global / CalLength_Global 'PCN3035 'PCN3640
            CalLineExist = True
          Else
'            ClearLineScreen.Ratio = 0.25 'PCN2193
            ConfigInfo.Ratio = 0.25 'PCN3035
          End If
        Case "RecordMode="
            RecordMode = value
        Case "TuningStyle="
            TuningStyle = value
        'vvvv PCNGL240403-1 ********************
        Case "IgnoreX1="
            IgnoreX1 = SafeCDbl(value)   'PCN2417 'PCN4616
        Case "IgnoreY1="
            IgnoreY1 = SafeCDbl(value)   'PCN2417 'PCN4616
        Case "IgnoreX2="
            IgnoreX2 = SafeCDbl(value)   'PCN2417 'PCN4616
        Case "IgnoreY2="
            IgnoreY2 = SafeCDbl(value)   'PCN2417 'PCN4616
        '^^^^ **********************************
        'vvvv PCN2639 ********************
        Case "IgnoreDistX1="
            IgnoreDistX1 = SafeCDbl(value)   'PCN4616 'PCN4616
        Case "IgnoreDistY1="
            IgnoreDistY1 = SafeCDbl(value)   'PCN2417 'PCN4616
        Case "IgnoreDistX2="
            IgnoreDistX2 = SafeCDbl(value)   'PCN2417 'PCN4616
        Case "IgnoreDistY2="
            IgnoreDistY2 = SafeCDbl(value)   'PCN2417 'PCN4616
        '^^^^ **********************************
        'vvvv PCNGL240403-1 ********************
'        Case "IPX="
'            ConfigInfo.PVShapeCentreX = Val(value)  'PCN2417 'PCN2820 'PCN4336
'        Case "IPY="
'            ConfigInfo.PVShapeCentreY = Val(value)  'PCN2417 'PCN2820 'PCN4336
        Case "IPGT="
            ConfigInfo.IPGradThres = SafeCDbl(value) 'PCN2417 'PCN2820 'PCN4616
        Case "IPDX="
            ConfigInfo.IPStDX = SafeCDbl(value) 'PCN2417 'PCN2820 'PCN4616
        Case "IPDY="
            ConfigInfo.IPStDY = SafeCDbl(value)  'PCN2417 'PCN2820 'PCN4616
        '^^^^ **********************************
        Case "PVGraphYRatio=" 'PCN2121
            PVGraphYRatio = SafeCDbl(value) 'PCN2121 'PCN2417 'PCN4616
        'vvvv PCN2773***************************
        Case "ProcessMethod="
            IPEnhancementAndIPProcessMethod.IPProcessMethod = value 'PCN2820
'            If ProcessMethod = "Type1" Then
'                ImageProcess.PVMethodType1.Tag = 1
'                Call setmethod(0) 'Type 1
'            Else
'                ImageProcess.PVMethodType2.Tag = 1
'                Call setmethod(1) 'Type 2
'            End If
        Case "Contrast="
            ConfigInfo.IPZone = SafeCDbl(value) 'PCN2820 'PCN4616
        Case "Enhancement="
            IPEnhancementAndIPProcessMethod.IPEnhancement = value 'PCN2820 'ID5395
            
        'vvvv PCN2769 ***********************************************
        Case "LimitCapMin=":  ConfigInfo.LimitCapacityL = SafeCDbl(value) 'PCN2820
        Case "LimitCapMax=":  ConfigInfo.LimitCapacityR = SafeCDbl(value) 'PCN2820
        'vvvv PCN4349 **************************
        Case "LimitOvalL=":   OvalityLimitL = SafeCDbl(value)
        Case "LimitOvalR=":   ConfigInfo.LimitOvality = SafeCDbl(value) 'PCN2820
                              OvalityLimitR = ConfigInfo.LimitOvality
        Case "LimitXYDiameterL=": ConfigInfo.LimitXYDiameterL = SafeCDbl(value)
        Case "LimitXYDiameterR=": ConfigInfo.LimitXYDiameterR = SafeCDbl(value)
        Case "LimitDiameterMaxL=": DiameterMaxLimitL = SafeCDbl(value)
        Case "LimitDiameterMaxR=": DiameterMaxLimitR = SafeCDbl(value)
        Case "LimitDiameterMinL=": GraphInfoContainer(PVMinDiameter).LimitL = SafeCDbl(value)
        Case "LimitDiameterMinR=": GraphInfoContainer(PVMinDiameter).LimitR = SafeCDbl(value)
        Case "LimitDiameterMedianL=": DiameterMedianLimitL = SafeCDbl(value)
                                      ConfigInfo.LimitDeltaL = DiameterMedianLimitL
        Case "LimitDiameterMedianR=": DiameterMedianLimitR = SafeCDbl(value)
                                      ConfigInfo.LimitDeltaR = DiameterMedianLimitR
        
            
'        Case "LimitDeltaMin="
'            ConfigInfo.LimitDeltaL = Val(value) 'PCN2820
'        Case "LimitDeltaMax="
'            ConfigInfo.LimitDeltaR = Val(value) 'PCN2820
        '^^^^ ***********************************
        '^^^^ *******************************************************
        'vvvv PCN2829 ***********************************************
        Case "PVGraphCapacityXScale=": PVGraphCapacityXScale = SafeCDbl(value)
        Case "PVGraphOvalityXScale=":  PVGraphOvalityXScale = SafeCDbl(value)
        Case "PVGraphDeltaXScale=":    PVGraphDeltaXScale = SafeCDbl(value)
        Case "PVGraphXYDiaXScale=":    PVGraphXYDiaXScale = SafeCDbl(value)
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''' PCN3540
        Case "PVGraphDiaMaxMinXScale=": PVGraphDiaMaxMinXScale = SafeCDbl(value)
        Case "PVGraphDiaMedianXScale=": PVGraphDiaMedianXScale = SafeCDbl(value)
        Case "PVGraphDiaMaxXScale=": PVGraphDiaMaxXScale = SafeCDbl(value)
        Case "PVGraphDiaMinXScale=": GraphInfoContainer(PVMinDiameter).XScale = SafeCDbl(value)
'PCN6458         Case "PVGraphInclinationXScale=": GraphInfoContainer(PVInclination).XScale = SafeCDbl(value) 'PCN6128
'PCN6458         If GraphInfoContainer(PVInclination).XScale = 0 Then GraphInfoContainer(PVInclination).XScale = 20
'        Case "PVGraphFractileXScale=": PVGraphFractileXScale = Val(value) 'PCN4235
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        '^^^^ *******************************************************
        'vvvv PCN3402 ***********************************************
        Case "PVGraphCapacityXOffset=": PVGraphCapacityXOffset = SafeCDbl(value)
        Case "PVGraphOvalityXOffset=":  PVGraphOvalityXOffset = SafeCDbl(value)
        Case "PVGraphDeltaXOffset=":    PVGraphDeltaXOffset = SafeCDbl(value)
        Case "PVGraphXYDiaXOffset=":    PVGraphXYDiaXOffset = SafeCDbl(value)
        '^^^^ *******************************************************
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''' PCN3540
        Case "PVGraphDiaMaxMinXOffset=": PVGraphDiaMaxMinXOffset = SafeCDbl(value)
        Case "PVGraphDiaMedianXOffset=": PVGraphDiaMedianXOffset = SafeCDbl(value)
        Case "PVGraphDiaMaxXOffset=": PVGraphDiaMaxXOffset = SafeCDbl(value) 'PCN4799
        Case "PVGraphDiaMinXOffset=": GraphInfoContainer(PVMinDiameter).XOffset = SafeCDbl(value)
'PCN6458         Case "PVGraphInclinationXOffset=": GraphInfoContainer(PVInclination).XOffset = SafeCDbl(value) 'PCN6128
        
        
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
'        Case "PVGraphFractileXOffset=": PVGraphFractileXOffset = Val(value) 'PCN4235
        
        'vvvv PCN2980 ************************
        Case "PVDiameterMethod="
            PVDiameterMethod = value
        '^^^^ ********************************
        Case "HASPLock="
            HASPLockActive = IIf(value = "true", True, False) 'PCN3197
        Case Else
      End Select
      Y = Y + 1
      SectionDetail = ConfigArray(Y)
    Loop
  End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'SectionHead = ConfigArray(Y)
'  If SectionHead = "[GraphSubTittle]" Then
'     Y = Y + 1
'     SectionDetail = ConfigArray(Y)
'     Do Until SectionDetail = "" Or Left(SectionDetail, 1) = "[" Or SectionDetail = "\\" 'End marker of data
'      If SectionDetail <> "" Then Call GetParam(SectionDetail, Parameter, value, PathName, FileName)
'      Select Case Parameter
'        Case "Summary_Flat=":           Summary_Flat = value
'        Case "Summary_MedianDiameter=": Summary_MedianDiameter = value
'        Case "Summary_Ovality=":        Summary_Ovality = value
'        Case "Summary_MaxDiameter=":    Summary_MaxDiameter = value
'        Case "Summary_XYDaimeter=":     Summary_XYDaimeter = value
'        Case "Summary_Capacity=":       Summary_Capacity = value
'        Case "Summary_Debris=":         Summary_Debris = value
'
'        Case "Analysis_Flat=":          Analysis_Flat = value
'        Case Else
'      End Select
'      Y = Y + 1
'      SectionDetail = ConfigArray(Y)
'    Loop
'  End If
  
  
  
  
  
  
'
  
'Summary_Flat=
'Summary_MedianDiameter=
'Summary_Ovality=
'Summary_MaxDiameter=
'Summary_XYDaimeter=
'Summary_Capacity=
'Summary_Debris=

'Analysis_Flat=
'Analysis_MedianDiameter=
'Analysis_Ovality=
'Analysis_MaxDiameter=
'Analysis_XYDaimeter=
'Analysis_Capacity=
'Analysis_Debris=
'Profile_Flat=
'Profile_MedianDiameter=
'Profile_Ovality=
'Profile_MaxDiameter=
'Profile_XYDaimeter=
'Profile_Capacity=
'Profile_Debris=
'Observations_Flat=
'Observations_MedianDiameter=
'Observations_Ovality=
'Observations_MaxDiameter=
'Observations_XYDaimeter=
'Observations_Capacity=
'Observations_Debris=
  
  SectionHead = ConfigArray(Y)
  If SectionHead = "[MTColors]" Then
     Y = Y + 1
     SectionDetail = ConfigArray(Y)
     Do Until SectionDetail = "" Or Left(SectionDetail, 1) = "[" Or SectionDetail = "\\" 'End marker of data
      If SectionDetail <> "" Then Call GetParam(SectionDetail, Parameter, value, PathName, FileName)
      Select Case Parameter
        Case "NormalDrawingColor="
            NormalDrawingColor = Val(value) 'PCN1931
        Case "SelectedObjectColor="
            SelectedObjectColor = Val(value) 'PCN1931
        Case "ModiCircleColor="
            ModiCircleColor = Val(value) 'PCN1931
        Case "ChosenModiCircleColor="
            ChosenModiCircleColor = Val(value) 'PCN1931
        Case "AreaFillingColor="
            AreaFillingColor = Val(value) 'PCN1931
        Case "ExtraObjectColor="
            ExtraObjectColor = Val(value) 'PCN1931
        Case "JointCircleColor="
            JointCircleColor = Val(value) 'PCN1931
        Case "TempDrawingColor="
            TempDrawingColor = Val(value) 'PCN1931
        Case "MovingObjectColor="
            MovingObjectColor = Val(value) 'PCN1931
        Case "RotatingObjectColor="
            RotatingObjectColor = Val(value) 'PCN1931
        Case "SelectionBoundaryColor="
            SelectionBoundaryColor = Val(value) 'PCN1931
        Case "TextSizeIndicatorColor="
            TextSizeIndicatorColor = Val(value) 'PCN1931
        Case Else
      End Select
      Y = Y + 1
      SectionDetail = ConfigArray(Y)
    Loop
  End If
  'PCN2111--------------------------------------------v
  SectionHead = ConfigArray(Y)
  If SectionHead = "[Regional Options]" Then
      Y = Y + 1
      SectionDetail = ConfigArray(Y)
      Do Until SectionDetail = "" Or Left(SectionDetail, 1) = "[" Or SectionDetail = "\\" 'End marker of data
        If SectionDetail <> "" Then Call GetParam(SectionDetail, Parameter, value, PathName, FileName)
        Select Case Parameter
            Case "Language="
                Language = value
            Case "ThreeDRenderingStyle="
                If value = "Software" Then value = 1
                If value = "Hardware" Then value = 0
                If value <> 0 And value <> 1 Then value = 1 'PCN4197 It can only be 0 or 1 now, that is 0 for auto, 1 for software
                ThreeDRenderingStyle = value 'PCN2266
'            Case "CaptureDevice="
'                CaptureDevice = value 'PCN2289
            'vvvv PCN3024 ******************************
            Case "PaperSize="
                Select Case value
                    Case "Letter"
                        Printer.PaperSize = vbPRPSLetter
                    Case "A4"
                        Printer.PaperSize = vbPRPSA4
                    Case "A5"
                        Printer.PaperSize = vbPRPSA5
                    Case Else
                        Printer.PaperSize = vbPRPSA4
                End Select
            '^^^^ **************************************
            'vvvv PCN3069 ******************************
            Case "ReportMarginTop="
                ReportMarginTop = SafeCDbl(value)
            Case "ReportMarginBottom="
                ReportMarginBottom = SafeCDbl(value)
            Case "ReportMarginLeft="
                ReportMarginLeft = SafeCDbl(value)
            Case "ReportMarginRight="
                ReportMarginRight = SafeCDbl(value)
            '^^^^ **************************************
            Case Else
        End Select
        Y = Y + 1
        SectionDetail = ConfigArray(Y)
      Loop
  End If '-------------------------------------------^
  'vvvv PCN2392 *******************************************
  SectionHead = ConfigArray(Y)
  If SectionHead = "[Fish Eye Distortion]" Then
      Y = Y + 1
      SectionDetail = ConfigArray(Y)
      Do Until SectionDetail = "" Or Left(SectionDetail, 1) = "[" Or SectionDetail = "\\" 'End marker of data
        If SectionDetail <> "" Then Call GetParam(SectionDetail, Parameter, value, PathName, FileName)
        Select Case Parameter
            Case "Fish_Distortion=": ConfigInfo.FishEyeDistortion = SafeCDbl(value)
            'Case "Fish_DistortionHorizontal": ConfigInfo.FishEyeDistortion = Val(value) 'PCN3687
            Case "Fish_DistortionHorizontal": ConfigInfo.FishEyeHorDistortion = SafeCDbl(value) 'PCN3687
            Case "Fish_Ratio=" 'PCN2495 --------------v
                ConfigInfo.FishEyeRatio = SafeCDbl(value)
                If ConfigInfo.FishEyeRatio = 0 Then
                    ConfigInfo.FishEyeRatio = 1
                End If
            '-----------------------------------------^
            'vvvv PCN2497 **********************************
            Case "Fish_CenterX=": ConfigInfo.FishEyeCenterX = SafeCDbl(value)
            Case "Fish_CenterY=": ConfigInfo.FishEyeCenterY = SafeCDbl(value)
            '^^^^ ******************************************
            'vvvv PCN3019 **********************************
            Case "Fish_OriginalWidth=": ConfigInfo.FishEyeOriginalWidth = SafeCDbl(value)
            Case "Fish_OriginalHeight=": ConfigInfo.FishEyeOriginalHeight = SafeCDbl(value)
            '^^^^ ******************************************
            'vvvv PCN3031 **********************************
            Case "Fish_Displayed=": FisheyeDisplayed = IIf(value = "True", True, False)
            '^^^^ ******************************************
            Case Else
        End Select
        Y = Y + 1
        SectionDetail = ConfigArray(Y)
      Loop
  End If
  
  '^^^^ **************************************************
  'vvvv PCN2639 *******************************************
  SectionHead = ConfigArray(Y)
  If SectionHead = "[Automatic Distance]" Then
      Y = Y + 1
      SectionDetail = ConfigArray(Y)
      Do Until SectionDetail = "" Or Left(SectionDetail, 1) = "[" Or SectionDetail = "\\" 'End marker of data
        If SectionDetail <> "" Then Call GetParam(SectionDetail, Parameter, value, PathName, FileName)
        Select Case Parameter
            Case "DistanceMethod="
                DistanceMethod = value
            Case Else
        End Select
        Y = Y + 1
        SectionDetail = ConfigArray(Y)
      Loop
  End If
  '^^^^ **************************************************
  'vvvv PCN2395 *******************************************
  SectionHead = ConfigArray(Y)
  If SectionHead = "[Video Settings]" Then
      Y = Y + 1
      SectionDetail = ConfigArray(Y)
      Do Until SectionDetail = "" Or Left(SectionDetail, 1) = "[" Or SectionDetail = "\\" 'End marker of data
        If SectionDetail <> "" Then Call GetParam(SectionDetail, Parameter, value, PathName, FileName)
        Select Case Parameter
            Case "VideoCaptureDevice="
                VideoCaptureDevice = Val(value)
            Case Else
        End Select
        Y = Y + 1
        SectionDetail = ConfigArray(Y)
      Loop
  End If
  '^^^^ **************************************************
   SectionDetail = "***" ' reset to NON blank
Next Y

'vvvv PCN2443 *******************************************
Dim strTemp As String
If INIRev <> INIVersion Then
    strTemp = DisplayMessage("ClearLine.ini VERSION ERROR. Expecting ")
    strTemp = strTemp & INIVersion & DisplayMessage(", ClearLine.ini is currently ")
    strTemp = strTemp & Format(INIRev, "###0.0") & DisplayMessage(" - This application may not work as designed.")
    MsgBox strTemp, vbCritical
End If
'^^^^ **************************************************

Exit Sub
Err_Handler:
Select Case Err
    Case 9 'Subscript out of range (end of file)
        Exit Sub
    Case Else
        MsgBox Err & "-ST2:" & Error$
End Select
End Sub

Function GetINI_ParameterInfoOnly(MyFile As String, Parameter As String, ReturnValue As String)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : GetINI_ParameterInfoOnly
'Created : 23 March 2004, PCN2639
'Updated :
'Prg By  : Geoff Logan
'Param   : MyFile - location of the INI file
'          Parameter - search for this parameter in the INI
'          ReturnValue - this is the value in the INI given to the parameter requested
'Desc    : Gets from the INI file a value for a specific parameter
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim ConfigLine As String
Dim PathName As String
Dim FileName As String
Dim FileParameter As String 'PCN3019
Dim FileReturnValue As String 'PCN3019


ReturnValue = ""

If MyFile = "" Then Exit Function 'PCN3809

If Dir(MyFile) = "" Then Exit Function

Dim FileNo As Integer
FileNo = FreeFile

Open MyFile For Input As #FileNo
Do While Not EOF(FileNo)
    Line Input #FileNo, ConfigLine
    If Len(ConfigLine) <> 0 Then
        'vvvv PCN3019 *****************************
        Call GetParam(ConfigLine, FileParameter, FileReturnValue, PathName, FileName)
        If Parameter = FileParameter Then
            ReturnValue = FileReturnValue
            Exit Do
        End If
        '^^^^ *************************************
    End If
Loop
Close #FileNo


Exit Function
Err_Handler:
Select Case Err
    Case 9 'Subscript out of range (end of file)
        Exit Function
    Case Else
        MsgBox Err & "-ST3:" & Error$
End Select
End Function
Function GetParam(ByVal MyString, Param, value, Optional PathName, Optional FileName)

On Error GoTo Err_Handler

Dim Loc As Integer, X As Integer

'path = False

Loc = InStr(MyString, "=")
If Loc <> 0 Then
  Param = Left(MyString, Loc)
  value = Trim(Mid(MyString, Loc + 1))
End If

If value = "" Then
  PathName = ""
  FileName = ""
  Exit Function
End If

If InStr(Param, "Path") <> 0 Then
  MyString = ""
  Do Until InStr(MyString, "\") <> 0
    MyString = Right(value, X)
    X = X + 1
  Loop
  PathName = Left(value, Len(value) - (X - 2))
  FileName = Right(value, (X - 2))
End If

Exit Function
Err_Handler:
MsgBox Err & "-ST4:" & Error$

End Function
Function ValidatePath(value, PathName, Optional FileName, Optional HaltApp, Optional ByVal Param)

On Error GoTo Err_Handler

Dim X As String
Dim retvalue As String 'PCN1916
Dim msg As String
Dim LogoLoadFail As Boolean 'PCNGL160103

LogoLoadFail = True   'PCNGL160103

If value <> "" Then 'PCNGL160103
    X = Dir(value)
    If X <> "" Then
        LogoLoadFail = False
        Exit Function
    End If
End If

LogoLoadHasFailed:
'    MsgBox DisplayMessage("Your Company Logo could not be found - please remember to load it.") 'PCNGL261202 'PCN2111 'PCNGL200803
    value = "" 'PCNGL261202


Exit Function
Err_Handler:
Select Case Err
    Case 52 'Bad file name (eg Z:\ when not connected to the network) 'PCNGL160103
        GoTo LogoLoadHasFailed
    Case Else
        MsgBox Err & "-ST5:" & Error$
End Select
End Function
Function DisplayDialog(Title, InitDir, value, Filter, retvalue)

On Error GoTo Err_Handler

'ClearLineProfilerV6.Dialog.CancelError = True
'ClearLineProfilerV6.Dialog.FileName = value
'ClearLineProfilerV6.Dialog.DialogTitle = Title
'ClearLineProfilerV6.Dialog.InitDir = InitDir
'ClearLineProfilerV6.Dialog.Filter = Filter
'ClearLineProfilerV6.Dialog.ShowOpen
'
'retvalue = ClearLineProfilerV6.Dialog.FileName

Exit Function

CancelPressed:
 retvalue = ""

Exit Function

Err_Handler:
Select Case Err
Case 32755
  Resume CancelPressed
Case Else
  MsgBox Err & "-ST6:" & Error$
End Select

End Function
Function INI_WriteBack(MyFile, Optional Parameter, Optional NewVal)
On Error GoTo Err_Handler
Dim X As Long 'PCN1916
Dim Y As Long
Dim SectionDetail As String
Dim ConfigLine As String

If SoftwareConfiguration = "Reader" Then Exit Function 'PCN???? somethnigs still try to write to INI

X = 1
SectionDetail = "***"
   
Config_LineCnt = 0

Dim FileNo As Integer

FileNo = FreeFile


Open MyFile For Input As #FileNo
Do While Not EOF(FileNo)
 Line Input #FileNo, ConfigLine
 Config_LineCnt = Config_LineCnt + 1
Loop
Close #FileNo

FileNo = FreeFile


ReDim ConfigArray(Config_LineCnt)
  
Open MyFile For Input As #FileNo
Do While Not EOF(FileNo)
 Line Input #FileNo, ConfigLine
 ConfigArray(X) = ConfigLine
 X = X + 1
Loop
Close #FileNo

' If param then find and replace in Array
If Not IsMissing(Parameter) Then
 'For Y = 1 To UBound(ConfigArray) - 1
 For Y = 1 To Config_LineCnt
    If InStr(ConfigArray(Y), Parameter) <> 0 And InStr(ConfigArray(Y), Parameter) < 2 Then
        ConfigArray(Y) = Parameter & NewVal
    End If
 Next Y
End If

' write value back to Config array


Open MyFile For Output As #FileNo
'For Y = 1 To UBound(ConfigArray) - 1
For Y = 1 To Config_LineCnt
  If ConfigArray(Y) = "\\" Then
    Print #FileNo, ConfigArray(Y)
    Exit For
  Else
    Print #FileNo, ConfigArray(Y)
  End If
Next Y
Close #FileNo

Exit Function
TidyUp:
Close #FileNo


Exit Function
Err_Handler:
Select Case Err
Case 55 ' File already open
    GoTo TidyUp
Case 75 ' Path / File Error
 'MsgBox DisplayMessage("Cannot Write to ClearLine.ini File. Check that file is not Write Protected."), vbExclamation 'PCN2111
 ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("Cannot Write to ClearLine.ini File. Check that file is not Write Protected."): ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
Case Else
 MsgBox Err & "-ST7:" & Error$
End Select
End Function

Function GetLineInformation(ByVal MyString, Param, value)

On Error GoTo Err_Handler

Dim Loc As Integer, X As Integer

Loc = InStr(MyString, ",")
If Loc <> 0 Then
  Param = Left(MyString, Loc - 1)
  value = Trim(Mid(MyString, Loc + 1))
End If

Exit Function
Err_Handler:
MsgBox Err & "-ST8:" & Error$

End Function


Private Function GetVal(str1, Pos) As Double
On Error GoTo Err_Handler
    Dim pos1 As Integer
    Dim i As Integer
    Dim SubStr As String
    
    pos1 = 0
    SubStr = str1
    For i = 0 To Pos
        SubStr = Right(SubStr, Len(SubStr) - pos1)
        pos1 = InStr(SubStr, ",")
        If pos1 = 0 Then
            GetVal = Val(SubStr)
        Else
            GetVal = Val(Left(SubStr, pos1 - 1))
        End If
    Next
Exit Function
Err_Handler:
Select Case Err
    Case Else
        MsgBox Err & "-ST9:" & Error$
        Resume Next
End Select

End Function
Private Function GetStr(str1, Pos) As String
On Error GoTo Err_Handler
    Dim pos1 As Integer
    Dim i As Integer
    Dim SubStr As String
    pos1 = 0
    SubStr = str1
    For i = 0 To Pos
        SubStr = Right(SubStr, Len(SubStr) - pos1)
        pos1 = InStr(SubStr, ",")
        If pos1 = 0 Then
            GetStr = SubStr
        Else
            GetStr = Left(SubStr, pos1 - 1)
        End If
    Next
Exit Function
Err_Handler:
Select Case Err
    Case Else
        MsgBox Err & "-STA:" & Error$
        Resume Next
End Select

End Function
Function PipeInfoArray_WriteBack(ImageDataFile, Optional Parameter As String, Optional NewVal)

On Error GoTo Err_Handler

' If param then find and replace in Array
If Not IsMissing(Parameter) Then
   If Left(Parameter, 1) = "[" Or Len(Parameter) = 0 Then
     PipeInfoArray(ArrayCnt) = Parameter
   Else
     PipeInfoArray(ArrayCnt) = Parameter & "," & NewVal
   End If
   ArrayCnt = ArrayCnt + 1
End If

Exit Function
Err_Handler:
Select Case Err
Case 75 ' Path / File Error
 'MsgBox DisplayMessage("Cannot Write to ClearLine.ini File. Check that file is not Write Protected."), vbExclamation 'PCN2111
 ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("Cannot Write to ClearLine.ini File. Check that file is not Write Protected."): ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
Case Else
 MsgBox Err & "-STB:" & Error$
End Select

End Function



Sub LoadCoreForms()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : LoadCoreForms
'Created : 16 December 2002,
'Updated : 18 November 2003, PCN2402
'Prg By  : Geoff Logan
'Param   :
'Desc    : In an ordered manor, load the core forms.
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

Load ClearLineProfilerV6
If SoftwareConfiguration = "Reader" Then
    ClearLineProfilerV6.Caption = ClearLineProfilerV6.Caption & " - " & DisplayMessage("Viewer")  'PCN4297
End If
ClearLineProfilerV6.WindowState = vbMaximized
ClearLineProfilerV6.Show


V6Splash.Show
DoEvents

'PCN3691 can't use string for concern of speed , 0 for normal Bitmap directdraw, 1 for picturebox draw
ScreenDrawing.ScreenDrawingType = 0
ScreenDrawing.ScreenDrawingOrientation = 0

'vvvv PCN4171 ***********************************
GraphInfoContainer(PVFlat).GraphType = "Flat"
GraphInfoContainer(PVMedianDiameter).GraphType = "MedianDiameter"
GraphInfoContainer(PVOvality).GraphType = "Ovality"
GraphInfoContainer(PVMaxDiameter).GraphType = "MaxDiameter"
GraphInfoContainer(PVMinDiameter).GraphType = "MinDiameter" 'PCN4333
GraphInfoContainer(PVXYDiameter).GraphType = "XYDiameter"
GraphInfoContainer(PVCapacity).GraphType = "Capacity"
GraphInfoContainer(PVYDiameter).GraphType = "YDiameter" 'PCN4296
GraphInfoContainer(PVDebris).GraphType = "Debris" 'PCN4461
'PCN6458 GraphInfoContainer(PVInclination).GraphType = "Inclination" 'PCN6128 added inclination graph
'PCN6458 GraphInfoContainer(PVDesignGradient).GraphType = "DesignGradient" 'PCN6178

GraphInfoContainer(PVFlat).PVXScaleUnits = ""
GraphInfoContainer(PVMedianDiameter).PVXScaleUnits = "Real"
GraphInfoContainer(PVOvality).PVXScaleUnits = "Per"
GraphInfoContainer(PVMaxDiameter).PVXScaleUnits = "Real"
GraphInfoContainer(PVMinDiameter).PVXScaleUnits = "Real" 'PCN4333
GraphInfoContainer(PVXYDiameter).PVXScaleUnits = "Per"
GraphInfoContainer(PVCapacity).PVXScaleUnits = "Per"
GraphInfoContainer(PVYDiameter).PVXScaleUnits = "Real"
GraphInfoContainer(PVDebris).PVXScaleUnits = "Real" 'PCN4461
'PCN6458 GraphInfoContainer(PVInclination).PVXScaleUnits = "Real" 'PCN6128
'PCN6458 GraphInfoContainer(PVDesignGradient).PVXScaleUnits = "Real" 'PCN6178


''PVGraphOrder(0) = "Flat"
''PVGraphOrder(1) = "MedianDiameter"
''PVGraphOrder(2) = "Ovality"
''PVGraphOrder(3) = "MaxDiameter"
''PVGraphOrder(4) = "XYDiameter"
''PVGraphOrder(5) = "Capacity"
''PVXScaleUnits(0) = ""
''PVXScaleUnits(1) = "Real"
''PVXScaleUnits(2) = "Per"
''PVXScaleUnits(3) = "Real"
''PVXScaleUnits(4) = "Per"
''PVXScaleUnits(5) = "Per"
'^^^^ *******************************************

'Disable form resizing events
StartupBypass = False

Load ClearLineScreen
ClearLineScreen.Show

V6Splash.ZOrder 0 'PCNGL130103



PVPageTop = 0
PVPageLeft = ClearLineScreen.width
'vvvv PCN2402 ************************************
If (IdentifyOperatingSystem = "Windows XP") Then
    PVPageHeight = 8800 '10550 'PCN4171
    PVPageWidth = 4400
Else
    PVPageHeight = 10730
    PVPageWidth = 4400
End If
'^^^^ ********************************************

'vvvv PCN4171 ********************************************

ClearLineTitle.Show

ControlsScreen.Show

ControlsMain.Show
'^^^^ ****************************************************
PVPageHeight = ControlsScreen.Top
'Load Observations 'PCNGL060103 'PCN4131 obs form removed
'Observations.Show 'PCN4131 obs form removed

'vvvv PCN4277 ******************************
Load OptionsPage

OptionsPage.Show
Load AutoTune

AutoTune.Show
Load PipelineDetails

PipelineDetails.Show
Load PrecisionVisionGraph

'PrecisionVisionGraph.Show
'If SoftwareConfiguration <> "Reader" Then Load SonarConfig 'PCN????

ControlsMain.MainViewSelected = 0
Call ControlsMain.ControlsDisplaySetup("DisplayPVGraph")
'^^^^ **************************************
    
'''''vvvv PCNGL300303-1 **************************
'''''Set the number of measurement decimal places
'''''For 'mm' set to 1
'''''For 'in' set to 2 (enable more accurate calibration and measurement)
''''If MeasurementUnits = "mm" Then
''''    Digits = 1
''''Else
''''    Digits = 2
''''End If
'''''^^^^ ****************************************

''''CurrentGraph = 1 ' 1:file 2:Grabedimage 'PCN4328 -removed

'MkDir App.Path & "\CBS" 'PCN2155
MkDir LocToSave & "CBS" 'PCN2155

'Disable form resizing events
StartupBypass = True

Load PVGraphsKeyForm 'PCN2990 PVGraphsKey Now a Form to float over
                     'All other forms, this was from displaying units
                     'instead of %

PVGraphsKeyForm.Visible = False 'PCN2990
PVGraphsKeyForm.Left = 7923 'PCN4920
PVGraphsKeyForm.Top = 5594
                
PVGraphsKeyForm.ZOrder 0
Call ScreenDrawing.FormTopMost(PVGraphsKeyForm.hwnd) 'PCN2990

ClearLineScreen.DimenResults.ZOrder 1
ClearLineScreen.AreaResults.ZOrder 1

ConfigInfo.FishEyeFlag = False
iFEScaleX = 10
iFESCaleY = 10

SemiEllipticalType = 0 'PCN3055 chooses what type off SemiElliptical should use
                       '0 is logans orrignal, 1 is Watercares, 2 is watercares mirror
Call InitialiseGraphStates 'PCN3373

PVScaleMarkerStFrame = 0 'PCN3373 start frame of PVScaleMarker
PVScaleMarkerFnFrame = 0 'PCN3373 start frame of PVScaleMarker
LastDataTime = 0 'PCNANT????

Exit Sub
Err_Handler:
    Select Case Err
        Case 75 'file/path access error
            Resume Next
        Case 76 'file/path access error
            Resume Next
        Case Else
            MsgBox Err & "-STC:" & Error$
    End Select
End Sub

Function SetAVIInitialised()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'SetAVIInitialised Function  Geoff Logan geofflogan@cbsys.co.nz
'
'Revision history"
'   V0.0    Geoff,    15/01/03     Building initial framework
'
'Description:
'           This function is very important for ensuring that the video functions (in C code)
'           are call when the AVI video file has been correctly initialised. This function must trap
'           most, if not all, the invalid C function inputs to minimise the likelihood of the
'           programs crashing. The C code crashes are extreme, hence the importance of this function.
'
'Purpose:
'           Set AVIInitialised after checking program operating conditions are suitable for AVI video functions.
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

'Check for a valid AVI file
If Len(VideoFileName) = 0 Then
    AVIInitialised = False
    Exit Function
ElseIf Dir(VideoFileName) = "" Then
    AVIInitialised = False
    Exit Function
Else
    AVIInitialised = True
End If

Exit Function
ErrorDetected:
    AVIInitialised = False

Exit Function
Err_Handler:
    Select Case Err
        Case 75 'file/path access error
            GoTo ErrorDetected
        Case 76 'file/path access error
            GoTo ErrorDetected
        Case Else
            MsgBox Err & "-STD:" & Error$
    End Select
End Function

Function CheckAVIInitialised() As Boolean
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'SetAVIInitialised Function  Geoff Logan geofflogan@cbsys.co.nz
'
'Revision history"
'   V0.0    Geoff,    15/01/03     Building initial framework
'
'Description:
'           This function is very important for ensuring that the video functions (in C code)
'           are call when the AVI video file has been correctly initialised. This function must trap
'           most, if not all, the invalid C function inputs to minimise the likelihood of the
'           programs crashing. The C code crashes are extreme, hence the importance of this function.
'
'Purpose:
'           Check program operating conditions are suitable for AVI video functions.
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

'Check for a valid AVI file
If Len(VideoFileName) = 0 Then
    AVIInitialised = False
    CheckAVIInitialised = False
    Exit Function
ElseIf Dir(VideoFileName) = "" Then
    AVIInitialised = False
    CheckAVIInitialised = False
    Exit Function
ElseIf AVIInitialised = False Then 'PCN1863
    CheckAVIInitialised = False 'PCN1863
    Exit Function
Else
    CheckAVIInitialised = True
End If

Exit Function
ErrorDetected:
    AVIInitialised = False
    CheckAVIInitialised = False

Exit Function
Err_Handler:
    Select Case Err
        Case 75 'file/path access error
            GoTo ErrorDetected
        Case 76 'file/path access error
            GoTo ErrorDetected
        Case Else
            'PCN 1982 LS 10/7/03
            'MsgBox DisplayMessage("Media File is not a valid file name.") 'PCN2111
            ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("Media File is not a valid file name."): ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
            VideoFileName = ""
            AVIInitialised = False
            Resume Next
            'MsgBox Err & " - " & Error$
    End Select
End Function


'Function SetMPGInitialised()
''~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
''SetMPGInitialised Function  Louise Shrimpton louiseS@cbsys.co.nz
''
''Revision history"
''   V0.0    Louise,    03/02/03     Building initial framework
''
''Description:
''           This function is very important for ensuring that the video functions (in C code)
''           are call when the MPG video file has been correctly initialised. This function must trap
''           most, if not all, the invalid C function inputs to minimise the likelihood of the
''           programs crashing. The C code crashes are extreme, hence the importance of this function.
''
''Purpose:
''           Set MPGInitialised after checking program operating conditions are suitable for MPG video functions.
''~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'On Error GoTo Err_handler
'
''Check for a valid MPG file
'If Len(MPGFileName) = 0 Then
'    MPGInitialised = False
'    Exit Function
'ElseIf Dir(MPGFileName) = "" Then
'    MPGInitialised = False
'    Exit Function
'Else
'    MPGInitialised = True
'End If
'
'Exit Function
'ErrorDetected:
'    MPGInitialised = False
'
'Exit Function
'Err_handler:
'    Select Case Err
'        Case 75 'file/path access error
'            GoTo ErrorDetected
'        Case 76 'file/path access error
'            GoTo ErrorDetected
'        Case Else
'            MsgBox Err & " - " & Error$
'    End Select
'End Function

Public Function IdentifyOperatingSystem() As String

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'IdentifyOperatingSystem() Function  Michelle Lindsay   michellelindsay@cbsys.co.nz
'
'Revision history"
'   V0.0    Michelle,    12/02/03     Building initial framework
'
'Description:
'           This function gets the operating system of the user's machine as a string
'           to enable the ClearLine Profiler screens to be optimally displayed.
'
'Purpose:
'           Get user's operating system as a string.
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  Dim rOsVersionInfo As OSVERSIONINFO
  Dim sOperatingSystem As String
  
  
  
  sOperatingSystem = "NONE"
  
  
  ' Pass the size of the structure into itself for the API call
  rOsVersionInfo.dwOSVersionInfoSize = Len(rOsVersionInfo)
  
  If GetVersionEx(rOsVersionInfo) Then
  
    Select Case rOsVersionInfo.dwPlatformId
    
      Case VER_PLATFORM_WIN32_NT
        
        If rOsVersionInfo.dwMajorVersion >= 5 Then
          If rOsVersionInfo.dwMinorVersion = 0 Then
            sOperatingSystem = "Windows 2000"
          Else
            sOperatingSystem = "Windows XP"
          End If
        Else
           sOperatingSystem = "Windows NT"
        End If
        
      Case VER_PLATFORM_WIN32_WINDOWS
        If rOsVersionInfo.dwMajorVersion >= 5 Then
           sOperatingSystem = "Windows ME"
        ElseIf rOsVersionInfo.dwMajorVersion = 4 And rOsVersionInfo.dwMinorVersion > 0 Then
           sOperatingSystem = "Windows 98"
        Else
           sOperatingSystem = "Windows 95"
        End If
        
      Case VER_PLATFORM_WIN32s
        sOperatingSystem = "Win32s"
        
    End Select
  End If

  IdentifyOperatingSystem = sOperatingSystem

End Function


Function SetCheckBoxTick(ContrlName As Control, SetON As Boolean) ', Optional SetupONLY As Boolean)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : SetCheckBoxTick
'Created : 15 March 2004, PCN2639
'Updated :
'Prg By  : Geoff Logan
'Param   : (None)
'Desc    :
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

If SetON Then
    ContrlName.Picture = LoadResPicture(104, vbResBitmap)  'CheckBoxON
    ContrlName.Tag = 1
Else
    ContrlName.Picture = LoadResPicture(103, vbResBitmap)  'CheckBoxOFF
    ContrlName.Tag = 0
End If


Exit Function
Err_Handler:
Select Case Err
    Case 438 'Object doesn't support method
        Exit Function
    Case Else
        MsgBox Err & "-STE:" & Error$
End Select
End Function

Function SavePVTuningToINI()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : SavePVTuningToINI
'Created : 11 February 2004, PCN2612
'Updated :
'Prg By  : Geoff Logan
'Param   : (None)
'Desc    : Saves the PV tuning settings to the INI file.
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

ConfigInfo.IPStDY = ConfigInfo.IPStDX 'PCN2820
'ConfigInfo.PVShapeCentreX = CalXParaYPara 'PCN2820
'ConfigInfo.PVShapeCentreY = ConfigInfo.PVShapeCentreX 'PCN2820
XAdjust = 1
'Save current settings to INI file
Call INI_WriteBack(MyFile, "IPX=", ConfigInfo.PVShapeCentreX) 'PCN2820 'PCN4336
Call INI_WriteBack(MyFile, "IPY=", ConfigInfo.PVShapeCentreY) 'PCN2820 'PCN4336
Call INI_WriteBack(MyFile, "IPGT=", ConfigInfo.IPGradThres) 'PCN2820
Call INI_WriteBack(MyFile, "IPDX=", ConfigInfo.IPStDX) 'PCN2820
Call INI_WriteBack(MyFile, "IPDY=", ConfigInfo.IPStDY) 'PCN2820
'Add further settings to INI file
Call INI_WriteBack(MyFile, "ProcessMethod=", Trim(IPEnhancementAndIPProcessMethod.IPProcessMethod)) 'PCN2820
Call INI_WriteBack(MyFile, "Contrast=", ConfigInfo.IPZone) 'PCN2820
Call INI_WriteBack(MyFile, "Enhancement=", Trim(IPEnhancementAndIPProcessMethod.IPEnhancement)) 'PCN2820
    
Exit Function
Err_Handler:
    MsgBox Err & "-STF:" & Error$
End Function

Function SetPVTuningProcessValues()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : SetPVTuningProcessValues
'Created : 12 May 2004, PCN2612
'Updated :
'Prg By  : Geoff Logan
'Param   : (None)
'Desc    : This function sets the process values in LaserLib for Precision Vision Tuning
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

Dim zoneValue As Double

If mediatype = Video Or mediatype = "Live" Or mediatype = StillImage Then
    'Method
'   PCN3017 no type1 or type2 anymore
'   If ConfigInfo.IPProcessMethod = "Type1" Then 'PCN2820
'        Call setmethod(0)
'    Else
'        Call setmethod(1) 'Type 2
'    End If
  
    zoneValue = CDbl(ConfigInfo.IPZone) 'PCN3017
    
    'Process setting
    Call setgradthreshold(ConfigInfo.IPGradThres) 'PCN2820
'    Call setstandarddeviation(ConfigInfo.PVShapeCentreX, ConfigInfo.PVShapeCentreY, ConfigInfo.IPStDX, ConfigInfo.IPStDY) 'PCN2820
    Call setstandarddeviation(ConfigInfo.IPStDX, 0, 0, 0) 'PCN3017
    'Call setprofilecontrast(ConfigInfo.IPZone) 'PCN2820 'PCN3017 remomved
    Call hough_setinsidezone(zoneValue) 'PCN3017
    Call hough_setoutsidezone(zoneValue * 2) 'PCN3017
    
    'Enhancement
    If Trim(IPEnhancementAndIPProcessMethod.IPEnhancement) = "Low" Then 'PCN2820
        Call setvideofiltertype(0)
    ElseIf Trim(IPEnhancementAndIPProcessMethod.IPEnhancement) = "High" Then 'PCN2820
        Call setvideofiltertype(1)
    ElseIf Trim(IPEnhancementAndIPProcessMethod.IPEnhancement) = "Mixed" Then 'PCN2820
        Call setvideofiltertype(3)
    Else
        Call setvideofiltertype(2) 'Standard
        IPEnhancementAndIPProcessMethod.IPEnhancement = "Standard" 'PCN2820
    End If
    
End If



Exit Function
Err_Handler:
    MsgBox Err & "-ST10:" & Error$
End Function


Function SetupVideoDisplayForPVTuning()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : SetupVideoDisplayForPVTuning
'Created : 12 May 2004, PCN2612
'Updated :
'Prg By  : Geoff Logan
'Param   : (None)
'Desc    : This function sets the video display for Precision Vision Tuning
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

If mediatype = Video Or mediatype = "Live" Or mediatype = StillImage Then 'PCN3194
    'Set Method, Process settings and Enhancement
    Call hough_processimageonoff(True)
    Call SetPVTuningProcessValues
    
    'Video setup as enhanced
    Call setimageanalysis(1)
    
    'Selection filter
    ShowGreenX = 0
    ShowGreenY = 0
    ShowProfileCandidates = 0 'PCN3017
 '   Call setselectionfilter(ShowGreenX, ShowGreenY)
    Call setprofilecandidates(ShowProfileCandidates) 'PCN3017
    
    'candidates
    ShowProf = 0 'PCN3017 Set it to default off
    Call setprofilecandidates(ShowProf)
    
    'Final profile ON
    Call setprofileoverlay(1)
    
    'Refresh video
    Call hough_processimageonoff(True)
    Call ClearLineScreen.RefreshVideoScreen 'PCN3017
    
End If

Exit Function
Err_Handler:
    MsgBox Err & "-ST11:" & Error$
End Function


Function SetupVideoDisplayAsNormal()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : SetupVideoDisplayAsNormal
'Created : 12 May 2004, PCN2612
'Updated :
'Prg By  : Geoff Logan
'Param   : (None)
'Desc    : This function sets the video display for Precision Vision Tuning
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

If mediatype = Video Or mediatype = "Live" Or mediatype = StillImage Then 'PCN3194
    'Set Method, Process settings and Enhancement
    Call SetPVTuningProcessValues
    
    'Video setup as normal
    Call setimageanalysis(0)
    
    'Selection filter
    ShowGreenX = 0
    ShowGreenY = 0
    ShowProfileCandidates = 0 'PCN3017
 '   Call setselectionfilter(ShowGreenX, ShowGreenY)
    Call setprofilecandidates(ShowProfileCandidates) 'PCN3017
    
    'candidates off
    ShowProf = 0
    Call setprofilecandidates(ShowProf)
    
    'Final profile off
    Call setprofileoverlay(0)
    
    'Refresh video
    If PVRecording = False Then Call hough_processimageonoff(False)
    Call ClearLineScreen.RefreshVideoScreen 'PCN3017
    
End If

Exit Function
Err_Handler:
    MsgBox Err & "-ST12:" & Error$
End Function

Function CalXParaYPara() As Double
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : CalXParaYPara
'Created : 11 February 2004, PCN2612
'Updated : 12 May 2004, PCN2612 - moved to Startup
'Prg By  : Geoff Logan
'Param   : (None)
'Desc    : This function defines the relationship between XPara and LaserWidth
'Usage   : PV Tuning
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

'CalXParaYPara = 10 ^ (ConfigInfo.IPStDX + 5) 'PCN2820 'PCN3017


Exit Function
Err_Handler:
    MsgBox Err & "-ST13:" & Error$
End Function

Function ConvertDistortionToZoom(DistortionValue As Integer) As Double

' PCN2648 (Antony van Iersel, 18 May 2004) '
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : ConvertDistortionToZoom
'Created : 18 May 2004, PCN2648
'Prg By  : Antony van Iersel
'Param   : Distortion Value (TransX)
'Desc    : Look up table, converting the TransX (Distortion Value) to
'          setorg2prcratio (Fish Eye Zoom Value)
'
'Usage   : Fish Eye Display - Zoom Factor
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler

    Dim ZoomLookUp(25) As Double
  
       
    ZoomLookUp(1) = 1.2
    ZoomLookUp(2) = 1.2
    ZoomLookUp(3) = 1.21
    ZoomLookUp(4) = 1.22
    ZoomLookUp(5) = 1.23
    ZoomLookUp(6) = 1.24
    ZoomLookUp(7) = 1.26
    ZoomLookUp(8) = 1.28
    ZoomLookUp(9) = 1.29
    ZoomLookUp(10) = 1.31
    ZoomLookUp(11) = 1.34
    ZoomLookUp(12) = 1.36
    ZoomLookUp(13) = 1.38
    ZoomLookUp(14) = 1.4
    ZoomLookUp(15) = 1.42
    ZoomLookUp(16) = 1.44
    ZoomLookUp(17) = 1.47
    ZoomLookUp(18) = 1.5
    ZoomLookUp(19) = 1.54
    ZoomLookUp(20) = 1.57
    ZoomLookUp(21) = 1.62
    ZoomLookUp(22) = 1.66
    ZoomLookUp(23) = 1.72
    ZoomLookUp(24) = 1.78
    ZoomLookUp(25) = 1.84
    
    ConvertDistortionToZoom = ZoomLookUp(DistortionValue)

Exit Function
Err_Handler:
    MsgBox Err & "-ST14:" & Error$
End Function


Function GetINI_ImageProcessInfo(MyFile As String, FileNo As Integer)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : GetINI_ImageProcessInfo
'Created : 19 May 2004, PCN2820
'Updated :
'Prg By  : Geoff Logan
'Param   : MyFile - location of the INI file
'          Parameter - search for this parameter in the INI
'          ReturnValue - this is the value in the INI given to the parameter requested
'Desc    : Gets from the INI file all the information relating to the ImageProcessing settings
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim ConfigLine As String
Dim PathName As String
Dim FileName As String
Dim Parameter As String
Dim ReturnValue As String

ReturnValue = ""

If Dir(MyFile) = "" Then Exit Function

Open MyFile For Input As #FileNo
Do While Not EOF(FileNo)
    Line Input #FileNo, ConfigLine
    If Len(ConfigLine) <> 0 Then
        Call GetParam(ConfigLine, Parameter, ReturnValue, PathName, FileName)
        Select Case Trim(Parameter)
'            Case "IPX="
'                ConfigInfo.PVShapeCentreX = Val(ReturnValue) 'PCN4336
'            Case "IPY="
'                ConfigInfo.PVShapeCentreY = Val(ReturnValue) 'PCN4336
            Case "IPGT="
                ConfigInfo.IPGradThres = SafeCDbl(ReturnValue)
            Case "IPDX="
                ConfigInfo.IPStDX = SafeCDbl(ReturnValue)
            Case "IPDY="
                ConfigInfo.IPStDY = SafeCDbl(ReturnValue)
            Case "ProcessMethod="
                IPEnhancementAndIPProcessMethod.IPProcessMethod = ReturnValue
            Case "Contrast="
                ConfigInfo.IPZone = SafeCDbl(ReturnValue)
            Case "Enhancement="
                IPEnhancementAndIPProcessMethod.IPEnhancement = ReturnValue
        End Select
    End If
Loop
Close #FileNo


Exit Function
Err_Handler:
Select Case Err
    Case 9 'Subscript out of range (end of file)
        Exit Function
    Case Else
        MsgBox Err & "-ST15:" & Error$
End Select
End Function

Function GetINI_LimitLineInfo(MyFile As String, FileNo As Integer)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Name    : GetINI_LimitLineInfo
'Created : 19 May 2004, PCN2820
'Updated :
'Prg By  : Geoff Logan
'Param   : MyFile - location of the INI file
'          Parameter - search for this parameter in the INI
'          ReturnValue - this is the value in the INI given to the parameter requested
'Desc    : Gets from the INI file all the information relating to the ImageProcessing settings
'Usage   :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo Err_Handler
Dim ConfigLine As String
Dim PathName As String
Dim FileName As String
Dim Parameter As String
Dim ReturnValue As String

ReturnValue = ""

If Dir(MyFile) = "" Then Exit Function

Open MyFile For Input As #FileNo
Do While Not EOF(FileNo)
    Line Input #FileNo, ConfigLine
    If Len(ConfigLine) <> 0 Then
        Call GetParam(ConfigLine, Parameter, ReturnValue, PathName, FileName)
        Select Case Parameter
        'vvvv PCN2769 ***********************************************
        Case "LimitCapMin="
            ConfigInfo.LimitCapacityL = SafeCDbl(ReturnValue) 'PCN2820
        Case "LimitCapMax="
            ConfigInfo.LimitCapacityR = SafeCDbl(ReturnValue) 'PCN2820
        'vvvv PCN4349 **************************
        Case "LimitOvalL=": OvalityLimitL = SafeCDbl(ReturnValue)
        Case "LimitOvalR=": ConfigInfo.LimitOvality = SafeCDbl(ReturnValue) 'PCN2820
                            OvalityLimitR = ConfigInfo.LimitOvality
        Case "LimitXYDiameterL=": ConfigInfo.LimitXYDiameterL = SafeCDbl(ReturnValue)
        Case "LimitXYDiameterR=": ConfigInfo.LimitXYDiameterR = SafeCDbl(ReturnValue)
        Case "LimitDiameterMaxL=": DiameterMaxLimitL = SafeCDbl(ReturnValue) 'PCN4799 was called MaxMin
        Case "LimitDiameterMaxR=": DiameterMaxLimitR = SafeCDbl(ReturnValue) 'PCN4799 was called MaxMin
        Case "LimitDiameterMinL=": GraphInfoContainer(PVMinDiameter).LimitL = SafeCDbl(ReturnValue) 'PCN4799 was called MaxMin
        Case "LimitDiameterMinR=": GraphInfoContainer(PVMinDiameter).LimitR = SafeCDbl(ReturnValue) 'PCN4799 was called MaxMin
        Case "LimitDiameterMedianL=": DiameterMedianLimitL = SafeCDbl(ReturnValue)
                                      ConfigInfo.LimitDeltaL = DiameterMedianLimitL
        Case "LimitDiameterMedianR=": DiameterMedianLimitR = SafeCDbl(ReturnValue)
                                      ConfigInfo.LimitDeltaR = DiameterMedianLimitR
            
'        Case "LimitDeltaMin="
'            ConfigInfo.LimitDeltaL = Val(ReturnValue) 'PCN2820
'        Case "LimitDeltaMax="
'            ConfigInfo.LimitDeltaR = Val(ReturnValue) 'PCN2820
        '^^^^ ***********************************
        '^^^^ *******************************************************
        End Select
    End If
Loop
Close #FileNo


Exit Function
Err_Handler:
Select Case Err
    Case 9 'Subscript out of range (end of file)
        Exit Function
    Case Else
        MsgBox Err & "-ST16:" & Error$
End Select
End Function

Function Validate_HASP_Lock() As Boolean    'PCN3197
'****************************************************************************************
'Name    : Validate_HASP_Lock
'Created : Nov 30 2004
'Updated :
'Prg By  : Eddie Jensen
'Param   :
'Desc    : detects the presence of the HASPLOCK.dll file and validate the HASP
'          envelope if detected
'Usage   :
'****************************************************************************************
On Error GoTo Err_Handler

Validate_HASP_Lock = False

If HASP_Lock_Validate = 1800.5622543 Then
' if the control enters this point then the HASP key and the dll have been located, and the
' HASP key contains the correct value
    Validate_HASP_Lock = True
End If

Exit Function
Err_Handler:

Select Case Err
    Case 53
        Exit Function
    Case 48
        Exit Function
    Case Else
        MsgBox Err & "-ST17:" & Error$
End Select
End Function

Sub DefineCircleShape()
'****************************************************************************************
' PCN3055
'Name    : DefineCircleShape
'Created : 9 Feb 2004
'Updated :
'Prg By  : Antony van Iersel
'Param   :
'Desc    : Fills in the ReferenceShape for Circle
'Usage   : The shapes are used for drawing and graph calculation
'****************************************************************************************
On Error GoTo Err_Handler

ReDim ReferenceShape(0)

ReferenceShape(0).Name = "Circle"
ReferenceShape(0).Use = "All"
ReferenceShape(0).CentreOffsetX = 0#
ReferenceShape(0).CentreOffsetY = 0#
ReferenceShape(0).NoArcs = 2
ReferenceShape(0).NoLines = 0

'Arc One'
ReferenceShape(0).Arcs(0).OriginX = 0
ReferenceShape(0).Arcs(0).OriginY = 0
ReferenceShape(0).Arcs(0).Radius = 1#
ReferenceShape(0).Arcs(0).StartAngle = 0#
ReferenceShape(0).Arcs(0).EndAngle = 180#

'Arc One'
ReferenceShape(0).Arcs(1).OriginX = 0
ReferenceShape(0).Arcs(1).OriginY = 0
ReferenceShape(0).Arcs(1).Radius = 1#
ReferenceShape(0).Arcs(1).StartAngle = 180#
ReferenceShape(0).Arcs(1).EndAngle = 360#

Exit Sub
Err_Handler:

Select Case Err
    Case 53
        Exit Sub
    Case 48
        Exit Sub
    Case Else
        MsgBox Err & "-ST18:" & Error$
End Select
End Sub

Sub DefineSemiEllipticalShape()
'****************************************************************************************
' PCN3055
'Name    : DefineSemiEllipticalShape
'Created : 14 December 2004
'Updated :
'Prg By  : Antony van Iersel
'Param   :
'Desc    : Fills in the ReferenceShape for SemiEllipticalShape
'Usage   : The shapes are used for drawing and graph calculation
'****************************************************************************************
On Error GoTo Err_Handler

ReDim Preserve ReferenceShape(1)

ReferenceShape(1).Name = "SemiElliptical"
ReferenceShape(1).Use = "All"
ReferenceShape(1).CentreOffsetX = 0#
ReferenceShape(1).CentreOffsetY = -0.5
ReferenceShape(1).NoArcs = 8
ReferenceShape(1).NoLines = 0

'Arc One'
ReferenceShape(1).Arcs(0).OriginX = -1.238095238
ReferenceShape(1).Arcs(0).OriginY = 0.010714286
ReferenceShape(1).Arcs(0).Radius = 2.238095238
ReferenceShape(1).Arcs(0).StartAngle = 359.72480002
ReferenceShape(1).Arcs(0).EndAngle = 37.58625399

'Arc Two'
ReferenceShape(1).Arcs(1).OriginX = -0.03047619
ReferenceShape(1).Arcs(1).OriginY = 0.940238095
ReferenceShape(1).Arcs(1).Radius = 0.714285714
ReferenceShape(1).Arcs(1).StartAngle = 37.58625399
ReferenceShape(1).Arcs(1).EndAngle = 87.54888086

'Arc Three'
ReferenceShape(1).Arcs(2).OriginX = 0.03047619
ReferenceShape(1).Arcs(2).OriginY = 0.940238095
ReferenceShape(1).Arcs(2).Radius = 0.714285714
ReferenceShape(1).Arcs(2).StartAngle = 92.45111914
ReferenceShape(1).Arcs(2).EndAngle = 142.41374601

'Arc Four'
ReferenceShape(1).Arcs(3).OriginX = 1.238095238
ReferenceShape(1).Arcs(3).OriginY = 0.010714286
ReferenceShape(1).Arcs(3).Radius = 2.238095238
ReferenceShape(1).Arcs(3).StartAngle = 142.41374601
ReferenceShape(1).Arcs(3).EndAngle = 180.27519998

'Arc Five'
ReferenceShape(1).Arcs(4).OriginX = -0.285714286
ReferenceShape(1).Arcs(4).OriginY = 0.003333333
ReferenceShape(1).Arcs(4).Radius = 0.714285714
ReferenceShape(1).Arcs(4).StartAngle = 180.27519998
ReferenceShape(1).Arcs(4).EndAngle = 202.91065938

'Need to be adjusted
'Arc Six'
ReferenceShape(1).Arcs(5).OriginX = 0#
ReferenceShape(1).Arcs(5).OriginY = 2.232142857
ReferenceShape(1).Arcs(5).Radius = 2.678571429
ReferenceShape(1).Arcs(5).StartAngle = 249.37471107
ReferenceShape(1).Arcs(5).EndAngle = 270#

'need to be adjusted
'Arc Seven'
ReferenceShape(1).Arcs(6).OriginX = 0#
ReferenceShape(1).Arcs(6).OriginY = 2.232142857
ReferenceShape(1).Arcs(6).Radius = 2.678571429
ReferenceShape(1).Arcs(6).StartAngle = 270#
ReferenceShape(1).Arcs(6).EndAngle = 290.63172371

'Arc Eight'
ReferenceShape(1).Arcs(7).OriginX = 0.285714286
ReferenceShape(1).Arcs(7).OriginY = 0.003333333
ReferenceShape(1).Arcs(7).Radius = 0.714285714
ReferenceShape(1).Arcs(7).StartAngle = 337.08934062
ReferenceShape(1).Arcs(7).EndAngle = 359.72480002

Exit Sub
Err_Handler:

Select Case Err
    Case 53
        Exit Sub
    Case 48
        Exit Sub
    Case Else
        MsgBox Err & "-ST19:" & Error$
End Select
End Sub

Sub DefineEggShape()
'****************************************************************************************
' PCN3055
'Name    : DefineSemiEllipticalShape
'Created : 14 December 2004
'Updated :
'Prg By  : Antony van Iersel
'Param   :
'Desc    : Fills in the ReferenceShape for SemiEllipticalShape
'Usage   : The shapes are used for drawing and graph calculation
'****************************************************************************************
On Error GoTo Err_Handler

ReDim Preserve ReferenceShape(2)

ReferenceShape(2).Name = "Egg"
ReferenceShape(2).Use = "All"
ReferenceShape(2).CentreOffsetX = 0#
ReferenceShape(2).CentreOffsetY = 0.5
ReferenceShape(2).NoArcs = 4
ReferenceShape(2).NoLines = 0



'Arc One'
ReferenceShape(2).Arcs(0).OriginX = 0
ReferenceShape(2).Arcs(0).OriginY = 0
ReferenceShape(2).Arcs(0).Radius = 1
ReferenceShape(2).Arcs(0).StartAngle = 0
ReferenceShape(2).Arcs(0).EndAngle = 180

'Arc Two'
ReferenceShape(2).Arcs(1).OriginX = 0
ReferenceShape(2).Arcs(1).OriginY = -1.5
ReferenceShape(2).Arcs(1).Radius = 0.5
ReferenceShape(2).Arcs(1).StartAngle = 207.811
ReferenceShape(2).Arcs(1).EndAngle = 332.18

'Arc Three'
ReferenceShape(2).Arcs(2).OriginX = 2
ReferenceShape(2).Arcs(2).OriginY = 0
ReferenceShape(2).Arcs(2).Radius = 3
ReferenceShape(2).Arcs(2).StartAngle = 180
ReferenceShape(2).Arcs(2).EndAngle = 217.5

'Arc Four'
ReferenceShape(2).Arcs(3).OriginX = -2
ReferenceShape(2).Arcs(3).OriginY = 0
ReferenceShape(2).Arcs(3).Radius = 3
ReferenceShape(2).Arcs(3).StartAngle = 323.5
ReferenceShape(2).Arcs(3).EndAngle = 0

Exit Sub
Err_Handler:

Select Case Err
    Case 53
        Exit Sub
    Case 48
        Exit Sub
    Case Else
        MsgBox Err & "-ST1A:" & Error$
End Select
End Sub
Sub DefineRinkerEllipseShape()
'****************************************************************************************
' PCN3055
'Name    : DefineSemiEllipticalShape
'Created : 14 December 2004
'Updated :
'Prg By  : Antony van Iersel
'Param   :
'Desc    : Fills in the ReferenceShape for SemiEllipticalShape
'Usage   : The shapes are used for drawing and graph calculation
'****************************************************************************************
On Error GoTo Err_Handler

ReDim Preserve ReferenceShape(3)

Dim ShapeScale As Double
ShapeScale = 1.630028382


ReferenceShape(3).Name = "RinkerEllipse"
ReferenceShape(3).Use = "All"
ReferenceShape(3).CentreOffsetX = 0# * ShapeScale
ReferenceShape(3).CentreOffsetY = -0.60555 * ShapeScale
ReferenceShape(3).NoArcs = 4
ReferenceShape(3).NoLines = 0



'Arc One'
ReferenceShape(3).Arcs(0).OriginX = 0 * ShapeScale
ReferenceShape(3).Arcs(0).OriginY = 0 * ShapeScale
ReferenceShape(3).Arcs(0).Radius = 1 * ShapeScale
ReferenceShape(3).Arcs(0).StartAngle = 64.90940421

ReferenceShape(3).Arcs(0).EndAngle = 115.09059579


'Arc Two'
ReferenceShape(3).Arcs(1).OriginX = -0.283530595 * ShapeScale
ReferenceShape(3).Arcs(1).OriginY = 0.605531716 * ShapeScale
ReferenceShape(3).Arcs(1).Radius = 0.331375847 * ShapeScale
ReferenceShape(3).Arcs(1).StartAngle = 115.09059579
ReferenceShape(3).Arcs(1).EndAngle = 244.90940421


'Arc Three'
ReferenceShape(3).Arcs(2).OriginX = 0# * ShapeScale
ReferenceShape(3).Arcs(2).OriginY = 1.211063433 * ShapeScale
ReferenceShape(3).Arcs(2).Radius = 1# * ShapeScale
ReferenceShape(3).Arcs(2).StartAngle = 244.90940421
ReferenceShape(3).Arcs(2).EndAngle = 295.09059579


'Arc Four'
ReferenceShape(3).Arcs(3).OriginX = 0.283530595 * ShapeScale
ReferenceShape(3).Arcs(3).OriginY = 0.605531716 * ShapeScale
ReferenceShape(3).Arcs(3).Radius = 0.331375847 * ShapeScale
ReferenceShape(3).Arcs(3).StartAngle = 295.09059579
ReferenceShape(3).Arcs(3).EndAngle = 64.90940421


Exit Sub
Err_Handler:

Select Case Err
    Case 53
        Exit Sub
    Case 48
        Exit Sub
    Case Else
        MsgBox Err & "-ST23:" & Error$
End Select
End Sub


Sub DefineEllipticalASTM_C507Shape()
'****************************************************************************************
' PCN3055
'Name    : DefineSemiEllipticalShape
'Created : 14 December 2004
'Updated :
'Prg By  : Antony van Iersel
'Param   :
'Desc    : Fills in the ReferenceShape for SemiEllipticalShape
'Usage   : The shapes are used for drawing and graph calculation
'****************************************************************************************
On Error GoTo Err_Handler

ReDim Preserve ReferenceShape(4)

Dim ShapeScale As Double
ShapeScale = 2

ReferenceShape(4).Name = "Elliptical ASTM C507"
ReferenceShape(4).Use = "All"
ReferenceShape(4).CentreOffsetX = 0# * ShapeScale
ReferenceShape(4).CentreOffsetY = 0 * ShapeScale
ReferenceShape(4).NoArcs = 4
ReferenceShape(4).NoLines = 0



'Arc One'
ReferenceShape(4).Arcs(0).OriginX = 0.225903614 * ShapeScale
ReferenceShape(4).Arcs(0).OriginY = 0 * ShapeScale
ReferenceShape(4).Arcs(0).Radius = 0.274096386 * ShapeScale
ReferenceShape(4).Arcs(0).StartAngle = 292.62
ReferenceShape(4).Arcs(0).EndAngle = 67.4

'Arc Two'
ReferenceShape(4).Arcs(1).OriginX = 0 * ShapeScale
ReferenceShape(4).Arcs(1).OriginY = -0.542168675 * ShapeScale
ReferenceShape(4).Arcs(1).Radius = 0.861445783 * ShapeScale
ReferenceShape(4).Arcs(1).StartAngle = 67.4
ReferenceShape(4).Arcs(1).EndAngle = 112.6


'Arc Three'
ReferenceShape(4).Arcs(2).OriginX = -0.225903614 * ShapeScale
ReferenceShape(4).Arcs(2).OriginY = 0 * ShapeScale
ReferenceShape(4).Arcs(2).Radius = 0.274096386 * ShapeScale
ReferenceShape(4).Arcs(2).StartAngle = 112.6
ReferenceShape(4).Arcs(2).EndAngle = 247.4


'Arc Four'
ReferenceShape(4).Arcs(3).OriginX = 0 * ShapeScale
ReferenceShape(4).Arcs(3).OriginY = 0.542168675 * ShapeScale
ReferenceShape(4).Arcs(3).Radius = 0.861445783 * ShapeScale
ReferenceShape(4).Arcs(3).StartAngle = 247.4
ReferenceShape(4).Arcs(3).EndAngle = 292.62


Exit Sub
Err_Handler:

Select Case Err
    Case 53
        Exit Sub
    Case 48
        Exit Sub
    Case Else
        MsgBox Err & "-ST23:" & Error$
End Select
End Sub

'Sub DefineLoafShape()
''****************************************************************************************
'' PCN3055
''Name    : DefineSemiEllipticalShape
''Created : 14 December 2004
''Updated :
''Prg By  : Antony van Iersel
''Param   :
''Desc    : Fills in the ReferenceShape for SemiEllipticalShape
''Usage   : The shapes are used for drawing and graph calculation
''****************************************************************************************
'On Error GoTo Err_Handler
'
'ReDim Preserve ReferenceShape(4)
'
'ReferenceShape(4).name = "Loaf"
'ReferenceShape(4).Use = "All"
'ReferenceShape(4).CentreOffsetX = 0#
'ReferenceShape(4).CentreOffsetY = 0.3
'ReferenceShape(4).NoArcs = 2
'ReferenceShape(4).NoLines = 2
'
'
'
''Arc One'
'ReferenceShape(4).Arcs(0).OriginX = 0
'ReferenceShape(4).Arcs(0).OriginY = 0
'ReferenceShape(4).Arcs(0).Radius = 1
'ReferenceShape(4).Arcs(0).StartAngle = 0#
'ReferenceShape(4).Arcs(0).EndAngle = 180#
'
''Arc Two'
'ReferenceShape(4).Arcs(1).OriginX = 0
'ReferenceShape(4).Arcs(1).OriginY = 2.5
'ReferenceShape(4).Arcs(1).Radius = 4.12310563
'ReferenceShape(4).Arcs(1).StartAngle = 255.96375653
'ReferenceShape(4).Arcs(1).EndAngle = 284.03624347
'
'
''Line One'
'ReferenceShape(4).Lines(0).StartX = -1
'ReferenceShape(4).Lines(0).StartY = 0
'ReferenceShape(4).Lines(0).EndX = -1
'ReferenceShape(4).Lines(0).EndY = -1.5
'
''Line Two'
'ReferenceShape(4).Lines(1).StartX = 1
'ReferenceShape(4).Lines(1).StartY = 0
'ReferenceShape(4).Lines(1).EndX = 1
'ReferenceShape(4).Lines(1).EndY = -1.5
'
'Exit Sub
'Err_Handler:
'
'Select Case Err
'    Case 53
'        Exit Sub
'    Case 48
'        Exit Sub
'    Case Else
'        MsgBox Err & "-ST1C:" & Error$
'End Select
'End Sub

Sub DefineBoxCulvertShape()
'****************************************************************************************
' PCN3055
'Name    : DefineSemiEllipticalShape
'Created : 14 December 2004
'Updated :
'Prg By  : Antony van Iersel
'Param   :
'Desc    : Fills in the ReferenceShape for SemiEllipticalShape
'Usage   : The shapes are used for drawing and graph calculation
'****************************************************************************************
On Error GoTo Err_Handler

ReDim Preserve ReferenceShape(5)

Dim ShapeScale As Double

ShapeScale = 1.592139239

ReferenceShape(5).Name = "BoxCulvert"
ReferenceShape(5).Use = "All"
ReferenceShape(5).CentreOffsetX = 0#
ReferenceShape(5).CentreOffsetY = 0#
ReferenceShape(5).NoArcs = 0
ReferenceShape(5).NoLines = 8

'Line One'
ReferenceShape(5).Lines(0).StartX = -0.62808577 * ShapeScale
ReferenceShape(5).Lines(0).StartY = 0.65 * ShapeScale
ReferenceShape(5).Lines(0).EndX = -0.5 * ShapeScale
ReferenceShape(5).Lines(0).EndY = 0.77808577 * ShapeScale


'Line Two'
ReferenceShape(5).Lines(1).StartX = -0.5 * ShapeScale
ReferenceShape(5).Lines(1).StartY = 0.77808577 * ShapeScale
ReferenceShape(5).Lines(1).EndX = 0.5 * ShapeScale
ReferenceShape(5).Lines(1).EndY = 0.77808577 * ShapeScale


'Line Three
ReferenceShape(5).Lines(2).StartX = 0.5 * ShapeScale
ReferenceShape(5).Lines(2).StartY = 0.77808577 * ShapeScale
ReferenceShape(5).Lines(2).EndX = 0.62808577 * ShapeScale
ReferenceShape(5).Lines(2).EndY = 0.65 * ShapeScale


'Line Four
ReferenceShape(5).Lines(3).StartX = 0.62808577 * ShapeScale
ReferenceShape(5).Lines(3).StartY = 0.65 * ShapeScale
ReferenceShape(5).Lines(3).EndX = 0.62808577 * ShapeScale
ReferenceShape(5).Lines(3).EndY = -0.65 * ShapeScale

'Line Five
ReferenceShape(5).Lines(4).StartX = 0.62808577 * ShapeScale
ReferenceShape(5).Lines(4).StartY = -0.65 * ShapeScale
ReferenceShape(5).Lines(4).EndX = 0.5 * ShapeScale
ReferenceShape(5).Lines(4).EndY = -0.77808577 * ShapeScale

'Line Six
ReferenceShape(5).Lines(5).StartX = 0.5 * ShapeScale
ReferenceShape(5).Lines(5).StartY = -0.77808577 * ShapeScale
ReferenceShape(5).Lines(5).EndX = -0.5 * ShapeScale
ReferenceShape(5).Lines(5).EndY = -0.77808577 * ShapeScale

'Line Seven
ReferenceShape(5).Lines(6).StartX = -0.5 * ShapeScale
ReferenceShape(5).Lines(6).StartY = -0.77808577 * ShapeScale
ReferenceShape(5).Lines(6).EndX = -0.62808577 * ShapeScale
ReferenceShape(5).Lines(6).EndY = -0.65 * ShapeScale

'Line Eight
ReferenceShape(5).Lines(7).StartX = -0.62808577 * ShapeScale
ReferenceShape(5).Lines(7).StartY = -0.65 * ShapeScale
ReferenceShape(5).Lines(7).EndX = -0.62808577 * ShapeScale
ReferenceShape(5).Lines(7).EndY = 0.65 * ShapeScale





Exit Sub
Err_Handler:

Select Case Err
    Case 53
        Exit Sub
    Case 48
        Exit Sub
    Case Else
        MsgBox Err & "-ST1D:" & Error$
End Select
End Sub

Sub DefineBarnShape()
'****************************************************************************************
' PCN3055
'Name    : DefineSemiEllipticalShape
'Created : 14 December 2004
'Updated :
'Prg By  : Antony van Iersel
'Param   :
'Desc    : Fills in the ReferenceShape for SemiEllipticalShape
'Usage   : The shapes are used for drawing and graph calculation
'****************************************************************************************
On Error GoTo Err_Handler

ReDim Preserve ReferenceShape(6)

Dim ShapeScale As Double

ShapeScale = 0.817661488


ReferenceShape(6).Name = "Barn"
ReferenceShape(6).Use = "All"
ReferenceShape(6).CentreOffsetX = 0#
ReferenceShape(6).CentreOffsetY = 0.3
ReferenceShape(6).NoArcs = 4
ReferenceShape(6).NoLines = 0





'Arc One'
ReferenceShape(6).Arcs(0).OriginX = 0 * ShapeScale
ReferenceShape(6).Arcs(0).OriginY = 0 * ShapeScale
ReferenceShape(6).Arcs(0).Radius = 1 * ShapeScale
ReferenceShape(6).Arcs(0).StartAngle = 12.89245751
ReferenceShape(6).Arcs(0).EndAngle = 167.10754249

'Arc Two'
ReferenceShape(6).Arcs(1).OriginX = 8.529838216 * ShapeScale
ReferenceShape(6).Arcs(1).OriginY = -1.952412 * ShapeScale
ReferenceShape(6).Arcs(1).Radius = 9.750431567 * ShapeScale
ReferenceShape(6).Arcs(1).StartAngle = 167.10754249
ReferenceShape(6).Arcs(1).EndAngle = 182.83156068

'Arc Three'
ReferenceShape(6).Arcs(2).OriginX = 0# * ShapeScale
ReferenceShape(6).Arcs(2).OriginY = 4.604007589 * ShapeScale
ReferenceShape(6).Arcs(2).Radius = 7.141123471 * ShapeScale
ReferenceShape(6).Arcs(2).StartAngle = 260.25534764
ReferenceShape(6).Arcs(2).EndAngle = 279.74465236

'Arc Four'
ReferenceShape(6).Arcs(3).OriginX = -8.529838216 * ShapeScale
ReferenceShape(6).Arcs(3).OriginY = -1.952412 * ShapeScale
ReferenceShape(6).Arcs(3).Radius = 9.750431567 * ShapeScale
ReferenceShape(6).Arcs(3).StartAngle = 357.16843932
ReferenceShape(6).Arcs(3).EndAngle = 12.89245751

Exit Sub
Err_Handler:

Select Case Err
    Case 53
        Exit Sub
    Case 48
        Exit Sub
    Case Else
        MsgBox Err & "-ST1E:" & Error$
End Select
End Sub

Sub DefineBarnDShape()
'****************************************************************************************
' PCN3055
'Name    : DefineSemiEllipticalShape
'Created : 14 December 2004
'Updated :
'Prg By  : Antony van Iersel
'Param   :
'Desc    : Fills in the ReferenceShape for SemiEllipticalShape
'Usage   : The shapes are used for drawing and graph calculation
'****************************************************************************************
On Error GoTo Err_Handler

ReDim Preserve ReferenceShape(7)

ReferenceShape(7).Name = "BarnD"
ReferenceShape(7).Use = "All"
ReferenceShape(7).CentreOffsetX = 0#
ReferenceShape(7).CentreOffsetY = 0.3
ReferenceShape(7).NoArcs = 4
ReferenceShape(7).NoLines = 0

'Arc One'
ReferenceShape(7).Arcs(0).OriginX = 0
ReferenceShape(7).Arcs(0).OriginY = 0
ReferenceShape(7).Arcs(0).Radius = 1
ReferenceShape(7).Arcs(0).StartAngle = 17.99
ReferenceShape(7).Arcs(0).EndAngle = 162.01

'Arc Two'
ReferenceShape(7).Arcs(1).OriginX = 5.499304473
ReferenceShape(7).Arcs(1).OriginY = -0.482827677
ReferenceShape(7).Arcs(1).Radius = 6.498818803
ReferenceShape(7).Arcs(1).StartAngle = 173#
ReferenceShape(7).Arcs(1).EndAngle = 187#

'Arc Three'
ReferenceShape(7).Arcs(2).OriginX = -5.499304473
ReferenceShape(7).Arcs(2).OriginY = -0.482827677
ReferenceShape(7).Arcs(2).Radius = 6.498818803
ReferenceShape(7).Arcs(2).StartAngle = 353#
ReferenceShape(7).Arcs(2).EndAngle = 7#

'Arc Four'
ReferenceShape(7).Arcs(3).OriginX = 0#
ReferenceShape(7).Arcs(3).OriginY = 3.002392373
ReferenceShape(7).Arcs(3).Radius = 4.381358676
ReferenceShape(7).Arcs(3).StartAngle = 257.46
ReferenceShape(7).Arcs(3).EndAngle = 282.54

Exit Sub
Err_Handler:

Select Case Err
    Case 53
        Exit Sub
    Case 48
        Exit Sub
    Case Else
        MsgBox Err & "-ST1F:" & Error$
End Select
End Sub
Sub DefineEggAShape()
'****************************************************************************************
' PCN3055
'Name    : DefineSemiEllipticalShape
'Created : 14 December 2004
'Updated :
'Prg By  : Antony van Iersel
'Param   :
'Desc    : Fills in the ReferenceShape for SemiEllipticalShape
'Usage   : The shapes are used for drawing and graph calculation
'****************************************************************************************
On Error GoTo Err_Handler

ReDim Preserve ReferenceShape(8)

ReferenceShape(8).Name = "Egg A"
ReferenceShape(8).Use = "All"
ReferenceShape(8).CentreOffsetX = 0#
ReferenceShape(8).CentreOffsetY = 0.3
ReferenceShape(8).NoArcs = 4
ReferenceShape(8).NoLines = 0



'Arc One'
ReferenceShape(8).Arcs(0).OriginX = 0
ReferenceShape(8).Arcs(0).OriginY = 0
ReferenceShape(8).Arcs(0).Radius = 1
ReferenceShape(8).Arcs(0).StartAngle = 355.8920588

ReferenceShape(8).Arcs(0).EndAngle = 184.1079412


'Arc Two'
ReferenceShape(8).Arcs(1).OriginX = 2.42396244
ReferenceShape(8).Arcs(1).OriginY = 0.161408948
ReferenceShape(8).Arcs(1).Radius = 3.43020596
ReferenceShape(8).Arcs(1).StartAngle = 184.1079412
ReferenceShape(8).Arcs(1).EndAngle = 212.032561

'Arc Three'
ReferenceShape(8).Arcs(2).OriginX = 0#
ReferenceShape(8).Arcs(2).OriginY = -1.34248644
ReferenceShape(8).Arcs(2).Radius = 0.57090598
ReferenceShape(8).Arcs(2).StartAngle = 212.032561
ReferenceShape(8).Arcs(2).EndAngle = 327.967439

'Arc Four'
ReferenceShape(8).Arcs(3).OriginX = -2.42396244
ReferenceShape(8).Arcs(3).OriginY = 0.161408948
ReferenceShape(8).Arcs(3).Radius = 3.43020596
ReferenceShape(8).Arcs(3).StartAngle = 327.967439
ReferenceShape(8).Arcs(3).EndAngle = 355.8920588

Exit Sub
Err_Handler:

Select Case Err
    Case 53
        Exit Sub
    Case 48
        Exit Sub
    Case Else
        MsgBox Err & "-ST20:" & Error$
End Select
End Sub

Sub DefineEggBShape()
'****************************************************************************************
' PCN3055
'Name    : DefineSemiEllipticalShape
'Created : 14 December 2004
'Updated :
'Prg By  : Antony van Iersel
'Param   :
'Desc    : Fills in the ReferenceShape for SemiEllipticalShape
'Usage   : The shapes are used for drawing and graph calculation
'****************************************************************************************
On Error GoTo Err_Handler

Dim RadiusScale As Single
RadiusScale = 1

ReDim Preserve ReferenceShape(9)

ReferenceShape(9).Name = "Egg B"
ReferenceShape(9).Use = "All"
ReferenceShape(9).CentreOffsetX = 0 * RadiusScale
ReferenceShape(9).CentreOffsetY = 0.3 * RadiusScale
ReferenceShape(9).NoArcs = 4
ReferenceShape(9).NoLines = 0



'Arc One'
ReferenceShape(9).Arcs(0).OriginX = 0 * RadiusScale
ReferenceShape(9).Arcs(0).OriginY = 0 * RadiusScale
ReferenceShape(9).Arcs(0).Radius = 1 * RadiusScale
ReferenceShape(9).Arcs(0).StartAngle = 347.6796452

ReferenceShape(9).Arcs(0).EndAngle = 192.3203548


'Arc Two'
ReferenceShape(9).Arcs(1).OriginX = 2.26331494 * RadiusScale
ReferenceShape(9).Arcs(1).OriginY = 0.49432482 * RadiusScale
ReferenceShape(9).Arcs(1).Radius = 3.3166682 * RadiusScale
ReferenceShape(9).Arcs(1).StartAngle = 192.3203548
ReferenceShape(9).Arcs(1).EndAngle = 215.9470533

'Arc Three'
ReferenceShape(9).Arcs(2).OriginX = 0# * RadiusScale
ReferenceShape(9).Arcs(2).OriginY = -1.14687634 * RadiusScale
ReferenceShape(9).Arcs(2).Radius = 0.52093292 * RadiusScale
ReferenceShape(9).Arcs(2).StartAngle = 215.9470533
ReferenceShape(9).Arcs(2).EndAngle = 324.0529467

'Arc Four'
ReferenceShape(9).Arcs(3).OriginX = -2.26331494 * RadiusScale
ReferenceShape(9).Arcs(3).OriginY = 0.49432482 * RadiusScale
ReferenceShape(9).Arcs(3).Radius = 3.3166682 * RadiusScale
ReferenceShape(9).Arcs(3).StartAngle = 324.0529467
ReferenceShape(9).Arcs(3).EndAngle = 347.6796452

Exit Sub
Err_Handler:

Select Case Err
    Case 53
        Exit Sub
    Case 48
        Exit Sub
    Case Else
        MsgBox Err & "-ST21:" & Error$
End Select
End Sub

Sub DefineCupCakeShape()
'****************************************************************************************
' PCN3055
'Name    : DefineSemiEllipticalShape
'Created : 14 December 2004
'Updated :
'Prg By  : Antony van Iersel
'Param   :
'Desc    : Fills in the ReferenceShape for SemiEllipticalShape
'Usage   : The shapes are used for drawing and graph calculation
'****************************************************************************************
On Error GoTo Err_Handler

Dim RadiusScale As Single
RadiusScale = 1.016949153

ReDim Preserve ReferenceShape(10)

Dim ShapeScale As Double

ShapeScale = 1.417434444

ReferenceShape(10).Name = "CupCake"
ReferenceShape(10).Use = "All"
ReferenceShape(10).CentreOffsetX = 0# * ShapeScale
ReferenceShape(10).CentreOffsetY = -0.3 * ShapeScale
ReferenceShape(10).NoArcs = 2
ReferenceShape(10).NoLines = 2



'Arc One'
ReferenceShape(10).Arcs(0).OriginX = 0 * ShapeScale
ReferenceShape(10).Arcs(0).OriginY = 0 * ShapeScale
ReferenceShape(10).Arcs(0).Radius = 1 * ShapeScale
ReferenceShape(10).Arcs(0).StartAngle = 45.1571138

ReferenceShape(10).Arcs(0).EndAngle = 134.8428862


'Arc Two'
ReferenceShape(10).Arcs(1).OriginX = 0 * ShapeScale
ReferenceShape(10).Arcs(1).OriginY = 3.0356757 * ShapeScale
ReferenceShape(10).Arcs(1).Radius = 3.68886751 * ShapeScale
ReferenceShape(10).Arcs(1).StartAngle = 259.46048417
ReferenceShape(10).Arcs(1).EndAngle = 280.53951583


'Line One'
ReferenceShape(10).Lines(0).StartX = 0.70516513 * ShapeScale
ReferenceShape(10).Lines(0).StartY = 0.70904312 * ShapeScale
ReferenceShape(10).Lines(0).EndX = 0.67474409 * ShapeScale
ReferenceShape(10).Lines(0).EndY = -0.59095688 * ShapeScale

'Line Two'
ReferenceShape(10).Lines(1).StartX = -0.70516513 * ShapeScale
ReferenceShape(10).Lines(1).StartY = 0.70904312 * ShapeScale
ReferenceShape(10).Lines(1).EndX = -0.67474409 * ShapeScale
ReferenceShape(10).Lines(1).EndY = -0.59095688 * ShapeScale



Exit Sub
Err_Handler:

Select Case Err
    Case 53
        Exit Sub
    Case 48
        Exit Sub
    Case Else
        MsgBox Err & "-ST22:" & Error$
End Select
End Sub


Sub DefineBulletShape()
'****************************************************************************************
' PCN3055
'Name    : DefineSemiEllipticalShape
'Created : 14 December 2004
'Updated :
'Prg By  : Antony van Iersel
'Param   :
'Desc    : Fills in the ReferenceShape for SemiEllipticalShape
'Usage   : The shapes are used for drawing and graph calculation
'****************************************************************************************
On Error GoTo Err_Handler
Dim RadiusScale As Double

RadiusScale = 1

ReDim Preserve ReferenceShape(11)

ReferenceShape(11).Name = "Bullet"
ReferenceShape(11).Use = "All"
ReferenceShape(11).CentreOffsetX = 0#
ReferenceShape(11).CentreOffsetY = 0#
ReferenceShape(11).NoArcs = 4
ReferenceShape(11).NoLines = 0



'Arc One'
ReferenceShape(11).Arcs(0).OriginX = 0
ReferenceShape(11).Arcs(0).OriginY = 0
ReferenceShape(11).Arcs(0).Radius = 0.560994868 / RadiusScale
ReferenceShape(11).Arcs(0).StartAngle = 51.29
ReferenceShape(11).Arcs(0).EndAngle = 128.71

'Arc Two'
ReferenceShape(11).Arcs(1).OriginX = -0.733122779 / RadiusScale
ReferenceShape(11).Arcs(1).OriginY = -0.914725622 / RadiusScale
ReferenceShape(11).Arcs(1).Radius = 1.733517568 / RadiusScale
ReferenceShape(11).Arcs(1).StartAngle = 0#
ReferenceShape(11).Arcs(1).EndAngle = 51.29

'Arc Three'
ReferenceShape(11).Arcs(2).OriginX = 0# / RadiusScale
ReferenceShape(11).Arcs(2).OriginY = 0# / RadiusScale
ReferenceShape(11).Arcs(2).Radius = 1.355309909 / RadiusScale
ReferenceShape(11).Arcs(2).StartAngle = 222.45
ReferenceShape(11).Arcs(2).EndAngle = 317.55


'Arc Four'
ReferenceShape(11).Arcs(3).OriginX = 0.733122779 / RadiusScale
ReferenceShape(11).Arcs(3).OriginY = -0.914725622 / RadiusScale
ReferenceShape(11).Arcs(3).Radius = 1.733517568 / RadiusScale
ReferenceShape(11).Arcs(3).StartAngle = 128.71
ReferenceShape(11).Arcs(3).EndAngle = 180#


Exit Sub
Err_Handler:

Select Case Err
    Case 53
        Exit Sub
    Case 48
        Exit Sub
    Case Else
        MsgBox Err & "-ST1B:" & Error$
End Select
End Sub

Sub DefineSquareShape()
'****************************************************************************************
' PCN3055
'Name    : DefineSemiEllipticalShape
'Created : 14 December 2004
'Updated :
'Prg By  : Antony van Iersel
'Param   :
'Desc    : Fills in the ReferenceShape for SemiEllipticalShape
'Usage   : The shapes are used for drawing and graph calculation
'****************************************************************************************
On Error GoTo Err_Handler

ReDim Preserve ReferenceShape(12)

Dim ShapeScale As Double

ShapeScale = 2


ReferenceShape(12).Name = "Square"
ReferenceShape(12).Use = "All"
ReferenceShape(12).CentreOffsetX = 0# * ShapeScale
ReferenceShape(12).CentreOffsetY = 0# * ShapeScale
ReferenceShape(12).NoArcs = 0
ReferenceShape(12).NoLines = 4

'Line One'
ReferenceShape(12).Lines(0).StartX = -0.5 * ShapeScale
ReferenceShape(12).Lines(0).StartY = 0.5 * ShapeScale
ReferenceShape(12).Lines(0).EndX = 0.5 * ShapeScale
ReferenceShape(12).Lines(0).EndY = 0.5 * ShapeScale


'Line Two'
ReferenceShape(12).Lines(1).StartX = 0.5 * ShapeScale
ReferenceShape(12).Lines(1).StartY = 0.5 * ShapeScale
ReferenceShape(12).Lines(1).EndX = 0.5 * ShapeScale
ReferenceShape(12).Lines(1).EndY = -0.5 * ShapeScale


'Line Three
ReferenceShape(12).Lines(2).StartX = 0.5 * ShapeScale
ReferenceShape(12).Lines(2).StartY = -0.5 * ShapeScale
ReferenceShape(12).Lines(2).EndX = -0.5 * ShapeScale
ReferenceShape(12).Lines(2).EndY = -0.5 * ShapeScale


'Line Four
ReferenceShape(12).Lines(3).StartX = -0.5 * ShapeScale
ReferenceShape(12).Lines(3).StartY = -0.5 * ShapeScale
ReferenceShape(12).Lines(3).EndX = -0.5 * ShapeScale
ReferenceShape(12).Lines(3).EndY = 0.5 * ShapeScale







Exit Sub
Err_Handler:

Select Case Err
    Case 53
        Exit Sub
    Case 48
        Exit Sub
    Case Else
        MsgBox Err & "-ST24:" & Error$
End Select
End Sub


Sub DefineMushroomShape()
'****************************************************************************************
' PCN3055
'Name    : DefineSemiEllipticalShape
'Created : 14 December 2004
'Updated :
'Prg By  : Antony van Iersel
'Param   :
'Desc    : Fills in the ReferenceShape for SemiEllipticalShape
'Usage   : The shapes are used for drawing and graph calculation
'****************************************************************************************
On Error GoTo Err_Handler

ReDim Preserve ReferenceShape(13)

Dim ShapeScale As Double
ShapeScale = 1.4914231



ReferenceShape(13).Name = "Mushroom"
ReferenceShape(13).Use = "All"
ReferenceShape(13).CentreOffsetX = 0# * ShapeScale
ReferenceShape(13).CentreOffsetY = 0# * ShapeScale
ReferenceShape(13).NoArcs = 4
ReferenceShape(13).NoLines = 0

'Arc One'
ReferenceShape(13).Arcs(0).OriginX = 0 * ShapeScale
ReferenceShape(13).Arcs(0).OriginY = 0.163958333 * ShapeScale
ReferenceShape(13).Arcs(0).Radius = 0.617291666666667 * ShapeScale
ReferenceShape(13).Arcs(0).StartAngle = 36.85243208
ReferenceShape(13).Arcs(0).EndAngle = 143.14756792


'Arc Two'
ReferenceShape(13).Arcs(1).OriginX = 0 * ShapeScale
ReferenceShape(13).Arcs(1).OriginY = 1.86 * ShapeScale
ReferenceShape(13).Arcs(1).Radius = 2.2 * ShapeScale
ReferenceShape(13).Arcs(1).StartAngle = 253.3
ReferenceShape(13).Arcs(1).EndAngle = 287


'Arc Three'
ReferenceShape(13).Arcs(2).OriginX = 0.21875 * ShapeScale
ReferenceShape(13).Arcs(2).OriginY = 0 * ShapeScale
ReferenceShape(13).Arcs(2).Radius = 0.890625 * ShapeScale
ReferenceShape(13).Arcs(2).StartAngle = 143.1475679
ReferenceShape(13).Arcs(2).EndAngle = 196.3


'Arc Four'
ReferenceShape(13).Arcs(3).OriginX = -0.21875 * ShapeScale
ReferenceShape(13).Arcs(3).OriginY = 0 * ShapeScale
ReferenceShape(13).Arcs(3).Radius = 0.890625 * ShapeScale
ReferenceShape(13).Arcs(3).StartAngle = 343.5
ReferenceShape(13).Arcs(3).EndAngle = 36.85243208

'Arc Five'
ReferenceShape(13).Arcs(4).OriginX = 0.476458333 * ShapeScale
ReferenceShape(13).Arcs(4).OriginY = 0.007916667 * ShapeScale
ReferenceShape(13).Arcs(4).Radius = 0.195416667 * ShapeScale
ReferenceShape(13).Arcs(4).StartAngle = 306.31206108
ReferenceShape(13).Arcs(4).EndAngle = 2.09885489


'Need to be adjusted
'Arc Six'
ReferenceShape(13).Arcs(5).OriginX = -0.476458333 * ShapeScale
ReferenceShape(13).Arcs(5).OriginY = 0.007916667 * ShapeScale
ReferenceShape(13).Arcs(5).Radius = 0.195416667 * ShapeScale
ReferenceShape(13).Arcs(5).StartAngle = 177.9014511
ReferenceShape(13).Arcs(5).EndAngle = 233.68793892






Exit Sub
Err_Handler:

Select Case Err
    Case 53
        Exit Sub
    Case 48
        Exit Sub
    Case Else
        MsgBox Err & "-ST25:" & Error$
End Select
End Sub

Sub DefineCOSRehab()
'****************************************************************************************
' PCN3055
'Name    : DefineCOSRehab
'Created : 17 Jan 2007
'Updated :
'Prg By  : Geoff Logan
'Param   :
'Desc    : Fills in the ReferenceShape for LA COS Rehab project
'Usage   : The shapes are used for drawing and graph calculation
'****************************************************************************************
On Error GoTo Err_Handler

ReDim Preserve ReferenceShape(14)

ReferenceShape(14).Name = "COSRehab"
ReferenceShape(14).Use = "All"
ReferenceShape(14).CentreOffsetX = 0#
ReferenceShape(14).CentreOffsetY = 0#
ReferenceShape(14).NoArcs = 4
ReferenceShape(14).NoLines = 0



'Arc One'
ReferenceShape(14).Arcs(0).OriginX = 0
ReferenceShape(14).Arcs(0).OriginY = 0
ReferenceShape(14).Arcs(0).Radius = 1
ReferenceShape(14).Arcs(0).StartAngle = 6.27729849
ReferenceShape(14).Arcs(0).EndAngle = 173.72270151

'Arc Two'
ReferenceShape(14).Arcs(1).OriginX = 2
ReferenceShape(14).Arcs(1).OriginY = -0.22 '-0.44
ReferenceShape(14).Arcs(1).Radius = 3.01206362
ReferenceShape(14).Arcs(1).StartAngle = 173.72270151
ReferenceShape(14).Arcs(1).EndAngle = 186.22729849

'Arc Three'
ReferenceShape(14).Arcs(2).OriginX = 0
ReferenceShape(14).Arcs(2).OriginY = -0.44 '-0.88
ReferenceShape(14).Arcs(2).Radius = 1
ReferenceShape(14).Arcs(2).StartAngle = 186.27729849
ReferenceShape(14).Arcs(2).EndAngle = 353.72270151

'Arc Four'
ReferenceShape(14).Arcs(3).OriginX = -2
ReferenceShape(14).Arcs(3).OriginY = -0.22 '-0.44
ReferenceShape(14).Arcs(3).Radius = 3.01206362
ReferenceShape(14).Arcs(3).StartAngle = 353.72270151
ReferenceShape(14).Arcs(3).EndAngle = 6.27729849

Exit Sub
Err_Handler:

Select Case Err
    Case 53
        Exit Sub
    Case 48
        Exit Sub
    Case Else
        MsgBox Err & "-ST26:" & Error$
End Select
End Sub


'Sub DefineCustomShape()
''****************************************************************************************
'' PCN3055
''Name    : DefineSemiEllipticalShape
''Created : 14 December 2004
''Updated :
''Prg By  : Antony van Iersel
''Param   :
''Desc    : Fills in the ReferenceShape for SemiEllipticalShape
''Usage   : The shapes are used for drawing and graph calculation
''****************************************************************************************
'On Error GoTo Err_Handler
'
'ReDim Preserve ReferenceShape(3)
'
'ReferenceShape(3).name = "Custom"
'ReferenceShape(3).Use = "All"
'ReferenceShape(3).CentreOffsetX = 0#
'ReferenceShape(3).CentreOffsetY = 0.5
'ReferenceShape(3).NoArcs = 4
'ReferenceShape(3).NoLines = 0
'
'
'
''Arc One'
'ReferenceShape(3).Arcs(0).OriginX = 0
'ReferenceShape(3).Arcs(0).OriginY = 0
'ReferenceShape(3).Arcs(0).Radius = 1
'ReferenceShape(3).Arcs(0).StartAngle = 0
'ReferenceShape(3).Arcs(0).EndAngle = 180
'
''Arc Two'
'ReferenceShape(3).Arcs(1).OriginX = 0
'ReferenceShape(3).Arcs(1).OriginY = -1.5
'ReferenceShape(3).Arcs(1).Radius = 0.5
'ReferenceShape(3).Arcs(1).StartAngle = 206.56505118
'ReferenceShape(3).Arcs(1).EndAngle = 333.4349489
'
''Arc Three'
'ReferenceShape(3).Arcs(2).OriginX = 2
'ReferenceShape(3).Arcs(2).OriginY = 0
'ReferenceShape(3).Arcs(2).Radius = 3
'ReferenceShape(3).Arcs(2).StartAngle = 180
'ReferenceShape(3).Arcs(2).EndAngle = 216
'
''Arc Four'
'ReferenceShape(3).Arcs(3).OriginX = -2
'ReferenceShape(3).Arcs(3).OriginY = 0
'ReferenceShape(3).Arcs(3).Radius = 3
'ReferenceShape(3).Arcs(3).StartAngle = 324
'ReferenceShape(3).Arcs(3).EndAngle = 0
'
'Exit Sub
'Err_Handler:
'
'Select Case Err
'    Case 53
'        Exit Sub
'    Case 48
'        Exit Sub
'    Case Else
'        MsgBox Err & " - " & error$
'End Select
'End Sub
'Sub DefineSemiEllipticalShapeWC()
''****************************************************************************************
'' PCN3055
''Name    : DefineSemiEllipticalShape Watercare Interpretation
''Created : 14 December 2004
''Updated :
''Prg By  : Antony van Iersel
''Param   :
''Desc    : Fills in the ReferenceShape for SemiEllipticalShape
''Usage   : The shapes are used for drawing and graph calculation
''****************************************************************************************
'On Error GoTo Err_Handler
'
'ReDim Preserve ReferenceShape(2)
'
'ReferenceShape(2).name = "SemiEllipticalWC"
'ReferenceShape(2).Use = "All"
'ReferenceShape(2).CentreOffsetX = 0#
'ReferenceShape(2).CentreOffsetY = -0.5
'ReferenceShape(2).NoArcs = 8
'ReferenceShape(2).NoLines = 0
'
''Arc One'
'ReferenceShape(2).Arcs(0).OriginX = 1.185424541
'ReferenceShape(2).Arcs(0).OriginY = 0
'ReferenceShape(2).Arcs(0).Radius = 2.189177915
'ReferenceShape(2).Arcs(0).startAngle = 139#
'ReferenceShape(2).Arcs(0).endAngle = 184#
'
'
''Arc Two'
'ReferenceShape(2).Arcs(1).OriginX = 0.028862017
'ReferenceShape(2).Arcs(1).OriginY = 0.931308586
'ReferenceShape(2).Arcs(1).Radius = 0.714285714
'ReferenceShape(2).Arcs(1).startAngle = 93#
'ReferenceShape(2).Arcs(1).endAngle = 132#
'
''Arc Three'
'ReferenceShape(2).Arcs(2).OriginX = -0.029096363
'ReferenceShape(2).Arcs(2).OriginY = 0.926462317
'ReferenceShape(2).Arcs(2).Radius = 0.714285714
'ReferenceShape(2).Arcs(2).startAngle = 49#
'ReferenceShape(2).Arcs(2).endAngle = 88#
'
''Arc Four'
'ReferenceShape(2).Arcs(3).OriginX = -1.185648669
'ReferenceShape(2).Arcs(3).OriginY = 0
'ReferenceShape(2).Arcs(3).Radius = 2.189576303
'ReferenceShape(2).Arcs(3).startAngle = 357#
'ReferenceShape(2).Arcs(3).endAngle = 42#
'
''Arc Five'
'ReferenceShape(2).Arcs(4).OriginX = 0.28775778
'ReferenceShape(2).Arcs(4).OriginY = -0.069460067
'ReferenceShape(2).Arcs(4).Radius = 0.707724034
'ReferenceShape(2).Arcs(4).startAngle = 344#
'ReferenceShape(2).Arcs(4).endAngle = 356#
'
''Arc Six'
'ReferenceShape(2).Arcs(5).OriginX = 0#
'ReferenceShape(2).Arcs(5).OriginY = 2.223472066
'ReferenceShape(2).Arcs(5).Radius = 2.678571429
'ReferenceShape(2).Arcs(5).startAngle = 270#
'ReferenceShape(2).Arcs(5).endAngle = 292#
'
''Arc Seven'
'ReferenceShape(2).Arcs(6).OriginX = 0#
'ReferenceShape(2).Arcs(6).OriginY = 2.223472066
'ReferenceShape(2).Arcs(6).Radius = 2.678571429
'ReferenceShape(2).Arcs(6).startAngle = 249#
'ReferenceShape(2).Arcs(6).endAngle = 270#
'
''Arc Eight'
'ReferenceShape(2).Arcs(7).OriginX = -0.279668166
'ReferenceShape(2).Arcs(7).OriginY = -0.061961005
'ReferenceShape(2).Arcs(7).Radius = 0.720847394
'ReferenceShape(2).Arcs(7).startAngle = 185#
'ReferenceShape(2).Arcs(7).endAngle = 198#
'
'Exit Sub
'Err_Handler:
'
'Select Case Err
'    Case 53
'        Exit Sub
'    Case 48
'        Exit Sub
'    Case Else
'        MsgBox Err & " - " & error$
'End Select
'End Sub

'Sub DefineSemiEllipticalShapeWCInverted()
''****************************************************************************************
'' PCN3055
''Name    : DefineSemiEllipticalShape Watercare Interpretation, Inverted
''Created : 14 December 2004
''Updated :
''Prg By  : Antony van Iersel
''Param   :
''Desc    : Fills in the ReferenceShape for SemiEllipticalShape
''Usage   : The shapes are used for drawing and graph calculation
''****************************************************************************************
'On Error GoTo Err_Handler
'
'ReDim Preserve ReferenceShape(3)
'
'ReferenceShape(3).name = "SemiEllipticalWCInverted"
'ReferenceShape(3).Use = "All"
'ReferenceShape(3).CentreOffsetX = 0#
'ReferenceShape(3).CentreOffsetY = -0.5
'ReferenceShape(3).NoArcs = 8
'ReferenceShape(3).NoLines = 0
'
''Arc One'
'ReferenceShape(3).Arcs(0).OriginX = -1.185424541
'ReferenceShape(3).Arcs(0).OriginY = 0
'ReferenceShape(3).Arcs(0).Radius = 2.189177915
'ReferenceShape(3).Arcs(0).startAngle = 356#
'ReferenceShape(3).Arcs(0).endAngle = 41#
'
'
''Arc Two'
'ReferenceShape(3).Arcs(1).OriginX = -0.028862017
'ReferenceShape(3).Arcs(1).OriginY = 0.931308586
'ReferenceShape(3).Arcs(1).Radius = 0.714285714
'ReferenceShape(3).Arcs(1).startAngle = 48#
'ReferenceShape(3).Arcs(1).endAngle = 87#
'
''Arc Three'
'ReferenceShape(3).Arcs(2).OriginX = 0.029096363
'ReferenceShape(3).Arcs(2).OriginY = 0.926462317
'ReferenceShape(3).Arcs(2).Radius = 0.714285714
'ReferenceShape(3).Arcs(2).startAngle = 92#
'ReferenceShape(3).Arcs(2).endAngle = 131#
'
''Arc Four'
'ReferenceShape(3).Arcs(3).OriginX = 1.185648669
'ReferenceShape(3).Arcs(3).OriginY = 0
'ReferenceShape(3).Arcs(3).Radius = 2.189576303
'ReferenceShape(3).Arcs(3).startAngle = 138#
'ReferenceShape(3).Arcs(3).endAngle = 183#
'
''Arc Five'
'ReferenceShape(3).Arcs(4).OriginX = -0.28775778
'ReferenceShape(3).Arcs(4).OriginY = -0.069460067
'ReferenceShape(3).Arcs(4).Radius = 0.707724034
'ReferenceShape(3).Arcs(4).startAngle = 184#
'ReferenceShape(3).Arcs(4).endAngle = 196#
'
''Arc Six'
'ReferenceShape(3).Arcs(5).OriginX = 0#
'ReferenceShape(3).Arcs(5).OriginY = 2.223472066
'ReferenceShape(3).Arcs(5).Radius = 2.678571429
'ReferenceShape(3).Arcs(5).startAngle = 248#
'ReferenceShape(3).Arcs(5).endAngle = 270#
'
''Arc Seven'
'ReferenceShape(3).Arcs(6).OriginX = 0#
'ReferenceShape(3).Arcs(6).OriginY = 2.223472066
'ReferenceShape(3).Arcs(6).Radius = 2.678571429
'ReferenceShape(3).Arcs(6).startAngle = 270#
'ReferenceShape(3).Arcs(6).endAngle = 291#
'
''Arc Eight'
'ReferenceShape(3).Arcs(7).OriginX = 0.279668166
'ReferenceShape(3).Arcs(7).OriginY = -0.061961005
'ReferenceShape(3).Arcs(7).Radius = 0.720847394
'ReferenceShape(3).Arcs(7).startAngle = 342#
'ReferenceShape(3).Arcs(7).endAngle = 355#
'
'Exit Sub
'Err_Handler:
'
'Select Case Err
'    Case 53
'        Exit Sub
'    Case 48
'        Exit Sub
'    Case Else
'        MsgBox Err & " - " & error$
'End Select
'End Sub




Sub InitialiseGraphStates()

'****************************************************************************************
' PCN3373
'Name    : InitialiseGraphStates
'Created : 16 March 2005
'Updated :
'Prg By  : Antony van Iersel
'Param   : None
'Desc    : Initialises the Graph States from ScreenDrawing, these are used
'          to keep track of the graph states so they can be updated erases etc
'****************************************************************************************
On Error GoTo Err_Handler

Dim i As Integer
For i = 0 To 5
    ImageGraphState(i).XScale = 1
    ImageGraphState(i).LeftLimitLine = 0
    ImageGraphState(i).RightLimitLine = 0
    ImageGraphState(i).CentreOffset = 0
    ImageGraphState(i).Left = 0
    ImageGraphState(i).Right = 200
    ImageGraphState(i).PreviousStartFrame = 0 'PCN???? ANT (was endframe)
    ImageGraphState(i).PreviousEndFrame = 0
    ImageGraphState(i).PreviousGraphType = ""
    ImageGraphState(i).PreviousSpeed = 0
    ImageGraphState(i).PreviousUnits = "Frames"
    ImageGraphState(i).GraphType = ""
    
    Set ImageGraphState(i).PictureImage = PrecisionVisionGraph.PVGraphImage(i)
Next i

ImageGraphState(6).XScale = 1
ImageGraphState(6).LeftLimitLine = 0
ImageGraphState(6).RightLimitLine = 0
ImageGraphState(6).CentreOffset = 0
ImageGraphState(6).Left = 0
ImageGraphState(6).Right = 200
ImageGraphState(6).PreviousStartFrame = 0 'PCN???? ANT (was endframe)
ImageGraphState(6).PreviousEndFrame = 0
ImageGraphState(6).PreviousGraphType = ""
ImageGraphState(6).PreviousSpeed = 0
ImageGraphState(6).PreviousUnits = "Frames"
ImageGraphState(6).GraphType = "Flat"
Set ImageGraphState(6).PictureImage = PrecisionVisionGraph.PrinterReportImage

ImageRulerState.XScale = 1
ImageRulerState.LeftLimitLine = 0
ImageRulerState.RightLimitLine = 0
ImageRulerState.CentreOffset = 0
ImageRulerState.Left = 0
ImageRulerState.Right = 200
ImageRulerState.PreviousStartFrame = 0 'PCN???? ANT (was endframe)
ImageRulerState.PreviousEndFrame = 0
ImageRulerState.PreviousGraphType = ""
ImageRulerState.PreviousUnits = "Frames"
Set ImageRulerState.PictureImage = PrecisionVisionGraph.PVYScaleImage

Exit Sub
Err_Handler:
        MsgBox Err & "-ST27:" & Error$
End Sub

Function ConvertPerToReal(ByVal Per As Double, ByVal ConvType As String) As Double

'****************************************************************************************
' PCN3373
'Name    : ConvertPerToReal
'Created : 20 April 2005
'Updated :
'Prg By  : Antony van Iersel
'Param   : Percentage to convert
'Desc    :
'****************************************************************************************
On Error GoTo Err_Handler

Dim ExpectedArea As Double



Select Case ConvType
    Case "Dia": ConvertPerToReal = ((100 + Per) / 100) * ExpectedDiameter
    Case "Rad": ConvertPerToReal = ((100 + Per) / 100) * (ExpectedDiameter / 2)
'    Case "Flat": ConvertPerToReal = (((100 + Per) / 100) - 1) * (ExpectedDiameter / 2)
    Case "Flat": ConvertPerToReal = (Per / 100) * (ExpectedDiameter / 2)
    Case "Area"
        If MeasurementUnits = "mm" Then
            ExpectedArea = (ExpectedDiameter / 20) * _
                           (ExpectedDiameter / 20) * _
                           PI
        Else
            ExpectedArea = (ExpectedDiameter / 2) * _
                           (ExpectedDiameter / 2) * _
                           PI
        End If
        ConvertPerToReal = Per / 100 * ExpectedArea
End Select

Exit Function
Err_Handler:
Select Case Err
    Case 13 'Invalid data 'PCN3268
        ConvertPerToReal = 0
    Case Else
        MsgBox Err & "-ST28:" & Error$
        Resume Next
End Select
End Function
Function ConvertRealToPer(ByVal unit As Double, ByVal ConvType As String) As Double

'****************************************************************************************
' PCN3373
'Name    : ConvertPerToReal
'Created : 20 April 2005
'Updated :
'Prg By  : Antony van Iersel
'Param   : Percentage to convert
'Desc    :
'****************************************************************************************
On Error GoTo Err_Handler

Dim ExpectedArea As Double
Dim Radius As Double

If ExpectedDiameter = 0 Then Exit Function
Radius = ExpectedDiameter / 2

Select Case ConvType

    Case "Dia": ConvertRealToPer = 100 * (unit - ExpectedDiameter) / ExpectedDiameter
    Case "Rad": ConvertRealToPer = 100 * (unit - Radius) / Radius
    Case "Area"
        If MeasurementUnits = "mm" Then
            ExpectedArea = (ExpectedDiameter / 20) * _
                           (ExpectedDiameter / 20) * _
                           PI
        Else
            ExpectedArea = (ExpectedDiameter / 2) * _
                           (ExpectedDiameter / 2) * _
                           PI
        End If
        ConvertRealToPer = 100 * (unit - ExpectedArea) / ExpectedArea
End Select

Exit Function
Err_Handler:
Select Case Err
    Case 13 'Invalid data 'PCN3268
        ConvertRealToPer = 0
    Case Else
        MsgBox Err & "-ST29:" & Error$
        Resume Next
End Select
End Function
Function ConvertPerToRealByGraph(ByVal Per As Double, ByVal Index As Integer, ByRef DisplayUnits As String) As Double

'****************************************************************************************
' PCN3373
'Name    : ConvertPerToReal
'Created : 20 April 2005
'Updated :
'Prg By  : Antony van Iersel
'Param   : Percentage to convert
'Desc    : When you want to convert a known percentage to a real unit measurement with a condition of the result
'        : relying on what type of graph it is, eg if its Capacity then result is in
'****************************************************************************************
On Error GoTo Err_Handler
Dim GraphType As String
Dim UnitsIndex As Integer
Dim NoOfGraphs As Integer
Dim UnitType As String

GraphType = ImageGraphState(Index).GraphType

'vvvv PCN4207 ************************************
''NoOfGraphs = UBound(PVGraphOrder)
''For UnitsIndex = 0 To NoOfGraphs
''    If PVGraphOrder(UnitsIndex) = GraphType Then
''        UnitType = PVXScaleUnits(UnitsIndex)
''        Exit For
''    End If
''Next UnitsIndex

UnitsIndex = GetGraphInfoIndex(Index)

'PCN5186 added following two lines
If ImageGraphState(0).GraphType = "XYDiameter" And MedianFlat Then
    ConvertPerToRealByGraph = Per
ElseIf GraphInfoContainer(UnitsIndex).PVXScaleUnits = "Real" Then
    ConvertPerToRealByGraph = ConvertPerToReal(Per, "Dia")
ElseIf GraphInfoContainer(UnitsIndex).PVXScaleUnits = "Area" Then
    ConvertPerToRealByGraph = ConvertPerToReal(Per, "Area")
Else
    ConvertPerToRealByGraph = Per
End If
'^^^^ ********************************************

''Select Case GraphType
''    Case "Ovality": ConvertPerToRealByGraph = Per: Exit Function 'PVXScaleUnits = "%": Exit Function  '
''    Case "Delta", _
''         "MaxMinDiameter", _
''         "MedianDiameter", _
''         "MaxDiameter", _
''         "XYDiameter": ConvertPerToRealByGraph = Per ' These graphs are allready measurements
''    Case "Capacity": ConvertPerToRealByGraph = ConvertPerToReal(Per, "Area")
''End Select

If MeasurementUnits = "mm" Then
    If GraphType = "Capacity" Then
        DisplayUnits = "cm"
    Else: DisplayUnits = "mm"
    End If
Else: DisplayUnits = "in"
End If
 
Exit Function
Err_Handler:
Select Case Err
    Case 13 'Invalid data 'PCN3268
        ConvertPerToRealByGraph = 0
    Case Else
        MsgBox Err & "-ST2A:" & Error$
End Select
End Function


Function ConvertRealToPerByGraph(ByVal Measurement As Double, ByVal Index As Integer, ByRef DisplayUnits As String) As Double

'****************************************************************************************
' PCN3373
'Name    : ConvertRealToPer
'Created : 16 August 2005
'Updated :
'Prg By  : Antony van Iersel
'Param   : Unit , Unit to convert to percent
'        : Index ,Graph Index to retrieve the type of graph to convert
'        : DisplyUnits, Type of units to display, at the moment it is all "%"
'Return  : Returns the converted value as a double
'Desc    : When you want to convert a real unit Measurement (mm, or inch etc) to a percentage of the expected
'        : diameter or radius
'****************************************************************************************
On Error GoTo Err_Handler

Dim GraphType As String
Dim ExpectedArea As Double

GraphType = ImageGraphState(Index).PreviousGraphType

'PCN5186 added folling lines and corisponding end if
If (GraphType = "XYDiameter" And MedianFlat) Then 'PCN6458 Or GraphType = "Inclination" Then 'PCN6128 added inclination
    ConvertRealToPerByGraph = Measurement
Else
    Select Case GraphType
        Case "Delta": ConvertRealToPerByGraph = ConvertRealToPer(Measurement, "Rad")
        Case "MaxMinDiameter", _
             "MedianDiameter", _
             "XYDiameter", _
             "MaxDiameter", _
             "MinDiameter": ConvertRealToPerByGraph = ConvertRealToPer(Measurement, "Dia")  'PCN4235
    
             
        Case "Capacity", "Ovality": ConvertRealToPerByGraph = Measurement 'These graphs are allread %
    End Select
End If


DisplayUnits = "%"

 
Exit Function
Err_Handler:
Select Case Err
    Case 13 'Invalid data 'PCN3268
        ConvertRealToPerByGraph = 0
    Case Else
        MsgBox Err & "-ST2B:" & Error$
End Select
End Function

Function ConvertUnitByGraph(ByVal Measurement As Double, ByVal Index As Integer, ByRef DisplayUnits As String) As Double

'****************************************************************************************
' PCN3373
'Name    : ConvertUnitsByGraph
'Created : 16 August 2005
'Updated :
'Prg By  : Antony van Iersel
'Param   : Unit , Unit to convert
'        : Index ,Graph Index to retrieve the type of graph to convert
'        : DisplyUnits, Type of units to display
'Return  : Returns the converted value as a double
'****************************************************************************************

On Error GoTo Err_Handler
Dim UnitsIndex As Integer
Dim NoOfGraphs As Integer
Dim UnitType As String


'vvvv PCN4207 ************************************
''NoOfGraphs = UBound(PVGraphOrder)
''For UnitsIndex = 0 To NoOfGraphs
''    If PVGraphOrder(UnitsIndex) = ImageGraphState(Index).PreviousGraphType Then
''        UnitType = PVXScaleUnits(UnitsIndex)
''        Exit For
''    End If
''Next UnitsIndex

UnitsIndex = GetGraphInfoIndex(Index)

'PCN5186 added the followin two lines
If ImageGraphState(0).GraphType = "XYDiameter" And MedianFlat Then
    ConvertUnitByGraph = ConvertRealToPerByGraph(Measurement, Index, DisplayUnits)
ElseIf GraphInfoContainer(UnitsIndex).PVXScaleUnits = "Real" Then
'    ConvertUnitByGraph = ConvertPerToRealByGraph(Measurement, Index, DisplayUnits)
    ConvertUnitByGraph = Measurement
    DisplayUnits = MeasurementUnits
Else
    ConvertUnitByGraph = ConvertRealToPerByGraph(Measurement, Index, DisplayUnits)
End If
'^^^^ ********************************************

Exit Function
Err_Handler:
Select Case Err
    Case 13 'Invalid data 'PCN3268
        ConvertUnitByGraph = 0
    Case Else
        MsgBox Err & "-ST2C:" & Error$
End Select

End Function

Function GetGraphInfoIndex(ByVal ImageGraphStateIndex As Integer) As Integer 'PCN4207
On Error GoTo Err_Handler
Dim GraphType As String
Dim UnitsIndex As Integer
Dim NoOfGraphs As Integer
Dim UnitType As String

GetGraphInfoIndex = 0

GraphType = ImageGraphState(ImageGraphStateIndex).GraphType

NoOfGraphs = UBound(GraphInfoContainer)
For UnitsIndex = 0 To NoOfGraphs
    If GraphInfoContainer(UnitsIndex).GraphType = GraphType Then
        GetGraphInfoIndex = UnitsIndex
        Exit Function
    End If
Next UnitsIndex


Exit Function
Err_Handler:
    MsgBox Err & "-ST2D:" & Error$
End Function

Function GetPVDVer() As Single
On Error GoTo Err_Handler
    Dim Ver As String
    Ver = ConfigInfo.PVDFileVersion
    If Ver = "" Then GetPVDVer = 0: Exit Function
    If Ver = "6.25" Then GetPVDVer = 6.25: Exit Function 'PCN4168 return wrong version number.
    GetPVDVer = CSng(Right(Ver, Len(Ver) - 1))
    
Exit Function
Err_Handler:
    GetPVDVer = 0
End Function

Function EnumWinProc(ByVal hwnd As Long, ByVal lParam As Long) As Long
On Error GoTo Error_handler

    Dim i As Integer
  Dim k As Long, sName As String
    i = 0
  
  If IsWindowVisible(hwnd) And GetParent(hwnd) = 0 Then
     sName = Space$(128)
     k = GetWindowText(hwnd, sName, 128)
     If k > 0 Then
        sName = Left$(sName, k)
        If lParam = 0 Then sName = UCase(sName)
        If sName Like sPattern Then
            i = i + 1
           hFind = hwnd
           EnumWinProc = 0
           If i = 2 Then Exit Function
        End If
     End If
  End If
  EnumWinProc = 1
  
Exit Function
Error_handler:
    Select Case Err
        Case Else: MsgBox Err & "-ST2E:" & Error$, vbExclamation
    End Select
End Function

Public Function FindWindowWild(sWild As String, Optional bMatchCase As Boolean = True) As Long
On Error GoTo Error_handler

  sPattern = sWild
  If Not bMatchCase Then sPattern = UCase(sPattern)
  EnumWindows AddressOf EnumWinProc, bMatchCase
  FindWindowWild = hFind
  
  Exit Function
Error_handler:
    Select Case Err
        Case Else: MsgBox Err & "-ST2F:" & Error$, vbExclamation
    End Select
End Function

Function GetCurrentLanguageSetting() As String
On Error GoTo Err_Handler
Dim INFFile As String
Dim FileNumber As Integer
Dim INIFileName As String
Dim INILanguage As String

GetCurrentLanguageSetting = "English"

If SoftwareConfiguration <> "Reader" Then    '
    'PCN6523 remove INF' INFFile = WindowsTempDirectory & "CBS\Clearline.inf"
    
    'PCN6523 remove INF' If Dir(INFFile) = "" Then
    'PCN6523 remove INF'     Exit Function
    'PCN6523 remove INF' End If
    
    'PCN6523 remove INF' FileNumber = FreeFile
    'PCN6523 remove INF' Open INFFile For Input As #FileNumber
     'PCN6523 remove INF' Do While Not EOF(FileNumber)
      'PCN6523 remove INF' Line Input #FileNumber, INIFileName
     'PCN6523 remove INF' Loop
    'PCN6523 remove INF' Close #FileNumber
    
    INIFileName = WindowsTempDirectory & "Clearline.ini" 'ID4601 "CBS\Clearline.ini"
    
    If Dir(INIFileName) = "" Then
        'First check if the INI is located in the Application
        'directory before asking the user for its location.
        'PCN6523 remove INF' INIFileName = App.Path & "\Clearline.ini"
        'PCN6523 remove INF' If Dir(INIFileName) = "" Then
            Exit Function
        'PCN6523 remove INF' End If
    End If
Else
        INIFileName = App.Path & "\deploy.ini"
        If Dir(INIFileName) = "" Then
            Exit Function
        End If
End If

Call GetINI_ParameterInfoOnly(INIFileName, "Language=", INILanguage)

If INILanguage <> "" Then
    GetCurrentLanguageSetting = INILanguage
End If

Exit Function
Err_Handler:
    MsgBox Err & "-ST30:" & Error$
End Function

Sub SetupSoftwareForFullConfiguration(strLoadPVDFile As String)
On Error GoTo Err_Handler
Dim ProductName As String
Dim ErrorStr As String
Dim INFFile As String
Dim INFReadONLY As Boolean
Dim X As Integer
Dim Y As Integer
Dim ScreenRes As Integer




IPD = False ' PCN4171

ReDim PipeObservations(0) 'To initialise size of observations array 'PCNGL130103

'vvvv PCN3115 *********************************************************
'Check to see if this application has been load with a command line PVD

strLoadPVDFile = Command()

'For testing, do not have this line un commented. Testing only.
'strLoadPVDFile = "CompanyName " & Chr(34) & "C:\Program Files\ClearLine Profiler\clpinterface.int" & Chr(34)
If Len(strLoadPVDFile) > 2 Then
    'Remove string quote marks , Only if the first character is a quote (Last bit added 22 Sept 2005, Antony)
    If Left(strLoadPVDFile, 1) = Chr(34) Then
        strLoadPVDFile = Mid(strLoadPVDFile, 2, Len(strLoadPVDFile) - 2)
    End If
End If
'^^^^ *****************************************************************
'For testing the Interface file, force the load to be "Granit"'''''''
'strLoadPVDFile = "Granit"                                           '

'vvvv PCN2212 ***************************
Dim ArmClass As CArmadillo
Dim ArmRegType As String


Set ArmClass = New CArmadillo

ArmRegType = ArmClass.ClearLineRegType

If Len(ArmClass.GetVersionNumber) = 0 Then
    'MsgBox DisplayMessage("Security failure. Unable to register application.")
    ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("Security failure. Unable to register application."): ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
    ArmRegType = ""
ElseIf ArmadilloVersion <> ArmClass.GetVersionNumber Then
    'MsgBox DisplayMessage("Security failure. Unable to register application.")
    ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("Security failure. Unable to register application."): ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
    ArmRegType = ""
ElseIf ArmClass.UserName = "" Or ArmClass.UserName = "Program Author" Then
    ArmRegType = ""
ElseIf ArmClass.UserKey = "" Then
    ArmRegType = ""
End If

'MsgBox "Key = " & ArmClass.UserKey
'MsgBox "Invalid Key = " & ArmClass.InvalidKey
'MsgBox "Valid Key = " & ArmClass.IsValidKey(ArmClass.UserName, ArmClass.UserKey)

Select Case ArmRegType
    Case "RegisteredWith3D"
        Registered = True
        ThreeDActivated = True
    Case "Registered", "SONAR"
        Registered = True
'        ThreeDActivated = False 'PCN2861
        ThreeDActivated = True 'PCN2861
    Case Else
'PCN4003 Fablock has to be disabled :( It doesn't work with the hongkong chinese
'        If Not StrongIsRegistered(ClearLineProfilerV5.FabLock1) Then 'Checking on registration of product PCNML310103
            Registered = False
            ThreeDActivated = True
'        Else
'            Registered = True
'            ThreeDActivated = True  'PCNANT
'        End If
End Select

'''''''''' WARNING ''''''''''''''''''''''''''''''''''''''
'For Testing PCN3816 ONLY dont forget to uncoment out....
MsgBox DisplayMessage("Dont forget to turn registration back on"): Registered = True

'ThreeDActivated = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'MsgBox "Registered as = " & ArmRegType

'vvvv PCN2230 *******************
'For 3D
'ThreeDActivated = False
ThreeDRunning = False
'^^^^ ***************************


ReDim WaterEgnoreList(180) 'PCN3219
PrinterClipOn = True 'PCN3691


'^^^^ ***************************************
RecordTimeIncrement = 100 ' Now number of frames captured before asking CPP to dump data' 5  '60 'In seconds, the time interval the C records for before drawing the graphs
DistanceMethod = "None" 'PCN2463

DrawShapeType = "Circle" 'PCN3055 Default draw type is Circle (Antony, 13 December 2004)
'DrawShapeType = "SemiElliptical"
PVDrawScreenRatio = 1

Call CheckRegionalSettings

Call LoadShapes

'Call DefineCircleShape 'PCN3055 circle needed (perfect circle)
'Call DefineSemiEllipticalShape 'PCN3055
'Call DefineEggShape 'PCN3055 sets up the shape definition for Semi Eliptical
'Call DefineRinkerEllipseShape
'Call DefineEllipticalASTM_C507Shape
'Call DefineBoxCulvertShape
'Call DefineBarnShape
'Call DefineBarnDShape
'Call DefineEggAShape
'Call DefineEggBShape
'Call DefineCupCakeShape
'Call DefineBulletShape
'Call DefineSquareShape
'Call DefineMushroomShape
'Call DefineCOSRehab 'PCNGL170107
'DefineSemiEllipticalShapeWCInverted 'PCN3055
Call InitialiseNumberPicArray 'PCN3691

'Getting the user's screen resolution PCN1876
X = Screen.width / 15
Y = Screen.height / 15
'If x = 800 And y = 600 Then MsgBox ("800 * 600") 'Testing ML050303
'If x = 1024 And y = 768 Then MsgBox ("1024 * 768") 'Testing ML050303
If X = 800 And Y = 600 Then ScreenRes = 800
If X = 1024 And Y = 768 Then ScreenRes = 1024
'

PVDSaved = False 'Set the variable for the recording of a .pvd file PCN1895

NoOfProfileSegments = 180 'PCNGL1812022 The number of segments within a single profile (usually 180, may increase to 360)

VideoFrame = 1
FramesPerSec = 25 'Frames per second (Use as Default) 'PCNGL130103
DrawAutoSnap = True 'Snaps to the nearest line or cirle 'PCNGL210103


MyFile = WindowsTempDirectory & "Clearline.ini" 'ID4601 'PCN6471 MyFile = App.Path & "\Clearline.ini"
'MyFile = WindowsTempDirectory & "CBS\Clearline.ini" 'PCN6471 MyFile = App.Path & "\Clearline.ini"

'INFFile = WindowsTempDirectory & "CBS\Clearline.inf" '6471 INFFile = App.Path & "\Clearline.inf"

' PCN6523 removing INF file
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''If Dir(INFFile) = "" Then                                                 '
''    Dim FileNo                                                            '
''    FileNo = FreeFile                                                     '
''    Open INFFile For Output As #FileNo                                    '
''                                                                          '
''    Print #FileNo, WindowsTempDirectory & "CBS\Clearline.ini"             '
''    Close #FileNo                                                         '
''End If                                                                    '
''                                                                          '
''FindInf:                                                                  '
''If Dir(INFFile) = "" Then                                                 '
''                                                                          '
''                                                                          '
''    ClearLineProfilerV6.Dialog.FileName = INFFile                         '
''    ClearLineProfilerV6.Dialog.DialogTitle = "ClearLine" 'PCN2111         '
''    ClearLineProfilerV6.Dialog.Filter = "Information File (*.inf)|*.inf"  '
''    ClearLineProfilerV6.Dialog.CancelError = True                         '
''    ClearLineProfilerV6.Dialog.ShowOpen                                   '
''    INFFile = ClearLineProfilerV6.Dialog.FileName                         '
''    GoTo FindInf                                                          '
''End If                                                                    '
''                                                                          '
''Open INFFile For Input As #1                                              '
'' Do While Not EOF(1)                                                      '
''  Line Input #1, MyFile                                                   '
'' Loop                                                                     '
''Close #1                                                                  '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

If Dir(MyFile) = "" Then
    Call CreateNewINI(MyFile)
End If

Call CheckAndUpdateINI

'FindIni:
'If Dir(MyFile) = "" Then
'    'vvvv PCN3320 ***************************************
'    'First check if the INI is located in the Application
'    'directory before asking the user for its location.
'    MyFile = WindowsTempDirectory & "CBS\Clearline.ini" 'PCN6471 MyFile = App.Path & "\Clearline.ini"
'    If Dir(MyFile) = "" Then
'        ClearLineProfilerV6.Dialog.FileName = MyFile
'        ClearLineProfilerV6.Dialog.DialogTitle = "ClearLine" 'PCN2111
'        ClearLineProfilerV6.Dialog.Filter = "Initialization File|ClearLine.ini"
'        ClearLineProfilerV6.Dialog.CancelError = True
'        ClearLineProfilerV6.Dialog.ShowOpen
'        MyFile = ClearLineProfilerV6.Dialog.FileName
'    End If
'    '^^^^ ***********************************************
'    INFReadONLY = False
'    Open INFFile For Output As #1
'    If Not INFReadONLY Then
'        Print #1, MyFile
'        Close #1
'    Else
'        'MsgBox DisplayMessage("The ClearLine.inf file is READONLY. You must update ClearLine.ini path manually."), vbCritical 'PCNGL170603 'PCN2111
'        ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("The ClearLine.inf file is READONLY. You must update ClearLine.ini path manually."): ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
'        Unload ClearLineProfilerV6
'        Exit Sub
'    End If
'    GoTo FindIni
'End If

'This will get the directory to save to:
LocToSave = Left(MyFile, Len(MyFile) - 13)

'Call CheckAndUpdateINI



'vvvvv ***************************** PCN3197

'this code was moved to below the HASP HL locking code

''vvvvv Setup application ***************************** 'PCNLS170203
'Load ClearLineProfilerV5
'ClearLineProfilerV5.WindowState = vbMaximized
'ClearLineProfilerV5.Show
'V6Splash.Show
'DoEvents
''^^^^^ ***********************************************

'^^^^******************************* PCN3197




'''''''''''''''''''''''''''''''''''''''''''''''
Dim UnitsString As String

Call GetINI_ParameterInfoOnly(MyFile, "MeasurementUnits=", UnitsString)
If UnitsString = "" Then
    Load InstallationForm
    While Not InstallationForm.Complete And IsOpen("InstallationForm") = True
        DoEvents
    Wend
    Unload InstallationForm
End If
''''''''''''''''''''''''''''''''''''''''''''''

Call LoadLanguageFromFile(ErrorStr) 'PCN4171

Call GetINI_ParameterInfoOnly(MyFile, "MeasurementUnits=", UnitsString)
If UnitsString = "" Then Exit Sub

Call GetINI_Information(MyFile)

'vvvv************* PCN3197 ******************************
If Registered = False And HASPLockActive = True Then
    If Validate_HASP_Lock = True Then
        Registered = True
        ThreeDActivated = True
    End If
End If



Exit Sub
Err_Handler:
    Select Case Err
        Case 286: Resume Next 'PCNANT
        Case 52 'Bad file or filename 'PCNGL091202
            Resume Next 'PCNGL091202
        Case Else
            MsgBox Err & "-ST31:" & Error$
    End Select
End Sub


Sub SetupSoftwareForReaderConfiguration()
On Error GoTo Err_Handler
Dim X As Integer
Dim Y As Integer
Dim ScreenRes As Integer



Registered = False
ThreeDActivated = True
ThreeDRunning = False
ThreeDRenderingStyle = 2 'Software

IPD = False ' PCN4171

ReDim PipeObservations(0) 'To initialise size of observations array 'PCNGL130103

ReDim WaterEgnoreList(180) 'PCN3219
PrinterClipOn = True 'PCN3691

RecordTimeIncrement = 100 ' Now number of frames captured before asking CPP to dump data' 5  '60 'In seconds, the time interval the C records for before drawing the graphs
DistanceMethod = "None" 'PCN2463

DrawShapeType = "Circle" 'PCN3055 Default draw type is Circle (Antony, 13 December 2004)
'DrawShapeType = "SemiElliptical"
PVDrawScreenRatio = 1

Call LoadShapes

'Call DefineCircleShape 'PCN3055 circle needed (perfect circle)
'Call DefineSemiEllipticalShape 'PCN3055
'Call DefineEggShape 'PCN3055 sets up the shape definition for Semi Eliptical
'Call DefineBulletShape
'Call DefineEllipticalASTM_C507Shape
'Call DefineBoxCulvertShape
'Call DefineBarnShape
'Call DefineBarnDShape
'Call DefineEggAShape
'Call DefineEggBShape
'Call DefineCupCakeShape
'Call DefineRinkerEllipseShape
'Call DefineSquareShape
'Call DefineMushroomShape
'Call DefineCOSRehab 'PCNGL170107
'DefineSemiEllipticalShapeWCInverted 'PCN3055
Call InitialiseNumberPicArray 'PCN3691

'Getting the user's screen resolution PCN1876
X = Screen.width / 15
Y = Screen.height / 15
'If x = 800 And y = 600 Then MsgBox ("800 * 600") 'Testing ML050303
'If x = 1024 And y = 768 Then MsgBox ("1024 * 768") 'Testing ML050303
If X = 800 And Y = 600 Then ScreenRes = 800
If X = 1024 And Y = 768 Then ScreenRes = 1024
'

PVDSaved = False 'Set the variable for the recording of a .pvd file PCN1895

NoOfProfileSegments = 180 'PCNGL1812022 The number of segments within a single profile (usually 180, may increase to 360)

VideoFrame = 1
FramesPerSec = 25 'Frames per second (Use as Default) 'PCNGL130103
DrawAutoSnap = True 'Snaps to the nearest line or cirle 'PCNGL210103

'vvvv PCN4241 !!!! *********
'LocToSave = App.Path & "\" 'PCN4240, the \ was missing on reader form the app.path
'LocToSave = "C:\"

LocToSave = WindowsTempDirectory 'PCN4568 when reading from a CD needs a readible place to store data., eg snapshot

ReadOnlyAppPath = App.Path & "\" 'PCN
'ReadOnlyAppPath = LocToSave
'^^^^ *****************

MeasurementUnits = "mm" 'by default, should get this from the PVD


Exit Sub
Err_Handler:
    Select Case Err
        Case 286: Resume Next 'PCNANT
        Case 52 'Bad file or filename 'PCNGL091202
            Resume Next 'PCNGL091202
        Case Else
            MsgBox Err & "-ST32:" & Error$
    End Select
End Sub

Sub CheckAndSetupInterface(strLoadPVDFile As String)
On Error GoTo Err_Handler
Dim InterfaceFileName As String
Dim value As String
Dim answer As Integer
 

'vvvv PCN3115 *******************************************************
'If this application has been openned with a PVD in the command line,
'then open this file

InterfaceFileName = Interfaces.ExtractInterfaceFileName(strLoadPVDFile) 'PCN3744
If Dir(InterfaceFileName) <> "" And InterfaceFileName <> "" Then
    Call SetupAsPerInterfaceFile(InterfaceFileName)
ElseIf strLoadPVDFile <> "" Then
    If Dir(strLoadPVDFile) <> "" Then
        If Format(Right(strLoadPVDFile, 4), ">") = ".PVD" Then
            Call OpenAnyFile(strLoadPVDFile)
        End If
    End If
End If


Exit Sub
Err_Handler:
    Select Case Err
        Case 286: Resume Next 'PCNANT
        Case 52 'Bad file or filename 'PCNGL091202
            Resume Next 'PCNGL091202
        Case Else
            MsgBox Err & "-ST33:" & Error$
    End Select
End Sub

Sub CheckAndSetupForUnRegDemoMode()
On Error GoTo Err_Handler
Dim value As String

If Registered = False Then                                  'But only if its not registered
    Call GetINI_ParameterInfoOnly(MyFile, "DemoLoad=", value)   'Check if DemoLoad is wanted
    If value = "yes" Then                                       'If yes then load either
    
        If ConfigInfo.Units = "mm" Then                                     'get the file for mm example
            Call GetINI_ParameterInfoOnly(MyFile, "DemoFile_mm=", value)    'or
        Else                                                                'get the file for inch example
            Call GetINI_ParameterInfoOnly(MyFile, "DemoFile_in=", value)
        End If
        
        If value <> "" Then
            value = App.Path & "\" & value  'Add the path name for the example file names
            If Dir(value) <> "" Then        'Check if it exists and if it does and
                If Format(Right(value, 4), ">") = ".PVD" Then   'its a pvd
                    
                    'before loading file, ask if this file is to be loaded on next run
               '     Answer = MsgBox(DisplayMessage("Do you want to load example .PVD on next run."), vbYesNo + vbQuestion)
               '     If Answer = vbNo Then Call INI_WriteBack(MyFile, "DemoLoad=", "no")
                    
                    Call OpenAnyFile(value)                     'then load it
                End If
            End If
        End If
        
    End If
End If


Exit Sub
Err_Handler:
    Select Case Err
        Case 286: Resume Next 'PCNANT
        Case 52 'Bad file or filename 'PCNGL091202
            Resume Next 'PCNGL091202
        Case Else
            MsgBox Err & "-ST34:" & Error$
    End Select
End Sub

Sub ReaderLoadPVD()
On Error GoTo Err_Handler
Dim strLoadPVDFile As String


strLoadPVDFile = Dir(App.Path & "\*.PVD")
If strLoadPVDFile = "" Then Exit Sub 'PCN4241 !!!!

strLoadPVDFile = App.Path & "\" & strLoadPVDFile
If Format(Right(strLoadPVDFile, 4), ">") = ".PVD" Then
    Call OpenAnyFile(strLoadPVDFile)
End If


Exit Sub
Err_Handler:
    Select Case Err
        Case 286: Resume Next 'PCNANT
        Case 52 'Bad file or filename 'PCNGL091202
            Resume Next 'PCNGL091202
        Case Else
            MsgBox Err & "-ST35:" & Error$
    End Select
End Sub

Sub GetSystem32Dir()
On Error GoTo Err_Handler

Dim fso
Dim DoesExist As Boolean
Dim FolderType As Integer
Dim SysFolder As Variant

FolderType = 1 'SystemFolder

Set fso = CreateObject("Scripting.FileSystemObject")

SysFolder = fso.GetSpecialFolder(FolderType)
SystemDir = CStr(SysFolder) & "\"

Exit Sub
Err_Handler:
    MsgBox Err & "-ST36:" & Error$
End Sub

Function SafeCDbl(ByVal TheString As String) As Double
On Error GoTo Err_Handler
    TheString = Trim(TheString)
    
    If Not IsNumeric(TheString) Then SafeCDbl = 0: Exit Function
    
    Dim Index As Integer
    Dim length As Integer
    Dim decPoint As Integer
    Dim character As String
    Dim devision As Double
    Dim value As Double
    Dim placing As Integer
    
    
    placing = 1
    
    length = Len(TheString)
    For Index = 1 To length
        character = Mid(TheString, Index, 1)
        If (character < "0" Or character > "9") And character <> "-" Then devision = (10 ^ placing) / 100: Exit For
        If character <> "-" Then placing = placing + 1
    Next Index
    If devision = 0 Then devision = (10 ^ placing) / 100
    For Index = 1 To length
        character = Mid(TheString, Index, 1)
        If character >= "0" And character <= "9" Then
            value = value + (CDbl(character) * devision)
            devision = devision / 10
        End If
    Next Index
    
    If Left(TheString, 1) = "-" Then value = value * -1
    
    SafeCDbl = value
    
Exit Function
Err_Handler:
    MsgBox Err & "-ST37:" & Error$
End Function

Sub CreateNewINI(ByVal FileName As String)
On Error GoTo Err_Handler

    Dim FileNo
    FileNo = FreeFile
    
    Open FileName For Output As #FileNo

Print #FileNo, "[Revision]"
Print #FileNo, "INIRevision=7.5"
Print #FileNo, "[Company Information]"
Print #FileNo, "CompanyName=CleanFlow Systems"
Print #FileNo, "PhoneNo=+64 9 4799901"
Print #FileNo, "FaxNo="
Print #FileNo, "RecordMode=PAL"
Print #FileNo, "TuningStyle=Manual"
Print #FileNo, "CompanyLogoPath=C:\Programs\ClearLine Profiler\Logo.JPG"
Print #FileNo, "MeasurementUnits="
Print #FileNo, "CalibrationDistance=200"
Print #FileNo, "CalibrationLineLength=298.015265530938"
Print #FileNo, "IPX=0"
Print #FileNo, "IPY=0"
Print #FileNo, "IPGT=44"
Print #FileNo, "IPDX=10"
Print #FileNo, "IPDY=10"
Print #FileNo, "IgnoreX1="
Print #FileNo, "IgnoreY1="
Print #FileNo, "IgnoreX2="
Print #FileNo, "IgnoreY2="
Print #FileNo, "IgnoreDistX1="
Print #FileNo, "IgnoreDistY1="
Print #FileNo, "IgnoreDistX2="
Print #FileNo, "IgnoreDistY2="
Print #FileNo, "PVGraphYRatio=10"
Print #FileNo, "ProcessMethod=Type2"
Print #FileNo, "Contrast=39"
Print #FileNo, "Enhancement=High"
Print #FileNo, "LimitCapMax=10"
Print #FileNo, "LimitOval=6.0"
Print #FileNo, "LimitDeltaMin=-10.0"
Print #FileNo, "LimitDeltaMax=10.0"
Print #FileNo, "LimitYDiameter=-10.0"
Print #FileNo, "LimitXDiameter=10.0"
Print #FileNo, "LimitDiameterMaxL=172.8"
Print #FileNo, "LimitDiameterMaxR=211.2"
Print #FileNo, "LimitDiameterMinL=170"
Print #FileNo, "LimitDiameterMinR=223"
Print #FileNo, "LimitDiameterMedianL=172.8"
Print #FileNo, "LimitDiameterMedianR=211.2"
Print #FileNo, "PVGraphCapacityXScale=14"
Print #FileNo, "PVGraphCapacityXOffset=0"
Print #FileNo, "PVGraphOvalityXScale=15"
Print #FileNo, "PVGraphOvalityXOffset=-25"
Print #FileNo, "PVGraphDeltaXScale=14"
Print #FileNo, "PVGraphDeltaXOffset=8"
Print #FileNo, "PVGraphXYDiaXScale=14"
Print #FileNo, "PVGraphXYDiaXOffset=-4"
Print #FileNo, "PVGraphDiaMaxMinXScale=34"
Print #FileNo, "PVGraphDiaMaxMinXOffset=8"
Print #FileNo, "PVGraphDiaMedianXScale=46"
Print #FileNo, "PVGraphDiaMedianXOffset=0"
Print #FileNo, "PVGraphDiaMaxXScale=45"
Print #FileNo, "PVGraphDiaMaxXOffset=-5"
Print #FileNo, "PVGraphDiaMinXScale=45"
Print #FileNo, "pvGraphDiaMinXOffset=0"
Print #FileNo, "PVDiameterMethod=XY"
Print #FileNo, "HASPLock=false"
Print #FileNo, "PipeDetailsAssetNo="
Print #FileNo, "PipeDetailsSiteID="
Print #FileNo, "PipeDetailsCity="
Print #FileNo, "PipeDetailsDate="
Print #FileNo, "PipeDetailsTime="
Print #FileNo, "PipeDetailsStNode="
Print #FileNo, "PipeDetailsStLoc="
Print #FileNo, "PipeDetailsFhNode="
Print #FileNo, "PipeDetailsFhLoc="
Print #FileNo, "PipeDetailsIntDiaExp="
Print #FileNo, "PipeDetailsOutDiaExp="
Print #FileNo, "PipeDetailsLength="
Print #FileNo, "PipeDetailsMaterial="
Print #FileNo, "KeyX="
Print #FileNo, "KeyY="
Print #FileNo, "[GraphSubTittle]"
Print #FileNo, "Summary_Flat="
Print #FileNo, "Summary_MedianDiameter="
Print #FileNo, "Summary_Ovality="
Print #FileNo, "Summary_MaxDiameter="
Print #FileNo, "Summary_MinDiameter="
Print #FileNo, "Summary_XYDiameter="
Print #FileNo, "Summary_Capacity="
Print #FileNo, "Summary_Debris="
Print #FileNo, "Analysis_Flat="
Print #FileNo, "Analysis_MedianDiameter="
Print #FileNo, "Analysis_Ovality="
Print #FileNo, "Analysis_MaxDiameter="
Print #FileNo, "Analysis_MinDiameter="
Print #FileNo, "Analysis_XYDiameter="
Print #FileNo, "Analysis_Capacity="
Print #FileNo, "Analysis_Debris="
Print #FileNo, "Profile_Flat="
Print #FileNo, "Profile_MedianDiameter="
Print #FileNo, "Profile_Ovality="
Print #FileNo, "Profile_MaxDiameter="
Print #FileNo, "Profile_MinDiameter="
Print #FileNo, "Profile_XYDiameter="
Print #FileNo, "Profile_Capacity="
Print #FileNo, "Profile_Debris="
Print #FileNo, "Observations_Flat="
Print #FileNo, "Observations_MedianDiameter="
Print #FileNo, "Observations_Ovality="
Print #FileNo, "Observations_MaxDiameter="
Print #FileNo, "Observations_MinDiameter="
Print #FileNo, "Observations_XYDiameter="
Print #FileNo, "Observations_Capacity="
Print #FileNo, "Observations_Debris="
Print #FileNo, "[MTColors]"
Print #FileNo, "NormalDrawingColor=5435552"
Print #FileNo, "SelectedObjectColor=16776960"
Print #FileNo, "ModiCircleColor=16711935"
Print #FileNo, "ChosenModiCircleColor=255"
Print #FileNo, "AreaFillingColor=65280"
Print #FileNo, "ExtraObjectColor=16777215"
Print #FileNo, "JointCircleColor=16777215"
Print #FileNo, "TempDrawingColor=65535"
Print #FileNo, "MovingObjectColor=16711680"
Print #FileNo, "RotatingObjectColor=65535"
Print #FileNo, "SelectionBoundaryColor=16777215"
Print #FileNo, "TextSizeIndicatorColor=16711680"
Print #FileNo, "[Regional Options]"
Print #FileNo, "Language=English"
Print #FileNo, "ThreeDRenderingStyle=Software"
Print #FileNo, "PaperSize=A4"
Print #FileNo, "ReportMarginTop=0"
Print #FileNo, "ReportMarginBottom=0"
Print #FileNo, "ReportMarginLeft=500"
Print #FileNo, "ReportMarginRight=500"
Print #FileNo, "[Fish Eye Distortion]"
Print #FileNo, "Fish_DistortionHorizontal=1.00676417295284"
Print #FileNo, "Fish_Distortion=182"
Print #FileNo, "Fish_Ratio=137.799999999998"
Print #FileNo, "Fish_CenterX=0"
Print #FileNo, "Fish_CenterY=0"
Print #FileNo, "Fish_OriginalWidth=352"
Print #FileNo, "Fish_OriginalHeight=264"
Print #FileNo, "Fish_Displayed=False"
Print #FileNo, "FecFileName=Profiler Demo 1.fec"
Print #FileNo, "CameraModel=Profiler Demo 1"
Print #FileNo, "[Automatic Distance]"
Print #FileNo, "DistanceMethod=StartFinishEstimate"
Print #FileNo, "[Video Settings]"
Print #FileNo, "VideoCaptureDevice=0"
Print #FileNo, "[Demo]"
Print #FileNo, "DemoLoad=yes"
Print #FileNo, "DemoFile_mm=Examples\Profiler Demo 1.pvd"
Print #FileNo, "DemoFile_in=Examples\Profiler Demo 1 inches.pvd"
Print #FileNo, "\\"

Close #FileNo

Exit Sub
Err_Handler:
    MsgBox Err & "-ST37.5:" & Error$

End Sub

Sub CheckAndUpdateINI()
On Error GoTo Err_Handler

    '[Revision]
    Dim StringINIRevision As String:                Call GetINI_ParameterInfoOnly(MyFile, "INIRevision=", StringINIRevision)
    '[Company Information]
    Dim StringCompanyName As String:                Call GetINI_ParameterInfoOnly(MyFile, "CompanyName=", StringCompanyName)
    Dim StringPhoneNo As String:                    Call GetINI_ParameterInfoOnly(MyFile, "PhoneNo=", StringPhoneNo)
    Dim StringFaxNo As String:                       Call GetINI_ParameterInfoOnly(MyFile, "FaxNo=", StringFaxNo)
    Dim StringRecordMode As String:                 Call GetINI_ParameterInfoOnly(MyFile, "RecordMode=", StringRecordMode)
    Dim StringTuningStyle As String:                Call GetINI_ParameterInfoOnly(MyFile, "TuningStyle=", StringTuningStyle)
    Dim StringCompanyLogoPath As String:            Call GetINI_ParameterInfoOnly(MyFile, "CompanyLogoPath=", StringCompanyLogoPath)
    Dim StringMeasurementUnits As String:           Call GetINI_ParameterInfoOnly(MyFile, "MeasurementUnits=", StringMeasurementUnits)
    Dim StringCalibrationDistance As String:        Call GetINI_ParameterInfoOnly(MyFile, "CalibrationDistance=", StringCalibrationDistance)
    Dim StringCalibrationLineLength As String:      Call GetINI_ParameterInfoOnly(MyFile, "CalibrationLineLength=", StringCalibrationLineLength)
    Dim StringIPX As String:                        Call GetINI_ParameterInfoOnly(MyFile, "IPX=", StringIPX)
    Dim StringIPY As String:                        Call GetINI_ParameterInfoOnly(MyFile, "IPY=", StringIPY)
    Dim StringIPGT As String:                       Call GetINI_ParameterInfoOnly(MyFile, "IPGT=", StringIPGT)
    Dim StringIPDX As String:                       Call GetINI_ParameterInfoOnly(MyFile, "IPDX=", StringIPDX)
    Dim StringIPDY As String:                       Call GetINI_ParameterInfoOnly(MyFile, "IPDY=", StringIPDY)
    Dim StringIgnoreX1 As String:                   Call GetINI_ParameterInfoOnly(MyFile, "IgnoreX1=", StringIgnoreX1)
    Dim StringIgnoreY1 As String:                   Call GetINI_ParameterInfoOnly(MyFile, "IgnoreY1=", StringIgnoreY1)
    Dim StringIgnoreX2 As String:                   Call GetINI_ParameterInfoOnly(MyFile, "IgnoreX2=", StringIgnoreX2)
    Dim StringIgnoreY2 As String:                   Call GetINI_ParameterInfoOnly(MyFile, "IgnoreY2=", StringIgnoreY2)
    Dim StringIgnoreDistX1 As String:               Call GetINI_ParameterInfoOnly(MyFile, "IgnoreDistX1=", StringIgnoreDistX1)
    Dim StringIgnoreDistY1 As String:               Call GetINI_ParameterInfoOnly(MyFile, "IgnoreDistY1=", StringIgnoreDistY1)
    Dim StringIgnoreDistX2 As String:               Call GetINI_ParameterInfoOnly(MyFile, "IgnoreDistX2=", StringIgnoreDistX2)
    Dim StringIgnoreDistY2 As String:               Call GetINI_ParameterInfoOnly(MyFile, "IgnoreDistY2=", StringIgnoreDistY2)
    Dim StringPVGraphYRatio As String:              Call GetINI_ParameterInfoOnly(MyFile, "PVGraphYRatio=", StringPVGraphYRatio)
    Dim StringProcessMethod As String:              Call GetINI_ParameterInfoOnly(MyFile, "ProcessMethod=", StringProcessMethod)
    Dim StringContrast As String:                   Call GetINI_ParameterInfoOnly(MyFile, "Contrast=", StringContrast)
    Dim StringEnhancement As String:                Call GetINI_ParameterInfoOnly(MyFile, "Enhancement=", StringEnhancement)
    Dim StringLimitCapMax As String:                Call GetINI_ParameterInfoOnly(MyFile, "LimitCapMax=", StringLimitCapMax)
    Dim StringLimitOval As String:                  Call GetINI_ParameterInfoOnly(MyFile, "LimitOval=", StringLimitOval)
    Dim StringLimitDeltaMin As String:              Call GetINI_ParameterInfoOnly(MyFile, "LimitDeltaMin=", StringLimitDeltaMin)
    Dim StringLimitDeltaMax As String:              Call GetINI_ParameterInfoOnly(MyFile, "LimitDeltaMax=", StringLimitDeltaMax)
    Dim StringLimitYDiameter As String:             Call GetINI_ParameterInfoOnly(MyFile, "LimitYDiameter=", StringLimitYDiameter)
    Dim StringLimitXDiameter  As String:            Call GetINI_ParameterInfoOnly(MyFile, "LimitXDiameter=", StringLimitXDiameter)
    Dim StringLimitDiameterMaxL As String:          Call GetINI_ParameterInfoOnly(MyFile, "LimitDiameterMaxL=", StringLimitDiameterMaxL)
    Dim StringLimitDiameterMaxR As String:          Call GetINI_ParameterInfoOnly(MyFile, "LimitDiameterMaxR=", StringLimitDiameterMaxR)
    Dim StringLimitDiameterMinL As String:          Call GetINI_ParameterInfoOnly(MyFile, "LimitDiameterMinL=", StringLimitDiameterMinL)
    Dim StringLimitDiameterMinR As String:          Call GetINI_ParameterInfoOnly(MyFile, "LimitDiameterMinR=", StringLimitDiameterMinR)
    Dim StringLimitDiameterMedianL As String:       Call GetINI_ParameterInfoOnly(MyFile, "LimitDiameterMedianL=", StringLimitDiameterMedianL)
    Dim StringLimitDiameterMedianR As String:       Call GetINI_ParameterInfoOnly(MyFile, "LimitDiameterMedianR=", StringLimitDiameterMedianR)
    Dim StringPVGraphCapacityXScale As String:      Call GetINI_ParameterInfoOnly(MyFile, "PVGraphCapacityXScale=", StringPVGraphCapacityXScale)
    Dim StringPVGraphCapacityXOffset As String:     Call GetINI_ParameterInfoOnly(MyFile, "PVGraphCapacityXOffset=", StringPVGraphCapacityXOffset)
    Dim StringPVGraphOvalityXScale As String:       Call GetINI_ParameterInfoOnly(MyFile, "PVGraphOvalityXScale=", StringPVGraphOvalityXScale)
    Dim StringPVGraphOvalityXOffset As String:      Call GetINI_ParameterInfoOnly(MyFile, "PVGraphOvalityXOffset=", StringPVGraphOvalityXOffset)
    Dim StringPVGraphDeltaXScale As String:         Call GetINI_ParameterInfoOnly(MyFile, "PVGraphDeltaXScale=", StringPVGraphDeltaXScale)
    Dim StringPVGraphDeltaXOffset As String:        Call GetINI_ParameterInfoOnly(MyFile, "PVGraphDeltaXOffset=", StringPVGraphDeltaXOffset)
    Dim StringPVGraphXYDiaXScale As String:         Call GetINI_ParameterInfoOnly(MyFile, "PVGraphXYDiaXScale=", StringPVGraphXYDiaXScale)
    Dim StringPVGraphXYDiaXOffset As String:        Call GetINI_ParameterInfoOnly(MyFile, "PVGraphXYDiaXOffset=", StringPVGraphXYDiaXOffset)
    Dim StringPVGraphDiaMaxMinXScale As String:     Call GetINI_ParameterInfoOnly(MyFile, "PVGraphDiaMaxMinXScale=", StringPVGraphDiaMaxMinXScale)
    Dim StringPVGraphDiaMaxMinXOffset As String:    Call GetINI_ParameterInfoOnly(MyFile, "PVGraphDiaMaxMinXOffset=", StringPVGraphDiaMaxMinXOffset)
    Dim StringPVGraphDiaMedianXScale As String:     Call GetINI_ParameterInfoOnly(MyFile, "PVGraphDiaMedianXScale=", StringPVGraphDiaMedianXScale)
    Dim StringPVGraphDiaMedianXOffset As String:    Call GetINI_ParameterInfoOnly(MyFile, "PVGraphDiaMedianXOffset=", StringPVGraphDiaMedianXOffset)
    Dim StringPVGraphDiaMaxXScale As String:        Call GetINI_ParameterInfoOnly(MyFile, "PVGraphDiaMaxXScale=", StringPVGraphDiaMaxXScale)
    Dim StringPVGraphDiaMaxXOffset As String:       Call GetINI_ParameterInfoOnly(MyFile, "PVGraphDiaMaxXOffset=", StringPVGraphDiaMaxXOffset)
    Dim StringPVGraphDiaMinXScale As String:        Call GetINI_ParameterInfoOnly(MyFile, "PVGraphDiaMinXScale=", StringPVGraphDiaMinXScale)
    Dim StringpvGraphDiaMinXOffset As String:       Call GetINI_ParameterInfoOnly(MyFile, "pvGraphDiaMinXOffset=", StringpvGraphDiaMinXOffset)
    Dim StringPVGraphInclinationXScale As String:   Call GetINI_ParameterInfoOnly(MyFile, "PVGraphInclinationXScale=", StringPVGraphInclinationXScale) 'PCN6128
    Dim StringPVGraphInclinationXOffset As String:  Call GetINI_ParameterInfoOnly(MyFile, "PVGraphInclinationXOffset=", StringPVGraphInclinationXOffset) 'PCN6128
   
    Dim StringPVDiameterMethod As String:           Call GetINI_ParameterInfoOnly(MyFile, "PVDiameterMethod=", StringPVDiameterMethod)
    Dim StringHASPLock As String:                   Call GetINI_ParameterInfoOnly(MyFile, "HASPLock=", StringHASPLock)
    Dim StringPipeDetailsAssetNo As String:         Call GetINI_ParameterInfoOnly(MyFile, "PipeDetailsAssetNo=", StringPipeDetailsAssetNo)
    Dim StringPipeDetailsSiteID As String:          Call GetINI_ParameterInfoOnly(MyFile, "PipeDetailsSiteID=", StringPipeDetailsSiteID)
    Dim StringPipeDetailsCity As String:            Call GetINI_ParameterInfoOnly(MyFile, "PipeDetailsCity=", StringPipeDetailsCity)
    Dim StringPipeDetailsDate As String:            Call GetINI_ParameterInfoOnly(MyFile, "PipeDetailsDate=", StringPipeDetailsDate)
    Dim StringPipeDetailsTime As String:            Call GetINI_ParameterInfoOnly(MyFile, "PipeDetailsTime=", StringPipeDetailsTime)
    Dim StringPipeDetailsStNode As String:          Call GetINI_ParameterInfoOnly(MyFile, "PipeDetailsStNode=", StringPipeDetailsStNode)
    Dim StringPipeDetailsStLoc As String:           Call GetINI_ParameterInfoOnly(MyFile, "PipeDetailsStLoc=", StringPipeDetailsStLoc)
    Dim StringPipeDetailsFhNode As String:          Call GetINI_ParameterInfoOnly(MyFile, "PipeDetailsFhNode=", StringPipeDetailsFhNode)
    Dim StringPipeDetailsFhLoc As String:           Call GetINI_ParameterInfoOnly(MyFile, "PipeDetailsFhLoc=", StringPipeDetailsFhLoc)
    Dim StringPipeDetailsIntDiaExp As String:       Call GetINI_ParameterInfoOnly(MyFile, "PipeDetailsIntDiaExp=", StringPipeDetailsIntDiaExp)
    Dim StringPipeDetailsOutDiaExp As String:       Call GetINI_ParameterInfoOnly(MyFile, "PipeDetailsOutDiaExp=", StringPipeDetailsOutDiaExp)
    Dim StringPipeDetailsLength As String:          Call GetINI_ParameterInfoOnly(MyFile, "PipeDetailsLength=", StringPipeDetailsLength)
    Dim StringPipeDetailsMaterial As String:        Call GetINI_ParameterInfoOnly(MyFile, "PipeDetailsMaterial=", StringPipeDetailsMaterial)
    Dim StringKeyX As String:                       Call GetINI_ParameterInfoOnly(MyFile, "KeyX=", StringKeyX)
    Dim StringKeyY As String:                       Call GetINI_ParameterInfoOnly(MyFile, "KeyY=", StringKeyY)
    Dim StringFlatType As String:                   Call GetINI_ParameterInfoOnly(MyFile, "FlatType=", StringFlatType)
    '[GraphSubTittle]
    Dim StringSummary_Flat As String:               Call GetINI_ParameterInfoOnly(MyFile, "Summary_Flat=", StringSummary_Flat)
    Dim StringSummary_MedianDiameter As String:     Call GetINI_ParameterInfoOnly(MyFile, "Summary_MedianDiameter=", StringSummary_MedianDiameter)
    Dim StringSummary_Ovality As String:            Call GetINI_ParameterInfoOnly(MyFile, "Summary_Ovality=", StringSummary_Ovality)
    Dim StringSummary_MaxDiameter As String:        Call GetINI_ParameterInfoOnly(MyFile, "Summary_MaxDiameter=", StringSummary_MaxDiameter)
    Dim StringSummary_MinDiameter As String:        Call GetINI_ParameterInfoOnly(MyFile, "Summary_MinDiameter=", StringSummary_MinDiameter)
    Dim StringSummary_XYDiameter As String:         Call GetINI_ParameterInfoOnly(MyFile, "Summary_XYDiameter=", StringSummary_XYDiameter)
    Dim StringSummary_Capacity As String:           Call GetINI_ParameterInfoOnly(MyFile, "Summary_Capacity=", StringSummary_Capacity)
    Dim StringSummary_Debris As String:             Call GetINI_ParameterInfoOnly(MyFile, "Summary_Debris=", StringSummary_Debris)
    Dim StringAnalysis_Flat As String:              Call GetINI_ParameterInfoOnly(MyFile, "Analysis_Flat=", StringAnalysis_Flat)
    Dim StringAnalysis_MedianDiameter As String:    Call GetINI_ParameterInfoOnly(MyFile, "Analysis_MedianDiameter=", StringAnalysis_MedianDiameter)
    Dim StringAnalysis_Ovality As String:           Call GetINI_ParameterInfoOnly(MyFile, "Analysis_Ovality=", StringAnalysis_Ovality)
    Dim StringAnalysis_MaxDiameter As String:       Call GetINI_ParameterInfoOnly(MyFile, "Analysis_MaxDiameter=", StringAnalysis_MaxDiameter)
    Dim StringAnalysis_MinDiameter As String:       Call GetINI_ParameterInfoOnly(MyFile, "Analysis_MinDiameter=", StringAnalysis_MinDiameter)
    Dim StringAnalysis_XYDiameter As String:        Call GetINI_ParameterInfoOnly(MyFile, "Analysis_XYDiameter=", StringAnalysis_XYDiameter)
    Dim StringAnalysis_Capacity As String:          Call GetINI_ParameterInfoOnly(MyFile, "Analysis_Capacity=", StringAnalysis_Capacity)
    Dim StringAnalysis_Debris As String:            Call GetINI_ParameterInfoOnly(MyFile, "Analysis_Debris=", StringAnalysis_Debris)
    Dim StringProfile_Flat As String:               Call GetINI_ParameterInfoOnly(MyFile, "Profile_Flat=", StringProfile_Flat)
    Dim StringProfile_MedianDiameter As String:     Call GetINI_ParameterInfoOnly(MyFile, "Profile_MedianDiameter=", StringProfile_MedianDiameter)
    Dim StringProfile_Ovality As String:            Call GetINI_ParameterInfoOnly(MyFile, "Profile_Ovality=", StringProfile_Ovality)
    Dim StringProfile_MaxDiameter As String:        Call GetINI_ParameterInfoOnly(MyFile, "Profile_MaxDiameter=", StringProfile_MaxDiameter)
    Dim StringProfile_MinDiameter As String:         Call GetINI_ParameterInfoOnly(MyFile, "Profile_MinDiameter=", StringProfile_MinDiameter)
    Dim StringProfile_XYDiameter As String:         Call GetINI_ParameterInfoOnly(MyFile, "Profile_XYDiameter=", StringProfile_XYDiameter)
    Dim StringProfile_Capacity As String:           Call GetINI_ParameterInfoOnly(MyFile, "Profile_Capacity=", StringProfile_Capacity)
    Dim StringProfile_Debris As String:             Call GetINI_ParameterInfoOnly(MyFile, "Profile_Debris=", StringProfile_Debris)
    Dim StringObservations_Flat As String:          Call GetINI_ParameterInfoOnly(MyFile, "Observations_Flat=", StringObservations_Flat)
    Dim StringObservations_MedianDiameter As String: Call GetINI_ParameterInfoOnly(MyFile, "Observations_MedianDiameter=", StringObservations_MedianDiameter)
    Dim StringObservations_Ovality As String:       Call GetINI_ParameterInfoOnly(MyFile, "Observations_Ovality=", StringObservations_Ovality)
    Dim StringObservations_MaxDiameter As String:   Call GetINI_ParameterInfoOnly(MyFile, "Observations_MaxDiameter=", StringObservations_MaxDiameter)
    Dim StringObservations_MinDiameter As String:   Call GetINI_ParameterInfoOnly(MyFile, "Observations_MinDiameter=", StringObservations_MinDiameter)
    Dim StringObservations_XYDiameter As String:    Call GetINI_ParameterInfoOnly(MyFile, "Observations_XYDiameter=", StringObservations_XYDiameter)
    Dim StringObservations_Capacity As String:      Call GetINI_ParameterInfoOnly(MyFile, "Observations_Capacity=", StringObservations_Capacity)
    Dim StringObservations_Debris As String:        Call GetINI_ParameterInfoOnly(MyFile, "Observations_Debris=", StringObservations_Debris)
    Dim StringFlat1KTitle As String:                Call GetINI_ParameterInfoOnly(MyFile, "Flat1KTitle=", StringFlat1KTitle)
    Dim StringFlatOvality1KTitle As String:         Call GetINI_ParameterInfoOnly(MyFile, "FlatOvality1KTitle=", StringFlatOvality1KTitle)
    '[MTColors]
    Dim StringNormalDrawingColor As String:         Call GetINI_ParameterInfoOnly(MyFile, "NormalDrawingColor=", StringNormalDrawingColor)
    Dim StringSelectedObjectColor As String:        Call GetINI_ParameterInfoOnly(MyFile, "SelectedObjectColor=", StringSelectedObjectColor)
    Dim StringModiCircleColor As String:            Call GetINI_ParameterInfoOnly(MyFile, "ModiCircleColor=", StringModiCircleColor)
    Dim StringChosenModiCircleColor As String:      Call GetINI_ParameterInfoOnly(MyFile, "ChosenModiCircleColor=", StringChosenModiCircleColor)
    Dim StringAreaFillingColor As String:           Call GetINI_ParameterInfoOnly(MyFile, "AreaFillingColor=", StringAreaFillingColor)
    Dim StringExtraObjectColor As String:           Call GetINI_ParameterInfoOnly(MyFile, "ExtraObjectColor=", StringExtraObjectColor)
    Dim StringJointCircleColor As String:           Call GetINI_ParameterInfoOnly(MyFile, "JointCircleColor=", StringJointCircleColor)
    Dim StringTempDrawingColor As String:           Call GetINI_ParameterInfoOnly(MyFile, "TempDrawingColor=", StringTempDrawingColor)
    Dim StringMovingObjectColor As String:          Call GetINI_ParameterInfoOnly(MyFile, "MovingObjectColor=", StringMovingObjectColor)
    Dim StringRotatingObjectColor As String:        Call GetINI_ParameterInfoOnly(MyFile, "RotatingObjectColor=", StringRotatingObjectColor)
    Dim StringSelectionBoundaryColor As String:     Call GetINI_ParameterInfoOnly(MyFile, "SelectionBoundaryColor=", StringSelectionBoundaryColor)
    Dim StringTextSizeIndicatorColor As String:     Call GetINI_ParameterInfoOnly(MyFile, "TextSizeIndicatorColor=", StringTextSizeIndicatorColor)
    '[Regional Options]
    Dim StringLanguage As String:                   Call GetINI_ParameterInfoOnly(MyFile, "Language=", StringLanguage)
    Dim StringThreeDRenderingStyle As String:       Call GetINI_ParameterInfoOnly(MyFile, "ThreeDRenderingStyle=", StringThreeDRenderingStyle)
    Dim StringPaperSize As String:                  Call GetINI_ParameterInfoOnly(MyFile, "PaperSize=", StringPaperSize)
    Dim StringReportMarginTop As String:            Call GetINI_ParameterInfoOnly(MyFile, "ReportMarginTop=", StringReportMarginTop)
    Dim StringReportMarginBottom As String:         Call GetINI_ParameterInfoOnly(MyFile, "ReportMarginBottom=", StringReportMarginBottom)
    Dim StringReportMarginLeft As String:           Call GetINI_ParameterInfoOnly(MyFile, "ReportMarginLeft=", StringReportMarginLeft)
    Dim StringReportMarginRight As String:          Call GetINI_ParameterInfoOnly(MyFile, "ReportMarginRight=", StringReportMarginRight)
    '[Fish Eye Distortion]
    Dim StringFish_DistortionHorizontal As String:  Call GetINI_ParameterInfoOnly(MyFile, "Fish_DistortionHorizontal=", StringFish_DistortionHorizontal)
    Dim StringFish_Distortion As String:            Call GetINI_ParameterInfoOnly(MyFile, "Fish_Distortion=", StringFish_Distortion)
    Dim StringFish_Ratio As String:                 Call GetINI_ParameterInfoOnly(MyFile, "Fish_Ratio=", StringFish_Ratio)
    Dim StringFish_CenterX As String:               Call GetINI_ParameterInfoOnly(MyFile, "Fish_CenterX=", StringFish_CenterX)
    Dim StringFish_CenterY As String:               Call GetINI_ParameterInfoOnly(MyFile, "Fish_CenterY=", StringFish_CenterY)
    Dim StringFish_OriginalWidth As String:         Call GetINI_ParameterInfoOnly(MyFile, "Fish_OriginalWidth=", StringFish_OriginalWidth)
    Dim StringFish_OriginalHeight As String:        Call GetINI_ParameterInfoOnly(MyFile, "Fish_OriginalHeight=", StringFish_OriginalHeight)
    Dim StringFish_Displayed As String:             Call GetINI_ParameterInfoOnly(MyFile, "Fish_Displayed=", StringFish_Displayed)
    Dim StringFecFileName As String:                Call GetINI_ParameterInfoOnly(MyFile, "FecFileName=", StringFecFileName)
    Dim StringCameraModel As String:                Call GetINI_ParameterInfoOnly(MyFile, "CameraModel=", StringCameraModel)
    '[Automatic Distance]
    Dim StringDistanceMethod As String:             Call GetINI_ParameterInfoOnly(MyFile, "DistanceMethod=", StringDistanceMethod)
    '[Video Settings]
    Dim StringVideoCaptureDevice As String:         Call GetINI_ParameterInfoOnly(MyFile, "VideoCaptureDevice=", StringVideoCaptureDevice)
    '[Demo]
    Dim StringDemoLoad As String:                   Call GetINI_ParameterInfoOnly(MyFile, "DemoLoad=", StringDemoLoad)
    Dim StringDemoFile_mm As String:                Call GetINI_ParameterInfoOnly(MyFile, "DemoFile_mm=", StringDemoFile_mm)
    Dim StringDemoFile_in As String:                Call GetINI_ParameterInfoOnly(MyFile, "DemoFile_in=", StringDemoFile_in)

    
    Dim FileNo
    FileNo = FreeFile
    
    Open MyFile For Output As #FileNo

    Print #FileNo, "[Revision]"
    Print #FileNo, "INIRevision=" & "7.9" 'PCN6025 'Added DeflectionOrNormal entry PCN6128 added inclination
    Print #FileNo, "[Company Information]"
    Print #FileNo, "CompanyName=" & StringCompanyName
    Print #FileNo, "PhoneNo=" & StringPhoneNo
    Print #FileNo, "FaxNo=" & StringFaxNo
    Print #FileNo, "RecordMode=" & StringRecordMode
    Print #FileNo, "TuningStyle=" & StringTuningStyle
    Print #FileNo, "CompanyLogoPath=" & StringCompanyLogoPath
    Print #FileNo, "MeasurementUnits=" & StringMeasurementUnits
    Print #FileNo, "CalibrationDistance=" & StringCalibrationDistance
    Print #FileNo, "CalibrationLineLength=" & StringCalibrationLineLength
    Print #FileNo, "IPX=" & StringIPX
    Print #FileNo, "IPY=" & StringIPY
    Print #FileNo, "IPGT=" & StringIPGT
    Print #FileNo, "IPDX=" & StringIPDX
    Print #FileNo, "IPDY=" & StringIPDY
    Print #FileNo, "IgnoreX1=" & StringIgnoreX1
    Print #FileNo, "IgnoreY1=" & StringIgnoreY1
    Print #FileNo, "IgnoreX2=" & StringIgnoreX2
    Print #FileNo, "IgnoreY2=" & StringIgnoreY2
    Print #FileNo, "IgnoreDistX1=" & StringIgnoreDistX1
    Print #FileNo, "IgnoreDistY1=" & StringIgnoreDistY1
    Print #FileNo, "IgnoreDistX2=" & StringIgnoreDistX2
    Print #FileNo, "IgnoreDistY2=" & StringIgnoreDistY2
    Print #FileNo, "PVGraphYRatio=" & StringPVGraphYRatio
    Print #FileNo, "ProcessMethod=" & StringProcessMethod
    Print #FileNo, "Contrast=" & StringContrast
    Print #FileNo, "Enhancement=" & StringEnhancement
    Print #FileNo, "LimitCapMax=" & StringLimitCapMax
    Print #FileNo, "LimitOval=" & StringLimitOval
    Print #FileNo, "LimitDeltaMin=" & StringLimitDeltaMin
    Print #FileNo, "LimitDeltaMax=" & StringLimitDeltaMax
    Print #FileNo, "LimitYDiameter=" & StringLimitYDiameter
    Print #FileNo, "LimitXDiameter=" & StringLimitXDiameter
    Print #FileNo, "LimitDiameterMaxL=" & StringLimitDiameterMaxL
    Print #FileNo, "LimitDiameterMaxR=" & StringLimitDiameterMaxR
    Print #FileNo, "LimitDiameterMinL=" & StringLimitDiameterMinL
    Print #FileNo, "LimitDiameterMinR=" & StringLimitDiameterMinR
    Print #FileNo, "LimitDiameterMedianL=" & StringLimitDiameterMedianL
    Print #FileNo, "LimitDiameterMedianR=" & StringLimitDiameterMedianR
    Print #FileNo, "PVGraphCapacityXScale=" & StringPVGraphCapacityXScale
    Print #FileNo, "PVGraphCapacityXOffset=" & StringPVGraphCapacityXOffset
    Print #FileNo, "PVGraphOvalityXScale=" & StringPVGraphOvalityXScale
    Print #FileNo, "PVGraphOvalityXOffset=" & StringPVGraphOvalityXOffset
    Print #FileNo, "PVGraphDeltaXScale=" & StringPVGraphDeltaXScale
    Print #FileNo, "PVGraphDeltaXOffset=" & StringPVGraphDeltaXOffset
    Print #FileNo, "PVGraphXYDiaXScale=" & StringPVGraphXYDiaXScale
    Print #FileNo, "PVGraphXYDiaXOffset=" & StringPVGraphXYDiaXOffset
    Print #FileNo, "PVGraphDiaMaxMinXScale=" & StringPVGraphDiaMaxMinXScale
    Print #FileNo, "PVGraphDiaMaxMinXOffset=" & StringPVGraphDiaMaxMinXOffset
    Print #FileNo, "PVGraphDiaMedianXScale=" & StringPVGraphDiaMedianXScale
    Print #FileNo, "PVGraphDiaMedianXOffset=" & StringPVGraphDiaMedianXOffset
    Print #FileNo, "PVGraphDiaMaxXScale=" & StringPVGraphDiaMaxXScale
    Print #FileNo, "PVGraphDiaMaxXOffset=" & StringPVGraphDiaMaxXOffset
    Print #FileNo, "PVGraphDiaMinXScale=" & StringPVGraphDiaMinXScale
    Print #FileNo, "pvGraphDiaMinXOffset=" & StringpvGraphDiaMinXOffset
    Print #FileNo, "PVGraphInclinationXScale=" & StringPVGraphInclinationXScale 'PCN6128
    Print #FileNo, "PVGraphInclinationXOffset=" & StringPVGraphInclinationXOffset 'PCN6128
    Print #FileNo, "PVDiameterMethod=" & StringPVDiameterMethod
    Print #FileNo, "HASPLock=" & StringHASPLock
    Print #FileNo, "PipeDetailsAssetNo=" & StringPipeDetailsAssetNo
    Print #FileNo, "PipeDetailsSiteID=" & StringPipeDetailsSiteID
    Print #FileNo, "PipeDetailsCity=" & StringPipeDetailsCity
    Print #FileNo, "PipeDetailsDate=" & StringPipeDetailsDate
    Print #FileNo, "PipeDetailsTime=" & StringPipeDetailsTime
    Print #FileNo, "PipeDetailsStNode=" & StringPipeDetailsStNode
    Print #FileNo, "PipeDetailsStLoc=" & StringPipeDetailsStLoc
    Print #FileNo, "PipeDetailsFhNode=" & StringPipeDetailsFhNode
    Print #FileNo, "PipeDetailsFhLoc=" & StringPipeDetailsFhLoc
    Print #FileNo, "PipeDetailsIntDiaExp=" & StringPipeDetailsIntDiaExp
    Print #FileNo, "PipeDetailsOutDiaExp=" & StringPipeDetailsOutDiaExp
    Print #FileNo, "PipeDetailsLength=" & StringPipeDetailsLength
    Print #FileNo, "PipeDetailsMaterial=" & StringPipeDetailsMaterial
    Print #FileNo, "KeyX=" & StringKeyX
    Print #FileNo, "KeyY=" & StringKeyY
    Print #FileNo, "FlatType=" & StringFlatType
    Print #FileNo, "[GraphSubTittle]"
    Print #FileNo, "Summary_Flat=" & StringSummary_Flat
    Print #FileNo, "Summary_MedianDiameter=" & StringSummary_MedianDiameter
    Print #FileNo, "Summary_Ovality=" & StringSummary_Ovality
    Print #FileNo, "Summary_MaxDiameter=" & StringSummary_MaxDiameter
    Print #FileNo, "Summary_MinDiameter=" & StringSummary_MinDiameter
    Print #FileNo, "Summary_XYDiameter=" & StringSummary_XYDiameter
    Print #FileNo, "Summary_Capacity=" & StringSummary_Capacity
    Print #FileNo, "Summary_Debris=" & StringSummary_Debris
    Print #FileNo, "Analysis_Flat=" & StringAnalysis_Flat
    Print #FileNo, "Analysis_MedianDiameter=" & StringAnalysis_MedianDiameter
    Print #FileNo, "Analysis_Ovality=" & StringAnalysis_Ovality
    Print #FileNo, "Analysis_MaxDiameter=" & StringAnalysis_MaxDiameter
    Print #FileNo, "Analysis_MinDiameter=" & StringAnalysis_MinDiameter
    Print #FileNo, "Analysis_XYDiameter=" & StringAnalysis_XYDiameter
    Print #FileNo, "Analysis_Capacity=" & StringAnalysis_Capacity
    Print #FileNo, "Analysis_Debris=" & StringAnalysis_Debris
    Print #FileNo, "Profile_Flat=" & StringProfile_Flat
    Print #FileNo, "Profile_MedianDiameter=" & StringProfile_MedianDiameter
    Print #FileNo, "Profile_Ovality=" & StringProfile_Ovality
    Print #FileNo, "Profile_MaxDiameter=" & StringProfile_MaxDiameter
    Print #FileNo, "Profile_MinDiameter=" & StringProfile_MinDiameter
    Print #FileNo, "Profile_XYDiameter=" & StringProfile_XYDiameter
    Print #FileNo, "Profile_Capacity=" & StringProfile_Capacity
    Print #FileNo, "Profile_Debris=" & StringProfile_Debris
    Print #FileNo, "Observations_Flat=" & StringObservations_Flat
    Print #FileNo, "Observations_MedianDiameter=" & StringObservations_MedianDiameter
    Print #FileNo, "Observations_Ovality=" & StringObservations_Ovality
    Print #FileNo, "Observations_MaxDiameter=" & StringObservations_MaxDiameter
    Print #FileNo, "Observations_MinDiameter=" & StringObservations_MinDiameter
    Print #FileNo, "Observations_XYDiameter=" & StringObservations_XYDiameter
    Print #FileNo, "Observations_Capacity=" & StringObservations_Capacity
    Print #FileNo, "Observations_Debris=" & StringObservations_Debris
    Print #FileNo, "Flat1KTitle=" & StringFlat1KTitle
    Print #FileNo, "FlatOvality1KTitle=" & StringFlatOvality1KTitle
    Print #FileNo, "[MTColors]"
    Print #FileNo, "NormalDrawingColor=" & StringNormalDrawingColor
    Print #FileNo, "SelectedObjectColor=" & StringSelectedObjectColor
    Print #FileNo, "ModiCircleColor=" & StringModiCircleColor
    Print #FileNo, "ChosenModiCircleColor=" & StringChosenModiCircleColor
    Print #FileNo, "AreaFillingColor=" & StringAreaFillingColor
    Print #FileNo, "ExtraObjectColor=" & StringExtraObjectColor
    Print #FileNo, "JointCircleColor=" & StringJointCircleColor
    Print #FileNo, "TempDrawingColor=" & StringTempDrawingColor
    Print #FileNo, "MovingObjectColor=" & StringMovingObjectColor
    Print #FileNo, "RotatingObjectColor=" & StringRotatingObjectColor
    Print #FileNo, "SelectionBoundaryColor=" & StringSelectionBoundaryColor
    Print #FileNo, "TextSizeIndicatorColor=" & StringTextSizeIndicatorColor
    Print #FileNo, "[Regional Options]"
    Print #FileNo, "Language=" & StringLanguage
    Print #FileNo, "ThreeDRenderingStyle=" & StringThreeDRenderingStyle
    Print #FileNo, "PaperSize=" & StringPaperSize
    Print #FileNo, "ReportMarginTop=" & StringReportMarginTop
    Print #FileNo, "ReportMarginBottom=" & StringReportMarginBottom
    Print #FileNo, "ReportMarginLeft=" & StringReportMarginLeft
    Print #FileNo, "ReportMarginRight=" & StringReportMarginRight
    Print #FileNo, "[Fish Eye Distortion]"
    Print #FileNo, "Fish_DistortionHorizontal=" & StringFish_DistortionHorizontal
    Print #FileNo, "Fish_Distortion=" & StringFish_Distortion
    Print #FileNo, "Fish_Ratio=" & StringFish_Ratio
    Print #FileNo, "Fish_CenterX=" & StringFish_CenterX
    Print #FileNo, "Fish_CenterY=" & StringFish_CenterY
    Print #FileNo, "Fish_OriginalWidth=" & StringFish_OriginalWidth
    Print #FileNo, "Fish_OriginalHeight=" & StringFish_OriginalHeight
    Print #FileNo, "Fish_Displayed=" & StringFish_Displayed
    Print #FileNo, "FecFileName=" & StringFecFileName
    Print #FileNo, "CameraModel=" & StringCameraModel
    Print #FileNo, "[Automatic Distance]"
    Print #FileNo, "DistanceMethod=" & StringDistanceMethod
    Print #FileNo, "[Video Settings]"
    Print #FileNo, "VideoCaptureDevice=" & StringVideoCaptureDevice
    Print #FileNo, "[Demo]"
    Print #FileNo, "DemoLoad=" & StringDemoLoad
    Print #FileNo, "DemoFile_mm=" & StringDemoFile_mm
    Print #FileNo, "DemoFile_in=" & StringDemoFile_in
    Print #FileNo, "\\"
    
    Close #FileNo
    
    Exit Sub
Err_Handler:
    MsgBox Err & "-ST37:" & Error$

End Sub

Sub LoadShapes()
On Error GoTo Err_Handler

Dim ShapeDirectory As String

ShapeDirectory = App.Path & "\Shape Files\"

Dim FileName As String
Dim SplitPath As String
Dim SplitName As String
Dim SplitExt As String

Dim i As Integer


FileName = Dir(ShapeDirectory & "*.shp")
While FileName <> ""
    Call SplitFilePath(FileName, SplitPath, SplitName, SplitExt)
    If LCase(SplitExt) = "shp" Then
        Call LoadShapeFile(ShapeDirectory & FileName)
    End If
    FileName = Dir
Wend


Exit Sub
Err_Handler:
    MsgBox Err & "-ST38:" & Error$
End Sub

Sub LoadShapeFile(ByVal FileName As String)
On Error GoTo Err_Handler

    
    Dim NoReferenceShapes As Integer
    
    Dim VersionNo As Single
    Dim user As String
    Dim NoOfArcs As Integer
    Dim NoOfLines As Integer

    Dim FileNo
    Dim InputString As String
    Dim ShapeScale As Double
    
    

    
    Dim LineNo As Integer
    Dim i As Integer
    Dim TabPos As Integer
    
    NoReferenceShapes = UBound(ReferenceShape)
    NoReferenceShapes = NoReferenceShapes + 1
NoREferenceShapesLoaded:
    ReDim Preserve ReferenceShape(NoReferenceShapes)
    
    FileNo = FreeFile
    
    Open FileName For Input As #FileNo
        
    Line Input #1, InputString: VersionNo = CSng(GetTabData(InputString, 1))
    Line Input #1, InputString: ReferenceShape(NoReferenceShapes).Name = GetTabData(InputString, 1)
    Line Input #1, InputString: ReferenceShape(NoReferenceShapes).Use = GetTabData(InputString, 1)
    Line Input #1, InputString: ReferenceShape(NoReferenceShapes).CentreOffsetX = CSng(GetTabData(InputString, 1))
    Line Input #1, InputString: ReferenceShape(NoReferenceShapes).CentreOffsetY = CSng(GetTabData(InputString, 1))
    Line Input #1, InputString: NoOfArcs = CSng(GetTabData(InputString, 1)): ReferenceShape(NoReferenceShapes).NoArcs = NoOfArcs
    Line Input #1, InputString: NoOfLines = CSng(GetTabData(InputString, 1)): ReferenceShape(NoReferenceShapes).NoLines = NoOfLines
    Line Input #1, InputString: ShapeScale = CDbl(GetTabData(InputString, 1))
    Line Input #1, InputString:
    
    If (NoOfArcs > 0) Then
        While InStr(1, InputString, "[Arcs]") = 0
            Line Input #1, InputString
        Wend
        Line Input #1, InputString
        For i = 1 To NoOfArcs
            Line Input #1, InputString
            With ReferenceShape(NoReferenceShapes).Arcs(i - 1)
                .Colour = vbGreen
                .OriginX = CSng(GetTabData(InputString, 1)) * ShapeScale
                .OriginY = CSng(GetTabData(InputString, 2)) * ShapeScale
                .Radius = CSng(GetTabData(InputString, 3)) * ShapeScale
                .StartAngle = CSng(GetTabData(InputString, 4))
                .EndAngle = CSng(GetTabData(InputString, 5))
            End With
        Next i
    End If
    
    If (NoOfLines > 0) Then
        While InStr(1, InputString, "[Lines]") = 0
            Line Input #1, InputString
        Wend
        Line Input #1, InputString
        For i = 1 To NoOfLines
        
            Line Input #1, InputString
            With ReferenceShape(NoReferenceShapes).Lines(i - 1)
                .Colour = vbGreen
                .StartX = CSng(GetTabData(InputString, 1)) * ShapeScale
                .StartY = CSng(GetTabData(InputString, 2)) * ShapeScale
                .EndX = CSng(GetTabData(InputString, 3)) * ShapeScale
                .EndY = CSng(GetTabData(InputString, 4)) * ShapeScale
            End With
        Next i
    End If
    Close #FileNo

Exit Sub
Err_Handler:
    If Err = 9 Then NoReferenceShapes = 0: GoTo NoREferenceShapesLoaded
    If Err = 13 Then Resume Next 'Dodgy data move onto next
    MsgBox Err & "-ST39:" & Error$
    
End Sub

Function GetTabData(ByVal StringLine As String, ByVal TabNo As Integer) As String
On Error GoTo Err_Handler

    Dim i As Integer
    Dim CountTab As Integer
    Dim NextTab As Integer
    
    CountTab = 0
    
    For i = 1 To TabNo
        CountTab = InStr(CountTab + 1, StringLine, Chr(9))
    Next i
    
    NextTab = InStr(CountTab + 1, StringLine, Chr(9))
    If NextTab = 0 Then
        GetTabData = Right(StringLine, Len(StringLine) - CountTab)
    Else
        GetTabData = Mid(StringLine, CountTab + 1, NextTab - CountTab - 1)
    End If

Exit Function
Err_Handler:
    MsgBox Err & "-ST40:" & Error$
End Function

Sub CheckRegionalSettings()
On Error GoTo Err_Handler

RegDecSymbol = GetRegionalSettings(LOCALE_SDECIMAL) 'NZ = "."
RegThousandSeperator = GetRegionalSettings(LOCALE_STHOUSAND) 'NZ = ","

If RegThousandSeperator <> "," Then
    Call SetRegionalSettings(LOCALE_STHOUSAND, ",")
    Call SetRegionalSettings(LOCALE_SDECIMAL, ".")
    'Call MsgBox("Regional number settings changed, settings will be reset when Profiler is closed", vbOKOnly)
End If

Exit Sub
Err_Handler:
    MsgBox Err & "-ST41:" & Error$
    
End Sub

Function GetRegionalSettings(ByVal SettingType As String) As String
On Error GoTo Err_Handler

      Dim Symbol As String
      Dim iRet1 As Long
      Dim iRet2 As Long
      Dim lpLCDataVar As String
      Dim Pos As Integer
      Dim Locale As Long
      
      Locale = GetUserDefaultLCID()

'LOCALE_SDATE is the constant for the date separator
'as stated in declarations
'for any other locale setting just change the constant

'Function can also be re-written to take the
'locale symbol being requested as a parameter
      
      iRet1 = GetLocaleInfo(Locale, SettingType, _
      lpLCDataVar, 0)
      Symbol = String$(iRet1, 0)
      
      iRet2 = GetLocaleInfo(Locale, SettingType, Symbol, iRet1)
      Pos = InStr(Symbol, Chr$(0))
      If Pos > 0 Then
           Symbol = Left$(Symbol, Pos - 1)
           'MsgBox "Regional Setting = " + Symbol
      End If
      
      GetRegionalSettings = Symbol

Exit Function
Err_Handler:
    MsgBox Err & "-ST42:" & Error$
End Function

Function SetRegionalSettings(ByVal SettingType As String, ByVal Symbol As String) As String 'Change the regional setting
On Error GoTo Err_Handler

      Dim iRet As Long
      Dim Locale As Long
      
'LOCALE_SDATE is the constant for the date separator
'as stated in declarations
'for any other locale setting just change the constant

'Function can also be re-written to take the
'locale information being set as a parameter

      Locale = GetUserDefaultLCID() 'Get user Locale ID
      iRet = SetLocaleInfo(Locale, SettingType, Symbol)
     
Exit Function
Err_Handler:
    MsgBox Err & "-ST43:" & Error$
End Function

'ID4834 every INI needs its own temp directory '18 Feb 2013
Function TrimOffStartPathCharacters(ByVal StartPath As String) As String
On Error GoTo Err_Handler

Dim cp As Integer

cp = InStr(1, StartPath, "\")

If cp < 2 Then TrimOffStartPathCharacters = "\": Exit Function

TrimOffStartPathCharacters = Right(StartPath, Len(StartPath) - cp + 1)


Exit Function
Err_Handler:
    MsgBox Err & "-ST44:" & Error$
End Function

    Sub CreateDirectory(ByVal TheDirectory As String)
        On Error GoTo Err_Handler

        Dim BackSlash As Integer
        Dim DirectoryBuild As String
        BackSlash = 3

        On Error Resume Next
        Do
            BackSlash = InStr(BackSlash + 1, TheDirectory, "\")
            If (BackSlash) = 0 Then Exit Do
            DirectoryBuild = Left(TheDirectory, BackSlash)
            MkDir (DirectoryBuild)
        Loop

        Exit Sub
Err_Handler:
        Select Case Err()
            Case Else
                MsgBox (Err.Number & "-GM48:" & Err.Description)
        End Select
    End Sub

