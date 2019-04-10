#include <windows.h>
#include <d3dx9.h>

#include <stdlib.h>
#include <atlbase.h>
#include <stdio.h>
#include "window.h"
#include "..\houghlibv2.0\CBSAlgebra.h"


const double D3DPIPEVERSION = 10.9; //Fixed 3D Capture
//10.8; Put back the export panel in three dee export stl export 
//10.7 Just egnores textures and modles that are not accesable
//10.6; //ver 6.0.0 release, this inlcudes colour code from VB and Auto and Custom settings
//10.5; //Adjust level of Detail for Three Dim, and new release
//10.4; //PCN3141 Direct X version detection
//10.3; PCN3111 pas thru units to 3D Pipe, PCN3112 auto select vertexMode
//10.2; PCN3085 was a problem when you loaded multiple videos. But a fault was also
									// found in the three D, a serious one. Now Fixed.
// 10.1 Bug fixes and customer version
// 10.0 PCN3017 move over to ClearLine Profiler 5.5
// 2.3 PCN2693 (9 August 2004, Antony) needed divide by multiplier to get real values
// 2.2 PCN2950 default 3D Pipe to colour / flat. Also added the click to flat to make it change
// 2.1 PCN2860 Multiplier Passed to allow greater presision when displaying capacity, ovality and delta
// 2.0 PCN2473 Language support added to C++ D3D Pipe  
// 1.9 Created to allow VB to confirm version capatibility.
// 1.5 PCN2453, now number of profile points pased differently. Was  long with a mulitplier now its a text string.
// 1.6 PCN2465 & PCN2461 & PCN2467, Memory Leaks reduced from 13Meg to 3k, Protection to stop VB accessing
//     Variables and functions when 3D scene is unloaded, Missing Texture files and Directory displayed
// 1.7 PCN2510 (Antony van Iersel, 23 December 2003) fixed colour limits select, was crashing when selected.
// 1.8 PCN2367 (Antony van Iersel, 18 Feburary 2004) added direction arrow
// 1.9 PCN2653 (Antony van Iersel, 26 Feburary 2004) crashes when clossing (on some machines)


//Global Variables
D3Dwindow *wind;
bool d3d_initialized=false;

void Msg(TCHAR *szFormat, ...)
{
    TCHAR szBuffer[512];

    va_list pArgs;
    va_start(pArgs, szFormat);
    _vstprintf(szBuffer, szFormat, pArgs);
    va_end(pArgs);

    MessageBox(NULL, szBuffer, TEXT("LaserProfiler Message"), MB_OK | MB_ICONERROR);
}

//-----------------------------------------------------------------------------
// External function-prototypes
//-----------------------------------------------------------------------------
extern HRESULT GetDXVersion( DWORD* pdwDirectXVersion, TCHAR* strDirectXVersion, int cchDirectXVersion );

void __stdcall d3d_getversion(double *ver){
/* G. Logan 2/7/03

Whenever changing the version of a DLL, ensure the number complies with the following guidelines.

The DLL may be changed more or less often than the VB version.
Do we want to update a user with a new version of the VB software every time we change the DLL version? Probably not.
So the VB software will except DLL version with the same major version number. That is if the VB DLL version is 1.0, the VB will accept the DLL version 1.0 to 1.9. The VB will not DLL versions <1.0 or >1.9.
E.g.: ClearLine Profiler's LaserLib.dll version = 1.0. Then ClearLine will accept LaserLib.dll version from 1.0 to 1.9
E.g.: Now for 3D Pipe "Threedee.dll" PN2240 9 October
Therefore, for a VB software with a DLL version number of 1.0, ALL DLLs with versions 1.0 to 1.9 MUST work on this VB software.
If the change in the DLL means it will not work on ALL VB software of the same major DLL version, then the DLL's version MUST increase the major DLL version.

	Therefore, for a VB software with a DLL version number of 1.0, ALL DLLs with versions 1.0 to 1.9 MUST work on this VB software.
	If the change in the DLL means it will not work on ALL VB software of the same major DLL version, then the DLL's version MUST increase the major DLL version.
*/
	#ifdef _DEBUG
		Msg("Warning!!! This is a debug version threedim.dll.");
	#endif
	*ver = D3DPIPEVERSION;
	    //TCHAR strResult[128];
}

//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// PCN3141
// Name: d3d_directxversion
// Created By: Antony van Iersel
//
// Description:	sets the reference vairable passed from vb to the current direct x
//              version installed. Any messages to be displayed by VB
// Input: None
// Output: None
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
void __stdcall d3d_directxversion(int *ver)
{
	DWORD dwDirectXVersion = 0;
    TCHAR strDirectXVersion[10];

	GetDXVersion( &dwDirectXVersion, strDirectXVersion, 10 );
	*ver = HIWORD(dwDirectXVersion);

//	Msg("DirectX version is %d.%d : %s\n",
//         HIWORD(dwDirectXVersion), LOWORD(dwDirectXVersion),
//         strDirectXVersion);
//	if(HIWORD(dwDirectXVersion<9)) Msg("DirectX version is current less than 9\nYour version is %d.\nProfile may not operate as expected",
//		HIWORD(dwDirectXVersion));
}

///////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////
///  Interface with VB
///////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////
/*
void __stdcall d3d_initialise(HINSTANCE &hInstance, 
							  HWND hwnd1, 
							  char *vertexMode, //0 for hardware, 1 for mixed, 2 for software
							  float *datax, //PCN2988 pass the x coordinate data
							  float *datay, //PCN2988 pass the y coordinate data
							  float *centrex,
							  float *centrey,
							  int frames, //PCN2453 (reload many times Crash)
							  char *path,	
							  int *pvColourRed,
							  int *pvColourGreen,
							  int *pvColourBlue,
							  double expectedRad,//PCN2693 Needed for colour calculations (Antony van Iersel, 15 March 2004)
							  int pVCalcMult, //PC(9 August 2004, Antony) needed divide by multiplier to get real values
							  int pVDataXYMult, //PCN2988 13 Sept 2004 its a devision needed for the XY data
							  int units)	//PCN3111 0 = metric 1 = imperial
{

	d3d_initialized=false;
	if(wind!=NULL) {delete wind; wind=NULL;} // PCN2465 (8 December,Antony van Iersel)

	long numberData; // PCN2453 (reload many times Crash)
	int verMode;
	if(!strcmp(vertexMode,"Hardware")) verMode=0;
	else if(!strcmp(vertexMode,"Mixed")) verMode=1;
	else verMode=2;

	numberData=frames*180; // PCN3017
	if(units==1) expectedRad*=25.40; //PCN3111 if units are imperial (units = 1) convert to mm's

	wind = new D3Dwindow(hwnd1, verMode, 
		datax, //PCN2988 pass the x coordiante data
		datay, //PCN2988 pass teh y coordinate data
		centrex,
		centrey,
		numberData,  // Minus one Frame
		path, 
		pvColourRed,
		pvColourGreen,
		pvColourBlue,
		expectedRad,//PCN2693 Needed for colour calculations (Antony van Iersel, 15 March 2004)
		pVCalcMult,
		pVDataXYMult, //PCN2988 13 Sept 2004 its a devision needed for the XY data
		units); //PCN3111 0 = metric, 1 = imperial
	if(wind->deviceCreated) d3d_initialized=true;
}
*/

void __stdcall d3d_initialise(HINSTANCE &hInstance, 
							  HWND hwnd1, 
							  char *vertexMode, //0 for hardware, 1 for mixed, 2 for software
							  float *datax, //PCN2988 pass the x coordinate data
							  float *datay, //PCN2988 pass the y coordinate data
							  float *centrex,
							  float *centrey,
							  int frames, //PCN2453 (reload many times Crash)
							  char *path,	
							  int *colourRed,
							  int *colourGreen,
							  int *colourBlue,
							  double expectedRad,//PCN2693 Needed for colour calculations (Antony van Iersel, 15 March 2004)
							  int pVCalcMult, //PC(9 August 2004, Antony) needed divide by multiplier to get real values
							  int pVDataXYMult, //PCN2988 13 Sept 2004 its a devision needed for the XY data
							  int units)	//PCN3111 0 = metric 1 = imperial
{
	d3d_initialized=false;
	if(wind!=NULL) {delete wind; wind=NULL;} // PCN2465 (8 December,Antony van Iersel)

	long numberData; // PCN2453 (reload many times Crash)
	int verMode;
	if(!strcmp(vertexMode,"Hardware")) verMode=0;
	else if(!strcmp(vertexMode,"Mixed")) verMode=1;
	else verMode=2;

	numberData=frames*180; // PCN3017
	if(units==1) expectedRad*=25.40; //PCN3111 if units are imperial (units = 1) convert to mm's

	wind = new D3Dwindow(hwnd1, verMode, 
		datax, //PCN2988 pass the x coordiante data
		datay, //PCN2988 pass teh y coordinate data
		centrex,
		centrey,
		numberData,  // Minus one Frame
		path,
		colourRed,
		colourGreen,
		colourBlue,
		expectedRad,//PCN2693 Needed for colour calculations (Antony van Iersel, 15 March 2004)
		pVCalcMult,
		pVDataXYMult, //PCN2988 13 Sept 2004 its a devision needed for the XY data
		units); //PCN3111 0 = metric, 1 = imperial
	if(wind->deviceCreated) d3d_initialized=true;
}

///////////////////////////////////////////////////////////////////////////
// Name: ded_setlanguge PCN2473
// Created By: Antony van Iersel
// Date: 27 Febuary 2004
// 
// Description:	Resets the pointer from the default english to the VB
//              language setting. 
// Revision: 11 March 2004... No long resets the pointer but copies the new
//           text across passed from VB
//
// Input: which line to replace (line), string form VB (text)
//
// Output: None
////////////////////////////////////////////////////////////////////////////

void __stdcall d3d_setlanguage(int line, char *text)
{
	if(wind==NULL) return; // capture if threedim not initialised then egnore
	if(wind->languageArray==NULL) return; // if not initialised then egnore

	strcpy(wind->languageArray[line].text,text);	
	wind->UpdatePanelText();
}

void __stdcall d3d_capture_window(char *fileName, HWND hWndCaptureHandle)
{
	if(wind==NULL) return; //PCN2461 (8 December 2003, Antony) Prevents calls if wind not decleared
	wind->WindowToBmp(fileName, hWndCaptureHandle);
}

void _stdcall d3d_destroy(void)
	{
	if(d3d_initialized) 
		{ 
		d3d_initialized=false; 
		if(wind!=NULL) { delete wind; wind = NULL; } //PCN2461 (17 March 2004, Antony)
		wind=NULL; //PCN2461 (8 December 2003, Antony) 
		} 
	}

void _stdcall d3d_refresh(void)
	{
	if(wind==NULL) return; //PCN2461 (8 December 2003, Antony)
	if(!d3d_initialized || !wind->deviceCreated) return;
	wind->RenderD3D();
	wind->Motion();
	
	}

void _stdcall d3d_left_button_down(int x, int y)
	{
	if(wind==NULL) return; //PCN2461 (8 December 2003, Antony)
	if(!d3d_initialized || !wind->deviceCreated) return;
	wind->leftX=x;
	wind->leftY=y;
	wind->rightBut=false; wind->leftBut=true;
	}

void _stdcall d3d_right_button_down(int x, int y)
	{
	if(wind==NULL) return; //PCN2461 (8 December 2003, Antony)
	if(!d3d_initialized || !wind->deviceCreated) return;
	wind->rightX=x;
	wind->rightY=y;
	wind->rightBut=true; wind->leftBut=false;
	}

void _stdcall d3d_mousemove_and_down(int x, int y)
	{
	if(wind==NULL) return; //PCN2461 (8 December 2003, Antony)
	if(!d3d_initialized || !wind->deviceCreated) return;
	wind->MouseMotion(x,y);
	}

void _stdcall d3d_keydown(int key)
	{
	if(wind==NULL) return; //PCN2461 (8 December 2003, Antony)
	if(!d3d_initialized || !wind->deviceCreated) return;
	wind->Keyboard(key);
	}

void _stdcall d3d_laser_focus(int focus)
	{
	if(wind==NULL) return; //PCN2461 (8 December 2003, Antony)
	if(!d3d_initialized || !wind->deviceCreated) return;
	wind->MoveLaserTo(focus);
	}

void _stdcall d3d_camselect(int cam)
{
	if(wind==NULL) return; //PCN2461 (8 December 2003, Antony)
	if(!d3d_initialized || !wind->deviceCreated) return;
	wind->SelectCamera(cam);
}
void _stdcall d3d_pipe_scale(int scale)
{
	if(wind==NULL) return; //PCN2461 (8 December 2003, Antony)
	if(!d3d_initialized || !wind->deviceCreated) return;
	wind->PipeDepthScale((float) scale);
}

void _stdcall d3d_play_speed(int speed)
{
	if(wind==NULL) return; //PCN2461 (8 December 2003, Antony)
	if(!d3d_initialized || !wind->deviceCreated) return;
	if(wind->p==NULL) return; //PCN2461 (8 December 2003, Antony)
	wind->p->laserSpeed=speed;
}

void _stdcall d3d_zoom_speed(int speed)
{
	if(wind==NULL) return; //PCN2461 (8 December 2003, Antony)
	if(!d3d_initialized || !wind->deviceCreated) return;
	wind->Zoom(speed);

}

void _stdcall d3d_reset(void)
{
	if(wind==NULL) return; //PCN2461 (8 December 2003, Antony)
	if(!d3d_initialized || !wind->deviceCreated) return;
	wind->ResetCameras();
}

void _stdcall d3d_scene_on_off(int select)
{
	if(wind==NULL) return; //PCN2461 (8 December 2003, Antony)
	if(!d3d_initialized || !wind->deviceCreated) return;
	if(select==1) wind->drawPipeTrueFalse=!wind->drawPipeTrueFalse;
	if(select==0) wind->drawScapeTrueFalse=!wind->drawScapeTrueFalse;
}

void _stdcall d3d_rotate_pipe_z(int deg)
{
	if(wind==NULL) return; //PCN2461 (8 December 2003, Antony)
	if(!d3d_initialized || !wind->deviceCreated) return;
	if(wind->camSelect==1) wind->pipeCam->RollTargetZ((float) deg);
}

void _stdcall d3d_rotate_pipe_y(int deg)
{
	if(wind==NULL) return; //PCN2461 (8 December 2003, Antony)
	if(!d3d_initialized || !wind->deviceCreated) return;
	if(wind->camSelect==1) wind->pipeCam->RollTargetY((float) deg);
}

void _stdcall d3d_unload_pipe()
{
	if(wind==NULL) return; //PCN2461 (8 December 2003, Antony)
	if(!d3d_initialized || !wind->deviceCreated) return;
}

void _stdcall d3d_pipe_texture(char *)
{
	if(wind==NULL) return; //PCN2461 (8 December 2003, Antony)
	if(!d3d_initialized || !wind->deviceCreated) return;
}

void _stdcall d3d_next_pipe_texture(void)
{
	if(wind==NULL) return; //PCN2461 (8 December 2003, Antony)
	if(!d3d_initialized || !wind->deviceCreated) return;
	if(wind->p==NULL) return; //PCN2461 (8 December 2003, Antony)
	if(wind->p->shadeOn) wind->p->TogglePipeShade();
	else wind->p->NextPipeTexture();
}

void _stdcall d3d_pvgraphtype(char *type, double minLimit, double maxLimit)
{
	if(wind==NULL) return; //PCN2461 (8 December 2003, Antony)
	if(!d3d_initialized || !wind->deviceCreated) return;
	if(wind->p==NULL) return; //PCN2461 (8 December 2003, Antony)
	//if(wind->p->shadeType==0) return; //PCN3825 (14 June 2004) Just because prev type was zero doesn't mean
										// that we didn't want to set a new type.
	if(!(strcmp(type,"Capacity"))) wind->p->SetShade(1, minLimit, maxLimit);
	else if(!(strcmp(type,"Ovality")) ) wind->p->SetShade(2, minLimit, maxLimit);
	else if(!(strcmp(type,"Delta"))   ) wind->p->SetShade(3, minLimit, maxLimit);
	else if(!(strcmp(type,"Flat3D"))  ) wind->p->SetShade(3, minLimit, maxLimit);// PCN2950 add colour shading if shade type is flat
	else wind->p->SetShade(0, 0, 0);
}

//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Name: d3d_pipe_colour_limts
// Created By: Antony van Iersel
// 
// Description:	Toggles pipe colouring from texture to colour limts and back
// Input: None
// Output: None
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
void _stdcall d3d_pipe_colour_limits(void)
{
	if(wind==NULL) return; //PCN2461 (8 December 2003, Antony)
	if(!d3d_initialized || !wind->deviceCreated) return;
	if(wind->p==NULL) return; //PCN2461 (8 December 2003, Antony)
	wind->p->TogglePipeShade();
}


//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Name: d3d_export PCN2376 14 November 2003
// Created By: Antony van Iersel
// 
// Description:	Calls the ExportPipe Function windows which in turn calls the 
// ExportPipe function in pipe all the while passing the filename to where
// the pipe exported data is to be saved, given from VB.
// Input: File Name on where to export pipe
// Output: None
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
void _stdcall d3d_export_stl(char *filename)
{
	if(wind==NULL) return; //PCN2461 (8 December 2003, Antony)
	if(!d3d_initialized || !wind->deviceCreated) return;
	wind->ExportPipe(filename);
}
