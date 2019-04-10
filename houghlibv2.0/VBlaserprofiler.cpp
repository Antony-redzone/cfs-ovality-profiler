
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// VBlaserprofiler 24 November 2003
// Created By: Louise Shrimpton
// 
// Description:	This is not a class, just a file containing all functions that 
//	interact with the Visual Basic code.
//
// Functionality:
//
// Inherited From:  None
//	 	
// 
//
// 
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#include "VBlaserprofiler.h"
#include <time.h>
#include <stdio.h>
#include <dshow.h>

double LASERPROFILERVERSION = 			16.5; // Out of bounds error on vob, resolution 352,480
										//16.4; // forcefully added elecard aspect ratio to registry if not there, also set elecard indexing to 3
										//16.3; // removed baby smoothing, changed im so that its only the size of the video, not xtra. Swaped height with width, added some more memory cleanup. Downsample, was introducing errors, spikes. 
											  // rewrote with simple averaging of good data, if less than two poings good, then make a hole.
										// 16.2;	
										// Release
										// 16, mean then average smoothing on 540 profile points
										// 15.9 killed the water green dot overlay
										// 15.8 It was showing the points all the time, now only when you hit the process view button
										// 15.7 Added 540 display
										// Take out the smarts, its screwing things up
										// 15.5 The ability to remove mask
										// 15.4 Force elecard to egnore aspect ration on profiler.exe and default entry on the regkey entry		
										// 15.3 Better ele card incorperation.
										// 15.2 Centre calc moved back to the pre-processing with lots of smarts, filling in whole data to reduce water level efect during recording
										// 15.1 Snapshot now has a flag to decide to snapshot with or without fisheye, not turn fisheye off totally for just a snapshot. PCN4596	
										// 15.0 Ablity to turn video sync on and off at loading time noVidoeSync 1 = true, 0 = false
										// 14.9 Ablility to lock donut
										// 14.8 Ablity to turn the centre calulation right off.
										// 14.7 Still trying to not deadlock it, put back sleep(100) in timeseek function
										// 14.6 If seeking then dont even try while in that state. Hopefully this will help stop the elusive deadlock on phils machine, possibly customer aswell
										// 14.5 Dead lock detection added, if for somereason laserlib goes deadlock, it will wait a secound and unlock itself.
										// 14.4 Video framerewind
										// 14.3 Caught a divide by error in FishEye Initialisation.
										// 14.2 Removed auto tracking counter. Debug tracking off.
										// 14.1 Debug tracking on, and I hope the last of the function calls that are not sub byref
										// 13.9 accadently deleted the setwaterlevel reference to vb, not put back in :):)
										// 13.8 Opps, I put the time stamp pause checking not in the dump the but video processing.
										// 13.7 Fixed time stamp error, was adding a pause frame to recording. lastTimeReocrded added.
										// 13.6 vb functions turned to subs
										// 13.5 Added Inverted FishEye calculation
										// 13.4 Blurring was detecting the video edge as a laser edge, now fixed
										// 13.3 Added bluring , not gausian, just an average blur. Removed over flow bug on variance distance.
										// 13.1; // Pre-testing to see if it does what it suppose to do.
										// 13.0 PCN3561 Crashing on video.
										// 11.3 PCN3533, deadlock
										// 11.2 5.6 beta release for show in america and germany
										// 11.1 Use side of laser profile to find centre horizontal, reflection for vertical
										// 11.0 Use Light in center to track centre
										// 10.9 Laser Tracking - inverse video selection - semieliptical shape
										// 10.8 PCN2781 IM is now initialised after video is initialised.
										// 10.7 PCN3122 Bulls Eye for Centre and Centre history. Laser width and Donut now has min limits
										// 10.6 Fatel flaw fixed, tried to delete page one with movie.width as index, after movie was deleted
										// 10.5; PCN3085 More memory leaks fixed,  		
										// 10.4; PCN2395 Ablility to select capture device, two new VB function calls added
										// 10.3; New Centre Two
										// 10.2; PCNXXXX new centre.
										// 10.1; Bug fixes customer release
										// 10.0 Version change for Clearline Profiler 5.5
										 // PCN2290 Distortion(Fish-Eye). This makes LaserProfilerVersion 5.4
										 // 5.4 -> 5.5 PCN2400(Adjust center, use variance, Progress bar), PCN2420, PCN2421
										 // 5.5 -> 5.6 PCN2461 (Majour memory leak fix)
										 // 5.6 -> 5.7 PCN2426(More Circles.Preparation only at the moment.)
										 //			   PCN2433(Refined Auto Calibration - Showing BW image, Using diagonal length for distortion processing, Variance(Straightness)
										 // 5.7 -> 5.8 PCN2575 Contrast Control on Auto Calibration Added
										 // 5.8 -> 5.9 PCN2488 The circles that are touching the edge of the image will be egnored 
										 //                    when auto calibrating the fish eye. (They may be Half Circles)
										 // 5.9 -> 6.0 PCN2612 Upgrade to Manual Tune, and Manual Tune user-interface.
										 // 6.0 -> 6.1 PCN2668 Every Frame is Processed, No missed frames.
										 // 6.1 -> 6.2 PCN2405 Finding true centre
										 // 6.2 -> 6.3 PCN2639 Distance Coutner made Active
										 // 6.3 -> 6.4 PCN2639 Distance Counter Stable and Auto Setting
										 // 6.4 -> 6.5 PCN2778 Last majour memory leak removed, Memory leak detection added
										 //					   and initHw added, using this to stop accesing laserlib
										 //					   instead of cheaking if hw is NULL.
										 // 6.5 -> 6.6 Water Level is now working and egnored,
										 //			   Ticker counter is working and stable,
										 //			   Profile goes thru a multi pass smooth.
										 // 6.6 -> 6.7 PCN2865 FrameAdvance added to video used to frame step in profiler
										 // 6.7 -> 6.8 PCN2874 Ticker counter now able to count in feet and meters.
										 // 6.8 -> 6.9 PCN2874 Ticker counter now able to count in feet and meters.
										 // 6.9 -> 10  Version change for Profiler 5.5

HWND hwnd;            // main window
HDC dc;

///// Debug Setup for dumping time values to file on exit, 12 November 2004
//
// FILE *f;
// double timeDump[10000];
// int timeDumpHead=0;
//
/////////////////////////////

FILE *logger; //PCN3251
int loggerOn = false; //PCN3251
char loggerFileName[30];

Laserprofiler *hw=NULL;
int hwInit=false;


void WhoAccesed(char *string)
{
	return; //Msg(string);
}



HINSTANCE apphinst;

/*id Msg(TCHAR *szFormat, ...)
	{
    TCHAR szBuffer[512];

    va_list pArgs;
    va_start(pArgs, szFormat);
    _vstprintf(szBuffer, szFormat, pArgs);
    va_end(pArgs);

    MessageBox(NULL, szBuffer, TEXT("LaserProfiler Message"), MB_OK | MB_ICONERROR);
	}
*/






//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Name: setprofileoverlay 11 Feb 2004
// Created By: Antony van Iersel
//
// Description: Set to overlay the pipe profile on the video display
// Input: 0 off
//		  1 100% of profile size
//		  2 105% of profile size
// Output: None
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
void __stdcall setprofileoverlay(int i)
{
	if(!hwInit) {WhoAccesed("setprofileoverlay");return;}		//PCN2704 (Antony van Iersel, 12 March) prevent acess from
						    //now PCN2778 (Antony van Iersel, 22 April 2004), checking for NULL not reliable enough, now switch
	if(!hw->vidInit) return;
	if(hw->mediaType==1) return; // VB if not intialised //PCN3194
	if(loggerOn)
		{
		logger = fopen(loggerFileName,"a+");
		fprintf(logger,"setprofileoverlay, recieved %i as int\n",i);
		fclose(logger);
		}
	hw->overlayProfile=i;
	if(i!=0) 		hw->radialProfile->SetShowInternalCircles(true);
	else hw->radialProfile->SetShowInternalCircles(false);
}

//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Name: setimageanalysis 11 Feb 2004
// Created By: Antony van Iersel
//
// Description: Set the video mode for the image analysis
// Input: 0 normal, or see thru overlay
//		  1 image inhance (gray video, used to display colour filtering (R,G,B or comb))
//		  2 black of the video to show overlays.
// Output: None
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
void __stdcall setimageanalysis(int i)
{
	if(!hwInit) {WhoAccesed("setimageanalysis");return;}		//PCN2704 (Antony van Iersel, 12 March) prevent acess from
							//now PCN2778 (Antony van Iersel, 22 April 2004), checking for NULL not reliable enough, now switch
	if(!hw->vidInit) return;
	if(hw->mediaType==1) return; // VB if not intialised //PCN3194

	if(loggerOn) 
	{
		logger=fopen(loggerFileName, "a+");
		fprintf(logger,"setimageanalysis, recieved %i as int\n",i);
		fclose(logger);
	}
	if(i==0) hw->radialProfile->SetShowVideoFilter(false);
	else hw->radialProfile->SetShowVideoFilter(true);
}


//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Name: getimageanalysis 15 Feb 2005
// Created By: Antony van Iersel
//
// Description: Gets the video mode for the image analysis
// Input: 0 normal,
//		  1 image inhance (gray video, used to display colour filtering (R,G,B or comb))
// Output: None
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
void __stdcall getimageanalysis(int &imageInhance)
{
	int state;
	if(!hwInit) {imageInhance = 0; return;}		//PCN2704 (Antony van Iersel, 12 March) prevent acess from
							//now PCN2778 (Antony van Iersel, 22 April 2004), checking for NULL not reliable enough, now switch
	if(!hw->vidInit) {imageInhance = 0; return ;}
	if(hw->mediaType==1) {imageInhance = 0; return;} // VB if not intialised //PCN3194

	if(loggerOn) 
	{
		logger=fopen(loggerFileName,"a+");
		fprintf(logger,"getimageanalysis ");
		fclose(logger);
	}

	state = hw->radialProfile->GetShowVideoFilter();

	if(loggerOn) 
	{
		logger=fopen(loggerFileName,"a+");
		fprintf(logger," sent %i as int\n",state);
		fclose(logger);
	}
	

	imageInhance = state;
}





//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Name: hough_showlaserwidth 30 August 2004
// Created By: Antony van Iersel
//
// Description: Shows the laser width of the laser before it desides its the laser or not
// Input: onff, 0 for off, 1 for on.
// Output: None
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
void __stdcall hough_showlaserwidth(int onoff)
{
	if(!hwInit) {WhoAccesed("showlaserwidth");return;}		//PCN2704 (Antony van Iersel, 12 March) prevent acess from
							//now PCN2778 (Antony van Iersel, 22 April 2004), checking for NULL not reliable enough, now switch
	if(!hw->vidInit) return;
	if(hw->mediaType==1) return; // VB if not intialised //PCN3194
	if(loggerOn) 
	{
		logger = fopen(loggerFileName,"a+");
		fprintf(logger,"hough_showlaserwidth, recieved %i as int\n",onoff);
		fclose(logger);
	}
	
	if(onoff!=0) hw->radialProfile->SetLaserWidthOverlayOn(true); //PCNAVI 30 August 2004
	else hw->radialProfile->SetLaserWidthOverlayOn(false);
}

//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Name: setvideofiltertype 13 Feb 2004
// Created By: Antony van Iersel
//
// Description: Set the video filter mode for the image analysis
// Input: 0 red
//		  1 green
//		  2 blue
//		  3 mix (eg green + blue)
// Output: None
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
void __stdcall setvideofiltertype(int i)
{
	if(!hwInit) {WhoAccesed("setvideofiltertype");return;}		//PCN2704 (Antony van Iersel, 12 March) prevent acess from
							//now PCN2778 (Antony van Iersel, 22 April 2004), checking for NULL not reliable enough, now switch
	if(!hw->vidInit) return;
	if(hw->mediaType==1) return; // VB if not intialised //PCN3194
	if(hw->radialProfile==NULL) return;

	if(loggerOn)
	{
		logger=fopen(loggerFileName,"a+");
		fprintf(logger,"setvideofiltertype, recieved %i as int\n",i);
		fclose(logger);
	}

	hw->videoFilterType=i;
	hw->radialProfile->SetFilterType(i);

}

// Sets the inside donut size
void __stdcall hough_setinsidezone(double i)
{
	if(!hwInit) {WhoAccesed("hough_setinsidezone");return;}		//PCN2704 (Antony van Iersel, 12 March) prevent acess from
							//now PCN2778 (Antony van Iersel, 22 April 2004), checking for NULL not reliable enough, now switch
	if(!hw->vidInit) return;
	if(hw->mediaType==1) return; // VB if not intialised //PCN3194
	if(hw->radialProfile==NULL) return;
	
	if(loggerOn) 
	{
		logger=fopen(loggerFileName,"a+");
		fprintf(logger,"hough_setinsidezone, recieved %f as double\n",i);
		fclose(logger);
	}

	hw->radialProfile->SetInternalRadius(1-(i/100));
}

// Sets the outside donut size
void __stdcall hough_setoutsidezone(double i)
{
	if(!hwInit) {WhoAccesed("hough_setoutsidezone");return;}		//PCN2704 (Antony van Iersel, 12 March) prevent acess from
							//now PCN2778 (Antony van Iersel, 22 April 2004), checking for NULL not reliable enough, now switch
	if(!hw->vidInit) return;
	if(hw->mediaType==1) return; // VB if not intialised //PCN3194
	if(hw->radialProfile==NULL) return;
	
	if(loggerOn) 
	{
		logger = fopen(loggerFileName,"a+");
		fprintf(logger,"hough_setoutsidezone, recieved %f as double\n",i);
		fclose(logger);
	}
	
	hw->radialProfile->SetExternalRadius(1+(i/100));
}

//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Name: setprofilecandidates 13 Feb 2004
// Created By: Antony van Iersel
//
// Description: Shows the pre selection filter as green.
// Input: x, 0 for off, 1 for on.
//		  y, 0 for off, 1 for on.
// Output: None
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
void __stdcall setprofilecandidates(int onoff)
{
	if(!hwInit) {WhoAccesed("setprofilecandidates");return;}		//PCN2704 (Antony van Iersel, 12 March) prevent acess from
							//now PCN2778 (Antony van Iersel, 22 April 2004), checking for NULL not reliable enough, now switch
	if(!hw->vidInit) return;
	if(hw->mediaType==1) return; // VB if not intialised //PCN3194

	if(loggerOn) 
	{
		logger = fopen(loggerFileName,"a+");
		fprintf(logger,"setprofilecandidates, recieved %i as int\n",onoff);
		fclose(logger);
	}

	if(onoff!=0) 
	{
		hw->radialProfile->SetShowProfileCandidatesOverlay(true); //PCNAVI 30 August 2004

	}
	else 
	{
		hw->radialProfile->SetShowProfileCandidatesOverlay(false);

	}
}
			   
// Sets the laser width
void __stdcall setstandarddeviation(double xwidth, double ywidth, double sdx, double sdy)
{
	if(!hwInit) {WhoAccesed("setstandarddeviation");return;}		//PCN2704 (Antony van Iersel, 12 March) prevent acess from
							//now PCN2778 (Antony van Iersel, 22 April 2004), checking for NULL not reliable enough, now switch
	if(!hw->vidInit) return;
	if(hw->mediaType==1) return; // VB if not intialised //PCN3194

	if(loggerOn) 
	{
		logger = fopen(loggerFileName,"a+");
		fprintf(logger,"setstandarddeviation , note: ywidth not used\n");
		fprintf(logger,"recieved xwidth %f as double\n",xwidth);
		fprintf(logger,"recieved ywidth %f as double\n",ywidth);
		fprintf(logger,"recieved sdx %f as double\n",sdx);
		fprintf(logger,"recieved sdy %f as double\n",sdy);
		fclose(logger);
	}
	

	hw->SD_X= sdx;
	hw->SD_Y= sdy;
	hw->radialProfile->SetLaserWidth((int) xwidth); //PCNAVI


}

// Sets the cutoff level, 
void __stdcall setgradthreshold(int th)
{
	if(!hwInit) {WhoAccesed("setgradthreshold");return;}		//PCN2704 (Antony van Iersel, 12 March) prevent acess from
							//now PCN2778 (Antony van Iersel, 22 April 2004), checking for NULL not reliable enough, now switch
	if(!hw->vidInit) return;
	if(hw->mediaType==1) return; // VB if not intialised //PCN3194

	if(loggerOn) 
	{
		logger = fopen(loggerFileName, "a+");
		fprintf(logger,"setgradthreshold, recieved %i as int\n",th);
		fclose(logger);
	}
	hw->radialProfile->SetCutoffLevel(th); //PCN????
}


//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Name: getpipesettings 5 December 2003
// Created By: Louise Shrimpton
// 
// Description:	Set to process a light or a dark pipe.  
// Input: 0 or 1
// Output: None
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
void __stdcall getpipesettings(double *XT, double *YT, int *GT, double *SDX, double *SDY, int *greenx, int *greeny, int *prof, int *col, int *percprofpnts, double *totalper, double *xadj){
	if(!hwInit) {WhoAccesed("getpipesettings");return;}		//PCN2704 (Antony van Iersel, 12 March) prevent acess from VB if not init
							//now PCN2778 (Antony van Iersel, 22 April 2004), checking for NULL not reliable enough, now switch

	if(loggerOn) 
	{
		logger = fopen(loggerFileName,"a+");
		fprintf(logger,"getpipesettings,\n");
		fprintf(logger,"passed XT %f as double\n",XT);
		fprintf(logger,"passed YT %f as double\n",YT);
		fprintf(logger,"passed GT %i as int\n",GT);
		fprintf(logger,"passed SDX %f as double\n",SDX);
		fprintf(logger,"passed SDY %f as double\n",SDY);
		fprintf(logger,"passed greenx %i as int\n",greenx);
		fprintf(logger,"passed greeny %i as int\n",greeny);
		fprintf(logger,"passed prof %i as int\n",prof);
		fprintf(logger,"passed col %i as int\n",col);
		fprintf(logger,"passed percprofpnts %i as int\n",percprofpnts);
		fprintf(logger,"passed totalper %f as double\n",totalper);
		fprintf(logger,"passed xadj %f as double\n",xadj);
		fclose(logger);
	}
	hw->getvariables(XT, YT, GT, SDX, SDY, greenx, greeny, prof, col, percprofpnts, totalper, xadj);
}

// Get the counter value from the counter class
void __stdcall getcounter(int &value)
{
	int answer;
	
	if(!hwInit) {value = 0;return;}		//PCN2704 (Antony van Iersel, 12 March) prevent acess from VB if not init
							//now PCN2778 (Antony van Iersel, 22 April 2004), checking for NULL not reliable enough, now switch
	if(!hw->vidInit)  answer = 0;
	else if(hw->mediaType==1) answer = 0; //PCN3194 If dont want to process
	else if(hw->IPD!=NULL) answer = hw->GetIPDDistance();
	//else if(hw->tickCount==NULL) answer= 0;
	else  answer = hw->tickCount->count; 
	
	if(loggerOn) 
	{
		logger = fopen(loggerFileName,"a+");
		fprintf(logger,"getcounter, returned %i as int\n",answer);
		fclose(logger);
	}
	
	value=answer;
}

//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// PCN2668
// Name: hough_getprocesstime
// Created By: Antony van Iersel, 24 August 2004
// 
// Description:	Sets the passed vairable to how long it takes to process the image.  
// Input:Sets passed vairable to how long it takes to process framecb
// Output: None
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ 
void __stdcall hough_getprocesstime(double &t)
{
	if(!hwInit) {WhoAccesed("getcounter");return;}		//PCN2704 (Antony van Iersel, 12 March) prevent acess from VB if not init
							//now PCN2778 (Antony van Iersel, 22 April 2004), checking for NULL not reliable enough, now switch
	if(!hw->vidInit) {t=0; return;}
	if(hw->mediaType==1) {t=0; return;} //PCN3194 If dont want to process

	if(loggerOn) 
	{
		logger = fopen(loggerFileName, "a+");
		fprintf(logger,"hough_getprocesstime, passed %f as double\n",t);
		fclose(logger);
	}
	
	t=hw->timeToProcessImage*1000;
}


void __stdcall setpipesettings(double XT, double YT, int GT, double SDX, double SDY, int greenx, int greeny, int prof, double xadj){//, int col, int textmid){
	if(!hwInit) {WhoAccesed("setpipesettings");return;}		//PCN2704 (Antony van Iersel, 12 March) prevent acess from VB if not init
							//now PCN2778 (Antony van Iersel, 22 April 2004), checking for NULL not reliable enough, now switch

	
	if(loggerOn) 
	{
		logger = fopen(loggerFileName,"a+");
		fprintf(logger,"setpipesettings,\n");
		fprintf(logger,"set XT %f as double\n",XT);
		fprintf(logger,"set YT %f as double\n",YT);
		fprintf(logger,"set GT %i as int\n",GT);
		fprintf(logger,"set SDX %f as double\n",SDX);
		fprintf(logger,"set SDY %f as double\n",SDY);
		fprintf(logger,"set greenx %i as int\n",greenx);
		fprintf(logger,"set greeny %i as int\n",greeny);
		fprintf(logger,"set prof %i as int\n",prof);
		fprintf(logger,"set xadj %f as double\n",xadj);
		fclose(logger);
	}
	hw->setvariables(XT, YT, GT, SDX, SDY, greenx, greeny, prof, xadj);
}

// Clear the profile buffere array
void __stdcall emptybuffer(){
	if(!hwInit) {WhoAccesed("emptybuffer");return;}		//PCN2704 (Antony van Iersel, 12 March) prevent acess from VB if not init
							//now PCN2778 (Antony van Iersel, 22 April 2004), checking for NULL not reliable enough, now switch
	if(loggerOn) 
	{
		logger = fopen(loggerFileName,"a+");
		fprintf(logger,"emptybuffer\n");
		fclose(logger);
	}
	hw->LPemptybuffer();
} 


// Video seek index by time
void _stdcall timeseek(double t) {
	long currentT;
	bool didASeek=false;
	if(!hwInit) {WhoAccesed("timeseek");return;}		//PCN2704 (Antony van Iersel, 12 March) prevent acess from
							//now PCN2778 (Antony van Iersel, 22 April 2004), checking for NULL not reliable enough, now switch
	if(!hw->vidInit) return ; // VB if not intialised
	if(hw->mediaType==2) return;
	if(hw->mediaType==1) return;
//	double videoRate;

	if(loggerOn)
	{
		logger=fopen(loggerFileName,"a+");
		fprintf(logger, "time seek recieved %f as double\n",t);
		fclose(logger);
	}
	
	LONGLONG currentTime;
	
	currentT=(long) (t*1000);

	if(hw->movie->lastSeek!=currentT) hw->movie->seekTime(t);
	hw->movie->lastSeek=(long) (t*1000);

	

	if(loggerOn)
	{
		logger = fopen(loggerFileName,"a+");

		currentTime = hw->movie->gettime();
		
		fprintf(logger,"Time seeked to %f, currentFrameGrabeTime is set to %f, %i called seeked\n",
		(double) currentTime/10000000,
		hw->currentFrameGrabeTime,
		didASeek); //PCN3251
		fclose(logger);
	}
}

// Pause video, really pauses the graph
void __stdcall videopause(void) {
	if(!hwInit) {WhoAccesed("videopause");return;}		//PCN2704 (Antony van Iersel, 12 March) prevent acess from
							//now PCN2778 (Antony van Iersel, 22 April 2004), checking for NULL not reliable enough, now switch
	if(!hw->vidInit) return; // VB if not intialised
	if(hw->mediaType==2) return;
	if(hw->mediaType==1) return;

	if(loggerOn) 
	{
		logger = fopen(loggerFileName , "a+");
		fprintf(logger,"videopause\n"); //PCN3251
		fclose(logger);
	}
	hw->movie->VideoRun(false);
	hw->movie->pause();
}

// Run the vidoe, really runs the graph
void __stdcall videorun(void) {
	if(!hwInit) {WhoAccesed("videorun");return;}		//PCN2704 (Antony van Iersel, 12 March) prevent acess from
							//now PCN2778 (Antony van Iersel, 22 April 2004), checking for NULL not reliable enough, now switch
	if(!hw->vidInit) return; // VB if not intialised
	if(hw->mediaType==1) return;
	if(loggerOn) 
	{
		logger = fopen(loggerFileName, "a+");
		fprintf(logger,"videorun\n");
		fclose(logger);
	}

	hw->movie->VideoRun(true);
	hw->movie->run();
}

//PCN2668 (3 March 2004 - Antony van Iersel) Now can step too next frame instead of using
//seek.

void __stdcall videoframeadvance(void)
{
	if(!hwInit) {WhoAccesed("videoframeadance\n");return;}		//PCN2704 (Antony van Iersel, 12 March) prevent acess from
							//now PCN2778 (Antony van Iersel, 22 April 2004), checking for NULL not reliable enough, now switch
	if(!hw->vidInit) return; // VB if not intialised
	if(hw->mediaType==2) return;
	if(hw->mediaType==1) return;
	
	if(loggerOn) 
	{
		logger = fopen(loggerFileName, "a+");
		fprintf(logger,"videoframeadvance"); //PCN3251
		fclose(logger);
	}
	videopause(); //PCN3533
	hw->movie->FrameAdvance();
}

void __stdcall videoframerewind(void)
{
	if(!hwInit) return;		//PCN2704 (Antony van Iersel, 12 March) prevent acess from
							//now PCN2778 (Antony van Iersel, 22 April 2004), checking for NULL not reliable enough, now switch
	if(!hw->vidInit) return; // VB if not intialised
	if(hw->mediaType==2) return;
	if(hw->mediaType==1) return;
	
	if(loggerOn) 
	{
		logger = fopen(loggerFileName, "a+");
		fprintf(logger,"videoframerewind"); //PCN3251
		fclose(logger);
	}
	videopause(); //PCN3533
	hw->movie->FrameRewind();
}

// Step to next video frame (not index by time or frame) used for Recording PVD
void __stdcall videostep(void)
{
	if(!hwInit) {WhoAccesed("videostep");return;}		//PCN2704 (Antony van Iersel, 12 March) prevent acess from
							//now PCN2778 (Antony van Iersel, 22 April 2004), checking for NULL not reliable enough, now switch
	if(!hw->vidInit) return; // VB if not intialised
	if(hw->mediaType==2) return;
	if(hw->mediaType==1) return;

	if(loggerOn) 
	{
		logger = fopen(loggerFileName, "a+");
		fprintf(logger,"videostep\n");
		fclose(logger);
	}

	hw->movie->step();
}



// Increase the video playback speed
void __stdcall increasespeed(void) {
	if(!hwInit) {WhoAccesed("increasespeed");return;}		//PCN2704 (Antony van Iersel, 12 March) prevent acess from
							//now PCN2778 (Antony van Iersel, 22 April 2004), checking for NULL not reliable enough, now switch
	if(!hw->vidInit) return; // VB if not intialised
	if(hw->mediaType==1) return;

	if(loggerOn) 
	{
		logger = fopen(loggerFileName,"a+");
		fprintf(logger, "increasespeed\n");
		fclose(logger);
	}
	hw->movie->ffwd();
}

// Decrease the video playback speed
void __stdcall decreasespeed(void) {
	if(!hwInit) {WhoAccesed("decreasespeed");return;}		//PCN2704 (Antony van Iersel, 12 March) prevent acess from
							//now PCN2778 (Antony van Iersel, 22 April 2004), checking for NULL not reliable enough, now switch
	if(!hw->vidInit) return; // VB if not intialised
	if(hw->mediaType==1) return;

	if(loggerOn) 
	{
		logger = fopen(loggerFileName,"a+");
		fprintf(logger, "decreasespeed\n");
		fclose(logger);
	}

	hw->movie->rwnd();
}

// Gets the total movie length in Time, units off seconds
void __stdcall getTime(double &retTime) {
	if(!hwInit) {retTime = 0;return;}		//PCN2704 (Antony van Iersel, 12 March) prevent acess from
							//now PCN2778 (Antony van Iersel, 22 April 2004), checking for NULL not reliable enough, now switch
	if(!hw->vidInit) {retTime = 0;return;}		//PCN2704 (Antony van Iersel, 12 March) prevent acess from // VB if not intialised
	if(hw->mediaType==1) {retTime = 0;return;}		//PCN2704 (Antony van Iersel, 12 March) prevent acess from

	LONGLONG time;
	if(loggerOn) 
	{
		logger = fopen(loggerFileName,"a+");
		fprintf(logger, "getTime ");
		fclose(logger);

	}
	time = hw->movie->videoLengthTime;

	retTime = (double)(time)/10000000;

	if(loggerOn)
	{
		logger = fopen(loggerFileName,"a+");
		fprintf(logger,"sent %f as double (Movie length)\n",retTime); //PCN3251
		fclose(logger);
	}
}				   

void __stdcall grabsnapshot(char *name, int registered, char *watermark,int fishEyeOn) {
	if(!hwInit) {WhoAccesed("grabsnapshot");return;}		//PCN2704 (Antony van Iersel, 12 March) prevent acess from
							//now PCN2778 (Antony van Iersel, 22 April 2004), checking for NULL not reliable enough, now switch
	if(!hw->vidInit) return; // VB if not intialised
	if(hw->mediaType==1) return;

	if(loggerOn)
	{
		logger = fopen(loggerFileName, "a+");
		fprintf(logger,"Snapshot\n"); //PCN3251
		fclose(logger);
	}
	
	hw->movie->grab(name,registered,watermark,fishEyeOn);
}


void __stdcall showrectangle(){
	if(!hwInit) {WhoAccesed("showrectangle");return;}		//PCN2704 (Antony van Iersel, 12 March) prevent acess from VB if not init
							//now PCN2778 (Antony van Iersel, 22 April 2004), checking for NULL not reliable enough, now switch
	
	if(loggerOn)
	{
		logger = fopen(loggerFileName, "a+");
		fprintf(logger, "showrectangle\n");
		fclose(logger);
	}
	hw->showrect();
}


//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// PCN2568 rewritten on 11 May 2004, Antony
// Name: setwaterlevel
// Created By: Original by lou
// Description: Takes the angles t1 and t2 and converts them to
//              water level values, wlLeft (water level Left) and wlRight (water level Right)
//				Converted to equivalent profile points at those angle, every
//				profile that comebetween these values can be egnored if wanted.
// Input: t1 left water level, t2 right water level
// Output: NONE
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


//void __stdcall setwaterlevel(double t1, double t2){
void __stdcall setwaterlevel(int *egnoreList)
{
	int i;
	if(hw==0) return;
	if(hw->radialProfile==0) return;
	if(loggerOn)
	{
		logger = fopen(loggerFileName, "a+");
		fprintf(logger, "setwaterlevel recieved ");
		for(i=0;i<PROFILE_SIZE;i++) fprintf(logger, "%i ",egnoreList[i]);
		fprintf(logger, "\n");
		fclose(logger);
	}
	
	for(i=0;i<PROFILE_SIZE;i++) hw->radialProfile->SetWaterLevel(i,egnoreList[i]);

}

void __stdcall setrectanglecoord(float xbottom, float ybottom, float xtop, float ytop, int setclear)
{
	float xb,yb,xt,yt;
	if(!hwInit) {WhoAccesed("setrectanglecoord");return;}		//PCN2704 (Antony van Iersel, 12 March) prevent acess from
							//now PCN2778 (Antony van Iersel, 22 April 2004), checking for NULL not reliable enough, now switch
	if(!hw->vidInit) return; // VB if not intialised

	if(loggerOn)
	{
		logger = fopen(loggerFileName, "a+");
		fprintf(logger, "setrectanglecoord sent %i, %i, %i,%i\n",xbottom, ybottom, xtop, ytop);
		fclose(logger);
	}


	if(xbottom > xtop)	{ xb = xtop; xt = xbottom;}
	else				{ xb = xbottom; xt = xtop;}
	
	if(ybottom > ytop)  { yb = ytop; yt = ybottom;}
	else				{ yb = ybottom; yt = ytop;}
	

	xb = xb / 100 * hw->movie->width;
	xt = xt / 100 * hw->movie->width;

	yb = yb / 100 * hw->movie->height;
	yt = yt / 100 * hw->movie->height;

	hw->radialProfile->SetTextMask(xb,yb,xt,yt, setclear);
}

void __stdcall hough_clearrectanglecoord(void)
{
	int i,j;
	if(!hwInit) {WhoAccesed("setrectanglecoord");return;}		//PCN2704 (Antony van Iersel, 12 March) prevent acess from
							//now PCN2778 (Antony van Iersel, 22 April 2004), checking for NULL not reliable enough, now switch
	if(!hw->vidInit) return; // VB if not intialised
	if(loggerOn)
	{
		logger = fopen(loggerFileName, "a+");
		fprintf(logger, "hough_clearrectanglecoord\n");
		fclose(logger);
	}


	for(i=0;i<hw->movie->height;i++)
		for(j=0;j<hw->movie->width;j++)
			hw->radialProfile->egnoreMask[i][j]=0;
}

//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// PCN2639 (29 March 2004)
// Name: setdistancerecrectangle
// Created By: Antony van Iersel
// Description: Sets the rectange for the distance counter
// Input: NONE
// Output: NONE
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

void __stdcall setdistancerectangle(int xbottom, int ybottom, int xtop, int ytop, 
									int units) // Units added for PCN2874 0 = meters, 1 = feet. (Ant, 8 June 2004)
{
///////////////////// PCN2639 ///////////////////////////////////////
	if(!hwInit) {WhoAccesed("setdistancerectangle");return;}		//PCN2704 (Antony van Iersel, 12 March) prevent acess from
							//now PCN2778 (Antony van Iersel, 22 April 2004), checking for NULL not reliable enough, now switch

	if(!hw->vidInit) return; // VB if not intialised
	if(hw->mediaType==1) return;

	if(loggerOn)
	{
		logger = fopen(loggerFileName, "a+");
		fprintf(logger, "setdistancerectangle recieved %i,%i, %i,%i,i%\n",xbottom, ybottom, xtop, ytop,units);
		fclose(logger);
	}


	int currentDistance=0, currentDirection=1;
	//if(hw->tickCount!=NULL) 
	//	{ 
	//	currentDistance=hw->tickCount->count; 
	//	currentDirection=hw->tickCount->direction;
	//	delete hw->tickCount; hw->tickCount=NULL;
	//	}
	//hw->tickCount = new Counter();
	//hw->tickCount->units=units; // PCN2874 (Antony van Iersel 8 June 2004)
	//hw->tickCount->count=currentDistance;
	//hw->tickCount->direction=currentDirection;
	//hw->tickCount->SetCounterMask(xtop,xbottom,ytop,ybottom);
	//hw->radialProfile->SetCounterMask(hw->tickCount->xLeft, hw->tickCount->yTop, hw->tickCount->xRight, hw->tickCount->yLower);
	
	
}

//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// PCN2639 (29 March 2004)
// Name: cleardistancerectangle
// Created By: Antony van Iersel
// Description: Sets the rectange for the distance counter and sets counter to true
// Input: NONE
// Output: NONE
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

void __stdcall cleardistancerectangle(void)
{
	return;
/*
	if(!hwInit) {WhoAccesed("cleardistancerectangle");return;}		//PCN2704 (Antony van Iersel, 12 March) prevent acess from
							//now PCN2778 (Antony van Iersel, 22 April 2004), checking for NULL not reliable enough, now switch
	if(!hw->vidInit) return; // VB if not intialised
	if(hw->tickCount==NULL) return;
	if(hw->mediaType==1) return;

	if(loggerOn)
	{
		logger = fopen(loggerFileName, "a+");
		fprintf(logger, "cleardistancerectangle\n");
		fclose(logger);
	}


	hw->tickCount->isSet=false;
	// PCN2864 (Antony van Iersel, 2 June 2004) Counter was disable but not cleared. Now //
	// the counter is set to 0,0,0,0. That is as good as cleared. /////////////////////////
	hw->tickCount->xLeft=0;	 ///////////////////////////////////////	
	hw->tickCount->xRight=0; //
	hw->tickCount->yTop=0;	 //
	hw->tickCount->yLower=0; //
	hw->radialProfile->SetCounterMask(0,0,0,0);
*/
}


//if i == 0 then clear, if i==1 then set for ignore centre (Now ignoreWaterLevel)
//j works the same way, but for the profile ignoreprofile (profileWaterLevel)
void __stdcall setwaterlevelbool(int i, int j){
	if(!hwInit) {WhoAccesed("setwaterlevelbool");return;}		//PCN2704 (Antony van Iersel, 12 March) prevent acess from VB if not init
							//now PCN2778 (Antony van Iersel, 22 April 2004), checking for NULL not reliable enough, now switch
	if(!hw->vidInit) return; // VB if not intialised 
	if(loggerOn)
	{
		logger = fopen(loggerFileName, "a+");
		fprintf(logger, "setwaterlevel recieved %i, %i\n",i,j);
		fclose(logger);
	}


	if(i) hw->radialProfile->SetIgnoreWaterLevel(true); //ignorecenter=true;
	else hw->radialProfile->SetIgnoreWaterLevel(false); //ignorecenter = false;
	if(j) hw->radialProfile->SetShowProfileWaterLevel(false); //ignoreprofile=true; 
	else hw->radialProfile->SetShowProfileWaterLevel(true); //ignoreprofile=false;
//	hw->ignoreprofile=false;
	
}

void __stdcall setdeviceinput(void)
{
	if(!hwInit) {WhoAccesed("setdeviceinput");return;}		//PCN2704 (Antony van Iersel, 12 March) prevent acess from
							//now PCN2778 (Antony van Iersel, 22 April 2004), checking for NULL not reliable enough, now switch
	if(!hw->vidInit) return; // VB if not intialised
	if(hw->mediaType==1) return;

	if(loggerOn)
	{
		logger = fopen(loggerFileName, "a+");
		fprintf(logger, "setdeviceinput\n");
		fclose(logger);
	}


	hw->movie->SetDeviceInput();

}




void __stdcall resizewindow(void){
	if(!hwInit) {WhoAccesed("resizewindow");return;}		//PCN2704 (Antony van Iersel, 12 March) prevent acess from
							//now PCN2778 (Antony van Iersel, 22 April 2004), checking for NULL not reliable enough, now switch
	if(!hw->vidInit) return; // VB if not intialised
	if(hw->mediaType==1) return;

	if(loggerOn)
	{
		logger = fopen(loggerFileName, "a+");
		fprintf(logger, "resizewindow\n");
		fclose(logger);
	}

	hw->movie->ResizeVideoWindow();
}



void __stdcall setrecprofstat(int i){
	if(!hwInit) {WhoAccesed("setrecprofstat");return;}		//PCN2704 (Antony van Iersel, 12 March) prevent acess from
							//now PCN2778 (Antony van Iersel, 22 April 2004), checking for NULL not reliable enough, now switch
	if(!hw->vidInit) return; // VB if not intialised
	if(hw->mediaType==1) return;

	if(loggerOn)
	{
		logger = fopen(loggerFileName, "a+");
		fprintf(logger, "setrecprofstat sent\n");
		fclose(logger);
	}

	if(i==1){
		hw->movie->recordprofileinfo = true;
		emptybuffer();

	}else{
		hw->movie->recordprofileinfo = false;
	}
}

void __stdcall setwindow(int i){
	if(!hwInit) {WhoAccesed("setwindow");return;	}	//PCN2704 (Antony van Iersel, 12 March) prevent acess from
							//now PCN2778 (Antony van Iersel, 22 April 2004), checking for NULL not reliable enough, now switch
	if(!hw->vidInit) return; // VB if not intialised
	if(hw->mediaType==1) return;
	
	if(loggerOn)
	{
		logger = fopen(loggerFileName, "a+");
		fprintf(logger, "setwindow recieved %i\n",i);
		fclose(logger);
	}


	if(i==1){
		hw->movie->window = true;
	}else{
		hw->movie->window = false;
	}
}

void __stdcall refreshframe(void){
	if(!hwInit) {WhoAccesed("refreshframe");return;}		//PCN2704 (Antony van Iersel, 12 March) prevent acess from
							//now PCN2778 (Antony van Iersel, 22 April 2004), checking for NULL not reliable enough, now switch
	if(!hw->vidInit) return; // VB if not intialised
	if(hw->mediaType==1) return;

	if(loggerOn) 
	{
		logger = fopen(loggerFileName, "a+");
		fprintf(logger,"refreshframe\n"); //PCN3251
		fclose(logger);
	}
	hw->movie->Refresh();
 }

void __stdcall optimize(int i){
	if(!hwInit) {WhoAccesed("optimize");return;}		//PCN2704 (Antony van Iersel, 12 March) prevent acess from VB if not init
							//now PCN2778 (Antony van Iersel, 22 April 2004), checking for NULL not reliable enough, now switch

	
	if(loggerOn)
	{
		logger = fopen(loggerFileName, "a+");
		fprintf(logger, "optimise recieved %i\n",i);
		fclose(logger);
	}

	if(i == 1) hw->optimized = 1;
	else hw->optimized = 0;
}

void __stdcall getcurrenttime(double &CurrentTime){
	if(!hwInit) {WhoAccesed("getcurrenttime"); CurrentTime = 0; return;}		//PCN2704 (Antony van Iersel, 12 March) prevent acess from
							//now PCN2778 (Antony van Iersel, 22 April 2004), checking for NULL not reliable enough, now switch
	if(!hw->vidInit) {CurrentTime = 0;return;} // VB if not intialised
	if(hw->mediaType==1) {CurrentTime = 0;return;}

	LONGLONG t;

	
	if(loggerOn)
	{
		logger = fopen(loggerFileName, "a+");
		fprintf(logger, "getcurrenttime ");
		fclose(logger);
	}

	t = hw->movie->gettime();

	if(loggerOn) 
	{
		logger = fopen(loggerFileName,"a+");
		fprintf(logger,"sent %f as double (Current Time), retrieved = %f\n",
		//hw->currentFrameGrabeTime,
		(double) t/10000000); //PCN3251
		fclose(logger);
	}
	CurrentTime =  ((double)t/10000000.0);
	//return hw->currentFrameGrabeTime;
	
}




//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Name: resetcounters 5 December 2003
// Created By: Louise Shrimpton
// 
// Description:  Sets the frame number to 0 and empties the profilebuffer  
//		
// Input: NONE
// Output: NONE
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
void __stdcall resetcounters(){
	if(!hwInit) {WhoAccesed("resetcounters");return;}		//PCN2704 (Antony van Iersel, 12 March) prevent acess from VB if not init
							//now PCN2778 (Antony van Iersel, 22 April 2004), checking for NULL not reliable enough, now switch
	if(!hw->vidInit) return; // VB if not intialised //PCN3194
	if(hw->mediaType==1) return;

	if(loggerOn)
	{
		logger = fopen(loggerFileName, "a+");
		fprintf(logger, "resetcounters\n");
		fclose(logger);
	}

	hw->frameno=0;
	hw->LoosingInfo = false;
	emptybuffer();
}


//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Name: getcenter 5 December 2003
// Created By: Louise Shrimpton
// 
// Description:  To retrive the current frame's center.  
//		
// Input: NONE
// Output: x and y become the center of the pipe for the current frame
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
void __stdcall getcenter(float *x, float *y){
	if(!hwInit) {WhoAccesed("getcenter");return;}		//PCN2704 (Antony van Iersel, 12 March) prevent acess from VB if not init
							//now PCN2778 (Antony van Iersel, 22 April 2004), checking for NULL not reliable enough, now switch

	if(!hw->vidInit) return; // VB if not intialised 
	if(loggerOn)
	{
		logger = fopen(loggerFileName, "a+");
		fprintf(logger, "getcentre\n");
		fclose(logger);
	}

	*x=0; //PCN3219
	*y=0; //PCN3219

	*x = hw->centerx;
	*y = hw->centery;
}


void __stdcall getversion(double *ver){ //PCN1970
/* G. Logan 2/7/03

Whenever changing the version of a DLL, ensure the number complies with the following guidelines.

The DLL may be changed more or less often than the VB version.
Do we want to update a user with a new version of the VB software every time we change the DLL version? Probably not.
So the VB software will except DLL version with the same major version number. That is if the VB DLL version is 1.0, the VB will accept the DLL version 1.0 to 1.9. The VB will not DLL versions <1.0 or >1.9.
E.g.: ClearLine Profiler's LaserLib.dll version = 1.0. Then ClearLine will accept LaserLib.dll version from 1.0 to 1.9

Therefore, for a VB software with a DLL version number of 1.0, ALL DLLs with versions 1.0 to 1.9 MUST work on this VB software.
If the change in the DLL means it will not work on ALL VB software of the same major DLL version, then the DLL's version MUST increase the major DLL version.
*/
//	if(loggerOn)
//	{
//		logger = fopen(loggerFileName, "a+");
//		fprintf(logger, "getversion\n");
//		fclose(logger);
//	}
	#ifdef _DEBUG
		Msg("Warning!!! This is a debug version laserlib.dll.");
	#endif

	*ver = LASERPROFILERVERSION;
}


//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Name: timestop 5 December 2003
// Created By: Louise Shrimpton
// 
// Description:  For use with the function getgroupedprofiledata.  Stops the video
//	to allow the VB to retrieve the profile data
//		
// Input: NONE
// Output: NONE
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
void __stdcall timestop(){
	if(!hwInit) {WhoAccesed("timestop");return;	}	//PCN2704 (Antony van Iersel, 12 March) prevent acess from
							//now PCN2778 (Antony van Iersel, 22 April 2004), checking for NULL not reliable enough, now switch
	if(!hw->vidInit) return; // VB if not intialised
	if(hw->mediaType==2) return;
	if(hw->mediaType==1) return;

	if(loggerOn) 
	{
		logger = fopen(loggerFileName, "a+");
		fprintf(logger,"timestop\n");//PCN3251
		fclose(logger);
	}
	videopause();

}

//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Name: timestart 5 December 2003
// Created By: Louise Shrimpton
// 
// Description:  For use with the function getgroupedprofiledata.  Starts the video
//	after it has been stopped to retrieve the data.
//		
// Input: NONE
// Output: NONE
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
void __stdcall timestart(){
	if(!hwInit) {WhoAccesed("timestart");return;}		//PCN2704 (Antony van Iersel, 12 March) prevent acess from
							//now PCN2778 (Antony van Iersel, 22 April 2004), checking for NULL not reliable enough, now switch
	if(!hw->vidInit) return; // VB if not intialised
	if(hw->mediaType==2) return;
	if(hw->mediaType==1) return;

	if(loggerOn) 
	{
		logger = fopen(loggerFileName,"a+");
		fprintf(logger,"timestart\n");
		fclose(logger);
	}
	videorun();

}




//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Name: initialise 5 December 2003
// Created By: Louise Shrimpton
// 
// Description:	Initialises the Dircet Show to allow video playback
// Input: the HINSTANCE from the VB, the handle of the window to display the video,
//	the name of the file and the device type (capture card?)
// Output: (As parameters) the height and width of the video, the x adjustment, the yadjustment,
//	the original width and height
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
void __stdcall initialise(HINSTANCE &hinst, 
						  HWND hwnd1, 
						  char *passedMediaType, 
						  char *name,
						  int *width, 
						  int *height, 
						  double *xadju, 
						  double *yadju, 
						  int *realheight, 
						  int *realwidth, 
						  int captureDevice,
						  int originalwidth,
						  int originalheight,
						  int noVideoSync)
{ // PCN2289 String added sDisplayType 20 oct 2003 PCN2395 select capture device
	
	    if(FAILED(CoInitialize(NULL))){
        Msg(TEXT("1022"));    
        return;
    }
	
	bool noSync;
	if(noVideoSync == 1 ) noSync = true;
	else noSync = false;
	
	if(loggerOn) Msg("Warning! Data loger is on, named C:\\LogFile + number.txt");
	if(loggerOn)
	{
		time_t ltime;
		time(&ltime);
		sprintf(loggerFileName,"C:\\LogFile%ld.txt",ltime);
		logger = fopen(loggerFileName,"a+");
		fprintf(logger, "Started video dll with file name %s\n",name);
		fclose(logger);
	}	

	if(loggerOn)
	{
		logger = fopen(loggerFileName,"a+");
		fprintf(logger, "About to load create laserprofiler\n",name);
		fclose(logger);
	}		
	hw=new Laserprofiler(name); //PCN3744 need to pass file name thru to check for IPD file


	if(hw == NULL) { Msg(TEXT("Not enough memory to create Laserprofiler.  Restart the application.")); return; }
	else hwInit=true;

	if(strcmp(passedMediaType,"Video") == 0) hw->mediaType=0;
	if(strcmp(passedMediaType,"Image") == 0) hw->mediaType=1;
	if(strcmp(passedMediaType,"Live") == 0) hw->mediaType=2;


	HRESULT hr=0;
	// Device type is now used for device select. nvidia is no longer needed seperatly
	//if(strcmp(sDeviceType,"nvidia")==0) hw->deviceType=1;// PCN2289 (1 is nvidia chipset)
	//else hw->deviceType=0;								 //         (0 is default       )
	hw->deviceType=0;

	//initialise semaphore for writing to bitmap
	hwnd=hwnd1;
	// Initialize COM
 //   if(FAILED(CoInitialize(NULL))){
 //       Msg(TEXT("1022"));    
 //       return;
 //   }
	hw->frameno = 0;
	dc=GetDC(NULL);
	if (hw->mediaType!=2) 
	{  //if it's not live, get the filename
		if((hw->vidInit) && (hw->mediaType!=1)) hw->movie->specifyfile(name);
	}
	


	if((hw->vidInit) && (hw->mediaType!=1))	
	{
		hr=hw->movie->CaptureVideo(hw->mediaType, captureDevice-1,noSync );

		if(hr != NULL){
			hwInit= false;
			strcpy(name, "erro"); // PCN2418 (21 November 2003, Antony van Iersel) VB has set aside
			delete hw;			  // 4 characters in the string name, so "error" does not if, now "erro"
			hw=NULL;
			return;
		}
		if(hw->vidInit) hw->movie->GetDimensions(width, height);
	
		if (FAILED (hr)) 
		{
			hwInit= false;
			Msg(TEXT("1023"));
			hw->movie->CloseInterfaces(hw->mediaType);
			DestroyWindow(hwnd);
			return;
		}
		else 
		{
			ShowWindow(hwnd, SW_SHOWNORMAL);
		}
	}
	else
	{
		hw->movie->height = *height;
		hw->movie->width = *width;
	}

	hw->getXAdjust(*width, *height, &hw->X_ADJUSTMENT, &hw->Y_ADJUSTMENT);
	*realheight = *height;
	*height = (int)(*height * hw->Y_ADJUSTMENT);
	*xadju = hw->Y_ADJUSTMENT;
	if(hw->xadjust) *yadju = hw->X_ADJUSTMENT;
	else *yadju = 0.0;
	hw->adjustsettings(*width, *height);
    
	
	hw->InitialiseIM();
	hw->radialProfile->Initialise(hw->im, hw->movie->width, hw->movie->height);
//	hw->radialProfile->centreLaser.SetVideoPointer(hw->im,hw->movie->height,hw->movie->width);
//	hw->radialProfile->centreLaser.SetLaserSize(5);
	
	if(loggerOn)
	{
		logger = fopen(loggerFileName,"a+");
		fprintf(logger, "About to Initialise Fisheye\n",name);
		fclose(logger);
	}		

	
//	PCN2904

	hw->FEye.Initialize(hw->movie->width,hw->movie->height,originalwidth,originalheight);
	if(loggerOn)
	{
		logger = fopen(loggerFileName,"a+");
		fprintf(logger, "video initialised and about to run live graph\n",name);
		fclose(logger);
	}	
	if(hw->mediaType==2) hw->movie->run(); //PCN???? If live then run the graph
	if(hw->mediaType==0) hw->movie->pause();

	if(loggerOn)
		{
			logger = fopen(loggerFileName,"a+");
			fprintf(logger, "Video successfully initialised\n",name);
			fclose(logger);
		}	

}

void msgbox(char *s) {
	MessageBox (HWND_DESKTOP, s, "Message", MB_OK | MB_ICONEXCLAMATION);
}


//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Name: gettotalnumframes 5 December 2003
// Created By: Louise Shrimpton
// 
// Description:	Returns the total number of frames in the profile buffer array
// Input: None
// Output: None
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
void __stdcall gettotalnumframes(int &numberOfFrames){
	int numberFrames;
	if(!hwInit) {numberOfFrames = 0;			return;}		//PCN2704 (Antony van Iersel, 12 March) prevent acess from VB if not init
	if(!hw->vidInit) {numberOfFrames = 0;		return;} // VB if not intialised //PCN3194
	if(hw->mediaType==1) {numberOfFrames = 0;	return;}

	if(loggerOn) 
	{
		logger = fopen(loggerFileName, "a+");
		fprintf(logger,"gettotalnumframes ");
		fclose(logger);
	}
	
	numberFrames = (int)(hw->in-hw->out);
	if(loggerOn) 
	{
		logger = fopen(loggerFileName, "a+");
		fprintf(logger,"sent %i as int (Number Frames in the profile array)\n",numberFrames); //PCN3251
		fclose(logger);
	}
	numberOfFrames = numberFrames;
}





//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Name: getgroupedprofiledataXY 5 December 2003
// Created By: Louise Shrimpton
// PCN6004 coordinates and centres now floats instead of ints
// 
// Description:	 this function will be called by the VB and will get all info for 
// the video once it has finished (or at a specific time increment, where the video 
// will be paused)
// Input: 
// Output:  An array X coordinates & Y coordinates, time information and the number of frames/1000.
//NOTE***  When calling this function in the VB, the data in the array goes from position 0 to
//position 179 (180 positions).
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
void __stdcall getgroupedprofiledataxy(float *xcoordinate, float *ycoordinate, float *xcentre, float *ycentre, double *time, int *numframes, int *distance){ //PCN2129 PCN2639 distance added (Antony van Iersel, 24 March 2004)
	if(!hwInit) {WhoAccesed("getgroupedprofiledataxy");return;}		//PCN2704 (Antony van Iersel, 12 March) prevent acess from
							//now PCN2778 (Antony van Iersel, 22 April 2004), checking for NULL not reliable enough, now switch
	if(!hw->vidInit) return; // VB if not intialised
	if(hw->mediaType==1) return;
	
	hw->movie->pause();

	if(loggerOn)
	{
		logger = fopen(loggerFileName, "a+");
		fprintf(logger , "getgroupeprofiledataxy\n");
		fclose(logger);
	}

	float x,y,cx,cy;
	vec2double screenCentre;

	double t; 
	int d;

	screenCentre.x = hw->movie->width/2;
	screenCentre.y = (hw->movie->height/2)*hw->Y_ADJUSTMENT;

	(*numframes) = (int)(hw->in-hw->out);
	for(int i=(int)0; i<(*numframes);i++){ //PCN2959 was <= was dumping upto and inclusive numframes, was one to many. 
		hw->profbuff.getprofilebuffer(i,NULL,&t,&cx,&cy,NULL,&d);//(i,profilebuffer[i].time; //PCN2959 was accessing array 0 - 1 and dumping rubish time on first frame
		cx=(cx - (float) screenCentre.x); //PCN6004 was * 10, but dont need that multiplier now
		cy=(cy - (float) screenCentre.y); //PCN6004 was * 10, but dont need that multiplier now
		
		for(int j=0;j<PROFILE_SIZE;j++) 
			{
			x=(hw->profbuff.getprofilebufferprofile(i,2,j)); //PCN6004 was * 10, but dont need that multiplier now
			y=(hw->profbuff.getprofilebufferprofile(i,3,j)); //PCN6004 was * 10, but dont need that multiplier now
			if((x==0) && (y==0))
				{
				xcoordinate[(i*PROFILE_SIZE)+j+1]=0;
				ycoordinate[(i*PROFILE_SIZE)+j+1]=0;
				}
			else
				{
				xcoordinate[(i*PROFILE_SIZE)+j+1]=(float) (x+cx); //PCN6004
				ycoordinate[(i*PROFILE_SIZE)+j+1]=(float) (y+cy); //PCN6004
				}
			}
		//find center, return coordinates of center
		//PCN3219 save centre of every profile
		xcentre[i] = (float) (-cx); //PCN6004
		ycentre[i] = (float) (-cy); //PCN6004
		time[i]=t;
		distance[i]=d;

//// Debug information, to get data to dump to file at exit.
//
//		if(timeDumpHead<10000) 
//		{
//			timeDump[timeDumpHead]=time[i];
//			timeDumpHead++;
//		}
//
////////////////////////////////////////////////////////////

	}
	hw->in =0;
	hw->out=0;

	if(loggerOn)
	{
		logger = fopen(loggerFileName, "a+");
		fprintf(logger , "profile data recieved\n");
		fclose(logger);
	}

}
	


//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// PCN2639
// Name: setdistancecounter 25 March 2004
// Created By: Antony van Iersel
// 
// Description:	Set the counter distance and direction
// Input: distance (Set the current counter distance), direction (0 for --, 1 for ++)
// Output: None
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

void __stdcall setdistancecounter(int distance, int direction) // 0 --, 1 ++
{
	return;
	/*
	if(!hwInit) return;
	if(!hw->vidInit) return;
	if(hw->tickCount==NULL) return;
	if(hw->mediaType==1) return;

	if(loggerOn)
	{
		logger = fopen(loggerFileName, "a+");
		fprintf(logger , "setdistancecounter recieved %i, %i\n",distance, direction);
		fclose(logger);
	}


	hw->tickCount->count=distance;
	hw->tickCount->direction=direction;
//	Msg("Distance to set is %i, direction is %i",(int) distance, (int) direction);
	*/
}

void __stdcall getdistancebuffer(int *buffer)
{
	return;
	/*
	if(!hwInit) return;
	if(!hw->vidInit) return;
	if(hw->tickCount==NULL) return;
	if(hw->mediaType==1) return;

	if(loggerOn)
	{
		logger = fopen(loggerFileName, "a+");
		fprintf(logger , "getdistancebuffer\n");
		fclose(logger);
	}

	*buffer=hw->tickCount->bufferValue;
	*/
}	


//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Name: uninitialiseLive 5 December 2003
// Created By: Louise Shrimpton
// 
// Description:	Must call when the VB closes live.  This uninitialises the 
//  Direct Show. 
// Input: None
// Output: None
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
void uninitialiseLive(){

	if(!hwInit) {WhoAccesed("uninitialiseLive");return;}		//PCN2704 (Antony van Iersel, 12 March) prevent acess from
							//now PCN2778 (Antony van Iersel, 22 April 2004), checking for NULL not reliable enough, now switch
	if(!hw->vidInit) return; // VB if not intialised
	if(hw->mediaType==1) return;

	if(loggerOn)
	{
		logger = fopen(loggerFileName, "a+");
		fprintf(logger , "uninitialiseLive\n");
		fclose(logger);
	}

	hwInit=false;
	hw->movie->CloseInterfaces(true);
	if(hw!=NULL) { delete hw; hw=NULL; }  //fix Friday 5/12/03  LS  Last fix!!
}


//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Name: uninitialise 5 December 2003
// Created By: Louise Shrimpton
// 
// Description:	Must call when the VB closes a video.  This uninitialises 
//	the Direct Show.
// Input: None
// Output: None
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
void uninitialise(){

	if(!hwInit) {WhoAccesed("unitialise");return;}		//PCN2704 (Antony van Iersel, 12 March) prevent acess from
							//now PCN2778 (Antony van Iersel, 22 April 2004), checking for NULL not reliable enough, now switch
	if(!hw->vidInit) return; // VB if not intialised //PCN3194

	if(loggerOn)
	{
		logger = fopen(loggerFileName, "a+");
		fprintf(logger , "uninitialise\n");
		fclose(logger);
	}

	hw->vidInit=false;
	hwInit=false;
	if(hw->mediaType!=1) hw->movie->CloseInterfaces(false); //PCN3194
	

	if(hw!=NULL) { delete hw; hw=NULL;}  //fix Friday 5/12/03  LS  Last fix!!

//// Used for debuging time, dumps time values on exit to file. 12 Nov 2004
//
//	f=fopen("C:\\TimeDump.txt","w");
//	int i;
//	for(i=0;i<timeDumpHead;i++)
//	{
//		fprintf(f,"Time: %f\n",timeDump[i]);
//	}
//	fclose(f);
//
///////////////////////////////////////////////////////////////////////////
	if(loggerOn)
	{
		logger = fopen(loggerFileName, "a+");
		fprintf(logger , "video closed, laserlib.dll unloaded\n");
		fclose(logger);
	}
	CoUninitialize();
	ReleaseDC(hwnd, dc);

}

Laserprofiler *getlp() // PCN2516
{	
	if(!hwInit) return NULL; // PCN2778 Antony van Iersel (1 June 2004)
	return hw;
}

void __stdcall turnfisheyeon(void){
	if(!hwInit) return;		//PCN2704 (Antony van Iersel, 25 July 2005) prevent acess from
		if(!hw->vidInit) return; // VB if not intialised
	if(loggerOn)
	{
		logger = fopen(loggerFileName, "a+");
		fprintf(logger, "turnfisheyeon\n");
		fclose(logger);
	}
	hw->FEye.TurnFishEyeOn();
}

void __stdcall turnfisheyeoff(void){
	if(!hwInit) return;		//PCN2704 (Antony van Iersel, 25 July 2005) prevent acess from
	if(!hw->vidInit) return; // VB if not intialised
	if(loggerOn)
	{
		logger = fopen(loggerFileName, "a+");
		fprintf(logger, "turnfisheyeoff\n");
		fclose(logger);
	}
	
	hw->FEye.TurnFishEyeOff();
}


void __stdcall settfactor(double TF){
	if(!hwInit) return;		//PCN2704 (Antony van Iersel, 25 July 2005) prevent acess from
	if(!hw->vidInit) return; // VB if not intialised
	if(loggerOn)
	{
		logger = fopen(loggerFileName, "a+");
		fprintf(logger, "settfactor recieved %f\n",TF);
		fclose(logger);
	}
	
	hw->FEye.SetTFactor(TF);
}

void __stdcall setfecentre(int x, int y){
	if(!hwInit) return;		//PCN2704 (Antony van Iersel, 25 July 2005) prevent acess from
	if(!hw->vidInit) return; // VB if not intialised
	if(loggerOn)
	{
		logger = fopen(loggerFileName, "a+");
		fprintf(logger, "setfecentre recieved x = %i, y = %i\n",x,y);
		fclose(logger);
	}
	
	hw->FEye.SetOffsets(x,y);
}

void __stdcall fisheyeison(int &FishEyeStatus){
	if(!hwInit) {FishEyeStatus = 0;return ;}		//PCN2704 (Antony van Iersel, 25 July 2005) prevent acess from
	if(!hw->vidInit) {FishEyeStatus = 0; return ;} // VB if not intialised
	if(loggerOn)
	{
		logger = fopen(loggerFileName, "a+");
		fprintf(logger, "fisheyeison sentback ");
		fclose(logger);
	}
	
	FishEyeStatus = hw->FEye.FishEyeStatus();

	if(loggerOn)
	{
		logger = fopen(loggerFileName, "a+");
		fprintf(logger, "%i as int\n",FishEyeStatus);
		fclose(logger);
	}
	
}

void __stdcall setimagesize(void){
	if(!hwInit) return;		//PCN2704 (Antony van Iersel, 25 July 2005) prevent acess from
	if(!hw->vidInit) return; // VB if not intialised
	if(loggerOn)
	{
		logger = fopen(loggerFileName, "a+");
		fprintf(logger, "setimagesize\n");
		fclose(logger);
	}
	
	hw->FEye.SetImageSize();
}

void __stdcall getscalevalue(double *Scale){
	if(!hwInit) {
		*Scale = 1;
		return;		//PCN2704 (Antony van Iersel, 25 July 2005) prevent acess from
	}

	if(!hw->vidInit) {
		*Scale = 1;
		return; // VB if not intialised 
	}
	if(loggerOn)
	{
		logger = fopen(loggerFileName, "a+");
		fprintf(logger, "getscalevalue sentback ");
		fclose(logger);
	}
	
	*Scale = hw->FEye.GetScale();

	if(loggerOn)
	{
		logger = fopen(loggerFileName, "a+");
		fprintf(logger, "%f as double\n",*Scale);
		fclose(logger);
	}
}

void __stdcall setscalevalue(double S){
	if(!hwInit) return;		//PCN2704 (Antony van Iersel, 25 July 2005) prevent acess from
	if(!hw->vidInit) return; // VB if not intialised
	if(loggerOn)
	{
		logger = fopen(loggerFileName, "a+");
		fprintf(logger, "setscalevalue recieved %f as double\n", S);
		fclose(logger);
	}
	
	hw->FEye.SetScale(S);
}

void __stdcall getimagesize(int &height, int &width){
	if(!hwInit) return;		//PCN2704 (Antony van Iersel, 25 July 2005) prevent acess from
	if(!hw->vidInit) return; // VB if not intialised
	if(loggerOn)
	{
		logger = fopen(loggerFileName, "a+");
		fprintf(logger, "getimagesize sentback ");
		fclose(logger);
	}
	
	height = hw->FEye.ImageHeight;
	width  = hw->FEye.ImageWidth;
	if(loggerOn)
	{
		logger = fopen(loggerFileName, "a+");
		fprintf(logger, "height = %i, width = %i as ints\n",height,width);
		fclose(logger);
	}
	
}

void __stdcall createmask(void){
	if(!hwInit) return;		//PCN2704 (Antony van Iersel, 25 July 2005) prevent acess from
	if(!hw->vidInit) return; // VB if not intialised
	if(loggerOn)
	{
		logger = fopen(loggerFileName, "a+");
		fprintf(logger, "createmask\n");
		fclose(logger);
	}
	
	hw->FEye.CreateMask();
}

void __stdcall calculatescale(void){
	if(!hwInit) return;		//PCN2704 (Antony van Iersel, 25 July 2005) prevent acess from
	if(!hw->vidInit) return; // VB if not intialised
	if(loggerOn)
	{
		logger = fopen(loggerFileName, "a+");
		fprintf(logger, "calculatescale\n");
		fclose(logger);
	}
	
	hw->FEye.SetDisplayScale();
}

void __stdcall setoriginalsize(int width, int height){
	if(!hwInit) return;		//PCN2704 (Antony van Iersel, 25 July 2005) prevent acess from
	if(!hw->vidInit) return; // VB if not intialised
	if(loggerOn)
	{
		logger = fopen(loggerFileName, "a+");
		fprintf(logger, "setoriginalsize recieved width = %i height = %i as ints\n",width,height);
		fclose(logger);
	}
	
	hw->FEye.SetOriginalSize(width,height);
}

void __stdcall livefisheye(int status){
	if(!hwInit) return;		//PCN2704 (Antony van Iersel, 25 July 2005) prevent acess from
	if(!hw->vidInit) return; // VB if not intialised
	if(loggerOn)
	{
		logger = fopen(loggerFileName, "a+");
		fprintf(logger, "livefisheye recieved status = %i as int\n",status);
		fclose(logger);
	}
	
	hw->LiveFishEye = status == 1 ? true : false;
}

void __stdcall hough_processimageonoff(bool onoff)
{
	if(!hwInit) return;
	if(!hw->vidInit) return; // VB if not intialised
	if(loggerOn)
	{
		logger = fopen(loggerFileName, "a+");
		fprintf(logger, "houghprocessimageonoff recieved onoff = %i as bool\n",onoff);
		fclose(logger);
	}
	
	if(onoff!=0) onoff = true;
	else onoff = false;
	hw->processingOn = onoff;	
}

void __stdcall transformoneimage(void){
	if(!hwInit) return;		//PCN2704 (Antony van Iersel, 25 July 2005) prevent acess from
	if(!hw->vidInit) return; // VB if not intialised
	if(loggerOn)
	{
		logger = fopen(loggerFileName, "a+");
		fprintf(logger, "transformoneimage\n");
		fclose(logger);
	}
	
	hw->FEye.Transform(hw->im);
}

void __stdcall hough_getcapturedevices(HWND hComboBox)
{
	HRESULT hr;
	ICreateDevEnum			*pDevEnum = 0;
	IEnumMoniker			*pEnum = 0;
	IMoniker				*pMoniker = 0;

	VARIANT varName;
	WCHAR friendlyname[180];

	int numberCaptureDevices=0;

	CoCreateInstance(CLSID_SystemDeviceEnum,
						  NULL,
						  CLSCTX_INPROC_SERVER,
						  IID_ICreateDevEnum,
						  (LPVOID*) &pDevEnum);

	hr = pDevEnum->CreateClassEnumerator(CLSID_VideoInputDeviceCategory,
										 &pEnum,
										 0);
	USES_CONVERSION;
	(long) SendMessage(hComboBox, 
					   CB_ADDSTRING, 
					   0,
					   (LPARAM) OLE2T(L"None"));
	if(hr!=S_OK)
	{
		if(pDevEnum!=0) pDevEnum->Release();
		if(pEnum!=0) pEnum->Release();
		return;
	}
	

	while (pEnum->Next(1, &pMoniker, NULL) == S_OK)
	{
		IPropertyBag *pPropBag;


		hr = pMoniker->BindToStorage(0,
									 0,
									 IID_IPropertyBag,
									 (void**) (&pPropBag));
		if (FAILED(hr))
		{
			pMoniker->Release();
			continue;
		}

		VariantInit(&varName);
		hr = pPropBag->Read(L"FriendlyName", &varName, 0);
		if(SUCCEEDED(hr)) 
		{
			wcscpy(friendlyname,varName.bstrVal);
			VariantClear(&varName);
			USES_CONVERSION;
			(long) SendMessage(hComboBox, 
							   CB_ADDSTRING, 
							   0,
							   (LPARAM) OLE2T(friendlyname));
		}
		pPropBag->Release();
		pMoniker->Release();
		numberCaptureDevices++;
	}
	if(numberCaptureDevices==0)
	{
		USES_CONVERSION;
		(long) SendMessage(hComboBox, 
						   CB_ADDSTRING, 
						   0,
						   (LPARAM) OLE2T(L"None"));
	}

	if(pDevEnum!=0) pDevEnum->Release();
	if(pEnum!=0) pEnum->Release();
}

void __stdcall hough_anycapturedevices(int &AnyCapture)
{
	HRESULT hr;
	ICreateDevEnum			*pDevEnum = 0;
	IEnumMoniker			*pEnum = 0;
	IMoniker				*pMoniker = 0;

	int numberCaptureDevices=0;

	CoCreateInstance(CLSID_SystemDeviceEnum,
						  NULL,
						  CLSCTX_INPROC_SERVER,
						  IID_ICreateDevEnum,
						  (LPVOID*) &pDevEnum);
	hr = pDevEnum->CreateClassEnumerator(CLSID_VideoInputDeviceCategory,
										 &pEnum,
										 0);
	if(hr!=S_OK)
	{
		if(pDevEnum!=0) pDevEnum->Release();
		if(pEnum!=0) pEnum->Release();
		AnyCapture = 0; return;
	}

	while (pEnum->Next(1, &pMoniker, NULL) == S_OK)
	{
		IPropertyBag *pPropBag;
		hr = pMoniker->BindToStorage(0,
									 0,
									 IID_IPropertyBag,
									 (void**) (&pPropBag));
		if (FAILED(hr))
		{
			pMoniker->Release();
			continue;
		}

		pPropBag->Release();
		pMoniker->Release();
		numberCaptureDevices++;
	}
	if(pDevEnum!=0) pDevEnum->Release();
	if(pEnum!=0) pEnum->Release();
	if(numberCaptureDevices==0) {AnyCapture = 0; return;}
	else {AnyCapture = 1; return;}
}

void __stdcall hough_checkforIPD(int &IPD)
{
	if(!hwInit)			{IPD = 0;return;}
	if(!hw->vidInit)	{IPD = 0;return;}
	if(hw->IPD==NULL)	{IPD = 0;return;}

	if(loggerOn)
	{
		logger = fopen(loggerFileName, "a+");
		fprintf(logger, "hough_checkforIPD sentback ");
		fclose(logger);
	}
	IPD = 1;

	if(loggerOn)
	{
		logger = fopen(loggerFileName, "a+");
		fprintf(logger, "%i \n",IPD);
		fclose(logger);
	}
	IPD = 1;
	return;
}

void _stdcall Hough_ProcessSingleImage(unsigned char *vbImage, int width, int height)
{
	if(!hwInit) return;
	if(hw->mediaType!=1) return; // VB if not intialised //PCN3194
	
	if(loggerOn)
	{
		logger = fopen(loggerFileName, "a+");
		fprintf(logger , "Hough_ProcessSingleImage\n");
		fclose(logger);
	}

	int i;
	int j;
	hw->framecb(0,0);

	for ( i = 0 ; i < hw->movie->width ; i++ ) 
		for ( j = 0 ; j < hw->movie->height ; j++ )
		{
			vbImage[(i + j * hw->movie->width) * 3] =     (char) hw->im[j][i].blue;  // blue;
			vbImage[(i + j * hw->movie->width) * 3 + 1] = (char) hw->im[j][i].green; // green;
			vbImage[(i + j * hw->movie->width) * 3 + 2] = (char) hw->im[j][i].red;   // red;
		}
}

void _stdcall Hough_InitialiseSingleImage(unsigned char *vbImage, int width, int height)
{
	if(!hwInit) return;
	if(hw->mediaType!=1) return; // VB if not intialised //PCN3194

	if(loggerOn)
	{
		logger = fopen(loggerFileName, "a+");
		fprintf(logger , "Hough_InitialiseSingleImage\n");
		fclose(logger);
	}

	hw->InitialiseSingleImage(vbImage, width, height);
}

void __stdcall hough_debugslider1(int value)
{
	if(loggerOn)
	{
		logger = fopen(loggerFileName, "a+");
		fprintf(logger , "hough_debugslider1\n");
		fclose(logger);
	}
	hw->blobBrightness=value;
}

void __stdcall hough_debugslider2(int value)
{
		if(loggerOn)
	{
		logger = fopen(loggerFileName, "a+");
		fprintf(logger , "hough_debugslider2\n");
		fclose(logger);
	}
}

void __stdcall hough_debugslider3(int value)
{
	if(!hwInit) return;
	if(!hw->vidInit) return;
	if(hw->radialProfile==NULL) return;

	if(loggerOn)
	{
		logger = fopen(loggerFileName, "a+");
		fprintf(logger , "hough_debugslider3\n");
		fclose(logger);
	}

	hw->FEye.SetYPosition(value);
	hw->FEye.CreateMask();
}

void __stdcall hough_debugslider4(int value)
{
		if(loggerOn)
	{
		logger = fopen(loggerFileName, "a+");
		fprintf(logger , "hough_debugslider4\n");
		fclose(logger);
	}
}

void __stdcall hough_debugcoordxy1(int x, int y)
{
		if(loggerOn)
	{
		logger = fopen(loggerFileName, "a+");
		fprintf(logger , "hough_debugcoordxy1\n");
		fclose(logger);
	}
}

void __stdcall hough_debugcoordxy2(int x, int y)
{
		if(loggerOn)
	{
		logger = fopen(loggerFileName, "a+");
		fprintf(logger , "hough_debugcoordxy2\n");
		fclose(logger);
	}
}

void __stdcall hough_debugcoordxy3(int x, int y)
{
		if(loggerOn)
	{
		logger = fopen(loggerFileName, "a+");
		fprintf(logger , "hough_debugcoordxy3\n");
		fclose(logger);
	}
}

void __stdcall hough_debugcoordxy4(int x, int y)
{
		if(loggerOn)
	{
		logger = fopen(loggerFileName, "a+");
		fprintf(logger , "hough_debugcoordxy4\n");
		fclose(logger);
	}
}

void __stdcall hough_debugcoordxy5(int x, int y)
{	if(loggerOn)
	{
		logger = fopen(loggerFileName, "a+");
		fprintf(logger , "hough_debugcoordxy5\n");
		fclose(logger);
	}
}

void __stdcall hough_debugbutton1(void)
{	if(loggerOn)
	{
		logger = fopen(loggerFileName, "a+");
		fprintf(logger , "hough_debugbutton1\n");
		fclose(logger);
	}
}

void __stdcall hough_debugbutton3(void)
{	if(loggerOn)
	{
		logger = fopen(loggerFileName, "a+");
		fprintf(logger , "hough_debugbutton3\n");
		fclose(logger);
	}
}

void __stdcall hough_SetYFishScale(double d)
{
	if(!hwInit) return;
	if(!hw->vidInit) return;
	if(hw->radialProfile==NULL) return;

	if(loggerOn)
	{
		logger = fopen(loggerFileName, "a+");
		fprintf(logger , "hough_SetYFishScale recieved %f as double\n",d);
		fclose(logger);
	}

	if (d<=0) d = 1;
	hw->FEye.SetYScale(d);
//	hw->FEye.SetDisplayScale();
//	hw->FEye.CreateMask();
}

void __stdcall hough_GetYFishScale(double &FishScale)
{
	if(!hwInit)					{FishScale = 1;return;}
	if(!hw->vidInit)			{FishScale = 1;return;}
	if(hw->radialProfile==NULL) {FishScale = 1;return;}

		if(loggerOn)
	{
		logger = fopen(loggerFileName, "a+");
		fprintf(logger , "hough_GetYFishScale sent back ");
		fclose(logger);
	}
	FishScale = hw->FEye.GetYScale();
	if(loggerOn)
	{
		logger = fopen(loggerFileName, "a+");
		fprintf(logger , "%f as double\n",FishScale);
		fclose(logger);
	}
}

void __stdcall hough_SetColourAdjust(double red, double green, double blue)
{
	if(!hwInit) return;
	if(!hw->vidInit) return;
	if(hw->mediaType==1) return; // VB if not intialised //PCN3194
	
	if(loggerOn)
	{
		logger = fopen(loggerFileName, "a+");
		fprintf(logger , "hough_SetColourAdjust recieved red = %f, green = %f, blue = %f as doubles\n",red,green,blue);
		fclose(logger);
	}

	
	hw->radialProfile->SetColourAdjust(red,green,blue);
	//hw->blobSizeThresh = (int) (red*red)*10;

}

void __stdcall hough_IsVideoRunning(int &running)
{
	if(!hwInit) {running=0; return;}// Video is not running
	if(!hw->vidInit) {running = 0; return;}

	if(loggerOn)
	{
		logger = fopen(loggerFileName, "a+");
		fprintf(logger , "hough_IsVideoRunning sentback ");
		fclose(logger);
	}

	if(hw->movie->IsVideoRunning()==false) running = 0;
	else running = 1;

	if(loggerOn)
	{
		logger = fopen(loggerFileName, "a+");
		fprintf(logger , "%i\n",running);
		fclose(logger);
	}
}

void __stdcall hough_centreOff(void)
{
	if(!hwInit) return;
	if(!hw->vidInit) return;
	if(hw->mediaType==1) return; // VB if not intialised //PCN3194

	hw->radialProfile->TurnCentreOff();
	
}

void __stdcall hough_centreOn(void)
{
	if(!hwInit) return;
	if(!hw->vidInit) return;
	if(hw->mediaType==1) return; // VB if not intialised //PCN3194

	hw->radialProfile->TurnCentreOn();

}

void __stdcall hough_setlaserpoint(int _x,int _y)
{
//	hw->radialProfile->centreLaser.SetLaserCentreCoord(_x,_y/hw->Y_ADJUSTMENT);
}

void __stdcall hough_lockdonut(double diameter)
{
	if(!hwInit) return;
	if(!hw->vidInit) return;
	if(hw->mediaType==1) return; // VB if not intialised //PCN3194

	if(diameter < 1) {hw->radialProfile->UnlockDonut();return;}
	
	hw->radialProfile->SetExpectedDiameter(diameter);
	hw->radialProfile->LockDonut();

}

void __stdcall hough_unlockdonut(void)
{
	if(!hwInit) return;
	if(!hw->vidInit) return;
	if(hw->mediaType==1) return; // VB if not intialised //PCN3194

	hw->radialProfile->UnlockDonut();

}

void __stdcall houge_AdjustContrastBright(double rgbScaler, double brightness)
{
	unsigned char lookup[256];
	double b;
	int i;


	float contrast =  (float) (((100 + rgbScaler) / 100) * ((100 + rgbScaler) / 100));
	for(i=0;i<256;i++)
	{

		b = i+brightness;
		b /= 256;
		b -=0.5;
		b *=contrast;
		b +=0.5;
		b *=256;
		if(b<0) b=0;
		if(b>255) b = 255;
		lookup[i]=(unsigned char) b;
	}
	hw->radialProfile->SetContrastBrightness(lookup);

}


