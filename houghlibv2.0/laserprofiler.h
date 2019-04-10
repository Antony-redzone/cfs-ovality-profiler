//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Class Name: Laserprofiler 24 November 2003
// Created By: Louise Shrimpton
// 
// Description:	
//
// Functionality:
//
// Inherited From:  None
//
// Other Objects Contained Within this Class: 
//	1) Video
//	2) Fisheye
//	
//
// Basic Behviours:
// 
// 
//
// 
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


#ifndef laserprofiler
#define laserprofiler

#include "profile.h"
#include <windows.h>
#include <math.h>
#include <fstream>
#include <stdio.h>
#include "video.h"
#include "profile.h"
#include <atlbase.h> // Fish-EYE added for feMsg method PCN2290 
#include "Counter.h"
#include "IPDInterface.h" //PCN3744, Adding IPDInterface
#include "RadialProcess.h"
#include "FishEyeTransform.h"


//Mutual exclusion is used for accessing the array of profile information
#define semaphore HANDLE
const int MAXSHOWXY=1000000;
const int MAXSHOWWATER=10000;


const int capturewidth=320;		// PCN2289 this is suppose to be default capture resolution
const int captureheight=240;	// PCN2289 this is suppose to be default capture resolution
const int captureframerate=15;

const int THRESH=0;

//This is used for finding the center of the circle.  Three edges points are taken, and the 
//center is found.  This is done N times, and the center is found by taking the x and y
//value that occurs the most.  The higher N is, the more accurate the center of the 
//circle is...  (in theory)
const int CIRCLESAMPLES=10000;

//This is the minimum number of profile points that can be found before it is considered
//a circle.  If the number of edge points found is less than this number, the red
//cross won't be drawn, and some other things won't happen.
const int MINCIRCLEPTS=10;

//This is the maximum number of edges that can fit into the array.  For larger images,
//this needs to be more.
const int MAXEDGES=100000;

struct pt {
	long x, y; //PCN2891 Increase accuracy, now long. (Antony, 18 June 2004) (Just in case)
	double strength;
};

struct houghpt {
	pt pos;
	long r; //PCN2891 Increase accuracy , now long. (Antony, 18 June 2004)
	long v; //PCN2891 Increase accuracy , now long. (Antony, 18 June 2004)
	int xcoord; //PCN2891 store x and y coordinate instead of a vector (Antony, 21 June 2004)
	int ycoord; //PCN2891 store x and y coordinate instead of a vector (Antony, 21 June 2004)
};

typedef struct _pixelXYRGB pixelXYRGB;
struct _pixelXYRGB
{
	vec2int coord;
	unsigned char r;
	unsigned char g;
	unsigned char b;
};

typedef unsigned uint;

class Laserprofiler {
public:
	Video *movie;			// for avi file
	IPDInterface *IPD; //PCN3744 adding IPDInterface
	int mediaType; //PCN????
	vec2double perfectCircle200[1800];
	vec2double perfectCircle201[1800];
	vec2double centre;

	double currentFrameGrabeTime;

	double timeToProcessImage;
	int origVideoHeight; //PCN2778 the video height and width change after gim is allocated 
	int origVideoWidth;  //        memory, so when it is deleted not enought is deleted  
	
	LensTransform FEye;
	bool LiveFishEye;

	int vidInit;
	Counter *tickCount;		// PCN2639 (24 March 2004, Antony van Iersel)
	RadialScan *radialProfile;
	bool xadjust;			//true if we need to adjust the width
	double Y_ADJUSTMENT;	//the adjustment of the file height to make it fit in the Clearline Profiler screen
	double X_ADJUSTMENT;    //the adjustment of the file width to make it fit in the Clearline Profiler screen
	int xb,xt,yb,yt;		//for the area not to look in, x bottom, x top, y bottom, y top
	int optimized;			//used to make the video play faster.  Only processed every other frame
	bool outsidetheta;
	float centerx; //PCN3219, change both from int to float
	float centery; //
	int deviceType;			// PCN2289
	bool LoosingInfo;
	int frameno; //PCN3289 not to be used (3 Feb 2005) but is needed because its part of a vairiable passed onprofile dump
	bool lightpipe;
	LONGLONG in;			//these are the positions of the frames in the buffer (profile buffer)
	LONGLONG out;
	
	//bool ignoreprofile    //PCN2568 removed replaced wiht profileWaterLevel;
	//bool ignorecenter;	//PCN2568 removed replaced with ignoreWaterLevel (Antony van Iersel, 11 May 2004)
	int ignoreWaterLevel;	//PCN2568 ignores water level in centre calculations
	int profileWaterLevel;	//PCN2568 when water level is on profile the water.
	semaphore profilelock;	//to allow mutual exclusion on the profile array

	//functions
	Laserprofiler(char *fileName);
	~Laserprofiler(void);

	void ProssesImage(void);

	int getredim(int x, int y);
	int getgreenim(int x, int y);
	int getblueim(int x, int y);
	void setim(int x, int y, int r, int g, int b);
	void setredim(int x, int y, int r);
	void setblueim(int x, int y, int b);
	void setgreenim(int x, int y, int g);

	// PCN2904
	void SwapFishEyePages(void);


	void framecb(double time,unsigned char *s);
//	void setwaterlevelthetas(double t1,double t2); removed, now set in VBlaserprofile (11 May 2004, Antony)
	void getframe(int i); // get frame i from the file
	void getprofile(); // get the profile
	void showwaterlevel(int i);
	void adjustsettings(double width, double height);
	void getXAdjust(int width, int height, double *x, double *y);
	float getprofile(int first, int second);  //access the profile array //PCN2891 was return int
	bool setprofile(int first, int second, float value); //PCN2891 was int
	void display(void); // put image into window
	void showrect();
	int isLight();
	void getvariables(double *XT, double *YT, int *GT, double *SDX, double *SDY, int *greenx, int *greeny, int *prof, int *col, int *percprofpnts, double *totalper, double *xadj);
	void setvariables(double XT, double YT, int GT, double SDX, double SDY, int greenx, int greeny, int prof, double xadj);
	void LPemptybuffer();
	void wait(semaphore h);
	void signal(semaphore h);
	void InitialiseIM(void);
	void InitialiseSingleImage(unsigned char *vbImagePointer, int width, int height); //PCN3194

	double DistOfTwoPoints(pt one, pt two);
	double DistOfTwoPointsFloat(double x1, double y1, double x2, double y2);
	int FindBlob(int y,int x, int count, int threshold);
	void ClearBlob(int size);
	void MarkBlob(int size);
	
	
	semaphore create(int v);
	profilebuffer profbuff;
	friend Video;
	int overlayProfile; // If 0, off, 1 profile size, 2 1.05% for profile
	int videomode; // If 0, then true overlay, if 1 Gray display, if 2 black out video;
	int videoFilterType;
	int profilerMethod; // If 0 normal profile, if 1 then threshold.
	int yshowgreen;  //make this constant!
	int xshowgreen;
	int showprof;
	int profilerThreshold;
	int GRAD_THRESHOLD;
	double SD_X;
	double SD_Y;
	double	X_THRESHOLD; 
	double	Y_THRESHOLD;
	pt	showX[MAXSHOWXY];
	pt	showY[MAXSHOWXY];
	pt	showWater[MAXSHOWWATER];

	int numSX;
	int numSY;
	int numCand;
	int numSW; //PCN1939 Number of points in showWater buffer to display;
	pixel **im; // for the image index PCN2639 made public to access  the image pointer
	unsigned char **FilterLooked;

	vec2int *FilterBlob;
	int blobBrightness;
	int blobSizeThresh;
	int CountCalls;
	
								//we dont want it to be profiled. (19 April 2006)
				// 23 Feb 2004
	unsigned char *stillImage; //PCN3194 This is where the original loaded image from VB is copied as not to have to
						// load it everytime the image is reprocessed. When a new image is loaded this will be
						// updated. This is to stop any contimaition from video processing. (22 August 2005, Antony)




	int wlLeft, wlRight; //PCN1939 Antony van Iersel, 6 May 2004.
						 // Water Level Left, Water Level Right.
	bool processingOn;
	int GetIPDDistance(void); //PCN3744 IPD interface
	////// PCN3284 ///////////////////////////
//	FILE *fileForTesting;				//
//	double distanceForTesting[10000];	//
//	double heightForTesting[10000];		//
//	int indexForTesting;				//
//////////////////////////////////////////
private:

	
//	double timesForTesting[500];

	int imWidth;
	int imHeight;

	vec2double showProfileOverlay[PROFILE_SIZE];
	pixelXYRGB *bullsEye;

	//pixel **im; // for the image index

	BITMAPINFO *bmi; // for the image
	houghpt max; // max from hough array

	int prevCx, prevCy; //PCN2405 5 May 2004, Antony, keep a track of previous centre
						// Michelles Idea, think its a good one.
	// int *xvals,*yvals,*rvals; PNC????
//	double theta1, theta2;  //these are the angles for the ignore water level
	
	int nedges;
//	int ndefedges;
    
	//prof profilebuffer[BUFFSIZE];  
	void FilterNoise(void);
	void gradientimage(void); // put image gradient in blue plane
	void Threshold();
	void gradientimageint(void); // put image gradient in blue plane (ints  only)


	void GrayVideo(void);
	void BlankVideo(void);
	void ShowGreen(void);
	void ShowProfileOverlay(void);
	void ShowProfileOverlayXY(void);
	void ShowWaterOverlay(void); //PCN1939 Antony van Iersel (6 May 04)
	void ShowTextBox(void); //Display Text Box
	void ShowEgnoreMask(void);
	void ShowCounterBox(void);
	void movegrad(void);
	bool getedges(int i, pt *p);  //get from edges array
	bool setedges(int i, pt *p);
	bool getedges(int i, int *x, int *y, double *strength=NULL);
	bool setedges(int i, int x, int y, int strength=0);
	void setshowx(int x,int y);
	void setshowy(int x,int y);
	void SetShowWater(int x, int y);
	bool getdefedges(int i, pt *p);  //get from defedges array
	bool setdefedges(int i, pt *p);
	bool getdefedges(int i, int *x, int *y, double *strength=NULL);
	bool setdefedges(int i, int x, int y, int strength=0);
	void GetT1T2(int r1,int g1, int b1, int r2, int g2, int b2, int &t1, int &t2); 
	bool textarea(int x,int y);

	void colourtext();
	void TrackCounter(void);
	float profile[4][PROFILE_SIZE]; // holds profile [0] is radius [1] is grey level // Was PROFILE_SIZE+10
	void SetPosProfile(int x, int y, double r, int i);
	pt posProf[PROFILE_SIZE];
	pt edges[MAXEDGES];    //this is for the definate circle points (ones >XThres & >YThres)
	pt defedges[MAXEDGES];   //this is for all the points, edge >0
	vec2double GetVector(vec2double coord);
	inline vec2double GetCoordinate(vec2double vector);
	void ShowBullsEye(void);
	void BuildBullsEye(void);
	bool IsInArray(pixelXYRGB *array, vec2int point);
	


	int showcol;
	int showwat;
	int numprofilepoints;
	int totalperc;

	int ignoretheta1, ignoretheta2;  //for ignoring the water level
	HWND hwnd;            // main window
	HDC dc;



};

bool getprofilebuffer(int i, int *frame=NULL, double *time=NULL, int *x=NULL, int *y=NULL, int *r=NULL);
float getprofilebufferprofile(int i, int first, int second);
int getprofileframeno(int i);


#endif