



//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// laserprofiler 24 November 2003
// Created By: Louise Shrimpton
// 
// Description:	This is where all the functionality of the image processing is
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
#include "CBSAlgebra.h"
#include "video.h"
#include "laserprofiler.h"
#include <time.h>
#include <dshow.h>

extern bool amSeeking;


int testGrid[20][2] =  {{40,34},{38,46},{38,58},{38,71}, // PCN2488 (AVI 6 Jan 2004)
						{50,33},{50,44},{50,57},{49,70}, // Setup for testing the Circle
						{62,32},{62,44},{61,56},{61,69}, // Grid scanning.
						{74,30},{73,43},{73,56},{73,68},
						{84,29},{86,42},{86,55},{86,68}}; 
int **resultGrid;	// PCN2488 (AVI 6 Jan 2004)






// PCN2320 -------------------v

//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Name: wait, signal, create 24 November 2003
// Created By: Louise Shrimpton
// 
// Description:	The functions are typical C functions for mutual exclusion.
//		To create a semaphore, call create.
//		To start mutual exclusion, call wait function.
//		To finish mutual exclusion, call signal. 
//
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
void Laserprofiler::wait(semaphore h) {
	WaitForSingleObject( h, MAXLONG);
}

void Laserprofiler::signal(semaphore h) {
	ReleaseSemaphore(h,1,NULL);
}

semaphore Laserprofiler::create(int v) {
	return CreateSemaphore(NULL,(long)v, MAXLONG, NULL);
}




//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Name: Accessing the im array 24 November 2003
// Created By: Louise Shrimpton
// 
// Description:	These functions allow access to the im array in laserprofiler. 
//   They are necessary for bounds checking the array and stop the "random crashing
//		problem".
//
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
int Laserprofiler::getblueim(int x, int y){
	if(x>=0 && x<movie->width && y>=0 && y<movie->height){
		return im[y][x].blue;
	}
	//Msg(TEXT("Accessing image out of bounds!!!"));
	return 0;
}

int Laserprofiler::getgreenim(int x, int y){
	if(x>=0 && x<movie->width && y>=0 && y<movie->height){
		return im[y][x].green;
	}
	//Msg(TEXT("Accessing image out of bounds!!!"));
	return 0;
}

int Laserprofiler::getredim(int x, int y){
	if(x>=0 && x<movie->width && y>=0 && y<movie->height){
		return im[y][x].red;
	}
	//Msg(TEXT("Accessing image out of bounds!!!"));
	return 0;
}

void Laserprofiler::setim(int x, int y, int r, int g, int b){
	if(x>=0 && x<movie->width && y>=0 && y<movie->height){
		im[y][x].red = (unsigned char) r;
		im[y][x].green = (unsigned char) g;
		im[y][x].blue = (unsigned char) b;
	}
	//Msg(TEXT("Setting image out of bounds!!!"));
}

void Laserprofiler::setredim(int x, int y, int r){
	if(x>=0 && x<movie->width && y>=0 && y<movie->height){
		im[y][x].red = (unsigned char) r;
	}
	//Msg(TEXT("Setting image out of bounds!!!"));
}

void Laserprofiler::setblueim(int x, int y, int b){
	if(x>=0 && x<movie->width && y>=0 && y<movie->height){
		im[y][x].blue = (unsigned char) b;
	}
	//Msg(TEXT("Setting image out of bounds!!!"));
}

void Laserprofiler::setgreenim(int x, int y, int g){
	if(x>=0 && x<movie->width && y>=0 && y<movie->height){
		im[y][x].green = (unsigned char) g;
	}
	//Msg(TEXT("Setting image out of bounds!!!"));
}

void Laserprofiler::ProssesImage(void)
{
//	overlayProfile=1;

	int i;
	
//	int centreScreenX = movie->width/2;
//	int centreScreenY = movie->height/2;

	/// Copy perfect circle //
//	for(i=0;i<1800;i++) 
//			{
//			setim((int) (perfectCircle200[i].x)+0.5+centreScreenX,(int) (perfectCircle200[i].y/Y_ADJUSTMENT)+0.5+centreScreenY,255,255,255);
//			setim((int) (perfectCircle201[i].x)+0.5+centreScreenX,(int) (perfectCircle201[i].y/Y_ADJUSTMENT)+0.5+centreScreenY,255,255,255);
//			}


//	if(tickCount!=NULL) tickCount->SetCounterPointer(im, movie->width, movie->height); //PCN3258
//	if(tickCount!=NULL) tickCount->Tick(); //PCN2639 Antony van Iersel (24 Feb 2004) //PCN3258


	radialProfile->Process(); // ratio 3:4


	// If fisheye is not being displayed then copy data for display //
	// before profile points or fisheyed							//
	if((FEye.FishEyeStatus() == OFF || LiveFishEye == false)  && overlayProfile!=0)			//
		for(i=0;i<PROFILE_SIZE;i++)	showProfileOverlay[i]=radialProfile->finalProfile[i].coordinate;
	//////////////////////////////////////////////////////////////////

	//Convert to none 3:4 ratio //////////////////////////////////////
	for(i=0;i<PROFILE_SIZE;i++)										//
	{																//
		radialProfile->finalProfile[i].coordinate.y/=Y_ADJUSTMENT;	//
	}																//
	//////////////////////////////////////////////////////////////////


	if(FEye.FishEyeStatus() == ON) 
		{
		for(i=0;i<PROFILE_SIZE;i++)
			{
			if(radialProfile->finalProfile[i].coordinate!=0)
				{
				FEye.ConvertPoint(radialProfile->finalProfile[i].coordinate);
				}
			}

		}
	
	//Convert to 3:4 ratio ///////////////////////////////////////////
	for(i=0;i<PROFILE_SIZE;i++)										//
	{																//
		radialProfile->finalProfile[i].coordinate.y*=Y_ADJUSTMENT;	//
//		if(radialProfile->finalProfile[i].coordinate!=0)
//			radialProfile->finalProfile[i].coordinate =
//				radialProfile->finalProfile[i].coordinate -
//				radialProfile->GetCentre();
	}																//
	//////////////////////////////////////////////////////////////////

	// If fisheye is not being displayed then copy data for display //
	// before profile points or fisheyed							//
	if((FEye.FishEyeStatus() == ON && LiveFishEye == true)  && overlayProfile!=0)			//
		for(i=0;i<PROFILE_SIZE;i++)	showProfileOverlay[i]=radialProfile->finalProfile[i].coordinate;
	//////////////////////////////////////////////////////////////////
	

	radialProfile->AdjustCentreFinalProfile(); // And removed final rough points;



	//PCN3233 IndexOffset added, if there is a level control then the profile is rotated,
	//the index will point to the most 6ocklock profile point and make this index 0
	//int indexOffset = (PROFILE_SIZE-radialProfile->GetMostVertical())+(PROFILE_SIZE/4);
	int indexOffset = 0;
	for(i=0;i<PROFILE_SIZE;i++) 
		{
		if(radialProfile->finalProfile[i].coordinate!=0)
			{
			profile[2][(i+indexOffset)%PROFILE_SIZE]=(float) radialProfile->finalProfile[i].coordinate.x;
			profile[3][(i+indexOffset)%PROFILE_SIZE]=(float) radialProfile->finalProfile[i].coordinate.y;
			}
		else
			{
			profile[2][(i+indexOffset)%PROFILE_SIZE]=0;
			profile[3][(i+indexOffset)%PROFILE_SIZE]=0;
			}
		}
	
	if (FEye.FishEyeStatus() == ON && LiveFishEye == true){
		FEye.Transform(im);
		FEye.CopyToVideo(im,movie->width,movie->height);
	}
	//PCN3013 adjust centre that is passed to VB thru fish eye if need

	centre = radialProfile->GetCentre();
//	if(FEye.FishEyeStatus() == ON) 
//	{
//		FEye.ConvertPoint(centre);
//	}

//	}
	///////////////////////////////////////////////////////////////////
	
	if(overlayProfile!=0) ShowProfileOverlayXY();
	ShowTextBox();
	ShowEgnoreMask();
//	ShowCounterBox();
//	if(tickCount!=NULL) tickCount->DrawSquare(tickCount->sxLeft,tickCount->sxRight);
	display(); // draw cross on image
//	radialProfile->ShowDrawLines();
//	radialProfile->ShowPutPixels();


			// Only for displaying purposes //////////////////
}


//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Name: Laserprofiler::framecb 24 November 2003
// Created By: Louise Shrimpton
// 
// Description:	This is the main function for the Image Processing.
//	Every frame captured from the DirectX is passed to this function for processing.
//	This is called by the samplegrabber for every frame
//
// Input: time is the time of the frame in the video
//		  s is a pointer to the pixel information for the frame of the video
//
// Output: none
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
void Laserprofiler::framecb(double time,unsigned char *s) {
	int j,i;
	double timeTemp;
	timeTemp = time;

	if(!amSeeking) time = (double) movie->gettime() / 10000000;
	if(timeTemp!=time)
	{
		__asm nop;
	}
	int perc; 
	double radius;
	
	if(mediaType!=1) //PCN3194
	{
		if(!vidInit) return;
		if(vidInit==0xdddddddd) return;


		//copy the frame data into a usable form (an array of pixels)

		for(j=0;j<movie->height;j++) 
			{// make an index	
			im[j]=(pixel *)s+j*movie->width;

		}

		//Only when movie snapeshot (grab) is not bussy
			
		// add entries in the array of video data past top and bottom;
		//this is a buffer
//PCN????
//		for(j=-16;j<0;j++) {
//			im[j]=im[0];
//			im[(movie->height-1)-j]=im[movie->height-1];

		//FilterNoise();
		//}
	}
	else
	{
	//	memcpy(stillImage,im,(movie->width*3)*movie->height);
		for ( i = 0 ; i < movie->width ; i++ ) 
			for ( j = 0 ; j < movie->height ; j++ )
			{
				im[j][i].blue	=  (unsigned char) stillImage[(i + j * movie->width) * 3];       // blue;
				im[j][i].green  =  (unsigned char) stillImage[(i + j * movie->width) * 3 + 1];   // green
				im[j][i].red    =  (unsigned char) stillImage[(i + j * movie->width) * 3 + 2];   // red
			}
	}
	

	if(processingOn)
	{
		
		wait(profilelock);  //only allow one frame to be accessing the arrays at once
		ProssesImage();
		centre=radialProfile->GetCentre();
		radius=radialProfile->GetAverageRadius(1);
	//Start of mutual exclusion

		if(movie->recordprofileinfo == true && movie->lastRecordedTime<time){
			movie->lastRecordedTime=time;
			profbuff.clear(in);

			//this doesn't go entirely with design.  Laserprofile ris a friend of prof and profilebuffer
			//to allow this, but because of tiem constraints, this is the quickest method.
			//I know it works.  LS
			memcpy(profbuff.buffer[in].profile,profile,sizeof(profile));

			// PCN2888 centre.x, y, radius replace "max.pos.x, y, max.radius" 'PCN3219 made centres float instead of int
			
			//PCN3744 (Antony van Iersel, 22 September 2005, 6:31 pm, third late night in a row :( ) 
			if(IPD!=NULL)			 profbuff.setprofilebuffer((int)in,frameno,time,(float) centre.x, (float) centre.y,(float) radius,(int) IPD->GetDistance(time*1000)); 
			//else if(tickCount!=NULL) profbuff.setprofilebuffer((int)in,frameno,time,(float) centre.x, (float) centre.y,(float) radius,tickCount->count*100); //PCN2639 (Antony van Iersel, 25 March 2004)
			else                     profbuff.setprofilebuffer((int)in,frameno,time,(float) centre.x, (float) centre.y,(float) radius,0);
			in=(in+1)%(profbuff.getbuffsize());
			frameno++; //PCN3289 Moved from below the movie->recordprofileinfo block to here
		}
		perc = ((numprofilepoints*100)/PROFILE_SIZE);
		totalperc = (totalperc+ perc);
		centerx = (float) centre.x; //PCN3219 - tracking centre, int now float for both
		centery = (float) centre.y; //PCN3219


		signal(profilelock);  //end of mutual exclusion
	}


//	timesForTesting[indexForTesting++]=currentTime;
//	if(indexForTesting>100)
//	{
//		for(int i=0;i<100;i++)
//			fprintf(fileForTesting,"%f\n",timesForTesting[i]);
//	}

}



//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Name: pxlrange 24 November 2003
// Created By: Louise Shrimpton
// 
// Description:	Make sure that i will always be between 0 and 255 (allowed values for a pixel)
// Input: i - the value to be tested
// Output:  returns the tested value
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
inline unsigned char pxlrange(int i) {
	if(i<0) return 0;
	if(i>255) return 255;
	return (unsigned char) i;
}


//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Name: getprofile 24 November 2003
// Created By: Louise Shrimpton
// 
// Description:	get the profile at [first][second] in the array
// Input: i - the position in profile
// Output:  returns true or false
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
float Laserprofiler::getprofile(int first, int second){  //access the profile array //PCN2891 was int
	if(first>=0 && first<4 && second>=0 && second < PROFILE_SIZE){ //PCN2891 was PROFILE_SIZE+10
		return profile[first][second];
	}
	return 0;
}

bool Laserprofiler::setprofile(int first, int second, float value){ //PCN2891 was int
	if(first>=0 && first<4 && second>=0 && second < PROFILE_SIZE){ //PCN2891 was PROFILE+10
		profile[first][second]= value;
		return true;
	}
	return false;
}

//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Name: GetT1T2 12 Feburaray 2004
// Created By: Antony van Iersel
// 
// Description:	Takes the Red Green and Blue inputs and result will either be single
//              colour or a mix

// Input: rgb 1, and rgb 2, mirror each other, its needed because the equation that
//		  uses this, uses a two pixel readings. eg 1 - 2
// Output: None, but the address is passed from t1 and t2, anything that changes there
//         is acutually changing the values of the function that calls this one.
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

void Laserprofiler::GetT1T2(int r1,int g1, int b1, int r2, int g2, int b2, int &t1, int &t2)
{
	if(videoFilterType==0)  // Red Filter
		{					// Removes all blue and green
		t1=r1;				// and returns only Red
		t2=r2;
		}
	if(videoFilterType==1) // Green Filter
		{				   // Same as above but returns Green
		t1=g1;
		t2=g2;
		}
	if(videoFilterType==2) // Blue Filter
		{				   // Same as above but returns Blue
		t1=b1;
		t2=b2;
		}
	if(videoFilterType==3) // Combination
		{
		t1=g1+b1;
		t2=g2+b2;
		}

	if(t1<0) t1=0; if(t1>255) t1=255;
	if(t2<0) t2=0; if(t2>255) t2=255;
}

//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Name: getXAdjust 24 November 2003
// Created By: Louise Shrimpton
// 
// Description:	Adjust the width and height of the frame to fit inside the ClearLine
// Profiler Screen.
// Input: The width and height of the video.  
// Output:  sets x and y to be the new height and width of the video
//			Also sets the xadjust variable in hw to true or false depending on whether
//			or not we've adjusted the width of the video.
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
void Laserprofiler::getXAdjust(int width, int height, double *x, double *y){
	if(height>width) xadjust = true;
	else xadjust = false;
	double ratio = 684.0/513.0;       //these are the dimensions of ClearLine Screen
//	double currratiox = (double)width/(double)height;
	double newheight = (double)width/ratio;
	*y = (double)newheight/(double)height;
//	double currratio = (double)height/(double)width;
	double newwidth = (double)height/ratio;
	*x = (double)newwidth/(double)width;
}

inline double Laserprofiler::DistOfTwoPoints(pt one, pt two)
{
	return sqrt(pow((double) (one.x-two.x),2)+pow((double) (one.y-two.y),2));
}

inline double Laserprofiler::DistOfTwoPointsFloat(double x1, double y1, double x2, double y2)
{
	return sqrt(pow(x1-x2,2)+pow(y1-y2,2));
}

//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Name: Laserprofiler::textarea 24 November 2003
// Created By: Louise Shrimpton 
// 
// Description:	Return true if (x,y) is a point in the image that shouldn't be looked at
//		as a possible profile point.  Checks to see if it is in the blocked out square (done
//		int VB). 
// Input: The x and y coordinates of the pixel in the image.
// Output:  Returns true if the pixel is in the critical area, and false otherwise.
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
bool Laserprofiler::textarea(int x,int y){
	if((x>=xb && x<=xt && y>=yb && y<=yt)){
		return true;
	}
	return false;
}

//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Name: Laserprofiler::display 25 November 2003
// Created By: Louise Shrimpton 
// 
// Description:	Draw a cross on the image
//				Uses the center and the radius of the pipe
//
// Input: none
// Output:  none
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
void Laserprofiler::display(void) {

	int i;

	double radius=radialProfile->GetAverageRadius(1);
	if(radius<5) return;

	vec2double centre;
	vec2int centreDisplay;
	vec2int left,right,up,down;
	
	centre=radialProfile->GetCentre();
	if(FEye.FishEyeStatus() == ON) centre.y=(centre.y/Y_ADJUSTMENT);
	else centre.y/=Y_ADJUSTMENT;


// Taking the cross away, making it a dot ///////////////////////////////////////////////
//
//	left.x=(int)  (centre.x-radius+0.5); left.y=(int)  (centre.y+0.5);
//	right.x=(int) (centre.x+radius+0.5); right.y=(int) (centre.y+0.5);
//	up.x=(int) (centre.x+0.5); up.y=(int) (centre.y-(radius/Y_ADJUSTMENT)+0.5);
//	down.x=(int) (centre.x+0.5); down.y=(int) (centre.y+(radius/Y_ADJUSTMENT)+0.5);

	centreDisplay.x=(int) (centre.x+0.5);
	centreDisplay.y=(int) (centre.y+0.5);

// And here is out dot //
	setim(centreDisplay.x, centreDisplay.y, 255,0,0);

	//	int i;

	for(i=1;i<bullsEye[0].coord.x;i++)
	{
		setim(bullsEye[i].coord.x+(int) (centre.x+0.5), 
			  bullsEye[i].coord.y+(int) (centre.y+0.5), 
			  bullsEye[i].r,
			  bullsEye[i].g,
			  bullsEye[i].b);
	}
	
//	if((centreDisplay.x>movie->width) || (centreDisplay.x<0)) return;
//	if((centreDisplay.y>movie->height) || (centreDisplay.y<0)) return;
//
//	for(x=left.x;x<right.x;x++)  setim(x,centreDisplay.y,255,0,0);
//	for(y=up.y;y<down.y;y++) setim(centreDisplay.x,y,255,0,0);
//		
///////////////////////////////////////////////////////////////////////////////////////		
	
}
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Name: Laserprofiler::showrect 26 November 2003
// Created By: Louise Shrimpton 
// 
// Description:	 When called, this will show the rectangle in the video that
//		should contain the text.  This will block it out and stop the area
//		from becoming pipe profile points.
//				
// Input: none
// Output:  none
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
void Laserprofiler::showrect(){
	if(showcol == 1) showcol = 0;
	else showcol = 1;
}


void Laserprofiler::ShowProfileOverlayXY(void)
{
	int i;
	double x,y;
	int xd, yd;
	vec2double coord;

	if(overlayProfile!=2) return;
	for(i=0;i<PROFILE_SIZE;i++)
		{
		if(showProfileOverlay[i]==0) continue;

		x = (int) ( showProfileOverlay[i].x );
		y = (int) ( showProfileOverlay[i].y );

		if(overlayProfile==2)
		{
			x-=centre.x; //Return points to oregin zero;
			y-=centre.y;  //Return points to oregin zero;
			coord = GetVector(vec2double(x,y));
			coord.y*=1.1;
			coord = GetCoordinate(coord);
			x=coord.x+centre.x;
			y=coord.y+centre.y;
		}

		xd=(int) (x+0.5);
		yd=(int) ((y/Y_ADJUSTMENT)+0.5);



		if(!radialProfile->IsWaterLevelOn() || !radialProfile->IsInWaterSection(i))
			{
			//////////////////////////////////////////////////////////////////////
			// PCN2608 (Antony van Iersel, 13 May 2004) Displayed Profile Point //
			 setim(xd,yd,0,255,0);		// no longer blue but green with a blue     //
			setim(xd-1,yd,0,0,255);	// boarder. Bounds checking done in setim   //
			setim(xd+1,yd,0,0,255);	//////////////////////////////////////////////
			setim(xd,yd+1,0,0,255);	//
			setim(xd,yd-1,0,0,255);	//
			setim(xd-1,yd+1,0,0,255);	//
			setim(xd-1,yd-1,0,0,255);	//
			setim(xd+1,yd+1,0,0,255);	//
			setim(xd+1,yd-1,0,0,255);	//
			//////////////////////////
			}
		
		if(radialProfile->IsWaterLevelOn() && radialProfile->IsInWaterSection(i))
			{
			//////////////////////////////////////////////////////////////////////
			// PCN2608 (Antony van Iersel, 13 May 2004) Displayed Profile Point //
			setim(xd,yd,0,0,255);		// no longer blue but green with a blue     //
			setim(xd-1,yd,0,255,0);	// boarder. Bounds checking done in setim   //
			setim(xd+1,yd,0,255,0);	//////////////////////////////////////////////
			setim(xd,yd+1,0,255,0);	//
			setim(xd,yd-1,0,255,0);	//
			setim(xd-1,yd+1,0,255,0);	//
			setim(xd-1,yd-1,0,255,0);	//
			setim(xd+1,yd+1,0,255,0);	//
			setim(xd+1,yd-1,0,255,0);	//
			//////////////////////////
			}
		}	
}

//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Name: Laserprofiler::ShowTextBox 27 May 2004
//       PCN2847
// Created By: Antony van Iersel 
// 
// Description:	Used to show where the General Text box is on the Video
//				
// Input: none
// Output:  none
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
void Laserprofiler::ShowTextBox(void)
{
	int x,y;
	for(x=xb;x<xt;x++)
		for(y=yb;y<yt;y++)
			setblueim(x,y,180);
}

//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Name: Laserprofiler::ShowEgnoreMask 19 April 2006
//       PCN2847
// Created By: Antony van Iersel 
// 
// Description:	Used to show where the egnore mask is on the video
//				
// Input: none
// Output:  none
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
void Laserprofiler::ShowEgnoreMask(void)
{
	int x,y;
	for(x=0;x<movie->width;x++)
		for(y=0;y<movie->height;y++)
			if(radialProfile->egnoreMask[y][x]>0) setblueim(x,y,180);
			
}


//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Name: Laserprofiler::ShowCounterBox 28 May 2004
//       PCN2847
// Created By: Antony van Iersel 
// 
// Description:	Used to show where the Counter Mask is on the Video
//				
// Input: none
// Output:  none
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
void Laserprofiler::ShowCounterBox(void)
{
	return;
	/*
	if(tickCount==NULL) return;
	int x,y;
	
	for(x=tickCount->xLeft;x<tickCount->xRight;x++)
		for(y=tickCount->yTop;y<tickCount->yLower;y++)
			setgreenim(x,y,120);
	*/
}


//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Name: Laserprofiler::BlankVideo 11 Feb 2004
// Created By: Antony van Iersel 
// 
// Description:	Used to make video black, used as to see more clearly overlays if
// needed, with out using it the original video with multiple overlays
//				
// Input: none
// Output:  none
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

void Laserprofiler::BlankVideo(void)
{
	for(int y=0;y<movie->height;y++)		//PCN3121 was 4 pixel boarder around the Blank  
		for(int x=0;x<movie->width;x++)     // video. eg +4 and -4 on limits
			setim(x,y,0,0,0);
}

//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Name: Laserprofiler::GrayVideo 11 Feb 2004
// Created By: Antony van Iersel 
// 
// Description:	Used to make video gray (Traditional Black and White), 
//				used to see what the image inhancement looks like before,
//				this is what the profiler would see to profile.
//				
// Input: none
// Output:  none
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

void Laserprofiler::GrayVideo(void)
{
	int t1,t2; //t2 is just a dummy, two are needed for profiler calculations,
			   //here we are only using one
	for(int y=0;y<movie->height;y++)	//PCN3121 removed boarder 
		for(int x=0;x<movie->width;x++) // from video.
			{
			GetT1T2(getredim(x,y),getgreenim(x,y),getblueim(x,y),0,0,0,t1,t2);
			setim(x,y,t1,t1,t1);
			}
}


//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Name: Laserprofiler::Laserprofiler 25 November 2003
// Created By: Louise Shrimpton 
// 
// Description:	Constructor for the laserprofiler.  Initializes all variables.  
//				
// Input: none
// Output:  none
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Laserprofiler::Laserprofiler(char *fileName) { //PCN3744 Filename added to check for IPD file

//	xvals = 0;
//	yvals = 0;
//	rvals = 0;
	IPD = new IPDInterface(fileName); //PCN3744 IPD interface
	if(IPD->Initialise()==-1) {delete IPD; IPD=NULL;}
	int i;
	currentFrameGrabeTime=0;
	mediaType = 0;
	im=0;			//PCN3194
	stillImage=0;	//PCN3194


// PCN3005
//	indexForTesting = 0;
//	fileForTesting=fopen("c:\\forTesting.txt","w");

	LiveFishEye = false;

	tickCount=NULL; //PCN3085 15 Oct 2004
	bullsEye=NULL;
	BuildBullsEye();


	
//	for(i=0;i<500;i++) timesForTesting[i]=0;

	//// Creating a perfect cirlce ////
	
	double deg=0;
	double degStep = (2*PI)/1800;

	for(i=0;i<1800;i++)
		{
		perfectCircle200[i].x=sin(deg)*200; perfectCircle200[i].y=cos(deg)*200;
		perfectCircle201[i].x=sin(deg)*201; perfectCircle201[i].y=cos(deg)*201;
		deg+=degStep;
		}


	//PCN2668 store how long it take to process image to adjust video time in VB accordingly
	timeToProcessImage=0; 
	
	profilelock=create(1);
	vidInit = false;	
	movie=new Video();
	vidInit = true;



	prevCx=0; prevCy=0;
//	theta1=0.0;
//	theta2=0.0;
	//PCN1939 wlLeft , wlRight. What profile points to egnore for water level.
	wlLeft=1; wlRight=PROFILE_SIZE-1;


	//Initialize all previous global variables
	deviceType = 0;	// PCN2289
	processingOn = true;
	GRAD_THRESHOLD=15;
	frameno=0;
	centerx = 0;
	centery = 0; 
	centre=0;
	LoosingInfo = false;
	yshowgreen=0;  
	xshowgreen=0;
	showprof =0;
	showcol = 1;
	showwat = 0;
	numprofilepoints = 0;
	totalperc = 0;
	optimized = 0;
	xadjust=false; 
	Y_ADJUSTMENT=1.0;
	X_ADJUSTMENT=1.0;
	xb=0;xt=0;yb=0;yt=0;  
	ignoretheta1=0; ignoretheta2=0;  
	outsidetheta=false;
	lightpipe = true; 
	SD_X=1.5;
	SD_Y=1.1;
	X_THRESHOLD=1*65536;
	Y_THRESHOLD=1*65536;
	in=0;
	out=0;

//	ignoreprofile = false;  
//	ignorecenter = false;  
	profileWaterLevel = false;
	ignoreWaterLevel = false;
	overlayProfile = 0;	// 0 is no show profile, 1 is overlay profile at 100% 2 is 105%
	videomode = 0;		// 0 to show overlays with video in background, 1 Gray Video, 2 Black Backgound
	videoFilterType = 2;// 0 Red, 1 Green, 2 Blue, 3 Mix
	profilerMethod = 0; // 0 for standard strength edge profile, 1 for threshold profile. 
	profilerThreshold = 197; // What brightness the Laser has to be to be a candidate.
	

	//tickCount = new Counter(); //PCN2848 Antony 28 May 2004
	//tickCount->SetCounterMask(0,0,0,0);	//PCN2848 Antony 28 May 2004

	radialProfile = new RadialScan(); //PCN2888
	radialProfile->FEye = &FEye;
	//resultGrid=fe->calGrid.ScanGrid((int **) testGrid,20); // PCN2488 (Antony van Iersel, 6 Jan 2003);
	blobSizeThresh = 0;
	blobBrightness = 20;
}

Laserprofiler::~Laserprofiler(void)
{

	int j;

//	Msg("Deleting Laserprofiler Class");
//	fclose(fileForTesting);
//	if(xvals!=NULL) { delete[] xvals; xvals=NULL; } //PCN3085 for the following [] was added to
//	if(yvals!=NULL) { delete[] yvals; yvals=NULL; }//remove all of the array.
//	if(rvals!=NULL) { delete[] rvals; rvals=NULL; }//.........................................
	for(j=0;j<origVideoHeight;j++)
	if(stillImage!=NULL) { delete[] stillImage; stillImage=NULL; }
	if(IPD!=NULL) {delete IPD; IPD=NULL;}
	
	if(vidInit) vidInit=false; 

	if(movie!=NULL) { delete movie; movie=NULL; }
	if(radialProfile!=NULL) { delete radialProfile; radialProfile=NULL; }
//PCN3085 remove page one from memory. Memory Leak.

	//if(tickCount!=NULL) { delete tickCount; tickCount=NULL;}//PCN3085 remove for memory leaks.
	if(bullsEye!=NULL) {delete[] bullsEye; bullsEye=NULL;};
//	if(FilterBlob!=NULL) {delete[] FilterBlob; FilterBlob=NULL;};
	
//	for(j=0;j<movie->height;j++) if(FilterLooked[j]!=NULL) {delete[] FilterLooked[j]; FilterLooked[j]=NULL;}
//	if(FilterLooked!=NULL) {delete[] FilterLooked; FilterLooked = NULL;}

	

	if(mediaType==1) 
	for(int i=0;i<movie->width;i++)
		delete[] im[i];
	
	delete[] im; im=0;
//	if(im!=0) 
//	{
////		for(i=0;i<imHeight;i++)
////			{
////			delete[] im[i];
////			}
//		delete[] im; im=0;
//	} //PCN????
	
//	DumpUnfreed();

}

//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// PCN2781
// Name: Laserprofiler::InitialiseIM
// Date: 2 December 2004
// Created By: Antony van Iersel
// 
// Description:	The initialisation of IM is moved from Laserprofer contructor to here,
//				the problem with having it in the constuctor is that the profiler didn't
//				know the image information yet when it tried to initialise it.
//				Now this is moved to after the video information is gathered.
//				
// Input: none
// Output: none
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

void Laserprofiler::InitialiseIM(void)
{
	if(!vidInit) return;


	int i,j;
	origVideoHeight=movie->height;
	origVideoWidth=movie->width;
//	xvals=new int[movie->width];
//	yvals=new int[movie->height];
//	rvals=new int[movie->width+movie->height];

	// give 32 pixels at top at bottom
//	im=new pixel *[movie->height+64];
//	for(i=0;i<movie->height+64;i++)
//		im[i] = new pixel[movie->width];
//	im+=32;
	
	im=new pixel *[movie->height];

	if(mediaType==1)
	for(i=0;i<movie->height;i++)
		im[i] = new pixel[movie->width];
	
	//delete[] im; im=0;	

	imHeight= movie->height;

	radialProfile->egnoreMaskHeight = movie->height;
	radialProfile->egnoreMask = new unsigned char *[movie->height];

	//FilterLooked = new unsigned char *[movie->height];
	for(i=0;i<movie->height;i++) 
	{
		radialProfile->egnoreMask[i] = new unsigned char[movie->width];
//		FilterLooked[i] = new unsigned char[movie->width];
	}

	for(i=0;i<movie->height;i++)
		for(j=0;j<movie->width;j++)
			radialProfile->egnoreMask[i][j]=0;

//	FilterBlob = new vec2int[(movie->height*movie->width)+1];


}

//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// PCN2781
// Name: Laserprofiler::InitialiseIM
// Date: 2 December 2004
// Created By: Antony van Iersel
// 
// Description:	The initialisation of IM is moved from Laserprofer contructor to here,
//				the problem with having it in the constuctor is that the profiler didn't
//				know the image information yet when it tried to initialise it.
//				Now this is moved to after the video information is gathered.
//				
// Input: none
// Output: none
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

void Laserprofiler::InitialiseSingleImage(unsigned char *vbImagePointer, int width, int height) //PCN3194
{
	long vbImageSize;

	vbImageSize = movie->height*(movie->width*3); // x3 is for each width point there is three colours
	stillImage = new unsigned char [vbImageSize];
	memcpy(stillImage,vbImagePointer,vbImageSize);
}


//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Name: Laserprofiler::getvariables 26 November 2003
// Created By: Louise Shrimpton 
// 
// Description:	This function can be called outside the laserprofiler, to get
//	the image processing variables.  The main purpose of this is to allow the user in VB
//	to see what the parameters for processing are.
//				
// Input: none
// Output:  XT, YT, GT, SDX, SDY, greenx, greeny, prof, col, percprofpnts, totalper, xadj
//			These variables are detailed in the Laserprofiler class declaration
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
void Laserprofiler::getvariables(double *XT, double *YT, int *GT, double *SDX, double *SDY, int *greenx, int *greeny, int *prof, int *col, int *percprofpnts, double *totalper, double *xadj){
	*XT = (int)X_THRESHOLD;  
	*YT = (int)Y_THRESHOLD;
	*GT = GRAD_THRESHOLD;
	*SDX = (double)SD_X;
	*SDY = (double)SD_Y;
	*greenx = xshowgreen;
	*greeny = yshowgreen;
	*prof = showprof;
	*col = 0;
	*percprofpnts = (int)((numprofilepoints*100)/PROFILE_SIZE);
	if(frameno==0) *totalper = 0;
	else *totalper = (float)totalperc/frameno;
	if(xadjust)  *xadj = (double)Y_ADJUSTMENT;
	else *xadj = (double)X_ADJUSTMENT;
}


//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Name: Laserprofiler::setvariables 26 November 2003
// Created By: Louise Shrimpton 
// 
// Description:	This function can be called outside the laserprofiler, to set
//	the image processing variables.  The main purpose of this is to allow the user in VB
//	to set the parameters for processing.
//				
// Input: XT, YT, GT, SDX, SDY, greenx, greeny, prof, col, percprofpnts, totalper, xadj
//			These variables are detailed in the Laserprofiler class declaration
// Output:  none
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
void Laserprofiler::setvariables(double XT, double YT, int GT, double SDX, double SDY, int greenx, int greeny, int prof, double xadj){
	if (XT != -1) X_THRESHOLD = (float)XT ;
	if (YT != -1) Y_THRESHOLD = (float)YT ;
	if (GT != -1) GRAD_THRESHOLD = GT;
	if (SDX != -1) SD_X= SDX;
	if (SDY != -1) SD_Y= SDY;
	if (prof == 1) {showprof = 1; radialProfile->SetShowProfileCandidatesOverlay(true);}
	else {showprof = 0; radialProfile->SetShowProfileCandidatesOverlay(false);}
	if(greenx == 1) xshowgreen = 1;
	else xshowgreen = 0;
	if(greeny==1) yshowgreen =1;
	else yshowgreen = 0;
}


//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Name: Laserprofiler::LPemptybuffer 26 November 2003
// Created By: Louise Shrimpton 
// 
// Description:	Empties the frame buffer of all the data that is in it.
//		The function doesn't actually remove the data, it just sets the index in the 
//		array to be 0.  This will mean any frame data found will write over the top 
//		of the current data in the array.  (The standard method of clearing arrays
//		in C)
//				
// Input: none
// Output:  none
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
void Laserprofiler::LPemptybuffer(){
	in = 0;
	out = 0;
}



//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Name: Laserprofiler::adjustsettings 26 November 2003
// Created By: Louise Shrimpton 
// 
// Description:	Adjusts the Standard deviation for the processing according to the 
//		size of the image.
//				
// Input: height and width of the image
// Output:  none
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
void Laserprofiler::adjustsettings(double width, double height){
	//adjusts the settings of the image processing depending on the height and width 
	//of the media
	//for width = 768;height=576;, use approximately SD_X = 0.75, SD_Y = 0.55
	//for width=300;height = 240;, use approximately SD_X=1.5; SD_Y=1.1;
	SD_X = 1.5;// / ((width/300.0) * 0.8); 
	SD_Y = 1.1;// / ((height/240.0) * 0.8);
}

inline vec2double Laserprofiler::GetCoordinate(vec2double vector)
{
	return vec2double(sin(vector.x)*vector.y,cos(vector.x)*vector.y);
}

vec2double Laserprofiler::GetVector(vec2double coord)
{
	double adj;
	double opp;

	double x=coord.x;
	double y=coord.y;

	double dist = sqrt((x*x) + (y*y)); 
	
	adj = fabs(coord.x);
	opp = fabs(coord.y);
	
	if((x==0) && (y==0)) return 0;
	if((x==0) && (y>0)) return 0;
	if((x==0) && (y<0)) return vec2double(PI				  ,dist);
	if((x>0) && (y==0)) return vec2double((PI/2)			  ,dist);
	if((x<0) && (y==0)) return vec2double(PI+(PI/2)		      ,dist);
	if((x>0) && (y>0))  return vec2double(       atan(adj/opp),dist); // + , +
	if((x>0) && (y<0))  return vec2double(PI-    atan(adj/opp),dist); // + , -
	if((x<0) && (y<0))  return vec2double(PI+    atan(adj/opp),dist); // - , -
	if((x<0) && (y>0))  return vec2double((2*PI)-atan(adj/opp),dist); // - , +
	return 0;
}

void Laserprofiler::ShowBullsEye(void)
{
//	int i;
//	for(i=1;i<bullsEye[0].x;i++)
//	{
//		setim(bullsEye[i].x+cx, bullsEye[i].y+cy, 255,0,0);
//	}
}


void Laserprofiler::BuildBullsEye(void)
{

	int i;
	int numberOfSteps = (360*10);
	double step=2*PI/numberOfSteps;
	vec2int coord;
	pixelXYRGB tempBullsEye[1000];
	double raid=0; //Fly buster :)
	tempBullsEye[0].coord.x=0;

	while(raid<(2*PI))
	{
		coord.x=(int) ((sin(raid)*5));
		coord.y=(int) ((cos(raid)*5));
		if(!IsInArray(&tempBullsEye[0], coord))
		{
			tempBullsEye[0].coord.x++;
			tempBullsEye[tempBullsEye[0].coord.x].coord=coord;
			tempBullsEye[tempBullsEye[0].coord.x].r=255;
			tempBullsEye[tempBullsEye[0].coord.x].g=0;
			tempBullsEye[tempBullsEye[0].coord.x].b=0;
		}
	raid+=step;
	}

for(i=-10;i<11;i++)
	{
		coord.x=i; coord.y=0;
		if(!IsInArray(&tempBullsEye[0], coord))
		{ 
			tempBullsEye[0].coord.x++; 
			tempBullsEye[tempBullsEye[0].coord.x].coord=coord; 
			tempBullsEye[tempBullsEye[0].coord.x].r=255;
			tempBullsEye[tempBullsEye[0].coord.x].g=0;
			tempBullsEye[tempBullsEye[0].coord.x].b=0;
		}
		coord.x=0; coord.y=i;
		if(!IsInArray(&tempBullsEye[0], coord))
		{ 
			tempBullsEye[0].coord.x++; 
			tempBullsEye[tempBullsEye[0].coord.x].coord=coord;
			tempBullsEye[tempBullsEye[0].coord.x].r=255;
			tempBullsEye[tempBullsEye[0].coord.x].g=0;
			tempBullsEye[tempBullsEye[0].coord.x].b=0;
		}
	}

	raid=0;
	while(raid<(2*PI))
	{
		coord.x=(int) ((sin(raid)*5.5));
		coord.y=(int) ((cos(raid)*5.5));
		if(!IsInArray(&tempBullsEye[0], coord))
		{
			tempBullsEye[0].coord.x++;
			tempBullsEye[tempBullsEye[0].coord.x].coord=coord;
			tempBullsEye[tempBullsEye[0].coord.x].r=90;
			tempBullsEye[tempBullsEye[0].coord.x].g=50;
			tempBullsEye[tempBullsEye[0].coord.x].b=50;
		}
	raid+=step;
	}

	raid=0;
	while(raid<(2*PI))
	{
		coord.x=(int) ((sin(raid)*4.5));
		coord.y=(int) ((cos(raid)*4.5));
		if(!IsInArray(&tempBullsEye[0], coord))
		{
			tempBullsEye[0].coord.x++;
			tempBullsEye[tempBullsEye[0].coord.x].coord=coord;
			tempBullsEye[tempBullsEye[0].coord.x].r=90;
			tempBullsEye[tempBullsEye[0].coord.x].g=50;
			tempBullsEye[tempBullsEye[0].coord.x].b=50;
		}
	raid+=step;
	}

	bullsEye = new pixelXYRGB[tempBullsEye[0].coord.x+2];
	for(i=0;i<tempBullsEye[0].coord.x+1;i++)
	{
		bullsEye[i]=tempBullsEye[i];
	}

}

bool Laserprofiler::IsInArray(pixelXYRGB *array, vec2int point)
{
	int i;
	if(array==NULL) return false;
	if(array[0].coord.x==0) return false;
	for(i=1;i<=array[0].coord.x;i++)
	{
		if(array[i].coord==point) return true;
	}
	return false;
}

int Laserprofiler::GetIPDDistance(void)
{
	LONGLONG t;
	double currentTime;
	
	t = movie->gettime();
	currentTime =(double) (t/10000); //PCN3251
	if(IPD!=NULL) return IPD->GetDistance(currentTime);
	return 0;
}

void Laserprofiler::FilterNoise()
{
	char threshold=100;
	int i,j;
	int BlobSize;


	threshold = blobBrightness;

	for ( i = 0 ; i < movie->width ; i++ ) 
		for ( j = 0 ; j < movie->height ; j++ )
		{
			FilterLooked[j][i] = 0;
		}

	for ( i = 0 ; i < movie->width ; i++)
		for(j = 0; j < movie->height ; j++)
		{
			BlobSize=FindBlob(j,i,0,threshold);
			CountCalls = 0;
			if((BlobSize <= blobSizeThresh) && (BlobSize > 0)) 
			{
				ClearBlob(BlobSize);
			}
			
			if(BlobSize> blobSizeThresh) MarkBlob(BlobSize);

			FilterLooked[j][i] = 1;
		}


}

int	Laserprofiler::FindBlob(int y, int x, int count,int threshold)
{	
	
	CountCalls++;
	if(y>=movie->height) {CountCalls--;return 0;}
	if(x>=movie->width) {CountCalls--;return 0;}
	if(x<0 || y<0) {CountCalls--;return 0;}
	if(FilterLooked[y][x] == 1) 
	{
		CountCalls--;return 0;
	}
	if(count>movie->width*movie->height) 
	{
		CountCalls--; return 0;
	}
	
	FilterLooked[y][x] = 1;
	if(CountCalls>30000) return 0;
	if(im[y][x].blue<threshold) {CountCalls-- ;return 0;}

	count++;
	FilterBlob[count].x=x;
	FilterBlob[count].y=y;
	
	count += FindBlob(y+1, x  , 0, threshold);
	count += FindBlob(y+1, x+1, 0, threshold);
	count += FindBlob(y+1, x-1, 0, threshold);

	count += FindBlob(y	 , x+1, 0, threshold);
	count += FindBlob(y	 , x-1, 0, threshold);
	count += FindBlob(y-1, x-1, 0, threshold);
	count += FindBlob(y-1, x,   0, threshold);
	count += FindBlob(y-1, x+1, 0, threshold);

	CountCalls--;
	return count;

	
}

void Laserprofiler::ClearBlob(int size)
{
	int count;

	if(size>movie->height*movie->width) return;
	for(count = 1;count<=size;count++)
	{
		im[FilterBlob[count].y][FilterBlob[count].x].red = 0;
		im[FilterBlob[count].y][FilterBlob[count].x].green = 0;
		im[FilterBlob[count].y][FilterBlob[count].x].blue = 0;
	}
}

void Laserprofiler::MarkBlob(int size)
{
	int count;
	if(size>movie->height*movie->width) return;
	


	for(count = 1;count<=size;count++)
	{
		
		im[FilterBlob[count].y][FilterBlob[count].x].red = 255;
		im[FilterBlob[count].y][FilterBlob[count].x].green = 0;
		im[FilterBlob[count].y][FilterBlob[count].x].blue = 0;
	}
}



