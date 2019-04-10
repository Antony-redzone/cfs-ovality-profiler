//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Profile 24 November 2003
// Created By: Louise Shrimpton
// 
// Description:	Storage for the profile information
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

#ifndef profileclass
#define profileclass

#include <windows.h>
#include <stdio.h>
class Laserprofiler;  //forward declaration to allow compilation


//This is the number of profile points that make up a circle.  At the moment, there is a profile
//point every 2 degrees, making the number of profile points 180.
const int PROFILE_SIZE = 180;

const int BUFFSIZE = 180*500;

//Each profile has a 3D array of profile points, a frame number, a time, 
//a center x and y coordinate and a radius
class prof {
public:
	prof(void);
	~prof(void);

	//get a particular position in the profile array
	bool getprofile(int first, int second, float *data); //PCN2891 was int
	
	//set a particular position in the profile array
	bool setprofile(int first, int second, float data); //PCN2891 was int

	//clear the profile array
	void clear();
	
	//the frame no
	int frame;

	//the time of the frame (in the video)
	double time;

	//center coordinates of the pipe and the radius
	float x,y,r; //PCN3219 need to track centres for adjustable water level so now make it float instead of int

	//distance buffer
	int distance; // PCN2639 25 (March 2004 , Antony van Iersel)

	friend Laserprofiler;
//private to enforce use of bounds checking get and set functions
private:
	//the array of profile points
	//profile[0]... is the radius
	//profile[1]... is the colour information (not implemented)
	//profile[2]... x coordinate (Was  (don't know - never used))
	//profile[3]... y coordinate

	//float profile[4][PROFILE_SIZE];  //0 is radius
	float *profile;
};




class profilebuffer{
public:
	profilebuffer();
	~profilebuffer();
	bool setprofilebuffer(int i, int frame, double time, float x, float y, float r, int d);
	bool getprofilebuffer(int i, int *frame=NULL, double *time=NULL, float *x=NULL, float *y=NULL, float data[]=NULL, int *distance=NULL);
	float getprofilebufferprofile(int i, int first, int second);
	int getprofileframeno(int i);
	bool getprofilebufferprofiledata(int i, int *data, int numdata);
	void clear(LONGLONG frame);
	int getbuffsize(){ return BUFFSIZE;}

	friend Laserprofiler;
private:
	prof *buffer; //PCN2779 not to be static, needs to be loaded dynamically
};

#endif