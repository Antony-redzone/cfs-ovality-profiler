//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// profile.cpp 24 November 2003
// Created By: Louise Shrimpton
// 
// Description:	This class is designed to handle the profile buffer imnformation
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

#include "profile.h"
#include "laserprofiler.h"

prof::prof(void)
{
	profile = new float[4*PROFILE_SIZE];
}

prof::~prof(void)
{
	delete[] profile;
}
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Name: prof::clear 24 November 2003
// Created By: Louise Shrimpton
// 
// Description:	Clears a profile in the array of profiles.
// Input: None
// Output: None
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
void prof::clear(){
	memcpy(profile,profile,sizeof(profile));
}


//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Name: prof::getprofile 24 November 2003
// Created By: Louise Shrimpton
// 
// Description:	get the profile point at profile[first][second].  
//puts the profile point in data.
//the profile point is the distance from the center
// Input: the first and second position of the array of profiles, profilebuffer
// Output: returns true if first and second is within the bounds of the array,
//			returns false otherwise.  Also, make data point to the integer value
//			at this position in the array
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
bool prof::getprofile(int first, int second, float *data){
	//check that it's within bounds
	// PCN2891 was 3 now 4 , PROFILE_SIZE+10 now just PROFILE_SIZE  
	if(first>=0 &&first<4 && second>=0 && second<PROFILE_SIZE){
		//*data = profile[first][second];
		*data = profile[second+(first*PROFILE_SIZE)];
		return true;
	}
	*data = 0;
	return false;
}


//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Name: prof::setprofile 24 November 2003
// Created By: Louise Shrimpton
// 
// Description:	set the profile point at profile[first][second]
// Input: the first and second position of the array profilebuffer and the
//		   integer to put here.
// Output: A return value of true or false
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
bool prof::setprofile(int first, int second, float data){
	//check that it's within the array bounds
	// PCN2891 was 3 now 4 , PROFILE_SIZE+10 now just PROFILE_SIZE  
	if(first>=0 &&first<4 && second>=0 &&second<PROFILE_SIZE){
		//profile[first][second] = data;
		profile[second+(first*PROFILE_SIZE)] = data;
		return true;
	}
	return false;
}



profilebuffer::profilebuffer()
{
	buffer = new prof[BUFFSIZE]; //PCN2779, needed to be loaded dynamically, this caused
								 // the memory allocation to find a better spot for memory
								 // avoiding the VB out of memory error.
}

// PCN2779, because we now create the buffer dynamically, we have to delete when finnish
profilebuffer::~profilebuffer(void)
{
	if(buffer!=0) {delete[] buffer; buffer=0;}
}



//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Name: setprofilebuffer 24 November 2003
// Created By: Louise Shrimpton
// 
// Description:	set the profile information at position i in the buffer profilebuffer
// Input: i - the position in profilebuffer
//		  frame - the frame number to be set at i
//		  time - the time of the recorded frame i
//		  x,y - the center of the pipe at this point
//		  r - the radius of this frame
// Output: A return value of true or false
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
bool profilebuffer::setprofilebuffer(int i, int frame, double time, float x, float y, float r, int d){
	//check that it's within bounds
	if(i>=0 && i<=BUFFSIZE){
		buffer[i].frame= frame;
		buffer[i].time= time;
		buffer[i].x=x;
		buffer[i].y=y;
		buffer[i].r=r;
		buffer[i].distance=d;
		return true;
	}
	return false;
}



//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Name: getprofilebuffer 24 November 2003
// Created By: Louise Shrimpton
// 
// Description:	get the profile information at position i in the buffer profilebuffer
// Input: i - the position in profilebuffer
// Output: frame - the frame number to be set at i
//		  time - the time of the recorded frame i
//		  x,y - the center of the pipe at this point
//		  r - the radius of this frame
//        A return value of true or false
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
bool profilebuffer::getprofilebuffer(int i, int *frame, double *time, float *x, float *y, float *r, int *distance){
	if(i>=0 && i<=BUFFSIZE){
		if(frame!=NULL) *frame = buffer[i].frame;
		if(time!=NULL) *time = buffer[i].time;
		if(x!=NULL) *x = buffer[i].x;
		if(y!=NULL) *y = buffer[i].y;
		if(r!=NULL) *r = buffer[i].r;
		if(distance!=NULL) *distance = buffer[i].distance;
		return true;
	}
	if(frame!=NULL) *frame = 0;
	if(time!=NULL) *time = 0.0;
	if(x!=NULL) *x = 0;
	if(y!=NULL) *y = 0;
	if(r!=NULL) *r = 0;
	return false;
}


//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Name: getprofilebufferprofile 24 November 2003
// Created By: Louise Shrimpton
// 
// Description:	get the actual profile points at position i in the buffer profilebuffer
// Input: i - the position in profile
//		  first - the first parameter for the array
//		  second - the second parameter for the array
// Output:  returns true or false
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
float profilebuffer::getprofilebufferprofile(int i, int first, int second){
	float data;
	if(i>=0 && i<=BUFFSIZE){
		buffer[i].getprofile(first,second,&data);
		return data;
	}
	return 0;
}


//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Name: getprofileframeno 24 November 2003
// Created By: Louise Shrimpton
// 
// Description:	get frame number at the ith position in the profilebuffer array
// Input: i - the position in profile
// Output:  returns true or false
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
int profilebuffer::getprofileframeno(int i){
	if(i>=0 && i<=BUFFSIZE){
		return buffer[i].frame;
	}
	return 0;
}


//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Name: clear 24 November 2003
// Created By: Louise Shrimpton
// 
// Description:	clear one frame's data 
// Input: frame - the frame in the buffer
// Output:  none
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
void profilebuffer::clear(LONGLONG frame){
	buffer[frame].clear();
}