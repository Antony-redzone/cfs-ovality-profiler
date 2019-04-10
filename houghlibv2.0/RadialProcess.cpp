#include <math.h>
#include <stdio.h>

#include "Video.h"
#include "FishEyeTransform.h"
#include "RadialProcess.h"



#define PI 3.1415926535897932384626433832795

RadialScan::RadialScan(void)
{
	egnoreMask = 0; //PCN????
	CentreOff = false;
	DonutLocked = false;
	expectedDiameter = 0;

	redAdjust = 0;
	greenAdjust = 0;
	blueAdjust = 0;
	
	int loop;
	validPointer=false;
	
	imHeight = 0;
	imWidth = 0; 

	mask.offset.x = 0;
	mask.offset.y = 0;
	mask.rayLength = 0;
	mask.averageRadius = 0;
	pickupLevel = 50;
	profileType = 0;
	cutoffLevel = 0;
	laserWidth = 8;
	noLaserWidthOverlay=0;
	noProfileCandidatesOverlay=0;
	noPutPixels=0;

	showLaserWidthOverlay=false;
	showInternalProfilePoints=false;
	showProfileCandidatesOverlay=false;
	showInternalCircle=false;
	showPutPixels=false;
	showDrawLines=false;
	showVideoFilter=false;
//	showLaserCentreCoord=false;


	internalRadius=0.66;
	externalRadius=1.33;

	wlLeft = 0;
	wlRight = 0;

	hasPrinted=false;

	counterMaskTopLeft.x=counterMaskBottomRight.x=0;
	counterMaskTopLeft.y=counterMaskBottomRight.y=0;
	ignoreWaterLevel = false;

	debug1=1;
	debug2=2;
	numberCentreDebug=0;
	averageCentreHead=0;
		copyBuffer = 0;

	for(loop=0;loop<MASK_RES;loop++) mask.rays[loop].rayAtoms=NULL;
	for(loop=0;loop<PROFILE_SIZE;loop++) egnoreList[loop]=0; //PCN3219 new water level, profile points marked 0 for accept
														//and 1 to egnore, this is passed from VB. The user interface
														//to set this up is from vb.
	
	CreateGausianMask();
	for(int i=0;i<256;i++) contrastBrightness[i]=i;
}

RadialScan::~RadialScan(void)
{
	int loop;
	for(loop=0;loop<MASK_RES;loop++) delete[] mask.rays[loop].rayAtoms; //PCN3085 [] was added to the delete because its deleting an array not a single item
	
	for(loop=0;loop<imHeight;loop++) delete[] copyBuffer[loop];
	delete[] copyBuffer;

	if(egnoreMask!=0)
	{
		for(loop=0;loop<egnoreMaskHeight;loop++)
		{
			delete[] egnoreMask[loop];
		}
		delete[] egnoreMask;
		egnoreMask = 0;
	}

			


//????
// Dumps the running average centres to file on exit.
//	FILE *f;
//	f=fopen("C:\\centreHistory.txt","w");
//	for(loop=0;loop<numberCentreDebug;loop++)
//	{
//		fprintf(f,"%f , %f, %s\n",centreDebugHistory[loop].x,centreDebugHistory[loop].y,&centreDebugHistoryNotes[loop][0]);
//	}
//	fclose(f);

}

void RadialScan::Initialise(pixel **imVideoPointer, int width, int height)
{
	int i;

	imHeight = height;
	imWidth = width;
	imVideo = imVideoPointer;
	imRatio = ((double) width*0.75) / (double) height;

	previousCentre.x = mask.offset.x =  width / 2;
	previousCentre.y = mask.offset.y = (int) (((float) height / 2) * imRatio);

	filterValue = 1;
	for(i=0;i<5;i++) centreHistory[i]=mask.offset;
	averageCentreHead=0;

	BuildMask();

	copyBuffer =new pixel *[imHeight];
	for(i=0;i<imHeight;i++) copyBuffer[i] = new pixel[imWidth];

	validPointer = true;

}

void RadialScan::Process(void)
{
	if(!validPointer) return;
	int loop;

	noLaserWidthOverlay=0;
	noProfileCandidatesOverlay=0;
	noPutPixels=0;
	noDrawLines=0;

	mask.offset.x=(double) imWidth/2;				// Initial mask centre for x
	mask.offset.y=(double) (imHeight/2)*imRatio;	// and y


	mask.averageRadius=0;	// Average radius on first pass is 0
	
	
	//centreLaser.SearchForLaser();
	
	
	RecalculateAverageCentre();
	
	mask.offset=averageCentreHistory; // Start the next centre from running average running centres
	if(mask.offset == centreHistory[0] && (mask.offset != vec2double(imWidth/2, (imHeight/2) * imRatio))) 
	{
		mask.offset.x=(double) imWidth/2;				// Initial mask centre for x
		mask.offset.y=(double) (imHeight/2)*imRatio;	// and y
	}

	if(CentreOff) 
	{
//		centreLaser.SearchForLaser();
		//GrabData(1);
		//for(loop=0;loop<MASK_RES;loop++) ProcessRay(&mask.rays[loop], 3, loop, false);
		//RecalculateCoordinates(1);
		//AdjustCentre(1);


		RecalculateAverageCentre();
		mask.offset.x=(double) imWidth/2;				// Initial mask centre for x
		mask.offset.y=(double) (imHeight/2)*imRatio;	// and y		
		

//		mask.offset.x=centreLaser.GetLaserCentreCoord().x ;//+ (double) imWidth/2;				// Initial mask centre for x
//		mask.offset.y=(centreLaser.GetLaserCentreCoord().y-30)*(imRatio) ;//+ (double) (imHeight/2)*imRatio;	// and y
		
		internalRadius = 0.26;
		
		mask.averageRadius=150;

	}

	if(!CentreOff)
	{
		// First Pass ////////////////////////////////////////////////////////////////////
		GrabData(12);																	//
		for(loop=0;loop<MASK_RES;loop+=12) ProcessRay(&mask.rays[loop], 1, loop, false);//
		RecalculateCoordinates(12);														//
		AdjustCentre(1);//
		
		if(!DonutLocked) mask.averageRadius=GetAverageRadius(12);										//
		else mask.averageRadius = expectedDiameter/2;
		//////////////////////////////////////////////////////////////////////////////////

		// Secound pass, finding a better centre /////////////////////////////////////////
		GrabData(12);																	//
		for(loop=0;loop<MASK_RES;loop+=12) ProcessRay(&mask.rays[loop], 2, loop, false);//
		RecalculateCoordinates(12);														//
		
		AdjustCentre(2); //PCN3233
		
		//PCN4539
		if(!DonutLocked) mask.averageRadius=GetAverageRadius(12);										//
		else mask.averageRadius = expectedDiameter/2;
		/////////////////////////////////////////////////////////////////////////////////
	}		
	int sampleSize;

	if(redAdjust==0) sampleSize = 1;
	else sampleSize = 3;

	//GrabData(3);
	//for(loop=0;loop<MASK_RES;loop+=3) ProcessRay(&mask.rays[loop], 3, loop, false);//
	//RecalculateCoordinates(3);														//
		
	//AdjustCentreWithSmartDataFill();


	// Final Pass ////////////////////////////////////////////////////////////////////
	GrabData(1);																	//
	for(loop=0;loop<MASK_RES;loop+=sampleSize) 
	{
		ProcessRay(&mask.rays[loop], 3, loop, false);	//
	}
	RecalculateCoordinates(sampleSize);														//
	RemoveRoughPoints(); //PCN2993													//
	RecalculateCoordinates(sampleSize);														//
	//////////////////////////////////////////////////////////////////////////////////

	DownSample();
	RemoveFinalRoughPoints(); //ANT
	if(!CentreOff) 
	{
		if(!DonutLocked) mask.averageRadius=GetAverageRadius(1);										//
		else mask.averageRadius = expectedDiameter/2;		
	}

	
//	BlankScreen();
//	ShowMask();
	if(showVideoFilter) ShowVideoFilter();
	
	if(showLaserWidthOverlay) ShowLaserWidthOverlay();
	if(showProfileCandidatesOverlay) ShowProfileCandidatesOverlay();
	if(showInternalCircle) ShowInternalCircle(mask.offset, mask.averageRadius*internalRadius, 255,128,0,false);
	if(showInternalCircle) ShowInternalCircle(mask.offset, mask.averageRadius*externalRadius, 255,128,0,false);

	if(showDrawLines) ShowDrawLines();
	if(showVideoFilter && !showProfileCandidatesOverlay) ShowProfilePoints(1,0,160,0,false); // Show working profile points
	//if(true) ShowProfilePoints(1,0,160,0,false); // Show working profile points
	if(showPutPixels) ShowPutPixels();

	
//	if(showLaserCentreCoord) centreLaser.DisplaySearchBox(-1,100,-1);
}

//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// PCN3156
// Name: RadialProfile::BuildRadiansTable
// Created By: Antony van Iersel, (12 November 2004)
// Description:	Fills in a lookup table for the radians that are going to be used
//              to create the mask, once the 3:4 angles are created they are adjusted
//				to the video image ratio
// Input: Double point to the passed preCalcualtedRadians table
// Output: None
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
void RadialScan::BuildRadiansTable(double *preCalculatedRadians)
{
	int loop;
	double directionStep=(2*PI)/MASK_RES;
	vec2double coord;
	double radian;

	coord.toVector(); // Turn coord into a vector (angle, distance)
	for(loop=0;loop<MASK_RES;loop++)
	{
		radian = ((loop+1)*directionStep)+PI;
		coord.x=radian; coord.y=1000;// Fill with vector value;
		coord = coord.toCoordinate(); // Convert Back to X Y coord;
		coord.y/=imRatio; // Adjust to display ratio;
		coord = coord.toVector(); // Convert Back into a vector;
		preCalculatedRadians[loop]=coord.x; // Store the new angle into the table
	}


}

void RadialScan::BuildMask(void)
{
	double *preCalculatedRadians; //PCN3156 need to precalculate the ray angles
	

	int			loop; 
	int			atom;	// Current Ray atom to fill.
	vec2int		nextPixel; // What pixel offset that is going to be put in the ray.
	vec2int		lastPixel; // Last pixel is compared to next pixel, to make sure there is no coppies along a sing ray.
	vec2double	rawPixel;  // Pixel coordiante before it is rounded to next and last pixel
	double		radiusStep = 0.1;
	double		radian;	// Radian angle of the ray;
	
	double		directionStep=(2*PI)/MASK_RES;	// Radians to next profile point
	double		radius;	// Current radius of the ray atom.
	vec2double	coord;	// Coordinate of pixel after its been scaled to 2:3 ratio
	double		sinRadian;
	double		cosRadian;
	

	if((imHeight==0) || (imWidth==0)) return; // If there is no height or width, then cant build mask

	preCalculatedRadians = new double[MASK_RES]; // to compensate the video ratio
	BuildRadiansTable(preCalculatedRadians);

	if(imHeight>imWidth) mask.rayLength = imHeight;//2; // We are assuming that the laser will never be larger
	else				 mask.rayLength = imWidth;//2;  // than the shortest screen attributes. eg height or width

	// Initialising the mask rays array to the length of mask.rayLength //
	for(loop=0;loop<MASK_RES;loop++)									//
		{																//
		mask.rays[loop].rayAtoms = new rayAtom[mask.rayLength];			//
		}																//
	//////////////////////////////////////////////////////////////////////

	lastPixel.x = nextPixel.x = 0;
	lastPixel.y = nextPixel.y = 0;

	// Loop thru all the rays finding there pixel cordinate value.
	for(loop=0;loop<MASK_RES;loop++)
		{
		radius = 25; // Starting the radius off at a set distance. This is the whole in the centre of the mask
		// Scan thru the ray and find all the pixels it crosses thru
		
		// Pi is added to the direction to rotate the mask 180 deg. Profile 0 is on the bottom not the top. 
		// and 1 is added to offset position by one profile;
		////// ^^^^^^^^^^^^^^-- Not any more.--^^^^^^^^^^^^

		//radian = ((loop+1)*directionStep)+PI; // Radian Angle of ray.
		radian = preCalculatedRadians[loop];
		sinRadian = sin(radian);
		cosRadian = cos(radian);

		for(atom=0;atom<mask.rayLength;atom++)
			{
			while(true) // Keep looping until you find the next pixel not looked at
				{
				rawPixel.x=sinRadian*radius; // Calculate the next rawPixel position
				rawPixel.y=cosRadian*radius; // for the x and y locations.
				nextPixel.x=(int) (rawPixel.x+0.5); // Round the rawPixel to find
				nextPixel.y=(int) (rawPixel.y+0.5); // the next absolute pixel location.
				if((nextPixel.x!=lastPixel.x) || (nextPixel.y!=lastPixel.y)) break; // If done this pixel then move to next one
				radius+=radiusStep;	// Look at the next radius unit.
				}
			lastPixel=nextPixel; // Last pixel looked at not needed to be looked at again.
			mask.rays[loop].rayAtoms[atom].coordScreen.x=nextPixel.x;	// Store pixel coordinate x
			mask.rays[loop].rayAtoms[atom].coordScreen.y=nextPixel.y;	// Store pixel coordinate y
			
			coord.x = (double) nextPixel.x;
			coord.y = (double) nextPixel.y * imRatio;
			mask.rays[loop].rayAtoms[atom].coord.x = coord.x; // Store pixel after ration 2:3
			mask.rays[loop].rayAtoms[atom].coord.y = coord.y; // for x and y;
			mask.rays[loop].rayAtoms[atom].radius = Hypot(coord.x,coord.y); // Store radial distance for that pixel
			mask.rays[loop].rayAtoms[atom].angle = GetAngle(coord.x, coord.y); // Store angle from 2:3 atom angle
			}		
		}
	delete[] preCalculatedRadians;
}


void RadialScan::DownSample(void)
{
	int i;
	vec2double zero;
	vec2double point;
//	mask.offset.y/=imRatio;

	zero.x=0; zero.y=0;
	for(i=0;i<PROFILE_SIZE;i++)	
		{
		point=GetProfilePointXY(i);
		if(point!=0) point=point+mask.offset;
		finalProfile[i].coordinate=point;
		}
//	mask.offset.y*=imRatio;
}

void RadialScan::GrabData(int sample) // Sample, 1 all rays, 2 every 2nd ray, 3 every 3rd ray. etc
{
	int loop;
	int atom;
	vec2int p; // Point to colour video image

	if(!validPointer) return;

	for(loop=0;loop<MASK_RES;loop+=sample) // Go thru the mask rays by step sample and retrive the profile point.
		{
		for(atom=0;atom<mask.rayLength;atom++) // Scan thru each ray atom and find retrieve raw image data
			{
			p.x = mask.rays[loop].rayAtoms[atom].coordScreen.x; // Get relative X and Y
			p.y = mask.rays[loop].rayAtoms[atom].coordScreen.y; // coordinate from ray
			p.x+=(int) (((float) mask.offset.x)+0.5); // Offset X and Y 
			p.y+=(int) (((float) mask.offset.y/imRatio)+0.5); // coordiante by centre of Laser
			if((p.x<0+redAdjust) || (p.x>=imWidth-redAdjust) || (p.y<2+redAdjust) || (p.y>=imHeight-redAdjust)) // y is limited too two to avoid video
				{														// scanline coruption of image processing
				break; // Boundry check, If its at the edge of the image no longer need to continue filling ray 
				}
			mask.rays[loop].rayAtoms[atom].processedImage = 
			mask.rays[loop].rayAtoms[atom].originalImage  = FilterVideoPixel(p.y,p.x);
			}
		mask.rays[loop].lastEntry=atom; // Mark where ray was cut off so don't have to check later in processing
		}
}

inline int	RadialScan::FilterVideoPixel(int y, int x) //PCN4141 pixel imPixel)

{
	int pixelValue; //PCN 4141
	int i,j;
	int totalR=0, totalG=0, totalB=0;
	pixel imPixel;
	int divideValue;
	


	if(x<0 || x>=imWidth) return 0; //ANT VOB
	if(y<0 || y>=imHeight) return 0; //ANT VOB

	if(redAdjust==0) imPixel = imVideo[y][x];
	else if((x>(int) redAdjust && x<imWidth-(int) redAdjust) && (y>(int) redAdjust && y<imHeight-(int) redAdjust))
	{
		divideValue = (int) (redAdjust*2)+1;
		divideValue*=divideValue;

		for(i=(int) -redAdjust;i<(int) (redAdjust+1);i++)
			for(j=(int) -redAdjust;j<(int) (redAdjust+1);j++)
			{

			totalR += (imVideo[y+i][x+j].red);
			totalG += (imVideo[y+i][x+j].green);
			totalB += (imVideo[y+i][x+j].blue);
			}
		
		totalR/=divideValue;
		totalG/=divideValue;
		totalB/=divideValue;

		if(totalR>255) totalR=255;
		if(totalG>255) totalG=255;
		if(totalB>255) totalB=255;
			
		imPixel.red = (unsigned char) totalR;
		imPixel.green = (unsigned char) totalG;
		imPixel.blue = (unsigned char) totalB;
		
	}
	else
	{
	imPixel = imVideo[y][x];
	}
	


	if(filterValue==0) return contrastBrightness[imPixel.red];
	else if(filterValue==1) return contrastBrightness[imPixel.green];
	else if(filterValue==2)
	{
		pixelValue = ((imPixel.red+imPixel.blue+imPixel.green));
		if(pixelValue>255) return pixelValue;
		return contrastBrightness[pixelValue];
	}
	else
	{
		pixelValue = (((imPixel.red*2)+imPixel.blue+imPixel.green));
		if(pixelValue>255) pixelValue = 255;
		return contrastBrightness[pixelValue];
	}
//	else //Red Invert
//	{
//		return ((255-imPixel.red)+(255-imPixel.blue)+(255-imPixel.green))/3;
//	}

//	else
//	{
//		//pixelValue = imPixel.red+imPixel.green+imPixel.blue;
//		pixelValue = (int) (((float) imPixel.red * redAdjust) + 
//					 ((float) imPixel.green*greenAdjust) +
//					 ((float) imPixel.blue*blueAdjust));
//		if(pixelValue>255) return 255;
//		return pixelValue;
//	}

}

double RadialScan::GetAverageProfile(int sample)
{
	int i,atomIndex;
	double avg=0;
	int count=0; //PCN2888 Opps forgot to initialise it (Antony 19 July 2004)
	
	for(i=0;i<MASK_RES;i+=sample)
		{
		if(mask.rays[i].profilePoint.coordinate!=0)
			{
			atomIndex=mask.rays[i].profilePoint.atomIndex;
			avg+=mask.rays[i].rayAtoms[atomIndex].originalImage;
			count++;
			}
		}
	return (count==0) ? 0:avg/count;
}

int RadialScan::Candidate(vec2double *edges,int in, int out, double greatestNeg, int &pos, int &neg,int sizeOfEdgeArray)
{
	vec2double  greatestEdge;
	vec2int		greatestAtom;

	greatestEdge.x=greatestEdge.y=0;
	greatestAtom.x=in; greatestAtom.y=out;

	int i,j;
	
	for(i=in,j=out;i<=out;i++,j--)
		{
		//This check was out by a = in a >= before this function was called. So now fixed before this function.
		//if((i<0) || (i>=sizeOfEdgeArray)) {Msg("Out of bounds, edge array RadialScan::Candidate i = %i",i);	 continue; } //PCN3561
		if(edges[i].x>greatestEdge.x) 
			{
			greatestEdge.x=edges[i].x; 
			greatestAtom.x=i;
			}
		//This check was out by a = in a >= before this function was called. So now fixed before this function.
		//if((j<0) || (j>=sizeOfEdgeArray)) {Msg("Out of bounds, edge array RadialScan::Candidate i = %i",j); continue;} //PCN35761
		if((edges[j].y<greatestEdge.y) && (edges[j].y<greatestNeg)) 
			{

			greatestEdge.y=edges[j].y; 
			greatestAtom.y=j;
			}
		}

	if((greatestAtom.y==out) || (greatestEdge.y==0)) {pos=0; neg=0; return false;}
	pos=greatestAtom.x;
	neg=greatestAtom.y;
	return true;
}


double RadialScan::ProcessRay(ray *singleRay, int pass, int loop, int display)
{
	int profileAtom = 0;
	int i;
	double strongestPosEdge=0;
	double strongestNegEdge=0;
	double strengthPos, strengthNeg;
	int atom;
	int countCandidates=0;

	vec2double *edges;
	vec2int *candidates;

	vec2int edgePair;
	vec2int profPair;
	
	edges	   = new vec2double[singleRay->lastEntry];
	candidates = new vec2int[singleRay->lastEntry];
	
	profPair=0;
	edgePair=0;

	int afar=laserWidth/2;
	int lookAfar; // PCN3122
	int lookAway; // PCN3122
	if(afar<0) afar=0;
	if(afar>singleRay->lastEntry) afar=singleRay->lastEntry-1;

	for(atom=1;atom<singleRay->lastEntry;atom++) // Scan for range of edge strength;
		{

		// Debug information
		 //if(showPutPixels)
		//	{ 
		//	AddPutPixels(vec2double(singleRay->rayAtoms[atom].coordScreen.x+mask.offset.x,
		//		         singleRay->rayAtoms[atom].coordScreen.y+mask.offset.y)
		//				 ,255,255,0);
		//	}

		if(IsInMask(singleRay->rayAtoms[atom].coordScreen)) { edges[atom]=0; continue; } 
		if((pass==3) && (singleRay->rayAtoms[atom].radius>mask.averageRadius*externalRadius)) {edges[atom]=0; continue;}
		if(singleRay->rayAtoms[atom].radius<mask.averageRadius*internalRadius) {edges[atom]=0; continue;}
		else 
			{
			//PCN3075 lookAfar and lookAway now added and adjusted to edge of image instead of droped when edge of image is passed.
			lookAfar = atom - afar; if(lookAfar<1) lookAfar=1;

			//(Antony, 27 June 2005)
			//PCN3561, the condition --v the >= was only > than, if lookaway was exactly lastEntry it would still let the ray be accessed out of bounds, very very occasounaly
			lookAway = atom + afar; if(lookAway>=singleRay->lastEntry) lookAway=singleRay->lastEntry-1;
			strengthPos = (double) (singleRay->rayAtoms[atom].processedImage - singleRay->rayAtoms[lookAfar].processedImage); // Find strength of edge for pos
			strengthNeg = (double) (singleRay->rayAtoms[lookAway].processedImage - singleRay->rayAtoms[atom].processedImage); // Find strength of edge for neg
			}
		if((strengthPos<cutoffLevel) || (strengthNeg>-cutoffLevel))	{ edges[atom]=0; continue;}
		if((pass==3) && showProfileCandidatesOverlay && (mask.averageRadius!=0)) AddProfileCandidatesOverlay(singleRay->rayAtoms[atom].coordScreen);

		if(strongestPosEdge<strengthPos) strongestPosEdge=strengthPos; // If current strength greater than maxPosStrength, make it new strength
		if(strongestNegEdge>strengthNeg) strongestNegEdge=strengthNeg; // If current strength less than maxNeg
		

		edges[atom].x = strengthPos;
		edges[atom].y = strengthNeg;
		}


	if(strongestPosEdge>cutoffLevel)
		{
		strongestPosEdge*=0.33;//(pickupLevel/100);
		strongestNegEdge*=0.33;//(pickupLevel/100);
		for(atom=1;atom<singleRay->lastEntry;atom++)
			{
			if(edges[atom].x>strongestPosEdge)
				{
				lookAfar = atom-afar; if(lookAfar<1) lookAfar=1;

				//(Antony, 27 June 2005)
				//PCN3561, the condition --v the >= was only > than, if lookaway was exactly lastEntry it would still let the ray be accessed out of bounds, very very occasounaly
				lookAway = atom+afar; if(lookAway>=singleRay->lastEntry) lookAway=singleRay->lastEntry-1;
				
				// Debug drawing information 12 Nov 2004 (Antony)
				// if(showPutPixels)
				//	{ 
				//	AddPutPixels(singleRay->rayAtoms[lookAfar].coordScreen,255,255,255);
				//	AddPutPixels(singleRay->rayAtoms[lookAway].coordScreen,255,255,255);
				//	}

				if(Candidate(edges,lookAfar, lookAway, strongestNegEdge, edgePair.x, edgePair.y,singleRay->lastEntry))
					{
					if((pass==3) && (singleRay->rayAtoms[edgePair.y].radius>(mask.averageRadius*externalRadius))) continue;
					if(singleRay->rayAtoms[edgePair.x].radius>(mask.averageRadius*internalRadius))
						{
						if(showLaserWidthOverlay && pass==3)
							{
							for(i=lookAfar;i<(lookAway);i++) AddLaserWidthOverlay(singleRay->rayAtoms[i].coordScreen);
							}
						candidates[countCandidates]=edgePair;	// Store possible profile neg pos edge
						atom+=afar;	// Skip the atoms its just checked.
						countCandidates++; // Count the number of candidates.
						}
					}
				}
			}
		}
	else countCandidates=0;

/////// Find best candidates //////////
	if(countCandidates==0) profileAtom=0; // If no candidates then no profile
	else{ 
		if(mask.averageRadius==0) profPair=candidates[0]; // if no radius then first run and get first profile candidate
		else 
			{
			profPair = FindBestCandidates(candidates,countCandidates,singleRay); // If average radius find best candidate
		}	
		profileAtom=(profPair.y+profPair.x)/2;			// Get middle of profile pos & neg edges
		}
///////////////////////////////////////

	vec2double pos,neg; // Profile coordinates for the best candidate neg and pos edge
	double posRadius, negRadius; // Profile radius for best candidate neg and pos edge

	double profileRadius;			// Final profile radius
	vec2double profileCoordinate;	// Final profile coordinate

/////// Final filter of profile points ///
	if(countCandidates>2) 
		{
	//	profileAtom=0;
//		for(int i=0;i<countCandidates;i++)
//			ShowAtom(loop,(candidates[i].x+candidates[i].y)/2,0,255,255,true);
		}

//	if(display) {ShowAtom(loop,profileAtom,0,0,255, true);}
	if(profileAtom!=0)
		{
		
		pos = singleRay->rayAtoms[profPair.x].coord;
		neg = singleRay->rayAtoms[profPair.y].coord;
		posRadius = singleRay->rayAtoms[profPair.x].radius; // Get true middle radius point and
		negRadius = singleRay->rayAtoms[profPair.y].radius; // not the one at profileAtom radus

		profileRadius = (posRadius + negRadius) / 2;
		profileCoordinate = (pos+neg) / 2; // Not very acurate better to use profile radius.
										   // This coordinate is recalculated by RecaluclateCoordinates funcion.
		}

	//////////////////////////////////////////////////////////////////////////////////////
	// Fill in the related data. (Radius) ,(Atom position), (coordinate) and (Angle)	//
	singleRay->profilePoint.atomIndex=profileAtom;	// Store in the ray position of the profile atom //
	if(profileAtom==0)																	// 
		{																				//
		singleRay->profilePoint.radius		= 0;										//
		singleRay->profilePoint.coordinate	= vec2double(0,0);							//
		singleRay->profilePoint.angle		= 0;										//
		}																				//														
	else																				//
		{	
		//
		singleRay->profilePoint.radius      =singleRay->rayAtoms[profileAtom].radius;				//
		singleRay->profilePoint.coordinate  =profileCoordinate;	// PCN3233, why was this =0 ???	
		singleRay->profilePoint.angle		=singleRay->rayAtoms[profileAtom].angle;				//
		} 																				//

	delete[] edges;
	delete[] candidates;
	return singleRay->profilePoint.radius; // return the radius size unit of the Atom.  //
	//////////////////////////////////////////////////////////////////////////////////////
}


vec2int RadialScan::FindBestCandidates(vec2int *candidates,int countCandidates, ray *singleRay)
{
	int closestToRadius;
	double closestDistance;
	int i;
	double absoluteDistance;
	double currentRadius;

	if(countCandidates==1) return candidates[0];

	if(1==1) //PCN4380
	{
		closestToRadius=0;
		currentRadius=(singleRay->rayAtoms[candidates[0].x].radius + 
					   singleRay->rayAtoms[candidates[0].y].radius)/2;
		closestDistance=fabs(currentRadius-mask.averageRadius);


		for(i=1;i<countCandidates;i++)
			{
			currentRadius=(singleRay->rayAtoms[candidates[i].x].radius + 
						   singleRay->rayAtoms[candidates[i].y].radius)/2;
			absoluteDistance=fabs(currentRadius-mask.averageRadius);
			if(absoluteDistance<closestDistance) 
				{
				closestDistance=absoluteDistance;
				closestToRadius=i;
				}
			}

		return candidates[closestToRadius];
	}
	else
	{
		return candidates[countCandidates-1];

		if(countCandidates==1) return candidates[0];
		closestToRadius=0;
		currentRadius=(singleRay->rayAtoms[candidates[0].x].radius + 
					   singleRay->rayAtoms[candidates[0].y].radius)/2;
		closestDistance=fabs(currentRadius-mask.averageRadius);


		for(i=1;i<countCandidates;i++)
			{
			currentRadius=(singleRay->rayAtoms[candidates[i].x].radius + 
						   singleRay->rayAtoms[candidates[i].y].radius)/2;
			absoluteDistance=fabs(currentRadius-mask.averageRadius);
			if(absoluteDistance<closestDistance) 
				{
				closestDistance=absoluteDistance;
				closestToRadius=i;
				}
			}

		return candidates[closestToRadius];
	}
}

void RadialScan::RecalculateCoordinates(int sample)
{
	double	angle;
	double	radius;

	int loop;
	for(loop=0;loop<MASK_RES;loop+=sample)
		{
		angle=mask.rays[loop].profilePoint.angle;
		radius=mask.rays[loop].profilePoint.radius;

		mask.rays[loop].profilePoint.coordinate.x = sin(angle)*radius;
		mask.rays[loop].profilePoint.coordinate.y = cos(angle)*radius;
		}
}

//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Name: PCN2568 RadialProfile::IsInWaterSection 1 June 2004
// Created By: Antony van Iersel
// Description:	Return true if the selected profile is in the selected
//              water section
// Input: Profile point the check.
// Output: If it is in selected water section return true.
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
int RadialScan::IsInWaterSection(int i)
{
	//PCN3219
	if(egnoreList[i]==0) return false;
	else return true;
}


void RadialScan::SmoothProfile(void)
{
	int i;
	
	int r1i,r3i;  
	double r1,r2,r3; //PCN2891 was int
	//float r=(float) max.r; //PCN2891 was int
	double profBuf[MASK_RES]; //PCN2891 was int
	double localProfile[MASK_RES];
	


//	float ll=(r*2)/4; // lower limit for radius //PCN2891 was int
//	float ul=(r*5)/4; // upper limit for radius // 6 was 5 //PCN2891 was int

	if(mask.averageRadius==0) GetAverageRadius(1);
	if(DonutLocked) mask.averageRadius=expectedDiameter/2; //PCN4539

//	numprofilepoints = 0;


	for(i=0;i<MASK_RES;i++) localProfile[i]=mask.rays[i].profilePoint.radius;

	for(int j=0;j<MASK_RES;j++) 
		{

		}

	//look for missing one point
	//PCN2795 (Antony van Iersel, 3 May 2004) ////////////////////////////////////
	//and PCN2737 removes any points that are on there own and fill in wholes	// 
	for(i=0;i<MASK_RES;i++)													//
		{																		//
		if(i==0) r1i=(MASK_RES-1);	else r1i=i-1;							//
		if(i==(MASK_RES-1)) r3i=0;	else r3i=i+1;							//
																				//
		r1=localProfile[r1i];														//
		r2=localProfile[i];														//
		r3=localProfile[r3i];														//
																				//
		if((r2==0) && (r3!=0) && (r1!=0)) localProfile[i]=(r1+r3)/2;				//
		if((r1==0) && (r3==0))													//
			localProfile[i]=0; // 11 May 2004, Antony, remove any points on the own.
		}																		//
	//////////////////////////////////////////////////////////////////////////////

	int loop;
	//PCN2737 Removing of rough points find if the nextdoor points are jaged or not, /////
	// if are then average it. (Antony 8 June, 2004, Sorry for lack of comments)		//
	// The more this is looped thru the smoother the profile gets and eliminates rough points.
	for(loop=0;loop<2;loop++) //DNF															//
		{																				//
		for(i=0;i<MASK_RES;i++) profBuf[i]=-1; // Clear the profile changes, all -1.//
																						//
		// Loops thru all points then store the ones that need to cange					//
		for(i=0;i<MASK_RES;i++)														//
			{																			//
			// Find its neighbouring points in the profile ring. Left r1i, right r3i	//
			if(i==0) r1i=(MASK_RES-1);	else r1i=i-1;								//
			if(i==(MASK_RES-1)) r3i=0;	else r3i=i+1;								//
																						//
			r1=localProfile[r1i];	// Retrieve profile point to left.						//
			r2=localProfile[i];	// Retrieve current profile point						//
			r3=localProfile[r3i];	// Retrieve profile point to right.						//
																						//
			if((r1==0) || (r2==0) || (r3==0)) continue; // If anyoff then are 0 get next points
																						//
			// If points look like -_- or _-_ then average the middle one				//
			if(((r1<r2) && (r3<r2)) || ((r1>r2) && (r3>r2)))							//
				profBuf[i]=(r1+r3)/2;													//
			}																			//
		// Loop thru and copy any changed profile points.								//
		for(i=0;i<MASK_RES;i++) if(profBuf[i]!=-1) localProfile[i]=profBuf[i];		//
		}																				//
	//////////////////////////////////////////////////////////////////////////////////////

	// Copy back the smoothed profile to the radial profile data
	for(i=0;i<MASK_RES;i++) mask.rays[i].profilePoint.radius = localProfile[i]; 

	RecalculateCoordinates(1); // Recalculate the Profile coordinates
}

double RadialScan::GetProfilePoint(int no)
{
	if(no>=MASK_RES) return 0;	// Make sure its in the range of scaned resolution.
	return mask.rays[no].profilePoint.radius; // Returns the radius of this particular Ray.
}

double RadialScan::GetProfilePointX(int no)
{	
	if(no>=PROFILE_SIZE) return 0;
	return GetProfilePointXY(no).x;
}

double RadialScan::GetProfilePointY(int no)
{
	if(no>=PROFILE_SIZE) return 0;
	return GetProfilePointXY(no).y;
}

vec2double RadialScan::GetProfilePointXY(int no)
{
	int j;//,avgCount=0;
	int index[3];
	int step=MASK_RES/PROFILE_SIZE;
	if(redAdjust!=0) return mask.rays[no*3].profilePoint.coordinate; //4141
	int goodCount=0;

	vec2double average[3];
	vec2double avgProfile;

	//return mask.rays[no*3].profilePoint.coordinate;

	avgProfile.x=0; avgProfile.y=0;
	if(no>=PROFILE_SIZE) return avgProfile;

	if(ignoreWaterLevel && !showProfileWaterLevel && IsInWaterSection(no)) return avgProfile;
	index[1]=no*step;
	
	if(index[1]==0)          index[0]=MASK_RES-1; else index[0]=index[1]-1;
	if(index[1]==MASK_RES-1) index[2]=0;          else index[2]=index[1]+1;
	

	vec2double a;
	vec2double b;
	vec2double c;

	a = mask.rays[index[0]].profilePoint.coordinate;
	b = mask.rays[index[1]].profilePoint.coordinate;
	c = mask.rays[index[2]].profilePoint.coordinate;

	for(j=0;j<3;j++)
		{
			if(mask.rays[index[j]].profilePoint.coordinate!=0)
			{
				avgProfile = avgProfile + mask.rays[index[j]].profilePoint.coordinate;
				goodCount++;
			}
		
		}

	if(goodCount<2) return vec2double(0);
	//vec2double goodreturn;
	//goodreturn = avgProfile/goodCount;
	//goodreturn = goodreturn.toVector();
	//if(goodreturn.y < 37)
	//{
	//	__asm nop
	//}
	
	return (avgProfile/goodCount);
}


double RadialScan::GetCenterX(void)
{
	return mask.offset.x;
}

double RadialScan::GetCenterY(void)
{
	return mask.offset.y;
}

double RadialScan::GetRadius(void)
{
	return mask.averageRadius;
}

double RadialScan::GetAverageRadius(int sample)
{
	int loop;
	int count=0;
	double totalRadius=0;
	double averageRadius;
	for(loop=0;loop<MASK_RES;loop+=sample)
		{
		if(mask.rays[loop].profilePoint.radius!=0)
			{
			totalRadius+=mask.rays[loop].profilePoint.radius; // Add all the radius points together;		
			count++;
			}
		}
	if(count==0) return 0;
	averageRadius=totalRadius/count;	// Divide by the number of radius, to get Average
	return averageRadius;					// Return Average Radius
}

void RadialScan::ShowMask(void)
{
	int loop;
	int atom;
	vec2int p; // Point to colour video image
	if(!validPointer) return;
	for(loop=0;loop<MASK_RES;loop++)
		{
		for(atom=0;atom<mask.rays[loop].lastEntry;atom++)
			{
			p.x = mask.rays[loop].rayAtoms[atom].coordScreen.x; // Get relative X and Y
			p.y = mask.rays[loop].rayAtoms[atom].coordScreen.y; // coordinate from ray
			p.x+=(int) (mask.offset.x+0.5); // Offset X and Y 
			p.y+=(int) (( mask.offset.y/imRatio)+0.5); // coordiante by centre of Laser
			if((p.x<0) || (p.x>=imWidth) || (p.y<0) || (p.y>=imHeight)) //ANT VOB 
				{
				/*Msg("Opps, out of bounds");*/ continue; // Boundry check
				}
			imVideo[p.y][p.x].blue+=50; // Make 
			imVideo[p.y][p.x].red+=50; // the pixel
			imVideo[p.y][p.x].green+=50;// white.
			}
		}
}


void RadialScan::ShowAtom(int rays, int atom, int red, int green, int blue, int rel)
{
	int x,y;

	red=(red>255)?255:red;
	green=(green>255)?255:green;
	blue=(blue>255)?255:blue;

	red=(red<0)?0:red;
	green=(green<0)?0:green;
	blue=(blue<0)?0:blue;

	if(rel)	{ x=(int) (mask.rays[rays].rayAtoms[atom].coordScreen.x+mask.offset.x+0.5);
			  y=(int) (mask.rays[rays].rayAtoms[atom].coordScreen.y+((double) mask.offset.y/imRatio)+0.5);
			}
	else	{ x=mask.rays[rays].rayAtoms[atom].coordScreen.x+(imWidth/2);
			  y=mask.rays[rays].rayAtoms[atom].coordScreen.y+(imHeight/2);
			}

	if((x < 0) || (y < 0)) return;
	if((x >= imWidth) || (y >= imHeight)) return; //ANT VOB
	imVideo[y][x].red=(unsigned char) red;
	imVideo[y][x].green=(unsigned char) green;
	imVideo[y][x].blue=(unsigned char) blue;
/*	imVideo[y-1][x-1].red=red;
	imVideo[y-1][x-1].green=green;
	imVideo[y-1][x-1].blue=blue;
	imVideo[y-1][x+1].red=red;
	imVideo[y-1][x+1].green=green;
	imVideo[y-1][x+1].blue=blue;
	imVideo[y+1][x+1].red=red;
	imVideo[y+1][x+1].green=green;
	imVideo[y+1][x+1].blue=blue;
	imVideo[y+1][x-1].red=red;
	imVideo[y+1][x-1].green=green;
	imVideo[y+1][x-1].blue=blue;
*/
}

void RadialScan::ShowProfilePoints(int sample,int red,int blue,int green, int centered)
{
	int loop;
	double x,y;
	int intx, inty;
	for(loop=0;loop<MASK_RES;loop+=sample)
	{
		x=mask.rays[loop].profilePoint.coordinate.x;
		y=mask.rays[loop].profilePoint.coordinate.y;
		if(centered) { x=x+(imWidth/2); y=  (y+(imHeight/2))/imRatio; }
		else		 { x=x+mask.offset.x; y=(y+mask.offset.y)/imRatio; }

		intx=(int) (x+0.5); inty=(int) (y+0.5);
		if((inty<0) || (inty>=imHeight)) return; //ANT VOB
		if((intx<0) || (intx>=imWidth)) return; //ANT VOB

		{
			if(IsWaterLevelOn() && IsInWaterSection(loop/3))
			{
				//////////////////////////////////////////////////////////////////////
				// PCN2608 (Antony van Iersel, 13 May 2004) Displayed Profile Point //
				setim(intx-1,inty,0,255,0);	// boarder. Bounds checking done in setim   //
				setim(intx+1,inty,0,255,0);	//////////////////////////////////////////////
				setim(intx,inty+1,0,255,0);	//
				setim(intx,inty-1,0,255,0);	//
				setim(intx-1,inty+1,0,255,0);	//
				setim(intx-1,inty-1,0,255,0);	//
				setim(intx+1,inty+1,0,255,0);	//
				setim(intx+1,inty-1,0,255,0);	//
				//////////////////////////
			}
			
			else

			{
				//////////////////////////////////////////////////////////////////////
				// PCN2608 (Antony van Iersel, 13 May 2004) Displayed Profile Point //

				setim(intx-1,inty,0,0,255);	// boarder. Bounds checking done in setim   //
				setim(intx+1,inty,0,0,255);	//////////////////////////////////////////////
				setim(intx,inty+1,0,0,255);	//
				setim(intx,inty-1,0,0,255);	//
				setim(intx-1,inty+1,0,0,255);	//
				setim(intx-1,inty-1,0,0,255);	//
				setim(intx+1,inty+1,0,0,255);	//
				setim(intx+1,inty-1,0,0,255);	//
				//////////////////////////
			}
		
		}
	}
	for(loop=0;loop<MASK_RES;loop+=sample)
	{
		x=mask.rays[loop].profilePoint.coordinate.x;
		y=mask.rays[loop].profilePoint.coordinate.y;
		if(centered) { x=x+(imWidth/2); y=  (y+(imHeight/2))/imRatio; }
		else		 { x=x+mask.offset.x; y=(y+mask.offset.y)/imRatio; }

		intx=(int) (x+0.5); inty=(int) (y+0.5);
		if((inty<0) || (inty>=imHeight)) return; //ANT VOB
		if((intx<0) || (intx>=imWidth)) return; //ANT VOB
		imVideo[(int) (inty)][(int) (intx)].blue=(unsigned char) blue;
		imVideo[(int) (inty)][(int) (intx)].red=(unsigned char) red;
		imVideo[(int) (inty)][(int) (intx)].green=(unsigned char) green;
		{
			if(IsWaterLevelOn() && IsInWaterSection(loop/3)) setim(intx,inty,0,0,255);	
			else setim(intx,inty,0,255,0);	// no longer blue but green with a blue     //

		
		}
	}

}


void RadialScan::setim(int x, int y, int r, int g, int b){
	if(x>=0 && x<imWidth && y>=0 && y<imHeight){
		imVideo[y][x].red = (unsigned char) r;
		imVideo[y][x].green = (unsigned char) g;
		imVideo[y][x].blue = (unsigned char) b;
	}
	//Msg(TEXT("Setting image out of bounds!!!"));
}

void RadialScan::SetOffset(int x, int y)
{
	mask.offset.x=x; // Offset the centre of the Mask
	mask.offset.y=y; // This would normally come from the Laser Center
}

inline double RadialScan::Hypot(double x, double y)
{
	return sqrt((x*x) + (y*y));
}

double RadialScan::GetAngle(double x,double y)
{
	double adj;
	double opp;
	
	adj = fabs(x);
	opp = fabs(y);
	
	if((x==0) && (y==0)) return 0;
	if((x==0) && (y>0)) return 0;
	if((x==0) && (y<0)) return PI;
	if((x>0) && (y==0)) return (PI/2);
	if((x<0) && (y==0)) return (PI+(PI/2));
	if((x>0) && (y>0))  return (       atan(adj/opp)); // + , +
	if((x>0) && (y<0))  return (PI-    atan(adj/opp)); // + , -
	if((x<0) && (y<0))  return (PI+    atan(adj/opp)); // - , -
	if((x<0) && (y>0))  return ((2*PI)-atan(adj/opp)); // - , +
	return 0;
}

void RadialScan::AdjustCentre(int pass)
{

	vec2double centre;
	centre=0;

	
	FindBestNeighbour(centre,512,pass);
	//FindMedianXCentre(centre);
	mask.offset.x=mask.offset.x+centre.x;
	mask.offset.y=mask.offset.y+centre.y;
	//PCN3122 if loose the centre use the running average to start the next centre
	if((mask.offset.x<0) || (mask.offset.y<0))  //
	{											//
		mask.offset=averageCentreHistory;		//
												//												
	}											//
	// 12 November 2004 (Antony) /////////////////



}

void RadialScan::AdjustCentreWithSmartDataFill(void)
{
	
	int i;
	for(i=0;i<180;i++)
	{
		if (ignoreWaterLevel) theCentre.egnoreList[i] = egnoreList[i];else theCentre.egnoreList[i]=0;
		theCentre.pvData[i] = mask.rays[i*3].profilePoint.coordinate;
	}


	theCentre.CalculateCentre();

	for(i=0;i<180;i++)
	{
		
	
		if(showPutPixels) AddPutPixels(vec2double(theCentre.pvData[i].x+mask.offset.x,
											      theCentre.pvData[i].y+mask.offset.y),255,255,255);
	}

	mask.offset.x += theCentre.pvCentreX;
	mask.offset.y += theCentre.pvCentreY;
	

}


void RadialScan::AdjustCentreFinalProfile(void)
{
	int i;
//	double radius;
//	SmoothLikeABabysBottom();

	vec2double centre;
	vec2double testCentre=0;


	centre.x=0;
	centre.y=0;


//	Do we need best centre here or at end at each dump?
//	FindBestNeighbour(centre,256,3); //PCN3122 Old final centre, still works rather good.
	
	mask.offset=mask.offset+centre;

// PCN3122 Keep track of the previous 5 centres
//
	if((mask.offset.x<0)       || (mask.offset.y<0) || 
	   (mask.offset.x>=imWidth) || (mask.offset.y>=imHeight)) // If can't find centre then use History //ANT VOB
	{										   // and add to history centre of screen
		centreHistory[averageCentreHead].x = (double) imWidth / 2;
		centreHistory[averageCentreHead].y = (((double) imHeight / 2) * imRatio);
		sprintf(&centreDebugHistoryNotes[numberCentreDebug][0],"Missed");
		mask.offset = averageCentreHistory; 
//		centreDebugHistory[numberCentreDebug]=mask.offset; //????
	}
	else 
	{
		centreHistory[averageCentreHead]=mask.offset;
	}
	
	if(numberCentreDebug<10000) numberCentreDebug++;
	
// Take the last centre and current centre and average them together to get used centre
	vec2double previousHistory;					// PCN3122 12 Nov 2004 (Antony)	//
	
	previousHistory = previousCentre;				//
	previousCentre  = mask.offset;
	if(previousCentre!=0)//
		//mask.offset = (mask.offset+previousHistory)/2;		//
	averageCentreHead=(averageCentreHead+1)%5;							//
//////////////////////////////////////////////////////////////////////////	

//	vec2double profileVector;

	for(i=0;i<PROFILE_SIZE;i++) 
	{
		if(finalProfile[i].coordinate!=0)
//		{
			finalProfile[i].coordinate=finalProfile[i].coordinate-mask.offset;
//			profileVector = finalProfile[i].coordinate.toVector();
//			finalProfile[i].angle=profileVector.x;
//			finalProfile[i].radius=profileVector.y;
//		}

	}
//	ShowPutPixels();
	
}

void RadialScan::FindBestNeighbour(vec2double &curPoint, double size, int pass)
{
	if(size<0.25) return;
	double variance;
	double closestVar;
	bool iTBN=false; // is there better neighbour
	vec2double closestPoint;
	vec2double lookingAt;

	closestPoint=curPoint;
	closestVar=GetCentreVariance(curPoint,pass);

	lookingAt.x=curPoint.x-size;
	lookingAt.y=curPoint.y-size;
	variance=GetCentreVariance(lookingAt,pass);
	if(variance<closestVar) {closestPoint=lookingAt; closestVar=variance; iTBN=true;}

	lookingAt.x=curPoint.x;
	lookingAt.y=curPoint.y-size;
	variance=GetCentreVariance(lookingAt,pass);
	if(variance<closestVar) {closestPoint=lookingAt; closestVar=variance; iTBN=true;}

	lookingAt.x=curPoint.x+size;
	lookingAt.y=curPoint.y-size;
	variance=GetCentreVariance(lookingAt,pass);
	if(variance<closestVar) {closestPoint=lookingAt; closestVar=variance; iTBN=true;}

	lookingAt.x=curPoint.x+size;
	lookingAt.y=curPoint.y;
	variance=GetCentreVariance(lookingAt,pass);
	if(variance<closestVar) {closestPoint=lookingAt; closestVar=variance; iTBN=true;}

	lookingAt.x=curPoint.x+size;
	lookingAt.y=curPoint.y+size;
	variance=GetCentreVariance(lookingAt,pass);
	if(variance<closestVar) {closestPoint=lookingAt; closestVar=variance; iTBN=true;}

	lookingAt.x=curPoint.x;
	lookingAt.y=curPoint.y+size;
	variance=GetCentreVariance(lookingAt,pass);
	if(variance<closestVar) {closestPoint=lookingAt; closestVar=variance; iTBN=true;}

	lookingAt.x=curPoint.x-size;
	lookingAt.y=curPoint.y+size;
	variance=GetCentreVariance(lookingAt,pass);
	if(variance<closestVar) {closestPoint=lookingAt; closestVar=variance; iTBN=true;}
	
	FindBestNeighbour(closestPoint, size/2,pass);
	curPoint=closestPoint;
}


//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Name: PCN2405 Laserprofiler::getCentreVariance 8 March 2004
// Created By: Antony van Iersel
// Description:	Return the maximum varience from a given point
//              and the posible profile points (blue overlay)
// Input: The current possible centre to look at
// Output:  The maximum variance.
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
double RadialScan::GetCentreVariance(vec2double p, int pass)
{
	int i;
	int centreEdgeSize=90;
	double variance=0; // max & min holds the current largest and smallest distance from
	double avgDist=0;
	int count=0;
	double distance[MASK_RES];
	int countGoodPoints;
	
	if(pass==1)
		{
		for(i=0;i<MASK_RES;i+=12)
			{
			if(mask.rays[i].profilePoint.radius!=0)
				{
				distance[count]=DistOfTwoPoints(p,mask.rays[i].profilePoint.coordinate);
				avgDist+=distance[count];
				count++;
				}
			}
		}
	if(pass==2)
		{
		for(i=0;i<MASK_RES;i+=12)
			{
			if(mask.rays[i].profilePoint.mark)
				if((ignoreWaterLevel && !IsInWaterSection(i/3)) || !ignoreWaterLevel)
					if(mask.rays[i].profilePoint.radius!=0)
						{
						distance[count]=DistOfTwoPoints(p,mask.rays[i].profilePoint.coordinate);
						avgDist+=distance[count];
						count++;
						}
			}
		}
	if(pass==3)
	{
		p=p+mask.offset;
		noPutPixels=0;
		// From 45 to 90 /////////////////////////////////////////////////////////////
		countGoodPoints=0;
		for(i=0;i<PROFILE_SIZE;i++)
		{
			//countGoodPoints++;
			//if(countGoodPoints>centreEdgeSize) break;
			if(finalProfile[i].mark)
				if((ignoreWaterLevel && !IsInWaterSection(i)) || !ignoreWaterLevel)
					if(finalProfile[i].coordinate!=0)
					{
						distance[count]=DistOfTwoPoints(p,finalProfile[i].coordinate);
						avgDist+=distance[count];
						count++;
						
//						AddPutPixels(finalProfile[i].coordinate,255,255,255);
						
					}
		}
/*
		// From 46 to 0 ////////////////////////////////////////////////////////////
		countGoodPoints=0;
		for(i=(PROFILE_SIZE/4)+1;i>=0;i--)
			{
			countGoodPoints++;
			if(countGoodPoints>centreEdgeSize) break;
			if(finalProfile[i].mark)
				if((ignoreWaterLevel && !IsInWaterSection(i)) || !ignoreWaterLevel)
					if(finalProfile[i].coordinate!=0)
						{
						distance[count]=DistOfTwoPoints(p,finalProfile[i].coordinate);
						avgDist+=distance[count];
						count++;
						
//						AddPutPixels(finalProfile[i].coordinate,255,255,255);
						
						}
			}

		// From 135 to 91 /////////////////////////////////////////////////////////////
		countGoodPoints=0;
		for(i=PROFILE_SIZE-(PROFILE_SIZE/4);i>PROFILE_SIZE/2;i--)
			{
			countGoodPoints++;
			if(countGoodPoints>centreEdgeSize) break;
			if(finalProfile[i].mark)
				if((ignoreWaterLevel && !IsInWaterSection(i)) || !ignoreWaterLevel)
					if(finalProfile[i].coordinate!=0)
						{
						distance[count]=DistOfTwoPoints(p,finalProfile[i].coordinate);
						avgDist+=distance[count];
						count++;
						
//						AddPutPixels(finalProfile[i].coordinate,255,255,255);
						
						}
			}

		// From 136 to 179 ////////////////////////////////////////////////////////////
		countGoodPoints=0;
		for(i=PROFILE_SIZE-(PROFILE_SIZE/4)+1;i<PROFILE_SIZE;i++)
			{
			countGoodPoints++;
			if(countGoodPoints>centreEdgeSize) break;
			if(finalProfile[i].mark)
				if((ignoreWaterLevel && !IsInWaterSection(i)) || !ignoreWaterLevel)
					if(finalProfile[i].coordinate!=0)
						{
						distance[count]=DistOfTwoPoints(p,finalProfile[i].coordinate);
						avgDist+=distance[count];
						count++;
						
//						AddPutPixels(finalProfile[i].coordinate,255,255,255);
						
						}
			}
		*/
	}


	if(count==0) return 0;

	avgDist/=(double) count;
	for(i=0;i<count;i++) variance+=fabs(avgDist-distance[i]);
	
	variance/=(double) count;
//	ShowPutPixels();
	return variance; // return the average variance
}

inline double RadialScan::DistOfTwoPoints(vec2double pt1, vec2double pt2)
{
	return sqrt(pow(pt1.x-pt2.x,2)+pow(pt1.y-pt2.y,2));
}

bool	RadialScan::IsInMask(vec2int point)
{
	
	point.x+=(int) (mask.offset.x+0.5);
	point.y+=(int) (((double) mask.offset.y/imRatio)+0.5);
	if(point.x<0 || point.x>=imWidth || point.y<0 || point.y>=imHeight) return true;
	if(egnoreMask[point.y][point.x]>0) return true;
	
	if((point.x>counterMaskTopLeft.x) && (point.x<counterMaskBottomRight.x) &&
	   (point.y>counterMaskTopLeft.y) && (point.y<counterMaskBottomRight.y))	return true;

	if((point.x>textMaskTopLeft.x)    && (point.x<textMaskBottomRight.x) &&
	   (point.y>textMaskTopLeft.y)    && (point.y<textMaskBottomRight.y))		return true;

return false;	
}

void RadialScan::SetCounterMask(int x1, int y1, int x2, int y2)
{
	counterMaskTopLeft.x	 = x1; counterMaskTopLeft.y		= y1; //(int) (((double) y1 / imRatio)+0.5);
	counterMaskBottomRight.x = x2; counterMaskBottomRight.y = y2; //(int) (((double) y2 / imRatio)+0.5);
}

void RadialScan::SetTextMask(float x1, float y1, float x2, float y2, int setclear)
{
	int x,y;
	int SetOrClear;

	if(setclear == 0) SetOrClear = 1;
	if(setclear == 1) SetOrClear = 0;

	if(x1<0) x1=0;	if(x2<0) x2=0;
	if(y1<0) y1=0;	if(y2<0) y2=0;

	if(x1>(imWidth-1)) x1 = (float) (imWidth-1); if(y1>(imHeight-1)) y1 = (float) (imHeight-1);
	if(x2>(imWidth-1)) x2 = (float) (imWidth-1); if(y2>(imHeight-1)) y2 = (float) (imHeight-1);

	for(y=(int) y1;y<=(int) y2 ;y++)
		for(x=(int) x1;x<=(int) x2;x++)
			egnoreMask[y][x]=SetOrClear;
//	textMaskTopLeft.x	  = x1;	textMaskTopLeft.y	  = y1;//(int) (((double) y1 / imRatio)+0.5);
//	textMaskBottomRight.x = x2; textMaskBottomRight.y = y2;//(int) (((double) y2 / imRatio)+0.5);
}

void RadialScan::BlankScreen(void)
{
	int x; int y;
	for(x=0;x<imWidth;x++)
		for(y=0;y<imHeight;y++)
			{
			imVideo[y][x].blue=0;
			imVideo[y][x].green=0;
			imVideo[y][x].red=0;
			}
}

//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Name: PCN2993 RadialScan::RemoveRoughPoints 20 August 2004
// Created By: Antony van Iersel
// Description:	Scan thru the hi res profile and removed any rough points, these are
//              points that stand out on there own
// Input: None
// Output: None
////////////////////////////////////////////////////////////////////////////////

void RadialScan::RemoveRoughPoints(void)
{
	double *radialCopy;
	double *averageProf;
	double sample5x[5];
	int countGood;
	double countTotal;

	int index;
	bool swaped;
	double t;

	

	


	int i;
	int j;



	radialCopy = new double[MASK_RES];
	averageProf = new double[MASK_RES];

	for(i=0;i<MASK_RES;i++) radialCopy[i] = mask.rays[i].profilePoint.radius;

	

	for(i=0;i<MASK_RES;i++)
	{

		

		countGood=0;
		for(j=0;j<5;j++) 
		{

			index = i+j;
			index = index%MASK_RES;

			if(radialCopy[index]!=0) countGood++;
			else
			{
				__asm nop
			}
			sample5x[j] = radialCopy[index]; 
			
		}
		
		
		if(countGood==5) countGood=2;
		else if(countGood==4) countGood=2;
		else if(countGood==3) countGood=3;
		else if(countGood==2) countGood=3;
		else if(countGood==1) countGood=4;
		else countGood=0;

		swaped=true;
		
		while(swaped)
		{
			swaped=false;
			for(j=0;j<4;j++)
			{
				if(sample5x[j]>sample5x[j+1]) {t = sample5x[j]; sample5x[j]=sample5x[j+1]; sample5x[j+1]=t; swaped=true;}
			}
		}
		index = (i+2)%MASK_RES;
		averageProf[index]=sample5x[countGood];
	}
	countTotal=0;
	int ia;
	int ib;
	int ic;
	int id;
	int ie;

	
	for(ia=0;ia<MASK_RES;ia++)
	{
		countTotal=0;
		countGood=0;
			

			ib=(ia+1)%MASK_RES;
			ic=(ia+2)%MASK_RES;
			id=(ia+3)%MASK_RES;
			ie=(ia+4)%MASK_RES;

			if(averageProf[ia]!=0)	{ countTotal+=averageProf[ia];	countGood++;}
			if(averageProf[ib]!=0)	{ countTotal+=averageProf[ib];	countGood++;}
			if(averageProf[ic]!=0)	{ countTotal+=averageProf[ic];	countGood++;}
			if(averageProf[id]!=0)	{ countTotal+=averageProf[id];	countGood++;}
			if(averageProf[ie]!=0)	{ countTotal+=averageProf[ie];	countGood++;}
		//if(countGood>2) mask.rays[index].profilePoint.radius = countTotal / (double) countGood;				
		if(countGood>2 && mask.rays[ic].profilePoint.coordinate != 0) mask.rays[ic].profilePoint.radius = countTotal / (double) countGood;
		else
		{
			mask.rays[ic].profilePoint.radius=0; 
			mask.rays[ic].profilePoint.atomIndex=0;
			mask.rays[ic].profilePoint.coordinate=0;
		}
	}
	



//	int one,two,three; // Raw index of the profile points
//	double *a,*b,*c;	// Profile points to check
//	double outlier = 12;//12;	// Normaly 4, How far apart do they have to be not to be rough
							// and multiplied by two for when they are rough
//	double varianceLeft;
//	double varianceRight;
//	double varianceTotal;
//	double varianceAverage;
//	int sampleSize; if(redAdjust==0) sampleSize=1; else sampleSize=3;

	// Scan thru the Hi Res Profile to find the rough points
//	for(i=0;i<MASK_RES;i+=sampleSize)
//	{
//		one=(i+(MASK_RES-sampleSize))%MASK_RES; // Find index on left, loop around from 0 to MASK_RES indes if -1
//		two=i; // index of rough point to adjust if needed
//		three=(i+sampleSize)%MASK_RES; // Find Index on Right, loop around from MASK_RES to 0 if MASK_RES
//
//		a=&(mask.rays[one].profilePoint.radius);   // pointer to the profile points
//		b=&(mask.rays[two].profilePoint.radius);   // speeds up access.
//		c=&(mask.rays[three].profilePoint.radius); //
//		// If the left and right profile points are close enough together and the centre is
//		// is outlier then make centre the average of the left and right.
//		varianceTotal	=fabs(*a-*c);
//		varianceAverage	=(fabs(*a+*c)/2)-*b;
//		varianceLeft	=fabs(*a-*b);
//		varianceRight	=fabs(*b-*c);
//
//		if((varianceLeft<=12) && (varianceRight<=12)) //Normally 2; try 12
//			{
//			if(showPutPixels) AddPutPixels(mask.rays[two].profilePoint.coordinate,255,255,0);
//			mask.rays[two].profilePoint.mark=true;
//			}
//		else mask.rays[two].profilePoint.mark=false;
//
//		if( (varianceTotal<outlier ) && (varianceAverage >(outlier*3) )) 
//			{
//			*b=(*a+*c)/2; // Average the left and right
//
//			// Dont forget to average the angle of the profile point aswell.
//			mask.rays[two].profilePoint.angle=(mask.rays[one].profilePoint.angle +
//											   mask.rays[three].profilePoint.angle)/2;
//			i++; // If and outlier skip, after adjusting skip to the next profile point.
//			}
//		else if((varianceLeft>outlier) && (varianceRight>outlier))
//			{
//			mask.rays[two].profilePoint.radius=0;
//			mask.rays[two].profilePoint.atomIndex=0;
//			mask.rays[two].profilePoint.coordinate=0;
//			//i++;
//			}
//	}

	delete[] radialCopy;
	delete[] averageProf;
}

void RadialScan::RemoveFinalRoughPoints(void)
{
	double *radialCopy;
	double *averageProf;
	double sample3x[3];
	int countGood;
	double countTotal;

	int index;
	bool swaped;
	double t;

	int i;
	int j;
	float angle;
	float radius;
	vec2double p;

	radialCopy = new double[MASK_RES];
	averageProf = new double[MASK_RES];

//	for(i=0;i<MASK_RES;i++) radialCopy[i] = mask.rays[i].profilePoint.radius;


	for(i=0;i<PROFILE_SIZE;i++)
		{
			if(finalProfile[i].coordinate==0) radialCopy[i]=0;
			else 
			{
				p = finalProfile[i].coordinate-mask.offset;
				p = p.toVector();
				finalProfile[i].angle=p.x;
				radialCopy[i]=p.y;
			}
		}


	for(i=0;i<PROFILE_SIZE;i++)
	{
		countGood=0;
		for(j=0;j<3;j++) 
		{

			index = i+j;
			index = index%PROFILE_SIZE;

			if(radialCopy[index]!=0) countGood++;
			else
			{
				__asm nop
			}
			sample3x[j] = radialCopy[index]; 
			
		}
		
		
		if(countGood==3) countGood=1;
		else if(countGood==2) countGood=1;
		else if(countGood==1) countGood=2;
		else countGood=0;

		swaped=true;
		index = (i+1)%PROFILE_SIZE;
		if(sample3x[1]!=0)
		{
			while(swaped)
			{
				swaped=false;
				for(j=0;j<2;j++)
				{
					if(sample3x[j]>sample3x[j+1]) {t = sample3x[j]; sample3x[j]=sample3x[j+1]; sample3x[j+1]=t; swaped=true;}
				}
			}
		
			averageProf[index]=sample3x[countGood];
		} else averageProf[index]=0;
		


	}

	countTotal=0;

	int ia;
	int ib;
	int ic;
	
	for(ia=0;ia<PROFILE_SIZE;ia++)
	{
		countTotal=0;
		countGood=0;
			

			ib=(ia+1)%PROFILE_SIZE;
			ic=(ia+2)%PROFILE_SIZE;

			if(averageProf[ia]!=0)	{ countTotal+=averageProf[ia];	countGood++;}
			if(averageProf[ib]!=0)	{ countTotal+=averageProf[ib];	countGood++;}
			if(averageProf[ic]!=0)	{ countTotal+=averageProf[ic];	countGood++;}

						
		if(countGood>0 && averageProf[ib]!=0)
		{
			finalProfile[ib].radius = (float) (countTotal / (double) countGood);
			radius = (float) finalProfile[ib].radius;
			angle  = (float) finalProfile[ib].angle;

			finalProfile[ib].coordinate.x = (sin(angle)*radius) + mask.offset.x;
			finalProfile[ib].coordinate.y = (cos(angle)*radius) + mask.offset.y;
		}
		else
		{
			finalProfile[ib].radius=0; 
			finalProfile[ib].coordinate=0;
		}
	}
	

//	int i;
//	double *a,*b,*c;	// Profile points to check
//	int one, two, three;
//	double outlier = 12;//12; 	// Normally 4 How far apart do they have to be not to be rough
							// and multiplied by two for when they are rough
//	double varianceLeft;
//	double varianceRight;

	//Reconstruct the radius data

//	for(i=0;i<PROFILE_SIZE;i++)
//		{
//		if(finalProfile[i].coordinate==0) finalProfile[i].radius=0;
//		else finalProfile[i].radius=DistOfTwoPoints(mask.offset,finalProfile[i].coordinate);
//		}

//	for(i=0;i<PROFILE_SIZE;i++)
//		{
//		one=(i+(MASK_RES-1))%PROFILE_SIZE; // Find index on left, loop around from 0 to MASK_RES indes if -1
//		two=i; // index of rough point to adjust if needed
//		three=(i+1)%PROFILE_SIZE; // Find Index on Right, loop around from MASK_RES to 0 if MASK_RES

//		a=&(finalProfile[one].radius);   // pointer to the profile points
//		b=&(finalProfile[two].radius);   // speeds up access.
//		c=&(finalProfile[three].radius); //
		// If the left and right profile points are close enough together and the centre is
		// is outlier then make centre the average of the left and right.
//		varianceLeft=fabs(*a-*b);
//		varianceRight=fabs(*b-*c);
//		if((varianceLeft<=12) || (varianceRight<=12)) //Normally 2 //:Try 12
//			{
			//if(showPutPixels) AddPutPixels(finalProfile[i].coordinate,0,255,255);
//			finalProfile[two].mark=true;
//			}
		
//		else finalProfile[two].mark=false; //PCN3122 mark water level to egnore centre calculations (12 Nov 2004, Antony)
//		if(ignoreWaterLevel && IsInWaterSection(two)) finalProfile[two].mark=false; //PCN????
//
//		if( (varianceLeft>outlier) && (varianceRight>outlier))
//			{
//			finalProfile[two].radius=0;
//			finalProfile[two].atomIndex=0;
//			finalProfile[two].coordinate=0;
//			}
//		else if((varianceLeft<outlier) && (varianceRight<outlier) && (*b==0))
//			{
//			*b=(*a+*c)/2;
//			finalProfile[two].coordinate=(finalProfile[one].coordinate+finalProfile[three].coordinate)/2;
//			}
//		}
}

void RadialScan::SmoothLikeABabysBottom()
{
	
	vec2double vector;
	double profBuf[PROFILE_SIZE];
	int i;
	int loop;
	int r1i,r3i;
	double r1,r2,r3;

	// Reconstruct radius and angles /////////////////////
	for(i=0;i<PROFILE_SIZE;i++)							//
	{													//
		vector = GetVector(finalProfile[i].coordinate);	//
		finalProfile[i].angle = vector.x;				//
		finalProfile[i].radius = vector.y;				//
	}													//
	/////////////////////////////////////////////////////

	//PCN2737 Removing of rough points find if the nextdoor points are jaged or not, /////
	// if are then average it. (Antony 8 June, 2004, Sorry for lack of comments)		//
	// The more this is looped thru the smoother the profile gets and eliminates rough points.
	for(loop=0;loop<5;loop++)															//
		{																				//
		for(i=0;i<PROFILE_SIZE;i++) profBuf[i]=-1; // Clear the profile changes, all -1.//
																						//
		// Loops thru all points then store the ones that need to cange					//
		for(i=0;i<PROFILE_SIZE;i++)														//
			{																			//
			// Find its neighbouring points in the profile ring. Left r1i, right r3i	//
			r1i=(i+(MASK_RES-1))%PROFILE_SIZE;
			r3i=(i+1)%PROFILE_SIZE;			
																						//
			r1=finalProfile[r1i].radius;	// Retrieve profile point to left.						//
			r2=finalProfile[i].radius;		// Retrieve current profile point						//
			r3=finalProfile[r3i].radius;	// Retrieve profile point to right.						//
																						//
			if((r1==0) || (r2==0) || (r3==0)) continue; // If anyoff then are 0 get next points
																						//
			// If points look like -_- or _-_ then average the middle one				//
			if(((r1<r2) && (r3<r2)) || ((r1>r2) && (r3>r2)))							//
				profBuf[i]=(r1+r3)/2;													//
			}																			//
		// Loop thru and copy any changed profile points.								//
		for(i=0;i<PROFILE_SIZE;i++) if(profBuf[i]!=-1) finalProfile[i].radius=profBuf[i];											//
		}																				//
	//////////////////////////////////////////////////////////////////////////////////////

	for(i=0;i<PROFILE_SIZE;i++)
	{
		finalProfile[i].coordinate = 
			GetCoordinate(vec2double(finalProfile[i].angle,finalProfile[i].radius));
		if(finalProfile[i].coordinate!=0) finalProfile[i].coordinate=finalProfile[i].coordinate+mask.offset;
	}

}

void RadialScan::AddLaserWidthOverlay(vec2int coord)
{
	if(noLaserWidthOverlay>=MASK_RES*100) return;
	laserWidthOverlay[noLaserWidthOverlay++]=coord;
}

void RadialScan::AddProfileCandidatesOverlay(vec2int coord)
{
	if(noProfileCandidatesOverlay>=MASK_RES*100) return;
	profileCandidatesOverlay[noProfileCandidatesOverlay++]=coord;
}

void RadialScan::AddPutPixels(vec2double coord, unsigned char red, unsigned char green, unsigned char blue)
{
	if(noPutPixels>=20000) return;
	putPixels[noPutPixels].coord=coord;
	putPixels[noPutPixels].red=red;
	putPixels[noPutPixels].green=green;
	putPixels[noPutPixels].blue=blue;
	noPutPixels++;
}

void RadialScan::AddDrawLines(vec2double coord, unsigned char red, unsigned char green, unsigned char blue)
{
	if((noDrawLines>=20000) || (noDrawLines<0)) return;
	drawLines[noDrawLines].coord=coord;
	drawLines[noDrawLines].red  =red;
	drawLines[noDrawLines].green=green;
	drawLines[noDrawLines].blue  =blue;
	noDrawLines++;
}

////////////////////////////////////////////////////////////////////////////////////
// Name: RadialScan::ShowLaserWidthOverlay
// Created By: Antony van Iersel
// Date: 27 August 2004
// Description:	Display the laser width overlay to show where the profiler is scanning
// Input: None
// Output: None
////////////////////////////////////////////////////////////////////////////////////

void RadialScan::ShowLaserWidthOverlay(void)
{
	int x,y,i;
	for(i=0;i<noLaserWidthOverlay;i++)
	{
		x=(int) (laserWidthOverlay[i].x+mask.offset.x+0.5);
		y=(int) (laserWidthOverlay[i].y+(mask.offset.y/imRatio)+0.5);
		if((x<0) || (x>=imWidth) || (y<0) || (y>=imHeight)) continue; //ANT VOB error
		imVideo[y][x].green=180;
		//imVideo[y][x].blue=imVideo[y][x].blue/2;
		//imVideo[y][x].red=imVideo[y][x].red/2;
	}
}

////////////////////////////////////////////////////////////////////////////////////
// Name: RadialScan::ShowProfileCandidatesOverlay
// Created By: Antony van Iersel
// Date: 27 August 2004
// Description:	Display the laser width overlay to show where the profiler is scanning
// Input: None
// Output: None
////////////////////////////////////////////////////////////////////////////////////

void RadialScan::ShowProfileCandidatesOverlay(void)
{
	int x,y,i;
	for(i=0;i<noProfileCandidatesOverlay;i++)
	{
		x=(int) (profileCandidatesOverlay[i].x+mask.offset.x+0.5);
		y=(int) (profileCandidatesOverlay[i].y+(mask.offset.y/imRatio)+0.5);
		if((x<0) || (x>=imWidth) || (y<0) || (y>=imHeight)) continue;  //ANT VOB error
		imVideo[y][x].red=255;
		imVideo[y][x].blue=66;
		imVideo[y][x].green=255;
	}
}

void RadialScan::ShowPutPixels(void)
{
	int x,y,i;
	for(i=0;i<noPutPixels;i++)
	{
		x=(int) (putPixels[i].coord.x+0.5);
		y=(int) ((putPixels[i].coord.y/imRatio)+0.5);

		if((x<0) || (x>=imWidth) || (y<0) || (y>=imHeight)) continue;  //ANT VOB error

		imVideo[y][x].red=putPixels[i].red;
		imVideo[y][x].green=putPixels[i].green;
		imVideo[y][x].blue=putPixels[i].blue;

//		if(imVideo[y][x].red<240) imVideo[y][x].red+=20;
//		if(imVideo[y][x].green<240) imVideo[y][x].green+=20;
//		if(imVideo[y][x].blue<240) imVideo[y][x].blue+=20;
	}

//	for(i=0;i<noPutPixels;i++)
//	{
//		x=(int) (putPixels[i].coord.x+mask.offset.x+0.5);
//		y=(int) (((putPixels[i].coord.y+mask.offset.y)/imRatio)+0.5);
//		if((x<0) || (x>imWidth) || (y<0) || (y>imHeight)) continue;
//		imVideo[y][x].red=putPixels[i].red;
//		imVideo[y][x].green=putPixels[i].green;
//		imVideo[y][x].blue=putPixels[i].blue;
//	}
}

void RadialScan::ShowDrawLines(void)
{
	vec2double coord_a;
	vec2double coord_b;
	unsigned char red,green,blue;
	

	int i;
	if(noDrawLines<2) return;

	for(i=0;i<noDrawLines-1;i++)
		{
		coord_a=drawLines[i].coord;
		coord_b=drawLines[i+1].coord;
		red=drawLines[i].red;
		green=drawLines[i].green;
		blue=drawLines[i].blue;
		DrawLine(coord_a,coord_b,red,green,blue);
		}
}

void RadialScan::ShowInternalCircle(vec2double cen, double size, int red, int green, int blue, int centered)
{
	double i;
	int x,y;
	vec2double coord;
	for(i=0;i<(2*PI);i+=(PI/540))
	{
		coord.x=sin(i)*size;
		coord.y=cos(i)*size;
		if(centered) { coord.x+=(imWidth/2); coord.y+=(imHeight/2); }
		else         { coord.x+=cen.x; coord.y+=cen.y; }
		y=(int) ((coord.y/imRatio)+0.5);
		x=(int) (coord.x+0.5);
		

		if((x<0) || (x>=imWidth) || (y<0) || (y>=imHeight)) continue; 
		imVideo[y][x].red=(unsigned char) red;
		imVideo[y][x].green=(unsigned char) green;
		imVideo[y][x].blue=(unsigned char) blue;
	}
}

void RadialScan::ShowVideoFilter(void)
{
	int x,y;
	unsigned char videoPixel;

	if(redAdjust==0)
	{
		for(x=0;x<imWidth;x++)
			for(y=0;y<imHeight;y++)
				imVideo[y][x].red = imVideo[y][x].green = imVideo[y][x].blue = FilterVideoPixel(y,x);
				
	}
	else
	{
		for(x=0;x<imWidth;x++)
			for(y=0;y<imHeight;y++)
			{
				videoPixel=FilterVideoPixel(y,x);
				copyBuffer[y][x].red = copyBuffer[y][x].green = copyBuffer[y][x].blue = videoPixel;
				
				
			}
		for(x=0;x<imWidth;x++)
			for(y=0;y<imHeight;y++)
				{
					imVideo[y][x] = copyBuffer[y][x];
				}



	}
	

}

void RadialScan::PrintProfile(char *file)
{
	int i;
	double radius;
	vec2double coord;
	double angle;
	double atom;



	FILE *f=NULL;
	f=fopen(file,"w");

	if(f!=NULL)
	{
		for(i=0;i<MASK_RES;i++)
		{
			radius = mask.rays[i].profilePoint.radius;
			coord = mask.rays[i].profilePoint.coordinate;
			angle = mask.rays[i].profilePoint.angle*(360/2/PI);
			atom = mask.rays[i].profilePoint.atomIndex;
			fprintf(f,"%f , %f , %f , %f , %i \n" , radius, coord.x, coord.y, angle, atom);

		}
	}		
	fclose(f);
}

void RadialScan::DrawLine(vec2double coord_a, vec2double coord_b, unsigned char red, unsigned char green, unsigned char blue)
{
	int x,y;
	double a,b;
	double x_size, y_size;
	double largest;
	double step_x;
	double step_y;
	double i;

	x_size=coord_b.x-coord_a.x;
	y_size=coord_b.y-coord_a.y;
	if(fabs(x_size)>fabs(y_size)) largest=fabs(x_size);
	else                          largest=fabs(y_size);

	step_x=x_size/largest;
	step_y=y_size/largest;

	a=coord_a.x;
	b=coord_a.y;

	for(i=0;i<largest;i++)
	{
		x = (int) (a+0.5);
		y = (int) (b+0.5);
		if((x>=imWidth) || (x<0) || (y>=imHeight) || (y<0)) //ANT VOB
		{
			a=step_x;
			b=step_y;
			continue;
		}
		imVideo[y][x].red = red;
		imVideo[y][x].green = green;
		imVideo[y][x].blue = blue;
		a+=step_x;
		b+=step_y;
	}
}

inline vec2double RadialScan::GetCoordinate(vec2double vector)
{
	return vec2double(sin(vector.x)*vector.y,cos(vector.x)*vector.y);
}

vec2double RadialScan::GetVector(vec2double coord)
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

////////////////////////////////////////////////////////////////////////////////////
// PCN3122
// Name: RadialScan::RecalculateAverageCentre
// Created By: Antony van Iersel
// Date: 12 November 2004
// Description:	Takes the five last centres and calculates teh average position
// Usage: This is used when the centre calculation fails, this is used instead.
// Input: None
// Output: None
////////////////////////////////////////////////////////////////////////////////////

void RadialScan::RecalculateAverageCentre(void)
{
	int i,count;
	count=0;
	averageCentreHistory=0;
	for(i=0;i<5;i++)
	{
		if(centreHistory[i]==0) continue;
		averageCentreHistory=averageCentreHistory+centreHistory[i];
		count++;
	}
	averageCentreHistory=averageCentreHistory/count;
}
/*
		int loop;
	//PCN2737 Removing of rough points find if the nextdoor points are jaged or not, /////
	// if are then average it. (Antony 8 June, 2004, Sorry for lack of comments)		//
	// The more this is looped thru the smoother the profile gets and eliminates rough points.
	for(loop=0;loop<5;loop++)															//
		{																				//
		for(i=0;i<PROFILE_SIZE;i++) profBuf[i]=-1; // Clear the profile changes, all -1.//
																						//
		// Loops thru all points then store the ones that need to cange					//
		for(i=0;i<PROFILE_SIZE;i++)														//
			{																			//
			// Find its neighbouring points in the profile ring. Left r1i, right r3i	//
			r1i=(i+(MASK_RES-1))%PROFILE_SIZE;
			r3i=(i+1)%PROFILE_SIZE;			
																						//
			r1=finalProfile[r1i].radius;	// Retrieve profile point to left.						//
			r2=finalProfile[i].radius;		// Retrieve current profile point						//
			r3=finalProfile[r3i].radius;	// Retrieve profile point to right.						//
																						//
			if((r1==0) || (r2==0) || (r3==0)) continue; // If anyoff then are 0 get next points
																						//
			// If points look like -_- or _-_ then average the middle one				//
			if(((r1<r2) && (r3<r2)) || ((r1>r2) && (r3>r2)))							//
				profBuf[i]=(r1+r3)/2;													//
			}																			//
		// Loop thru and copy any changed profile points.								//
		for(i=0;i<PROFILE_SIZE;i++) if(profBuf[i]!=-1) finalProfile[i].radius=profBuf[i];											//
		}																				//
	//////////////////////////////////////////////////////////////////////////////////////

	for(i=0;i<PROFILE_SIZE;i++)
	{
		finalProfile[i].coordinate.x=sin(finalProfile[i].angle)*finalProfile[i].radius;
		finalProfile[i].coordinate.y=cos(finalProfile[i].angle)*finalProfile[i].radius;
	}
*/

void RadialScan::CreateGausianMask(void)
{
GausianMask[0][0] = 0.00000067; GausianMask[1][0] = 0.00002292; GausianMask[2][0] = 0.00019117; GausianMask[3][0] =	0.00038771; GausianMask[4][0] =	0.00019117; GausianMask[5][0] =	0.00002292; GausianMask[6][0] =	0.00000067;
GausianMask[0][1] = 0.00002292; GausianMask[1][1] =	0.00078633; GausianMask[2][1] =	0.00655965; GausianMask[3][1] =	0.01330373; GausianMask[4][1] =	0.00655965; GausianMask[5][1] =	0.00078633; GausianMask[6][1] =	0.00002292;
GausianMask[0][2] = 0.00019117; GausianMask[1][2] =	0.00655965; GausianMask[2][2] =	0.05472157; GausianMask[3][2] =	0.11098164; GausianMask[4][2] =	0.05472157; GausianMask[5][2] =	0.00655965; GausianMask[6][2] =	0.00019117;
GausianMask[0][3] = 0.00038771; GausianMask[1][3] = 0.01330373; GausianMask[2][3] =	0.11098164; GausianMask[3][3] =	0.22508352; GausianMask[4][3] =	0.11098164; GausianMask[5][3] =	0.01330373; GausianMask[6][3] =	0.00038771;
GausianMask[0][4] = 0.00019117; GausianMask[1][4] =	0.00655965; GausianMask[2][4] =	0.05472157; GausianMask[3][4] =	0.11098164; GausianMask[4][4] =	0.05472157; GausianMask[5][4] =	0.00655965; GausianMask[6][4] =	0.00019117;
GausianMask[0][5] = 0.00002292; GausianMask[1][5] =	0.00078633; GausianMask[2][5] =	0.00655965; GausianMask[3][5] =	0.01330373; GausianMask[4][5] =	0.00655965; GausianMask[5][5] =	0.00078633; GausianMask[6][5] =	0.00002292;
GausianMask[0][6] = 0.00000067; GausianMask[1][6] =	0.00002292; GausianMask[2][6] =	0.00019117; GausianMask[3][6] =	0.00038771; GausianMask[4][6] =	0.00019117; GausianMask[5][6] =	0.00002292; GausianMask[6][6] =	0.00000067;
	
}

void RadialScan::FindMedianXCentre(vec2double &centre)
{

	int i;
	int countGoodPoints = -1;
	double aveX=0;
	double aveY=0;
	double xvalues[180];
	double yvalues[180];
	double temp;
	bool swaped = true;
	
	for(i=0;i<180;i++)
	{
		if(!finalProfile[i].mark) continue;
		if(finalProfile[i].coordinate==0) continue;
		countGoodPoints++;
		xvalues[countGoodPoints] = finalProfile[i].coordinate.x;
		yvalues[countGoodPoints] = finalProfile[i].coordinate.y;
	}

	if(countGoodPoints == -1) {centre.x=0; centre.y=0; return;}
	
	while(swaped)
	{
		swaped = false;
		for(i=1;i<=countGoodPoints;i++)
		{
			if(xvalues[i]<xvalues[i-1]) 
				{
					temp = xvalues[i];
					xvalues[i]=xvalues[i-1];
					xvalues[i-1]=temp;
					swaped = true;
				}
		}
	}
	swaped = true;
	while(swaped)
	{
		swaped = false;
		for(i=1;i<=countGoodPoints;i++)
		{
			if(yvalues[i]<yvalues[i-1]) 
			{
				temp = yvalues[i];
				yvalues[i]=yvalues[i-1];
				yvalues[i-1]=temp;
				swaped = true;
			}

		}
	}
	
	if(countGoodPoints > 6) centre.x = (xvalues[3]+xvalues[countGoodPoints-3])/2;
	//if(countGoodPoints > 6) x = (xvalues[3])+35;
	//if(countGoodPoints > 9) x = (xvalues[countGoodPoints-6])-35;

	centre.y = 0;
}

