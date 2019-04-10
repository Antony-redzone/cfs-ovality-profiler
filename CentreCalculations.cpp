#include <windows.h>
#include <stdio.h>


#include "Shapes.h"
#include "CentreCalculations.h"



Centre::Centre(float *_pvDataX, 
			   float *_pvDataY,
			   float *_pvCentreX,
			   float *_pvCentreY,
			   int *_egnoreList,
			   int _fromFrame,
			   int _toFrame,
			   int _waterLevelCentre,
			   int _edgeCentre)
{
	hwnd=0;
	pvDataX = _pvDataX;
	pvDataY = _pvDataY;
	pvCentreX = _pvCentreX;
	pvCentreY = _pvCentreY;
	passedEgnoreList = _egnoreList;
	fromFrame = _fromFrame;
	toFrame = _toFrame;

	shapeRadius   = 0;
	shapeCentreX  =	0;
	shapeCentreY  =	0;
	shapeRotation =	0;


	screenWidth = 0;
	screenHeight = 0;
	screenRatio = 0;
	screenCentre.x = 0;
	screenCentre.y = 0;

	shape = 0;

	if(_waterLevelCentre==1) waterLevelCentre = true;
	else waterLevelCentre = false;

	if(_edgeCentre!=0) getEdgeCentre = true;
	else getEdgeCentre = false;
	
	if(_waterLevelCentre==2) reduceWaterLevelData = true;
	else reduceWaterLevelData = false;

	if(_waterLevelCentre==3) cutWaterLevel = true;
	else cutWaterLevel = false;

	removeWaterLevelData = false;
	

}

Centre::Centre(ReferenceShape_V10 *_Shape,
		   double _ShapeRadius,
		   double _ShapeCentreX,
		   double _ShapeCentreY,
		   double _ShapeRotation,
		   float *_pvDataX, 
		   float *_pvDataY,
		   float *_pvCentreX,
		   float *_pvCentreY,
		   int *_egnoreList,
		   int _fromFrame,
		   int _toFrame,
		   HWND _hwnd,
		   float _screenWidth,
		   float _screenHeight,
		   double _screenRatio)
{
	hwnd = _hwnd;
	pvDataX = _pvDataX;
	pvDataY = _pvDataY;
	pvCentreX = _pvCentreX;
	pvCentreY = _pvCentreY;
	passedEgnoreList = _egnoreList;
	fromFrame = _fromFrame;
	toFrame = _toFrame;

	shapeRadius   = _ShapeRadius;
	shapeCentreX  =	_ShapeCentreX;
	shapeCentreY  =	_ShapeCentreY;
	shapeRotation =	_ShapeRotation;


	screenWidth = _screenWidth;
	screenHeight = _screenHeight;
	screenRatio = _screenRatio;
	screenCentre.x = _screenWidth / 2;
	screenCentre.y = _screenHeight / 2;

	shape = new Shapes(_Shape,
					   _ShapeRadius,
					   _ShapeCentreX,
					   _ShapeCentreY,
					   _ShapeRotation);

	waterLevelCentre = false;
	getEdgeCentre = false;
	cutWaterLevel = false;
	removeWaterLevelData = false;

}

Centre::~Centre(void)
{
	if(shape!=0) {delete shape; shape = 0;}
}

void Centre::CalculateCentre(void)
{
	
	int i;
	useWaterLevel=false;
	for(i=0;i<180;i++) if(passedEgnoreList[i]==1) {useWaterLevel=true; break;}

	for(currentFrame=(fromFrame-1);currentFrame<toFrame;currentFrame++) 
		CalculateCentre(currentFrame);
}

void Centre::CalculateCentre(long frameNumber)
{
	double centreX;
	double centreY;

	indexProfileOne = (frameNumber*180)+1; // Profile points are from 1 to 180 inclusive
	
	//if(useWaterLevel) CreateFilteredPoints(indexProfileOne);	//Create the filtered point to calculate the Ovality
	//esle CreateFakePoints(frameNumber);

	CreateFilteredPoints(indexProfileOne);




	MarkRoughPoints();
	
	
	if(!getEdgeCentre || (waterLevelCentre == false)) CentreCalculation(centreX,centreY);
	if(getEdgeCentre) CentreCalculationTwo(centreX,centreY);
	
	pvCentreX[frameNumber] -= (float) (centreX);
	pvCentreY[frameNumber] -= (float) (centreY);


	//pvCentreY[frameNumber] = pvCentreY[frameNumber]; //PCNANT Centre Changed

	if(waterLevelCentre) pvCentreY[frameNumber] = 0;
}

void Centre::CentreCalculationTwo(double &x, double &y)
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
		if(fakePoints[i]==0) continue;
		countGoodPoints++;
		xvalues[countGoodPoints] = fakePoints[i].x;
		yvalues[countGoodPoints] = fakePoints[i].y;
	}

	if(countGoodPoints == -1) {x=0; y=0; return;}
	
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
	
	if(countGoodPoints > 6) x = (xvalues[3]+xvalues[countGoodPoints-3])/2;
	//if(countGoodPoints > 6) x = (xvalues[3])+35;
	//if(countGoodPoints > 9) x = (xvalues[countGoodPoints-6])-35;

	y = y;


		
}

void Centre::CreateFakePoints(long frameNumber)
{
	long index=(frameNumber*180)+1;
	long pointIndex;
	long pointLoop;

	for(pointLoop=0;pointLoop<180;pointLoop++)
	{
		egnoreList[pointLoop] = passedEgnoreList[pointLoop];
		pointIndex=pointLoop+index;
//		if(egnoreList[pointLoop]==1) {pvDataX[pointIndex]=0; pvDataY[pointIndex]=0;}
		if((pvDataX[pointIndex]==0) && (pvDataY[pointIndex]==0)) fakePoints[pointLoop]=0;
		else 
		{
			fakePoints[pointLoop].x = pvDataX[pointIndex] + pvCentreX[frameNumber];
			fakePoints[pointLoop].y = pvDataY[pointIndex] + pvCentreY[frameNumber];
		}
	}
}

void Centre::CentreCalculation(double &x, double &y)
{
	vec2double centre;

	centre.x=0;
	centre.y=0;


	FindBestNeighbour(centre,64); //PCN3122 Old final centre, still works rather good.
	x = centre.x;
	y = centre.y;
}

void Centre::FindBestNeighbour(vec2double &curPoint, double size)
{
	if(size<0.015625) return;
	double variance;
	double closestVar;
	bool iTBN=false; // is there better neighbour
	vec2double closestPoint;
	vec2double lookingAt;

	closestPoint=curPoint;
	closestVar=GetCentreVariance(curPoint);

	lookingAt.x=curPoint.x-size;
	lookingAt.y=curPoint.y-size;
	variance=GetCentreVariance(lookingAt);
	if(variance<closestVar) {closestPoint=lookingAt; closestVar=variance; iTBN=true;}

	lookingAt.x=curPoint.x;
	lookingAt.y=curPoint.y-size;
	variance=GetCentreVariance(lookingAt);
	if(variance<closestVar) {closestPoint=lookingAt; closestVar=variance; iTBN=true;}

	lookingAt.x=curPoint.x+size;
	lookingAt.y=curPoint.y-size;
	variance=GetCentreVariance(lookingAt);
	if(variance<closestVar) {closestPoint=lookingAt; closestVar=variance; iTBN=true;}

	lookingAt.x=curPoint.x+size;
	lookingAt.y=curPoint.y;
	variance=GetCentreVariance(lookingAt);
	if(variance<closestVar) {closestPoint=lookingAt; closestVar=variance; iTBN=true;}

	lookingAt.x=curPoint.x+size;
	lookingAt.y=curPoint.y+size;
	variance=GetCentreVariance(lookingAt);
	if(variance<closestVar) {closestPoint=lookingAt; closestVar=variance; iTBN=true;}

	lookingAt.x=curPoint.x;
	lookingAt.y=curPoint.y+size;
	variance=GetCentreVariance(lookingAt);
	if(variance<closestVar) {closestPoint=lookingAt; closestVar=variance; iTBN=true;}

	lookingAt.x=curPoint.x-size;
	lookingAt.y=curPoint.y+size;
	variance=GetCentreVariance(lookingAt);
	if(variance<closestVar) {closestPoint=lookingAt; closestVar=variance; iTBN=true;}
	
	FindBestNeighbour(closestPoint, size/2);
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
double Centre::GetCentreVariance(vec2double p)
{
	int i;
	double variance=0; // max & min holds the current largest and smallest distance from
	double avgDist=0;
	int count=0;
	double distance[180];
	vec2double ortho;
	double orthoDistance;

	if(shape==0)
	{
		for(i=0;i<180;i++)
			{
			//if(egnoreList[i]==0) PCN4363 Opps, this should never have been in, the water level is filled in now 
				if(fakePoints[i]!=0)
					{
					distance[count]=DistOfTwoPoints(p,fakePoints[i]);
					avgDist+=distance[count];
					count++;
					}
			}
	}else
	{
		for(i=0;i<180;i++)
			{
			if(egnoreList[i]==0)
				if(fakePoints[i]!=0)
					{
					shape->ProfileRefShapeDistCalc((float) fakePoints[i].x,(float) -fakePoints[i].y,
													&ortho.x, &ortho.y, &orthoDistance);
					distance[count]=fabs(orthoDistance);
					avgDist+=distance[count];
					count++;
					}
			}
	}
	if(count==0) return 0;

	avgDist/=(double) count;
	for(i=0;i<count;i++) variance+=fabs(avgDist-distance[i]);
	
	variance/=(double) count;
	return variance; // return the average variance
}

inline double Centre::DistOfTwoPoints(vec2double pt1, vec2double pt2)
{
	return sqrt(pow(pt1.x-pt2.x,2)+pow(pt1.y-pt2.y,2));
}

void Centre::MarkRoughPoints(void)
{
	int i;
	double *a,*b,*c;	// Profile points to check
	int one, two, three;
	double outlier = 120;//12; 	// Normally 4 How far apart do they have to be not to be rough
							// and multiplied by two for when they are rough
	double varianceLeft;
	double varianceRight;

	//Reconstruct the radius data

	for(i=0;i<180;i++)
		{
		if(fakePoints[i]==0) radiusPoints[i]=0;
		else radiusPoints[i]=DistOfTwoPoints(0,fakePoints[i]);
		}

	for(i=0;i<180;i++)
		{
		one=(i+(180-1))%180; // Find index on left, loop around from 0 to MASK_RES indes if -1
		two=i; // index of rough point to adjust if needed
		three=(i+1)%180; // Find Index on Right, loop around from MASK_RES to 0 if MASK_RES

		a=&(radiusPoints[one]);   // pointer to the profile points
		b=&(radiusPoints[two]);   // speeds up access.
		c=&(radiusPoints[three]); //
		// If the left and right profile points are close enough together and the centre is
		// is outlier then make centre the average of the left and right.
		varianceLeft=fabs(*a-*b);
		varianceRight=fabs(*b-*c);
		if((varianceLeft<=120) || (varianceRight<=120)) //Normally 2 //:Try 12 *10 for Clearline calcs
			{
			//if(showPutPixels) AddPutPixels(finalProfile[i].coordinate,0,255,255);
			//egnoreList[two]=1;
			}
		else egnoreList[two]=1;
		
		//else [two].mark=false; //PCN3122 mark water level to egnore centre calculations (12 Nov 2004, Antony)
		//if(ignoreWaterLevel && IsInWaterSection(two)) finalProfile[two].mark=false; //PCN????

		}
}

void Centre::CreateFilteredPoints(long index)
{
	long point;
	long pointIndex;
	vec2double coord;

//	if(currentFrame == 813)
//	{
//		__asm nop
//	}

	for(point=0;point<180;point++)
	{
		pointIndex = point+index;

		if(pvDataX[pointIndex]==0 && pvDataY[pointIndex]==0) holes[point]=1;
		else holes[point]=0;

		egnoreList[point] = passedEgnoreList[point];

		if(pvDataX[pointIndex]==0 && pvDataY[pointIndex]==0) fakePoints[point]=0;
		else if(egnoreList[point]==1) fakePoints[point]=0;
		else fakePoints[point] = vec2double(pvDataX[pointIndex]+pvCentreX[currentFrame],pvDataY[pointIndex]+pvCentreY[currentFrame]);
		if((fakePoints[point].x>10000) || (fakePoints[point].x<-10000)) 
		{
			fakePoints[point].x = 0;
			fakePoints[point].y = 0;
			holes[point]=1;
		}
	}
	


	CreateFilledHoles(); // Where there is a whole fill in as best as we can to avoid false reading
	


//	if(hwnd!=0) 
//		for(point=0;point<180;point++)
//		{
//				DrawLine(fakePoints[point], vec2double(0,0), Color::Yellow);
//		}

//////////////////////////////////////////////////////////////////////////
//																		//			
//	Testing, put the data back as pvDataY to show resulting fill in		//
/*
	for(point=0;point<180;point++)										//
	{		
	
		//if(hwnd!=0) DrawCircle(fakePoints[point],1,Color::Black);
		//if(point==89 && hwnd!=0) DrawCircle(fakePoints[point],2,Color::Black);

		pointIndex = point+index;										//
		if(((pvDataX[pointIndex]==0) && pvDataY[pointIndex]==0) || egnoreList[point]==1)			//
		{	
			if((fakePoints[point].x!=0) || (fakePoints[point].y!=0))
			{
				if(fakePoints[point].x >=0) pvDataX[pointIndex] = (float) fakePoints[point].x + 20000;	//
				else pvDataX[pointIndex] = (float) fakePoints[point].x - 20000;
				pvDataY[pointIndex] = (float) fakePoints[point].y;
				pvDataX[pointIndex] = pvDataX[pointIndex] - pvCentreX[currentFrame];
				pvDataY[pointIndex] = pvDataY[pointIndex] - pvCentreY[currentFrame];

			}
			else
			{
				pvDataX[pointIndex] = 0;
				pvDataY[pointIndex] = 0;
			}

		}																//	
	}	
*/

//	for(point=0;point<180;point++)
//	{
//		if(egnoreList[point]==1) 
//		{
//			pointIndex = point + index; 
//			pvDataX[pointIndex]=0;
//			pvDataY[pointIndex]=0;
//		}
//	}
//
//																		//																			
//////////////////////////////////////////////////////////////////////////


//	//Find first good point;
//	for(point=0;point<180;point++)
//	{
//		radiusPoints[point] = fakePoints[point].toVector().y;
//	}
//
//	for(point=1;point<179;point++)
//	{
//		
//		filteredPoints[point] = FilterThreePoints(radiusPoints[point-1],
//												 radiusPoints[point],
//												 radiusPoints[point+1]);
//	}
//
//	filteredPoints[0] = FilterThreePoints(radiusPoints[179],
//										 radiusPoints[0],
//										 radiusPoints[1]);
//
//	filteredPoints[179] = FilterThreePoints(radiusPoints[178],
//										 radiusPoints[179],
//										 radiusPoints[0]);

}



void Centre::CreateFilledHoles(void)
{
	int i,j;
	double averageShift=0;
	double numberOfPairs=0;
	double shift=0;
	vec2double a,b;

	if(currentFrame == 1265) 
	{
		__asm nop
	}

	for(i=136,j=0;i<223;i++,j++) botPoints[j] = i%180;

	//Fill it any missing top points
	for(i=0;i<180;i++) if(egnoreList[i]==1) fakePoints[i]=0;

	for(i=0;i<180;i++)
	{
		b = fakePoints[i]; if(b==0) FindTopWhole(i);
	}
	if(cutWaterLevel) return; //PCN4538 actually just filling in the data as old, but not mirror the top

	if(reduceWaterLevelData) for(i=30;i<150;i++) egnoreList[i]=0;

	for(i=0;i<180;i++) if(egnoreList[i]==1) fakePoints[i]=0;

	if(removeWaterLevelData) return; //PCN4538	


	for(i=1;i<86;i++) 
	{
		if(holes[botPoints[i]]==1) 
		{
			fakePoints[botPoints[i]]=0; //Fill in the original whole but only on the buttom
		}
	}

	for(i=0;i<87;i++)
	{
		b = fakePoints[botPoints[i]];
		if(b==0)
		{
			i = FindHole(i);
		}
	}
	//if(reduceWaterLevelData) for(i=0;i<180;i++) egnoreList[i] = passedEgnoreList[i];

}

void Centre::FindTopWhole(int i)
{
	vec2double a;
	int j=0;
	int left =i ,right = i;
	for(left = i ;j<180;left++,j++)	
	{
		a = fakePoints[left%180]; 
		if(a!=0) break; 
	}
	if(j>175) return; //Not enough data to fix
	j=0;
	for(right = i;j<180;right--,j++) 
	{ 
		a = fakePoints[(right+180)%180]; 
		if(a!=0) break; 
	}
	FillTopWhole(left%180, (right+180)%180);
}

void Centre::FillTopWhole(int left, int right)
{
	vec2double a;
	vec2double b;
	vec2double fill;

	int numberHoles;
	int hole;
	int i;
	double grad;
	double sliceSize;

	if(left < right) numberHoles = (left+180) - right;
	else  numberHoles = left-right;
	if(numberHoles > 175 || numberHoles == 0) 
	{
		return; //Not enough data to fill
	}

	a = fakePoints[left];
	b = fakePoints[right];
	
	a = a.toVector();
	b = b.toVector();

	
	grad = (a.y-b.y)		/ numberHoles;
	if(b.x>a.x) sliceSize = ((a.x+(2*PI))-b.x) / numberHoles;
	else sliceSize = (a.x - b.x) / numberHoles;
	
	for(hole=right+1,i=1;i<numberHoles;hole++,i++)
	{
		fill.x = b.x+(i*sliceSize);
		fill.y = b.y+(i*grad);
		fakePoints[hole%180]=fill.toCoordinate();
	}
}

int Centre::FindHole(int i)
{
	
	vec2double a,b;
	int left=0, right;

	vec2double targetLeft;
	vec2double targetRight;
	vec2double sourceLeft;
	vec2double sourceRight;
	double shift;

	right = i;
	for(;i<87;i++)
	{
		a = fakePoints[botPoints[i]];
		if(a!=0) break;
	}

	left = i;
	if(left==87) return left;
	if(right==0) return left;
	right+=-1;

	sourceLeft  = fakePoints[botPoints[left]];
	sourceRight = fakePoints[botPoints[right]];

	shift = 0;
	targetRight = GetProfileIntersection(vec2double(sourceRight.x+shift,sourceRight.y),
											 vec2double(sourceRight.x+shift,sourceRight.y*-1));
	if(targetRight==0)
	{
		targetRight.x=sourceRight.x;
		targetRight.y=0;
		__asm nop
	}

	shift = 0;
	targetLeft  = GetProfileIntersection(vec2double(sourceLeft.x+shift,sourceLeft.y),
											 vec2double(sourceLeft.x, sourceLeft.y *-1));
	if(targetLeft==0)
	{
		targetLeft.x=sourceLeft.x;
		targetLeft.y=0;
	}
	

	//if(hwnd!=0)
	//{
	//	DrawCircle(sourceLeft,4,Color::Maroon);
	//	DrawCircle(sourceRight,4,Color::Fuchsia);
	//	DrawLine(fakePoints[botPoints[left]], vec2double(0,0), Color::Green);
	//	DrawLine(fakePoints[botPoints[right]], vec2double(0,0), Color::Blue);
	//	DrawCircle(targetRight,7,Color::Blue);
	//	DrawCircle(targetLeft,7,Color::Green);
	//}
	

	FillHole(right,left,targetLeft,targetRight);
	return left;

}

void Centre::FillHole(int right, int left, vec2double rightHeight, vec2double leftHeight)
{
	vec2double leftCoord;
	vec2double rightCoord;
	vec2double holeCoord;
	vec2double feedCoord;
	double pointsYGrad;
	double pointsXGrad;
	double topYGrad;
	double topXGrad;
	double noPoints;
	double pointHeight;
	double topHeight;
	double intersectHeight;
	double scaleAdjust;
	vec2double topLength;
	vec2double bottomLength;

	int i;
	
	if(left<=right) return; //If there is no whole return.

	noPoints = left - right;

	rightCoord = fakePoints[botPoints[right]];
	leftCoord = fakePoints[botPoints[left]];

	if(rightCoord.y==leftCoord.y) pointsYGrad=0;
	else pointsYGrad = (leftCoord.y-rightCoord.y) / noPoints;

	if(rightCoord.x<=leftCoord.x) return; // there is no whole return.
	pointsXGrad = (leftCoord.x - rightCoord.x) / noPoints;


	if(leftHeight.y == rightHeight.y) topYGrad = 0;
	else topYGrad = (rightHeight.y - leftHeight.y ) / noPoints;

	if(leftHeight.x == rightHeight.x) topXGrad = 0;
	else topXGrad = (leftHeight.x - rightHeight.x) / noPoints;

	//if(hwnd!=0)
	//{
	//	DrawLine(rightCoord,leftCoord,Color::Gray);
	//	DrawLine(rightHeight,leftHeight,Color::Gray);
	//}

	topLength = (rightHeight-leftHeight);
	topLength.toVector();

	bottomLength = (leftCoord - rightCoord);
	bottomLength.toVector();
	if(bottomLength.y!=0) scaleAdjust = fabs(topLength.y/bottomLength.y);
	else scaleAdjust = 1;

	for(i = right+1; i<left;i++)
	{
		holeCoord.x = ((i-right) * pointsXGrad) + rightCoord.x;
		feedCoord.x = leftHeight.x - ((i-right) * topXGrad);//    + rightHeight.x;
		intersectHeight = GetHorizontalIntersection(feedCoord.x);

		
		
		//if(hwnd!=0) DrawCircle(vec2double(feedCoord.x,intersectHeight),4,Color::Black);

		pointHeight = (((i-right) * pointsYGrad) + rightCoord.y);

		if(intersectHeight == 0 ) topHeight = 0;
		else topHeight = (intersectHeight - leftHeight.y) - ((i-right)*topYGrad);
		holeCoord.y=pointHeight - (topHeight/1);
		fakePoints[botPoints[i]]=holeCoord;
		//if(hwnd!=0) DrawCircle(holeCoord,2,Color::Navy);

	}
}

double Centre::GetHorizontalIntersection(double x)
{
	int i;
	vec2double a,b;
	double xd,yd;
	double xi,yi;
	double grad;

	for(i=44;i<134;i++)
	{
		a = fakePoints[i];
		b = fakePoints[i+1];
		if((a!=0) && (b!=0))
		{
			if(x==a.x) return a.y;
			if(x==b.x) return b.y;
			if(a.x==b.x) continue;
			if((x>a.x) && (x<b.x))
			{
				if(a.y==b.y) return a.y;
				yd=b.y-a.y;
				xd=b.x-a.x;
				grad = yd/xd;
				xi=x-a.x;
				yi=grad*xi;
				yi+=a.y;
				return yi;
			}
		}
	}
	return 0;
}

vec2double Centre::GetProfileIntersection(vec2double point)
{
	int i;
	vec2double a,b;
	vec2double intersect;
	bool section;
	bool orig;


	for(i=44;i<134;i++)
	{
		a = fakePoints[i];
		b = fakePoints[i+1];
		::Intersection().TwoLines(point,vec2double(0,0),a,b,intersect,section,orig);
		if(orig) return intersect;
	}
	return vec2double(0,0);
}

vec2double Centre::GetProfileIntersection(vec2double pointa, vec2double pointb)
{
	int i;
	vec2double a,b;
	vec2double intersect;
	bool section;
	bool orig;


	for(i=44;i<134;i++)
	{
		a = fakePoints[i];
		b = fakePoints[i+1];
		::Intersection().TwoLines(pointa,pointb,a,b,intersect,section,orig);
		if(orig) return intersect;
	}
	return vec2double(0,0);
}


//void Centre::InitGraphics(void)
//{
//	 hdc = GetDC(hwnd);
//
//	 Gdiplus::GdiplusStartupInput gdiplusStartupInput;
//	 Gdiplus::GdiplusStartup(&m_gdiplusToken, &gdiplusStartupInput, NULL);
//
//	 graphics = new Graphics(hdc);
//}

//void Centre::DrawLine(vec2double a, vec2double b, Color colour)
//{
//	a.y*=-1;
//	b.y*=-1;
//
//	a = a / screenRatio; a = a + screenCentre;
//	b = b / screenRatio; b = b + screenCentre;
//
//	Pen linePen(colour, 1);
//	
 //
//	
//	graphics->DrawLine(&linePen, (int) a.x, (int) a.y, (int) b.x, (int) b.y);
//}

//void Centre::DrawCircle(vec2double a,double size, Color colour)
//{
//	a.y*=-1;
//	a = a / screenRatio; a = a + screenCentre;
//	size = size / screenRatio;
//	a.x -=(size/2);
//	a.y -=(size/2);
//
//	Pen cirPen(colour, 1);
//	graphics->DrawArc(&cirPen, (int) a.x, (int) a.y, (int) size, (int) size, 0,360);
//}
