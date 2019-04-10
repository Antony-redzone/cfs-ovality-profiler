#include <windows.h>
#include <stdio.h>

#include "Centre.h"

Centre::Centre()
{

}



Centre::~Centre(void)
{

}


void Centre::CalculateCentre(void)
{
	double centreX;
	double centreY;
	int i;

	CreateFilteredPoints();
	MarkRoughPoints();
	
	CentreCalculation(centreX,centreY);
	
	pvCentreX = centreX;
	pvCentreY = centreY;

	for (int i=0;i<180;i++)
	{
	   fakePoints[i].x -= centreX;
	   fakePoints[i].y -= centreY;
	}

	ReSpreadProfile();

	
	CentreCalculation(centreX,centreY);

	pvCentreX = (float) (centreX)+pvCentreX;
	pvCentreY = (float) (centreY)+pvCentreY;

	for (i=0;i<180;i++)
	{
	   pvData[i].x = fakePoints[i].x + pvCentreX;
	   pvData[i].y = fakePoints[i].y + pvCentreY;
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
		}
}

void Centre::CreateFilteredPoints(void)
{
	long point;
	vec2double coord;

	for(point=0;point<180;point++)
	{

		if(pvData[point].x==0 && pvData[point].y==0) 
		{
			holes[point]=1;
		}
		else 
		{
			holes[point]=0;
		}

		//egnoreList[point] = passedEgnoreList[point];

		if(pvData[point].x==0 && pvData[point].y==0) 
		{
			fakePoints[point]=0;
		}
		else if(egnoreList[point]==1) 
		{
			fakePoints[point]=0;

		}
		else 
		{
			fakePoints[point] = vec2double(pvData[point].x,pvData[point].y);
		}
	}
	CreateFilledHoles(); // Where there is a whole fill in as best as we can to avoid false reading
}

void Centre::CreateFilledHoles(void)
{
	int i,j;
	double averageShift=0;
	double numberOfPairs=0;
	double shift=0;
	vec2double a,b;

	for(i=136,j=0;i<223;i++,j++) botPoints[j] = i%180;

	//Fill it any missing top points
	for(i=0;i<180;i++) if(egnoreList[i]==1) fakePoints[i]=0;

	for(i=0;i<180;i++)
	{
		b = fakePoints[i]; if(b==0) FindTopWhole(i);
	}
	//if(cutWaterLevel) return; //PCN4538 actually just filling in the data as old, but not mirror the top

	//if(reduceWaterLevelData) for(i=30;i<150;i++) egnoreList[i]=0;

	for(i=0;i<180;i++) if(egnoreList[i]==1) fakePoints[i]=0;

	//if(removeWaterLevelData) return; //PCN4538	


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

void Centre::ReSpreadProfile(void)
{

	vec2double TempPoints[180];
	int i;
	double Gradiant;
	vec2double point;

	Gradiant = (PI*2)/180;
	
	
	for(i=0;i<180;i++)
	{
		point = point.toVector();
		point.x = (Gradiant * i) + PI;
		if((point.x)> (2*PI)) point.x -=(2*PI);
		point.y = 2000;
		point = point.toCoordinate();

		TempPoints[i]=GetProfileIntersection360deg(point);
	}

	for(i=0;i<180;i++)
	{
		fakePoints[i]=TempPoints[i];
	}

}

vec2double Centre::GetProfileIntersection360deg(vec2double point)
{
	int i;
	vec2double a,b;
	vec2double intersect;
	bool section;
	bool orig;


	for(i=0;i<179;i++)
	{
		a = fakePoints[i];
		b = fakePoints[i+1];
		::Intersection().TwoLines(point,vec2double(0,0),a,b,intersect,section,orig);
		if(section && orig) return intersect;
	}
	a = fakePoints[179];
	b = fakePoints[0];
	::Intersection().TwoLines(point,vec2double(0,0),a,b,intersect,section,orig);
	if(section && orig) return intersect;

	return vec2double(0,0);
}


