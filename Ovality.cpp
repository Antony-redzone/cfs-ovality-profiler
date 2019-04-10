#include "Ovality.h"
#include <math.h>
#include "Common.h"

Ovality::Ovality(float *_pvDataX, 
				   float *_pvDataY,
				   float *_pvCentreX,
				   float *_pvCentreY,
				   int *_egnoreList,
				   float *_pvOvalityFullData,
				   int _pvDataXYMultiplier, 
				   int _fromFrame,
				   int _toFrame,
 			   	   HWND _hwnd,
				   float _screenWidth,
				   float _screenHeight,
				   double _screenRatio)
{
	pvDataX=_pvDataX;
	pvDataY=_pvDataY;
	pvCentreX=_pvCentreX;
	pvCentreY=_pvCentreY;
	pvOvalityFullData = _pvOvalityFullData;
	pvDataXYMultiplier = _pvDataXYMultiplier;
	fromFrame = _fromFrame;
	toFrame= _toFrame;
	egnoreList = _egnoreList;
	hwnd = _hwnd;
	screenWidth = _screenWidth;
	screenHeight = _screenHeight;
	screenRatio = _screenRatio;
	screenCentre.x = _screenWidth / 2;
	screenCentre.y = _screenHeight / 2;
}


Ovality::~Ovality(void)
{
}

void Ovality::CalculateOvality(void)
{
	//if(hwnd != 0) InitGraphics();
	for(currentFrame=(fromFrame-1);currentFrame<toFrame;currentFrame++) 
	{
		if(currentFrame==1148)
		{
			__asm nop;
		}
		CalculateFrameOvality();
	}

}

void Ovality::CalculateFrameOvality(void)
{

	long indexProfileOne;
	double calculatedOvality;


	indexProfileOne = (currentFrame*180)+1; // Profile points are from 1 to 180 inclusive
	CreateFilteredPoints(indexProfileOne);	//Create the filtered point to calculate the Ovality
	
	calculatedOvality = OvalityCalculation();
	pvOvalityFullData[currentFrame]=(float) calculatedOvality;

}

void Ovality::CreateFilteredPoints(long index)
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
		else if(egnoreList[point]==1) holes[point]=1;
		else holes[point]=0;


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


	//Find first good point;
	for(point=0;point<180;point++)
	{
		radiusPoints[point] = fakePoints[point].toVector();
	}

	for(point=1;point<179;point++)
	{
		
		filteredPoints[point] = FilterThreePoints(radiusPoints[point-1].y,
												 radiusPoints[point].y,
												 radiusPoints[point+1].y);
	}

	filteredPoints[0] = FilterThreePoints(radiusPoints[179].y,
										 radiusPoints[0].y,
										 radiusPoints[1].y);

	filteredPoints[179] = FilterThreePoints(radiusPoints[178].y,
										 radiusPoints[179].y,
										 radiusPoints[0].y);

}


double Ovality::OvalityCalculation(void)
{
	const int NumGoodForTrue = 45; //Number of true reading needed to give a non estimated ovality
	const int NumInvaForBad = 45;  //Number of invalid readings needed to give a invalid ovality

	int i;
	double segRad;
	double oppSegRad;
	double dia; // Diameter of the two points, either true or calculated
	double maxDiaTrue = 0; // Maximum Diameter with no holes in profile
	double minDiaTrue = 100000000; // Minimum Diameter with no holes in profile
	double maxDiaEsti = 0; // Maximum Diameter with esitmated diameter values;
	double minDiaEsti = 100000000; // Minimum Diameter with estimated diamerer values;
	double diaTolTrue = 0; // Total of true diameter distances
	double diaTolEsti = 0; // Total of estimated diameter distances
	double meanDiameter;
	double ovality;

	int noOfTrue = 0; // Number of good diameters
	int noOfEsti = 0; // Number of estimated diameteres
	int noOfInva = 0; // Number of invalid points, no dai.

	for(i=0;i<90;i++)
	{
		segRad = filteredPoints[i];			// Get radius distance
		oppSegRad = filteredPoints[i+90];	// Get the opposit radius distance
		if((segRad>0) && (oppSegRad>0))	    // If both valid data, as not zero total true values
		{
			dia = segRad+oppSegRad;				 // Diameter is both opposite radius added together
			noOfTrue++;							 // Incriment no the true daimeter readings
			diaTolTrue+=dia;					 // Add that diameter to the running true total
			if(dia>maxDiaTrue) maxDiaTrue = dia; // If its the largest diameter so far then remember that
			if(dia<minDiaTrue) minDiaTrue = dia; // If its smallest diameter so far then remember that
		}
		else if ((segRad==0) && (oppSegRad==0)) noOfInva++;  // if both sides are zero then cant calculate and add
															 // a inciment toal invalid readings else
		else	// find out which one is valid and use the twice its radius as diameter
		{
			if(segRad==0) dia = oppSegRad * 2; 
			else          dia = segRad    * 2;
			noOfEsti++;							 // Incrimint no of estimated diameter readingns
			diaTolEsti+=dia;					 // Add that estimated reading tot he running estimate total
			if(dia>maxDiaEsti) maxDiaEsti = dia; // Rember if its the largest estimated diameter
			if(dia<minDiaEsti) minDiaEsti = dia; // Remember if its smallest estimated diamter
		}
	}		

	int badPoints;

	badPoints = 0;
	//count no of bad
	for(i=0;i<90;i++)
	{
		if((egnoreList[i]==1) || holes[i]==1 || egnoreList[i+90]==1 || holes[i+90]) badPoints++;
	}
	
	if(noOfTrue>NumGoodForTrue) // If we have 45 good reading then true readnig will be used for Ovality
	{
		meanDiameter =  diaTolTrue / noOfTrue;						// Calculation of the mean diamter

		
		if(fabs(meanDiameter - maxDiaTrue) > fabs(meanDiameter - minDiaTrue))	ovality = 100 * (maxDiaTrue - meanDiameter) / meanDiameter; // Ovality Calculation
		else ovality = 100 * (meanDiameter - minDiaTrue) / meanDiameter; // Ovality Calculation
		
		
		if(badPoints>NumInvaForBad) ovality*=-1;		
		return ovality; 
	}
	if(noOfInva>NumInvaForBad) return -100000;

	meanDiameter = (diaTolTrue+diaTolEsti) / (noOfTrue+noOfEsti); // Calculation of the mean
	
	
	
	if(maxDiaTrue > 0)  
	{	
		if(fabs(meanDiameter - maxDiaTrue) > fabs(meanDiameter - minDiaTrue))	ovality = -100 * (maxDiaTrue - meanDiameter) / meanDiameter; // Ovality Calculation with true max dia
		else ovality = -100 * (meanDiameter - minDiaTrue) / meanDiameter; // Ovality Calculation with true max dia
	}
	else				
	{
		if(fabs(meanDiameter - maxDiaEsti) > fabs(meanDiameter - minDiaEsti)) ovality = -100 * (maxDiaEsti - meanDiameter) / meanDiameter; // Ovality Calculation with estimated max dia
		else ovality = -100 * (meanDiameter - minDiaEsti) / meanDiameter; // Ovality Calculation with true max dia
	}

	return ovality;
}

void Ovality::CreateFilledHoles(void)
{
	int i,j;
	double averageShift=0;
	double numberOfPairs=0;
	double shift=0;
	vec2double a,b;

	for(i=136,j=0;i<223;i++,j++) botPoints[j] = i%180;

	//Fill it any missing top points
	//for(i=0;i<180;i++) if(egnoreList[i]==1) fakePoints[i]=0;

	for(i=0;i<180;i++)
	{
		b = fakePoints[i]; if(b==0) FindTopWhole(i);
	}


	//for(i=0;i<180;i++) if(egnoreList[i]==1) fakePoints[i]=0;
	
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

}

void Ovality::FindTopWhole(int i)
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

void Ovality::FillTopWhole(int left, int right)
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
	if(numberHoles >175 || numberHoles == 0) 
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

int Ovality::FindHole(int i)
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

void Ovality::FillHole(int right, int left, vec2double rightHeight, vec2double leftHeight)
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
	//  DrawLine(rightCoord,leftCoord,Color::Gray);
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

double Ovality::GetHorizontalIntersection(double x)
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

vec2double Ovality::GetProfileIntersection(vec2double point)
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

vec2double Ovality::GetProfileIntersection(vec2double pointa, vec2double pointb)
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


//void Ovality::InitGraphics(void)
//{
//	 hdc = GetDC(hwnd);
//
//	 Gdiplus::GdiplusStartupInput gdiplusStartupInput;
//	 Gdiplus::GdiplusStartup(&m_gdiplusToken, &gdiplusStartupInput, NULL);
//
//	 graphics = new Graphics(hdc);
//}

//void Ovality::DrawLine(vec2double a, vec2double b, Color colour)
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

//void Ovality::DrawCircle(vec2double a,double size, Color colour)
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