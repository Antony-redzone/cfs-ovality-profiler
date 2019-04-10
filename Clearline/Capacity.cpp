#include "Capacity.h"

Capacity::Capacity(float *_pvDataX, 
				   float *_pvDataY,
				   float *_pvCentreX,
				   float *_pvCentreY,
				   int *_egnoreList,
				   float *_pvCapacityFullData,
				   int _pvDataXYMultiplier, 
				   double _pvExpectedDiameter,
				   int _fromFrame,
				   int _toFrame)
{
	pvDataX=_pvDataX;
	pvDataY=_pvDataY;
	pvCentreX=_pvCentreX;
	pvCentreY=_pvCentreY;
	pvCapacityFullData = _pvCapacityFullData;
	pvDataXYMultiplier = _pvDataXYMultiplier;
	pvExpectedDiameter = _pvExpectedDiameter;
	fromFrame = _fromFrame;
	toFrame= _toFrame;
	egnoreList = _egnoreList;
}

Capacity::~Capacity(void)
{
}

void Capacity::CalculateCapacity(void)
{

	for(currentFrame=(fromFrame-1);currentFrame<toFrame;currentFrame++) 
		CalculateFrameCapacity();

}

void Capacity::CalculateFrameCapacity(void)
{

	long indexProfileOne;
	double calculatedArea;
	double trueArea;
	double percentage;

	indexProfileOne = (currentFrame*180)+1; // Profile points are from 1 to 180 inclusive
	//if(CreateFakePoints(indexProfileOne)==-1) {pvCapacityFullData[currentFrame]=0; return;}
	CreateFilteredPoints(indexProfileOne);

	//pvCapacityFullData starts at index 1 to number of frames inclusive.
	//pvCapacityFullData[frameNumber+1]=0;
	calculatedArea=AreaCalculation();
	trueArea=((pvExpectedDiameter/2)*(pvExpectedDiameter/2))*PI;
	percentage=(calculatedArea/trueArea*100)-100;
	if(percentage==-100) percentage = -1000000000;
									  

	pvCapacityFullData[currentFrame]= (float) percentage;
}

int Capacity::CreateFakePoints(long index)
{
	long point;
	long pointIndex;
	long pointLoop;
	long firstGoodPoint=-1;

	vec2double coord;
	//Find first good point;
	for(point=0;point<180;point++)
	{
		pointIndex=point+index;
		if(egnoreList[point]==1) continue;
		if((pvDataX[pointIndex]==0) && (pvDataY[pointIndex]==0)) continue;
		firstGoodPoint=pointIndex;
		break;
	}
	if(firstGoodPoint==-1)  return -1; // No good points
	
	for(pointLoop=firstGoodPoint;pointLoop<(firstGoodPoint+181);pointLoop++)
	{
		point=pointLoop%180;
		pointIndex=point+index;

		if(((pvDataX[pointIndex]!=0) || (pvDataY[pointIndex]!=0)) && egnoreList[point]==0)
		{
			fakePoints[point]=vec2double(pvDataX[pointIndex]+pvCentreX[currentFrame],pvDataY[pointIndex]+pvCentreY[currentFrame]);
			continue;
		}
		else {
			coord = fakePoints[(point+179)%180];
			coord = coord.toVector();
			GetRightSideHeight(index,coord.y,pointLoop); //Fills the hole recursively
		}
	}
	return 0;
}

double Capacity::GetRightSideHeight(long index, double leftHandHeight, long &pointLoop)
{
	
	double rightHandHeight;
	long passedPoint;
	long pointIndex;
	long point;
	double dataX,dataY;
	vec2double coord;

	passedPoint=pointLoop%180;
	pointLoop++;
	point=pointLoop%180;
	pointIndex=index+point;

	dataX=pvDataX[pointIndex];
	dataY=pvDataY[pointIndex];
	if(egnoreList[point]==1) {dataX=0; dataY=0;}

	if((dataX!=0) || (dataY!=0))
	{
		fakePoints[point]=vec2double(dataX+pvCentreX[currentFrame], dataY+pvCentreY[currentFrame]);
		coord=fakePoints[point];
		coord=coord.toVector();
		rightHandHeight = coord.y;
	}
	else rightHandHeight = GetRightSideHeight(index, leftHandHeight, pointLoop);
	coord.y=(leftHandHeight+rightHandHeight)/2;
	coord.x=passedPoint*PI/90;
	coord=coord.toCoordinate();
	fakePoints[passedPoint]=coord;

	return rightHandHeight;
}

double Capacity::AreaCalculation(void)
{
	double total=0;
	Area trig;

	
	long index;
	for(index=0;index<179;index++)
	{
		total+=trig.Triangle(vec2double(0,0),fakePoints[index]/pvDataXYMultiplier,fakePoints[index+1]/pvDataXYMultiplier);
	}

	total+=trig.Triangle(vec2double(0,0),fakePoints[179]/pvDataXYMultiplier,fakePoints[0]/pvDataXYMultiplier);
	return  (total);
}

void Capacity::CreateFilteredPoints(long index)
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
	
		if(hwnd!=0) DrawCircle(fakePoints[point],1,Color::Black);
		if(point==89 && hwnd!=0) DrawCircle(fakePoints[point],2,Color::Black);
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



void Capacity::CreateFilledHoles(void)
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

void Capacity::FindTopWhole(int i)
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

void Capacity::FillTopWhole(int left, int right)
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

int Capacity::FindHole(int i)
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
	


	FillHole(right,left,targetLeft,targetRight);
	return left;

}

void Capacity::FillHole(int right, int left, vec2double rightHeight, vec2double leftHeight)
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

		pointHeight = (((i-right) * pointsYGrad) + rightCoord.y);

		if(intersectHeight == 0 ) topHeight = 0;
		else topHeight = (intersectHeight - leftHeight.y) - ((i-right)*topYGrad);
		holeCoord.y=pointHeight - (topHeight/1);
		fakePoints[botPoints[i]]=holeCoord;


	}
}

double Capacity::GetHorizontalIntersection(double x)
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

vec2double Capacity::GetProfileIntersection(vec2double point)
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

vec2double Capacity::GetProfileIntersection(vec2double pointa, vec2double pointb)
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


	


