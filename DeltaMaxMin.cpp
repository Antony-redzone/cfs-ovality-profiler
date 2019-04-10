#include "DeltaMaxMin.h"
#include <math.h>


DeltaMaxMin::DeltaMaxMin(float *_pvDataX, 
				   float *_pvDataY,
				   float *_pvCentreX,
				   float *_pvCentreY,
				   int *_egnoreList,
				   double *_pvDeltaFullMax,
				   double *_pvDeltaFullMin,
				   short *_pvDeltaSegFullMax,
				   short *_pvDeltaSegFullMin,
				   int _pvDataXYMultiplier, 
				   int _fromFrame,
				   int _toFrame)
{
	pvDataX=_pvDataX;
	pvDataY=_pvDataY;
	pvCentreX=_pvCentreX;
	pvCentreY=_pvCentreY;
	pvDeltaFullMax = _pvDeltaFullMax;
	pvDeltaFullMin = _pvDeltaFullMin;
	pvDeltaSegFullMax = _pvDeltaSegFullMax;
	pvDeltaSegFullMin = _pvDeltaSegFullMin;
	pvDataXYMultiplier = _pvDataXYMultiplier;
	fromFrame = _fromFrame;
	toFrame= _toFrame;
	egnoreList = _egnoreList;

}

DeltaMaxMin::~DeltaMaxMin(void)
{
}

void DeltaMaxMin::CalculateDeltaMaxMin(void)
{
	for(currentFrame=(fromFrame-1);currentFrame<toFrame;currentFrame++) 
		CalculateFrameDeltaMaxMin();

}

void DeltaMaxMin::CalculateFrameDeltaMaxMin(void)
{
	long indexProfileOne;
	double calculatedMaxDelta;
	double calculatedMinDelta;
	int maxSegment;
	int minSegment;

	if(currentFrame==182)
	{
		__asm nop;
	}

	indexProfileOne = (currentFrame*180)+1; // Profile points are from 1 to 180 inclusive
	CreateFilteredPoints(indexProfileOne);	//Create the filtered point to calculate the DeltaMaxMin
	MaxMinRadiusCalculation(maxSegment, minSegment);
	

	if(maxSegment==-1) { calculatedMaxDelta=-1000000000; maxSegment = -1;}
	else {
		calculatedMaxDelta = radiusPoints[maxSegment].y / pvDataXYMultiplier;;
		maxSegment = maxSegment + 1;
	}

	if(minSegment==-1) { calculatedMinDelta=-1000000000; minSegment = -1;}
	else {
		calculatedMinDelta = radiusPoints[minSegment].y / pvDataXYMultiplier;
		minSegment = minSegment + 1;
		}
	
	pvDeltaFullMax[currentFrame] = calculatedMaxDelta;
	pvDeltaFullMin[currentFrame] = calculatedMinDelta;
	pvDeltaSegFullMax[currentFrame]=(short) maxSegment;
	pvDeltaSegFullMin[currentFrame]=(short) minSegment;
}



void DeltaMaxMin::CreateFilteredPoints(long index)
{
	long point;
	long pointIndex;
	vec2double coord;

	//Find first good point;
	for(point=0;point<180;point++)
	{
		pointIndex = point+index;
		if(pvDataX[pointIndex]==0 && pvDataY[pointIndex]==0) radiusPoints[point]=0;
		else if(egnoreList[point]==1) radiusPoints[point]=0;
		else radiusPoints[point] = vec2double(pvDataX[pointIndex]+pvCentreX[currentFrame],pvDataY[pointIndex]+pvCentreY[currentFrame]).toVector();
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

void DeltaMaxMin::MaxMinRadiusCalculation(int &maxSegment, int &minSegment)
{
	int i;

	_indexedRadius max;
	_indexedRadius min;
	double currentRadius;

	min.index=-1;
	
	for(i=0;i<180;i++)
	{
		if(filteredPoints[i]==0) continue;
		min.radius = max.radius = filteredPoints[i];
		min.index = max.index = i;
		break;
	}

	for(;i<180;i++)
	{
		currentRadius = filteredPoints[i];
		if((currentRadius < min.radius) && (currentRadius!=0)) {min.radius = currentRadius; min.index = i;}
		if(currentRadius > max.radius) {max.radius = currentRadius; max.index = i;}
	}

	if(min.index==-1) 
	{
		maxSegment=-1;
		minSegment=-1;
	}
	else
	{
		maxSegment = max.index;
		minSegment = min.index;
	}
}

double DeltaMaxMin::FilterThreePoints(double left, double centre, double right)
{
	double leftThreshold;
	double rightThreshold;

	if(centre==0) return 0;

	leftThreshold = fabs((centre - left) / centre);
	rightThreshold = fabs((centre - right) / centre);
	if((leftThreshold < 0.05) && rightThreshold < 0.05) return (left + centre + right)/3;
	return 0;
}
