#include "Percentile.h"
#include <math.h>

Percentile::Percentile(float *_pvDataX, 
				   float *_pvDataY,
				   float *_pvCentreX,
				   float *_pvCentreY,
				   int *_egnoreList,
				   float *_pvPercentileFullData,
				   int _pvDataXYMultiplier, 
				   int _fromFrame,
				   int _toFrame)
{
	pvDataX=_pvDataX;
	pvDataY=_pvDataY;
	pvCentreX=_pvCentreX;
	pvCentreY=_pvCentreY;
	pvPercentileFullData = _pvPercentileFullData;
	pvDataXYMultiplier = _pvDataXYMultiplier;
	fromFrame = _fromFrame;
	toFrame= _toFrame;
	egnoreList = _egnoreList;

}

Percentile::~Percentile(void)
{
}

void Percentile::CalculatePercentile(void)
{

	for(currentFrame=(fromFrame-1);currentFrame<toFrame;currentFrame++) 
		CalculateFramePercentile();

}

void Percentile::CalculateFramePercentile(void)
{
	long indexProfileOne;
	float calculatedPercentile;
	
	indexProfileOne = (currentFrame*180)+1; // Profile points are from 1 to 180 inclusive
	CreateDiameterPoints(indexProfileOne);	//Create the filtered point to calculate the Percentile
	calculatedPercentile = PercentileCalculation();

	if(calculatedPercentile==0) calculatedPercentile=-1000000000;
	
	
	pvPercentileFullData[currentFrame] = calculatedPercentile;
}

void Percentile::CreateDiameterPoints(long index)
{
	long point;
	long pointIndex;
	vec2double coordA,coordB;

	numberOfValidPairs=0;
	for(point=0;point<90;point++)
	{
		pointIndex = point+index;
		if(pvDataX[pointIndex]==0 && pvDataY[pointIndex]==0) coordA=0;
		else if(egnoreList[point]==1) coordA=0;
		else coordA = vec2double(pvDataX[pointIndex]+pvCentreX[currentFrame],pvDataY[pointIndex]+pvCentreY[currentFrame]).toVector();
		
		if(pvDataX[pointIndex+90]==0 && pvDataY[pointIndex+90]==0) coordB=0;
		else if(egnoreList[point+90]==1) coordB=0;
		else coordB = vec2double(pvDataX[pointIndex+90]+pvCentreX[currentFrame],pvDataY[pointIndex+90]+pvCentreY[currentFrame]).toVector();
		
		if((coordA == 0 ) || (coordB == 0)) continue;
		diameterPoints[numberOfValidPairs] = (float) (coordA.y + coordB.y) / pvDataXYMultiplier;
		numberOfValidPairs++;
	}
}


float Percentile::PercentileCalculation()
{
	if(numberOfValidPairs==0) return 0;
	if(numberOfValidPairs<3) return diameterPoints[0];

	int i;
	float temp;
	//Sort diameters

	bool swaped=true;
	while(swaped)
	{
		swaped=false;
		for(i=0;i<numberOfValidPairs-1;i++)
		{
			if(diameterPoints[i]>diameterPoints[i+1])
			{
				swaped=true;
				temp = diameterPoints[i];
				diameterPoints[i]=diameterPoints[i+1];
				diameterPoints[i+1]=temp;
			}
		}
	}
	return diameterPoints[(int) ((float) numberOfValidPairs/0.95)];
}