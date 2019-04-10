#include <string.h>
#include "Flat3D.h"



	


Flat3D::Flat3D(float *_pvDataX, 
		      float *_pvDataY,
			  float *_pvCentreX,
			  float *_pvCentreY,
			  int *_egnoreList,
		      int _pvDataXYMultiplier, 
		      int *_pvColourRed, 
		      int *_pvColourGreen,
		      int *_pvColourBlue,
		      double _pvExpectedDiameter,
		      double _pvRedLimit,
		      double _pvBlueLimit,
		      int _fromFrame,
			  int _toFrame,
			  float *_offsetX, 
			  float *_offsetY,
			  ReferenceShape_V10 *_Shape,
			  double _ShapeRotation,
			  double *_graphData,
			  int _shadingType) // PCN4484 _offsetX and Y were changeed from double to float * PCN4974
{
	pvDataX				= _pvDataX;
	pvDataY				= _pvDataY;
	pvCentreX			= _pvCentreX;
	pvCentreY			= _pvCentreY;
	pvDataXYMultiplier	= _pvDataXYMultiplier;
	pvColourRed			= _pvColourRed;
	pvColourGreen		= _pvColourGreen;
	pvColourBlue		= _pvColourBlue;
	pvExpRad			= _pvExpectedDiameter/2;
	pvRedLimit			= _pvRedLimit;
	pvBlueLimit			= _pvBlueLimit;
	fromFrame			= _fromFrame;
	toFrame				= _toFrame;
	offset.x			= _offsetX[0];
	offset.y			= _offsetY[0];
	egnoreList			= _egnoreList;
	offsetX = _offsetX;
	offsetY = _offsetY;
	graphData			= _graphData; //PCN4974
	shadingType			= _shadingType; //PCN4974

	int i;
	char shapeCompare[]="Circle";

	simpleCircle = true;
	for(i=0;i<6;i++) 
	{
		if(shapeCompare[i]!=_Shape->name[i]) {simpleCircle=false; break;}
		
	}
	if(pvExpRad==0) useMedianFlat = true; // PCN4974
	else useMedianFlat = false; //PCN4974


	//if(!simpleCircle) _offsetX*=-1; PCN4484 its an array now, so it need to be changed at the calculation

	shape = new Shapes(_Shape,
					   pvExpRad,
						_offsetX[0],
						_offsetY[0],
						_ShapeRotation);
	
	SetColourLimitsAndGradiants();

	
}

void Flat3D::CalculateFlat3D(long fromFrame, long toFrame)
{

	for(currentFrame=(fromFrame-1);currentFrame<toFrame;currentFrame++)
		CalculateFrameFlat3D();
}

void Flat3D::CalculateFrameFlat3D(void)
{
	vec2double profilePoint;
	double xyMultiplier = pvDataXYMultiplier;
	double normalisedPercent;
	long indexColour;
	long indexProfile;
	long profileNo;

	if(useMedianFlat) pvExpRad = graphData[currentFrame] / 2;  //PCN4974
	for(profileNo=1;profileNo<=180;profileNo++)
	{
		// Need to add +1 becuase when VB passed the array size for the colour
		// as 0 to 180 inclusive not 0 to 179 inclusive like the profilepoints
		indexColour = (profileNo)+(currentFrame*181); 
														
		indexProfile = profileNo+(currentFrame*180);
		if(pvDataX[indexProfile]==0 && pvDataY[indexProfile]==0) profilePoint=0;
		else if(egnoreList[profileNo-1]==1) profilePoint=0;
		else profilePoint=vec2double((double) (pvDataX[indexProfile]+pvCentreX[currentFrame])/ xyMultiplier,
									(double) (pvDataY[indexProfile]+pvCentreY[currentFrame])/ xyMultiplier);
		if(profilePoint==0) 
		{
			pvColourBlue[indexColour]=0;
			pvColourRed[indexColour]=0;
			pvColourGreen[indexColour]=0;
			continue;
		}
		if(simpleCircle) 
		{
			profilePoint.x -= offsetX[currentFrame];
			profilePoint.y += offsetY[currentFrame];
			normalisedPercent = (100 * (pvExpRad-profilePoint.length()) )/pvExpRad;
		}
		else
		{

			profilePoint.y*=-1;
			normalisedPercent = 100 * (-shape->ProfileRefShapeDistCalc(profilePoint))/pvExpRad;
		}

		if(shadingType==0) SetFlat3DColourGradiant(indexColour,normalisedPercent); //PCN4974, add the shadding select, shading type 1 was allways there but not used anymore
		if(shadingType==1) SetFlat3DColour(indexColour,normalisedPercent); //PCN4974, add the shadding select, shading type 1 was allways there but not used anymore
		if(pvColourRed[indexColour]<0) pvColourRed[indexColour]=0;
		if(pvColourGreen[indexColour]<0) pvColourGreen[indexColour]=0;
		if(pvColourBlue[indexColour]<0) pvColourBlue[indexColour]=0;
		if(pvColourRed[indexColour]>255) pvColourRed[indexColour]=255;
		if(pvColourGreen[indexColour]>255) pvColourGreen[indexColour]=255;
		if(pvColourBlue[indexColour]>255) pvColourBlue[indexColour]=255;

	}
}



void Flat3D::SetFlat3DColour(long index, double normalisedPercent)
{
	double blueLimit;
	double redLimit;


	blueLimit = pvBlueLimit*-1;
	redLimit = pvRedLimit;

	if(normalisedPercent<0)
	{
		normalisedPercent*=-1;
		if(normalisedPercent>redLimit)
		{	//RED
			pvColourRed[index]=255;
			pvColourGreen[index]=0;
			pvColourBlue[index]=0;
		}
		else if(normalisedPercent>(2*redLimit/3))
		{	//ORANGE
			pvColourRed[index]=255;
			pvColourGreen[index]=150;
			pvColourBlue[index]=0;
		}
		else if(normalisedPercent>(redLimit/3))
		{	//YELLOW
			pvColourRed[index]=255;
			pvColourGreen[index]=255;
			pvColourBlue[index]=20;
		}
		else
		{	//WHITE
			pvColourRed[index]=255;
			pvColourGreen[index]=255;
			pvColourBlue[index]=255;
		}
	}


	else if(normalisedPercent>0)
	{
		
		if(normalisedPercent>blueLimit)
		{	//BLUE
			pvColourRed[index]=40;
			pvColourGreen[index]=73;
			pvColourBlue[index]=111;
		}
		else if(normalisedPercent>(2*blueLimit/3))
		{	//AQUA
			pvColourRed[index]=90;
			pvColourGreen[index]=155;
			pvColourBlue[index]=204;
		}
		else if(normalisedPercent>(blueLimit/3))
		{	//GREEN
			pvColourRed[index]=181;
			pvColourGreen[index]=224;
			pvColourBlue[index]=238;
		}
		else
		{	//White
			pvColourRed[index]=255;
			pvColourGreen[index]=255;
			pvColourBlue[index]=255;
		}
	}
	else
	{
		pvColourBlue[index]=255;
		pvColourRed[index]=255;
		pvColourGreen[index]=255;
	}
}

void Flat3D::SetFlat3DColourGradiant(long index, double normalisedPercent)
{
	double blueLimit;
	double redLimit;
	double ratio;
	double red;
	double blue;
	double green;

	blueLimit = pvBlueLimit*-1;
	redLimit = pvRedLimit;
	ratio = redLimit/4;

	if(normalisedPercent<0)
	{
	
		normalisedPercent*=-1;
		if(normalisedPercent>redLimit)
		{	//RED
			pvColourRed[index]=255;
			pvColourGreen[index]=0;
			pvColourBlue[index]=0;
		}
		else if(normalisedPercent>(3*redLimit/4))
		{	//ORANGE

			green = normalisedPercent - (3*redLimit/4);
			green = (green/ratio)*150;

			pvColourRed[index]=255;
			pvColourGreen[index]=150 - (int) green;
			pvColourBlue[index]=0;
		}
		else if(normalisedPercent>(2*redLimit/4))
		{	//YELLOW

			green = normalisedPercent - (2 * redLimit/4);
			green = (green/ratio) * 105;

			blue = normalisedPercent - (2 * redLimit / 4);
			blue = (blue/ratio) * 20;


			pvColourRed[index]=255;
			pvColourGreen[index]=255- (int) green;
			pvColourBlue[index]=20- (int) blue;
		}
		else if(normalisedPercent>(1*redLimit/4))
		{	//YELLOW
			blue = normalisedPercent - (1*redLimit/4);
			blue = (blue/ratio) * 235;
			

			pvColourRed[index]=255;
			pvColourGreen[index]=255;
			pvColourBlue[index]=255- (int) blue;
		}
		else
		{	//WHITE
			pvColourRed[index]=255;
			pvColourGreen[index]=255;
			pvColourBlue[index]=255;
		}
	}


	else if(normalisedPercent>0)
	{
		
		if(normalisedPercent>blueLimit)
		{	//BLUE
			


			pvColourRed[index]=40;
			pvColourGreen[index]=73;
			pvColourBlue[index]=111;
		}
		else if(normalisedPercent>(3*blueLimit/4))
		{	//AQUA
			red = normalisedPercent - (3*blueLimit/4);
			red = (red/ratio) * 50;

			green = normalisedPercent - (3*blueLimit/4);
			green = (green/ratio) * 82;

			blue = normalisedPercent - (3*blueLimit/4);
			blue = (blue/ratio) * 93;


			pvColourRed[index]=90 - (int) red;
			pvColourGreen[index]=155 - (int) green;
			pvColourBlue[index]=204- (int) blue;
		}
		else if(normalisedPercent>(2*blueLimit/4))
		{	//GREEN
			red = normalisedPercent - (2*blueLimit/4);
			red = (red/ratio) * 91;

			green = normalisedPercent - (2*blueLimit/4);
			green = (green/ratio) * 69;

			blue = normalisedPercent - (2*blueLimit/4);
			blue = (blue/ratio) * 34;

			pvColourRed[index]=181 - (int) red;
			pvColourGreen[index]=224 - (int) green;
			pvColourBlue[index]=238- (int) blue;
		}
		else if(normalisedPercent>(1*blueLimit/4))
		{	//GREEN
			red = normalisedPercent-(1*blueLimit/4);
			red = (red/ratio) * 74;

			green = normalisedPercent-(1*blueLimit/4);
			green = (green/ratio) * 31;

			blue = normalisedPercent-(1*blueLimit/4);
			blue = (blue/ratio) * 17;

			pvColourRed[index]=255- (int) red;
			pvColourGreen[index]=255- (int) green;
			pvColourBlue[index]=255- (int) blue;
		}

		else
		{	//White
			pvColourRed[index]=255;
			pvColourGreen[index]=255;
			pvColourBlue[index]=255;
		}
	}
	else
	{
		pvColourBlue[index]=255;
		pvColourRed[index]=255;
		pvColourGreen[index]=255;
	}
}

void Flat3D::TestPattern(void)
{
	long index;
	long i,j;
	int red,green,blue;
	int modulate=0;
	green=0;
	blue=0;
	for(j=0;j<toFrame;j++)
	{
		if((j%10)==0) modulate+=10;
		if(modulate>255) modulate=0;
		green=modulate;
		for(i=1;i<181;i++)
		{
			index = i+(j*181);
			if(i<45) red=63;
			if(i>45 && i<90) red = 125;
			if(i>90 && i < 135) red = 188;
			if(i>135) red=255;
			pvColourRed[index]=red;
			pvColourGreen[index]=green;
			pvColourBlue[index]=blue;
		}
	}
}

void Flat3D::SetColourLimitsAndGradiants(void)
{
	redRamp.RedGradiant = 1;
	
}
