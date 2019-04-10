#ifndef _Flat3D_
#define _Flat3D_

#include "Shapes.h"
#include "..\houghlibv2.0\CBSAlgebra.h"


class Flat3D 
{
public:
	Flat3D(float *_pvDataX, 
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
		   int _fromeFrame,
		   int _toFrame,
		   float *_offsetX,
		   float *_offsetY,
		   ReferenceShape_V10 *_Shape,
		   double _ShapeRotation,
		   double* _graphData,
		   int _shadingType); //PCN4484 _offsetX and Y changed from double to float *  //PCN4974

	void CalculateFlat3D(void);
	void CalculateFlat3D(long fromFrame, long toFrame);
	void TestPattern(void);

private:
	struct _ColourRamp
	{
		double RedUpperClip;
		double RedLowerClip;
		double RedGradiant;

		double GreenUpperClip;
		double GreenLowerClip;
		double GreenGradiant;
		
		
		double BlueUpperClip;
		double BlueLowerClip;
		double BlueGradiant;
	};

	_ColourRamp redRamp;
	_ColourRamp blueRamp;

	float *pvDataX;
	float *pvDataY;
	float *pvCentreX;
	float *pvCentreY;
	double pvDataXYMultiplier;
	int *pvColourRed;
	int *pvColourGreen;
	int *pvColourBlue;
	double pvExpRad;
	double pvRedLimit;
	double pvBlueLimit;
	long fromFrame;
	long toFrame;
	vec2double offset; //PCN3567
	float *offsetX;
	float *offsetY;
	long currentFrame;
	bool simpleCircle;
	bool useMedianFlat; //PCN4974
	int shadingType; //PCN4974

	Shapes *shape;

	int *egnoreList;
	double *graphData; //PCN4974
	
	void CalculateFrameFlat3D(void);
	
	
	void SetFlat3DColour(long index, double normalisedPercent);
	void SetFlat3DColourGradiant(long index, double normalisedPercent);
	void SetColourLimitsAndGradiants(void);
	
};

#endif