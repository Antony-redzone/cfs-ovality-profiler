#include "..\houghlibv2.0\CBSAlgebra.h"


class DeltaMaxMin
{
public:
	DeltaMaxMin(float *_pvDataX, 
 		    float *_pvDataY,
			float *_pvCentreX,
			float *_pvCenterY,
			int *_egnoreList,
			double *_pvDeltaFullMax,
			double *_pvDeltaFullMin,
			short *_pvDeltaSegFullMax,
			short *_pvDeltaSegFullMin,
			int _pvDataXYMultiplier, 
		    int _fromFrame,
			int _toFrame);

	~DeltaMaxMin(void);
	void CalculateDeltaMaxMin(void);
private:
	float *pvDataX; 
	float *pvDataY;
	float *pvCentreX;
	float *pvCentreY;
	int pvDataXYMultiplier;
	double *pvDeltaFullMax;
	double *pvDeltaFullMin;
	short *pvDeltaSegFullMax;
	short *pvDeltaSegFullMin;
	
	int fromFrame;
	int toFrame;

	long currentFrame;

	double filteredPoints[180];
	vec2double radiusPoints[180];

	struct _indexedRadius
	{
		double radius;
		int index;
	};

	int *egnoreList;

	void CalculateFrameDeltaMaxMin(void);
	void CreateFilteredPoints(long index);
	void MaxMinRadiusCalculation(int &maxSegment, int &minSegment);
	double FilterThreePoints(double left, double centre, double right);
};