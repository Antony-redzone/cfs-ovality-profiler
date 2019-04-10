#include "..\houghlibv2.0\CBSAlgebra.h"

class XYDiameterMaxMin
{
public:
	XYDiameterMaxMin(float *_pvDataX, 
 		    float *_pvDataY,
			float *_pvCentreX,
			float *_pvCentreY,
			int *_egnoreList,
			double *_pvXYDiameterFullMax,
			double *_pvXYDiameterFullMin,
			short *_pvXYDiameterSegFullMax,
			short *_pvXYDiameterSegFullMin,
			int _pvDataXYMultiplier, 
		    int _fromFrame,
			int _toFrame);

	~XYDiameterMaxMin(void);
	void CalculateXYDiameterMaxMin(void);
private:
	float *pvDataX; 
	float *pvDataY;
	float *pvCentreX;
	float *pvCentreY;
	int pvDataXYMultiplier;
	double *pvXYDiameterFullMax;
	double *pvXYDiameterFullMin;
	short *pvXYDiameterSegFullMax;
	short *pvXYDiameterSegFullMin;

	
	int fromFrame;
	int toFrame;

	long currentFrame;

	double filteredPoints[180];
	vec2double radiusPoints[180];
	vec2double fakePoints[180];

	int *egnoreList;

	struct _indexedDiameter
	{
		double diameter;
		int index;
	};

	int botPoints[91];


	void CalculateFrameXYDiameterMaxMin(void);
	
	double DiameterCalculation(int segment);
	void CreateFilteredPoints(void);//long index);
	void MaxMinDiameterCalculation(int &maxSegment, int &minSegment);
	double FilterThreePoints(double left, double centre, double right);

	void CreateFakePoints(long index);
	void CreateFilledHoles(void);
	void FindTopWhole(int i);
	void FillTopWhole(int left, int right);
	int FindHole(int i);
	void FillHole(int right, int left, vec2double rightHeight, vec2double leftHeight);
	double GetHorizontalIntersection(double x);
	vec2double GetProfileIntersection(vec2double point);
	vec2double GetProfileIntersection(vec2double pointa, vec2double pointb);

};