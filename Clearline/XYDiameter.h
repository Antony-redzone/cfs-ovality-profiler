#include "..\houghlibv2.0\CBSAlgebra.h"

class XYDiameter
{
public:
	XYDiameter(float *_pvDataX, 
 		    float *_pvDataY,
			float *_pvCentreX,
			float *_pvCentreY,
			int *_egnoreList,
			double *_pvXDiameterFullData,
			double *_pvYDiameterFullData,
			int _pvDataXYMultiplier, 
		    int _fromFrame,
			int _toFrame);

	~XYDiameter(void);
	void CalculateXYDiameter(void);
private:
	float *pvDataX; 
	float *pvDataY;
	float *pvCentreX;
	float *pvCentreY;
	int pvDataXYMultiplier;
	double *pvXDiameterFullData;
	double *pvYDiameterFullData;
	
	int fromFrame;
	int toFrame;

	long currentFrame;

	double filteredPoints[180];
	vec2double radiusPoints[180];

	int *egnoreList;

	int botPoints[91];
	vec2double fakePoints[180];

	void CalculateFrameXYDiameter(void);
	void CreateRadiusPoints(void);//long index);
	double DiameterCalculation(int segment);

	void CreateFilteredPoints(long index);
	void CreateFilledHoles(void);
	void FindTopWhole(int i);
	void FillTopWhole(int left, int right);
	int FindHole(int i);
	void FillHole(int right, int left, vec2double rightHeight, vec2double leftHeight);
	double GetHorizontalIntersection(double x);
	vec2double GetProfileIntersection(vec2double point);
	vec2double GetProfileIntersection(vec2double pointa, vec2double pointb);

};