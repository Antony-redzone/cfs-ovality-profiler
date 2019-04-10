#include "..\houghlibv2.0\CBSAlgebra.h"

class Capacity
{
public:
	Capacity(float *_pvDataX, 
			   float *_pvDataY,
			   float *_pvCentreX,
			   float *_pvCentreY,
			   int *_egnoreList,
			   float *_pvCapacityFullData,
			   int _pvDataXYMultiplier, 
			   double _pvExpectedDiameter,
			   int _fromFrame,
			   int _toFrame);
	~Capacity(void);
	void CalculateCapacity(void);
private:
	float *pvDataX; 
	float *pvDataY;
	float *pvCentreX;
	float *pvCentreY;
	float *pvCapacityFullData;
	int pvDataXYMultiplier;
	double pvExpectedDiameter;
	int fromFrame;
	int toFrame;
	long currentFrame;

	vec2double fakePoints[180];

	int *egnoreList;
	int holes[180];
	int botPoints[91];

	void CalculateFrameCapacity(void);
	int CreateFakePoints(long index);
	double GetRightSideHeight(long index, double leftSideHeight, long &pointLoop);
	double AreaCalculation(void);
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