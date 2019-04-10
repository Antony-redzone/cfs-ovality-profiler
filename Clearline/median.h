#include "..\houghlibv2.0\CBSAlgebra.h"

class Median
{
public:
	Median(float *_pvDataX, 
 		    float *_pvDataY,
			float *_pvCentreX,
			float *_pvCentreY,
			int *_egnoreList,
			double *_pvMedianFullData,
			int _pvDataXYMultiplier, 
		    int _fromFrame,
			int _toFrame);

	~Median(void);
	void CalculateMedian(void);
private:
	float *pvDataX; 
	float *pvDataY;
	float *pvCentreX;
	float *pvCentreY;
	int pvDataXYMultiplier;
	double *pvMedianFullData;
	
	int fromFrame;
	int toFrame;

	long currentFrame;

	int numberOfValidPairs;
	double diameterPoints[90];

	int *egnoreList;
		int holes[180];

	double filteredPoints[180];
	vec2double fakePoints[180];  //PCN4314
	vec2double radiusPoints[180];

	int botPoints[91];


	void CalculateFrameMedian(void);

	void CreateDiameterPoints(void); //Replaced long index with void PCN4314
	double MedianCalculation(void);
	int CreateFakePoints(long index); //PCN4314
	double GetRightSideHeight(long index, double leftHandHeight, long &pointLoop); //PCN4314

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