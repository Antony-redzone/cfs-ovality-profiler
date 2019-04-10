#include "..\houghlibv2.0\CBSAlgebra.h"

class Percentile
{
public:
	Percentile(float *_pvDataX, 
 		    float *_pvDataY,
			float *_pvCentreX,
			float *_pvCentreY,
			int *_egnoreList,
			float *_pvPercentileFullData,
			int _pvDataXYMultiplier, 
		    int _fromFrame,
			int _toFrame);

	~Percentile(void);
	void CalculatePercentile(void);
private:
	float *pvDataX; 
	float *pvDataY;
	float *pvCentreX;
	float *pvCentreY;
	int pvDataXYMultiplier;
	float *pvPercentileFullData;
	
	int fromFrame;
	int toFrame;

	long currentFrame;

	int numberOfValidPairs;
	float diameterPoints[90];

	int *egnoreList;

	void CalculateFramePercentile(void);
	void CreateDiameterPoints(long index);
	float PercentileCalculation(void);
};