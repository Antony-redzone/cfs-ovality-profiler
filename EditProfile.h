#include "..\houghlibv2.0\CBSAlgebra.h"

class EditProfile
{
public:
	EditProfile(float *_pvDataX, 
			   float *_pvDataY,
			   float *_pvCentreX,
			   float *_pvCentreY,
			   float *_pvCentAdjX,
			   float *_pvCentAdjY,
			   int *_egnoreList,
			   int _fromFrame,
			   int _toFrame,
			   double _diameter,
			   float *_graphData);
	~EditProfile(void);
	void FilterCrap(void);
	void CreatePerfectPipe(void);
	void FillPerfectPoints(void);
	void FillInNoData(void);
	void StampProfileData(void);
private:
	float *pvDataX; 
	float *pvDataY;
	float *pvCentreX;
	float *pvCentreY;
	float *pvCentAdjX;
	float *pvCentAdjY;
	int *egnoreList;
	double cutoffDifference;
	double cutoffHeight;
	float *graphData;
	
	int fromFrame;
	int toFrame;

	int topDirection;

	long currentFrame;
	double diameter;

	vec2double fakePoints[180];
	vec2double fakePointsOther[180];
	vec2double radialPoints[180];
	vec2double currentCentre;


	void CreateFilteredPoints(long index);
	void FrameFilterCrap(void);
	void FrameCreatePerfectPipe(void);
	void StoreProfile(void);
	void FilterProfile(void);
	void FilterPoint(int i,vec2double topPoint);
	void CreateProfile(void);
	void ReSpreadProfile(void);
	vec2double FindMedianTopPoint(int startPoint);
	vec2double GetProfileIntersection(vec2double point);
	void FrameFill(void);
	void ReOrderProfile(void);

	void FrameStamp(void);

	
};