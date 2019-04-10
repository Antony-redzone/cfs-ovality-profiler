#include <windows.h>
#include "..\houghlibv2.0\CBSAlgebra.h"


//#include <Gdiplus.h>
//using namespace Gdiplus; 
//#pragma comment(lib, "gdiplus.lib")

class Centre
{
public:
	Centre();
	~Centre(void);
	void CalculateCentre(void);
	bool waterLevelCentre;
	bool getEdgeCentre;
	double pvCentreX;
	double pvCentreY;
	vec2double pvData[180];
	int egnoreList[180];

private:



	int *passedEgnoreList;

	int holes[180];

	long indexProfileOne;
	vec2double fakePoints[180];

	double filteredPoints[180];
	double radiusPoints[180];
	long currentFrame;

	double shapeRadius;
	double shapeCentreX;
	double shapeCentreY;
	double shapeRotation;


	int botPoints[91];
	bool useWaterLevel;
	bool cutWaterLevel;
	bool removeWaterLevelData;
	bool reduceWaterLevelData;


	void CentreCalculationTwo(double &x, double &y);
	void CentreCalculation(double &x, double &y);
	inline double DistOfTwoPoints(vec2double pt1, vec2double pt2);
	void FindBestNeighbour(vec2double &curPoint, double size);
	double GetCentreVariance(vec2double p);
	void MarkRoughPoints(void);

	void CreateFilteredPoints(void);
	void CreateFilledHoles(void);
	void FindTopWhole(int i);
	void FillTopWhole(int left, int right);
	int  FindHole(int i);
	void FillHole(int right, int left, vec2double rightHeight, vec2double leftHeight);
	double GetHorizontalIntersection(double x);
	vec2double GetProfileIntersection(vec2double point);
	vec2double GetProfileIntersection(vec2double pointa, vec2double pointb);
	void ReSpreadProfile(void);
	vec2double GetProfileIntersection360deg(vec2double point);
	
	
};
