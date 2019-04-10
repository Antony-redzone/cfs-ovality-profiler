#include <windows.h>
#include "..\houghlibv2.0\CBSAlgebra.h"


//#include <Gdiplus.h>
//using namespace Gdiplus; 
//#pragma comment(lib, "gdiplus.lib")

class Centre
{
public:
	Centre(float *_pvDataX, 
		   float *_pvDataY,
		   float *_pvCentreX,
		   float *_pvCentreY,
		   int *_egnoreList,
		   int _fromFrame,
		   int _toFrame,
		   int _waterLevelCentre,
		   int _edgeCentre);

	Centre(ReferenceShape_V10 *_Shape,
		   double _ShapeRadius,
		   double _ShapeCentreX,
		   double _ShapeCentreY,
		   double _ShapeRotation,
		   float *_PVDataX,
		   float *_PVDataY,
		   float *_PVCentreX,
		   float *_PVCentreY,
		   int *_EgnoreList,
		   int _fromFrame,
		   int _toFrame,
		   HWND _hwnd,
		   float _screenWidth,
		   float _screenHeight,
		   double _screenRatio);

	~Centre(void);
	void CalculateCentre(void);
	bool waterLevelCentre;
	bool getEdgeCentre;

private:
//	ULONG_PTR m_gdiplusToken;
	HWND hwnd;
	float screenWidth;
	float screenHeight;
	double screenRatio;
	HDC hdc;
	vec2double screenCentre;
//	Graphics *graphics;


	float *pvDataX;
	float *pvDataY;
	float *pvCentreX;
	float *pvCentreY;
	int *passedEgnoreList;
	int egnoreList[180];
	int holes[180];
	int fromFrame;
	int toFrame;
	long indexProfileOne;
	vec2double fakePoints[180];
	double filteredPoints[180];
	double radiusPoints[180];
	long currentFrame;

	Shapes *shape;
	double shapeRadius;
	double shapeCentreX;
	double shapeCentreY;
	double shapeRotation;


	int botPoints[91];
	bool useWaterLevel;
	bool cutWaterLevel;
	bool removeWaterLevelData;
	bool reduceWaterLevelData;

	void CalculateCentre(long frameNumber);
	void CentreCalculationTwo(double &x, double &y);
	void CentreCalculation(double &x, double &y);
	void CreateFakePoints(long frameNumber);
	inline double DistOfTwoPoints(vec2double pt1, vec2double pt2);
	void FindBestNeighbour(vec2double &curPoint, double size);
	double GetCentreVariance(vec2double p);
	void MarkRoughPoints(void);

	void CreateFilteredPoints(long index);
	void CreateFilledHoles(void);
	void FindTopWhole(int i);
	void FillTopWhole(int left, int right);
	int  FindHole(int i);
	void FillHole(int right, int left, vec2double rightHeight, vec2double leftHeight);
	double GetHorizontalIntersection(double x);
	vec2double GetProfileIntersection(vec2double point);
	vec2double GetProfileIntersection(vec2double pointa, vec2double pointb);
//	void InitGraphics(void);
//	void DrawLine(vec2double a, vec2double b, Color colour);
//	void DrawCircle(vec2double a,double size, Color colour);
	
	
};