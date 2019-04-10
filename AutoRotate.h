#include "Shapes.h"
#include "..\houghlibv2.0\CBSAlgebra.h"
#include <windows.h>
//#define ULONG_PTR ULONG
//#include <Gdiplus.h>
//using namespace Gdiplus; 
//#pragma comment(lib, "gdiplus.lib")

class AutoRotate
{
public:
	AutoRotate(ReferenceShape_V10 *_Shape,
				double _ShapeRadius,
				double _ShapeCentreX,
				double _ShapeCentreY,
				double _ShapeRotation,
				float *_PVDataX,
				float *_PVDataY,
				float *_PVCentreX,
				float *_PVCentreY,
				int *_EgnoreList,
				HWND _hwnd,
				float _screenWidth,
				float _screenHeight,
				double _screenRatio);
	~AutoRotate(void);
	void CalculateRotation(int _FromFrame, int _ToFrame);
private:
//	ULONG_PTR m_gdiplusToken;
	HWND hwnd;
	float screenWidth;
	float screenHeight;
	double screenRatio;
	HDC hdc;
	vec2double screenCentre;
//	Graphics *graphics;

	double *rotationWeight;
	vec2double fakePoints[180];
	int currentFrame;
	Shapes *shape;
	double shapeRadius;
	double shapeCentreX;
	double shapeCentreY;
	double shapeRotation;
	float *pvDataX;
	float *pvDataY;
	float *pvCentreX;
	float *pvCentreY;
	int *egnoreList;

	void AutoRotateFrame(void);
	void CreateFakePoints(void);
	double GetWeight(double rot);
	void RotateProfile(double rot);
	void CopyFakePointsBack(void);
	void FindBestRotation(double &curPoint, double step);

//	void InitGraphics(void);
//	void DrawLine(vec2double a, vec2double b, Color colour);
//	void DrawCircle(vec2double a,double size, Color colour);
	int GetMostVertical(void);
	void AdjustTopIndex(void);

};