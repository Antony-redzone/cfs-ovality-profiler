#include <windows.h>
#include "..\houghlibv2.0\CBSAlgebra.h"
//#define ULONG_PTR ULONG
//#include <Gdiplus.h>
//using namespace Gdiplus; 
//#pragma comment(lib, "gdiplus.lib")








class Ovality
{
public:
	Ovality(float *_pvDataX, 
 		    float *_pvDataY,
			float *_pvCentreX,
			float *_pvCentreY,
			int *_egnoreList,
			float *_pvOvalityFullData,
			int _pvDataXYMultiplier, 
		    int _fromFrame,
			int _toFrame,
			HWND _hwnd,
			float _screenWidth,
			float _screenHeight,
			double _screenRatio);

	~Ovality(void);
	void CalculateOvality(void);
private:
	float *pvDataX; 
	float *pvDataY;
	float *pvCentreX;
	float *pvCentreY;
	float *pvOvalityFullData;
	int pvDataXYMultiplier;
	int fromFrame;
	int toFrame;

	long currentFrame;
//	Graphics *graphics;
//	ULONG_PTR m_gdiplusToken;

	double filteredPoints[180];
	vec2double radiusPoints[180];
	vec2double fakePoints[180];

	int *egnoreList;
	int holes[180];

	int botPoints[91];

	HWND hwnd;
	float screenWidth;
	float screenHeight;
	double screenRatio;
	HDC hdc;
	vec2double screenCentre;

	void CalculateFrameOvality(void);
	void CreateFilteredPoints(long index);
	double OvalityCalculation(void);
	void CreateFilledHoles(void);
	double GetHorizontalIntersection(double x);
	int FindHole(int i);
	void FillHole(int right, int left, vec2double rightHeight, vec2double leftHeight);
	vec2double GetProfileIntersection(vec2double point);
	vec2double GetProfileIntersection(vec2double pointa, vec2double pointb);
//	void InitGraphics(void);
//	void DrawLine(vec2double a, vec2double b, Color colour);
//	void DrawCircle(vec2double a,double size, Color colour);
	void FindTopWhole(int i);
	void FillTopWhole(int left, int right);
};