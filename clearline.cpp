#include <atlbase.h>
//#include <windows.h>


#include <stdio.h>
#include <io.h>
#include <fcntl.h>
#include <sys/types.h>
#include <sys/stat.h>


/*
#define ULONG_PTR ULONG
#include <gdiplus.h>
using namespace Gdiplus; 
#pragma comment(lib, "gdiplus.lib")
*/


#include "Flat3d.h"
#include "Capacity.h"
#include "Ovality.h"
#include "XYDiameter.h"
#include "XYDiameterMaxMin.h"
#include "DeltaMaxMin.h"
#include "Median.h"
#include "LoadPVD.h"
#include "CentreCalculations.h"
#include "Shapes.h"
#include "FilterGraph.h"
#include "AutoRotate.h"
#include "Fractile.h"
#include "Percentile.h"

#include "EmbeddedFile.h"
#include "ZipProfile.h"
#include "EditProfile.h"


//ULONG_PTR m_gdiplusToken;


void Msg(TCHAR *szFormat, ...)
	{
    TCHAR szBuffer[512];

    va_list pArgs;
    va_start(pArgs, szFormat);
    _vstprintf(szBuffer, szFormat, pArgs);
    va_end(pArgs);

    MessageBox(NULL, szBuffer, TEXT("LaserProfiler Message"), MB_OK | MB_ICONERROR);
	};


////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

double VERSION =		4.6; //inclination lay of pipe calculation v Richard
						//4.5; //inclination lay of pipe calculation
						//4.4; //inclination graph calculation
						//4.3; //inverted the delfection graph, so greater than median is possible %, lessthan is negative %
							// also out by one on the delfection graph, shoud have been <toFrames, not <= to frames
						//4.2; //in max min calc, max and min were not initialised
						//4.1; Deflection graphs
						//4.0; //PCN4974 added median flat
						//3.9; // water level bug in true diameter. Now true dimater, ovality, and capacity, there water level is treated like a hole.
							 // abs - was used for smooth, but capacity goes -ve, not marker for capaicyt to not use abs
						// 3.8 Release
						// 3.7 Medain / True egnoreing water level
						// 3.6 Ovality change, if ovality min diamter, is greater deflection than max diamter then use min diamter to calculate ovality
						// 3.5 Fixed medain diameter to egnore water level when holes filled for fill whole algorithim
						// 3.4 Tweeked the fill whole from median diameter
						// 3.3 Bug in smoothing, accumilated error, total on avereagre was not initialise to zero
						// 3.2 Just added if the data is not valid, egnore
						// Smoothing
						// 3.0 Added fill in for max min calcs
						// 2.9; ????	
						// Ovality reflecs fill even if there is no water, same for centre calc
						// 2.8 Changing the fill.
						// Added a reference centre shift. And put a whole fill algorithms to True (median) Diameter
						// 2.7 Add whole fill in for the X/Y graph  
						// Note for below: didn't make any difference at the end, windows 2000 drew incorectly like it was in 16bit colour
						// 2.5 Removed GDI interfaces, this causes bad error's with windows 2000 
						// 2.4 Different centre type selectable. (Edge, water or normal)
						// 2.3 Filtered graphs added
						// 2.2 Capacity now has a better whole fill in, just like the rest of the graph calulations	
						// 2.1 replaced quicksort, which caused overflow, with a simple sort. Reinstated median calulation
						// 2.0 Overflow index in Quicksort, Quicksort was max index not max number of array
						// 1.9 Whole fill in for median diameter.
						// 1.8 Fixed shap flat, when moved left, it moved right
						// 1.7 anyith percentile ( actually its fratile when you choose any number) but its for 90% percentile
						// 1.6 The move and insert file was appended using "a+" should have used "r+" to write to middle of file
						// 1.5; Auto rotate more complete
						// 1.4;
						// 1.3;	// When loading a file now loads into a float before transfering data to array. Didn't fix the 
						// overflow eror that VB get on ExpectedDiameter when bad data is passed thru
						//1.2; //divide by zero error that was not picked up, it was in ovality and centre calcs
						//1.1; //Water level was not passed to centre calculation
						//1.0; //The start of version tracking for Clearline.dll


# include <algorithm>

//template <class itemType, class indexType=int>
//template <class itemType, class indexType>

void QuickSort(float a[], int l, int r)
{

  static float m;
  static int j;
  static int i;


  if(r > l) {
    m = a[r]; i = l-1; j = r;
    for(;;) {
      while(a[++i] < m);
      while(a[--j] > m);
      if(i >= j) break;
      std::swap(a[i], a[j]);
    }
    std::swap(a[i],a[r]);
    QuickSort(a,l,i-1);
    QuickSort(a,i+1,r);
  }

}

void MySort(float a[], int size)
{

	bool swaped;
	int i;
	
	swaped = true;
	while(swaped)
	{
		swaped = false;
		for(i=0;i<size-1;i++)
		{
			if(a[i]>a[i+1]) 
			{
				swaped = true; 
				std::swap(a[i],a[i+1]);
			}
		}
		if(swaped)
		{
			swaped = false;
			for(i=size-1;i>0;i--)
			{
				if(a[i-1]>a[i])
				{
					swaped = true;
					std::swap(a[i-1],a[i]);
				}
			}
		}

	}

}






void __stdcall clearline_getversion(double *ver)
{
//	Msg("Warning, this is not normal Clearline.dll, this is for the Pro Pipe stuff!!!!");
	*ver=VERSION;
}

void __stdcall clearLine_SmoothOutGraphSingle(float *_dataGraphSource, float *_dataGraphDest, int _numberFrames, int abs) //0 for not absolute, 1 for absolute (0 for capacity)
{
	float sampleData[5];
	float *copy = new float[_numberFrames];
	float total;
	int i;
	int j;
	int tot;
	int adjust;

	copy = new float[_numberFrames];
	for(i=0;i<_numberFrames;i++)
	{
		adjust=0;
		if(i==0) adjust=2;
		if(i==1) adjust=1;
		if(i==_numberFrames-1) adjust = -2;
		if(i==_numberFrames-2) adjust = -1;

		if(abs == 0)
		{
			for(j=0;j<5;j++) 
			{
				sampleData[j]=(_dataGraphSource[i+j-2+adjust]);
			}
		}
		if(abs != 0)
		{
			for(j=0;j<5;j++) 
			{
				sampleData[j]=fabs(_dataGraphSource[i+j-2+adjust]);
			}
		}

		MySort(sampleData,5);
		copy[i]=sampleData[2];
	}

	for(i=0;i<_numberFrames;i++)
	{
		tot=0;
		total=0;
		adjust=0;
		if(i==0) adjust=2;
		if(i==1) adjust=1;
		if(i==_numberFrames-1) adjust = -2;
		if(i==_numberFrames-2) adjust = -1;
		for(j=i-2;j<i+3;j++)
		{
			if(_dataGraphDest[j+adjust]!=-1000000000)
			{
				total += copy[j+adjust];tot++;
			}
		}
		if(tot==0) 
		{
			_dataGraphDest[i] = -1000000000;
		}
		else _dataGraphDest[i] = total /=tot;
	}
	delete[] copy;

}

void __stdcall clearLine_SmoothOutGraphDouble(double *_dataGraphSource, float *_dataGraphDest, int _numberFrames, int abs) //0 for not absolute, 1 for absolute (0 for capacity)
{
	float sampleData[5];
	float *copy = new float[_numberFrames];
	float total;
	int i;
	int j;
	int tot;
	int adjust;

	copy = new float[_numberFrames];
	for(i=0;i<_numberFrames;i++)
	{
		if((i==_numberFrames-1) || (i == _numberFrames-2))
		{
			__asm nop
		}
		adjust=0;
		if(i==0) adjust=2;
		if(i==1) adjust=1;
		if(i==_numberFrames-1) adjust = -2;
		if(i==_numberFrames-2) adjust = -1;
		
		if(abs == 0)
		{
			for(j=0;j<5;j++) 
			{
				sampleData[j]=(_dataGraphSource[i+j-2+adjust]);
			}
		}
		if(abs != 0)
		{
			for(j=0;j<5;j++) 
			{
				sampleData[j]=fabs(_dataGraphSource[i+j-2+adjust]);
			}
		}
		
		MySort(sampleData,5);
		copy[i]=sampleData[2];
	}

	for(i=0;i<_numberFrames;i++)
	{
		if((i==_numberFrames-1) || (i == _numberFrames-3))
		{
			__asm nop
		}
		tot=0;
		total=0;
		adjust=0;
		if(i==0) adjust=2;
		if(i==1) adjust=1;
		if(i==_numberFrames-1) adjust = -2;
		if(i==_numberFrames-2) adjust = -1;
		for(j=i-2;j<i+3;j++)
		{
			if(_dataGraphDest[j+adjust]!=-1000000000)
			{
				total += copy[j+adjust];tot++;
			}
		}
		if(tot==0) 
		{
			_dataGraphDest[i] = -1000000000;
		}
		else _dataGraphDest[i] = total /=tot;
	}
	delete[] copy;
}


void __stdcall clearLine_FilterGraphSingle(float *_dataGraph, int _numberFrames, float _slope)
{

	int i;
	float left;
	float gradientLeft;
	float gradientRight;

	left = _dataGraph[0];


	for(i=1;i<_numberFrames-1;i++)
	{
//		if(i==1748)
//		{
//			__asm nop;
//		}
		if((left-_dataGraph[i])==0) {left = _dataGraph[i]; continue;}
		gradientLeft = 1 / (left-_dataGraph[i]);
		if(fabs(gradientLeft) > _slope) {left = _dataGraph[i]; continue;}

		if((_dataGraph[i]-_dataGraph[i+1])==0) {left = _dataGraph[i]; continue;}
		gradientRight = 1/(_dataGraph[i]-_dataGraph[i+1]);
		if(fabs(gradientRight) > _slope) {left = _dataGraph[i]; continue;}

		if((gradientLeft > 0) && (gradientRight > 0)) {left = _dataGraph[i]; continue;}
		if((gradientLeft < 0) && (gradientRight < 0)) {left = _dataGraph[i]; continue;}

		left = _dataGraph[i];

		_dataGraph[i]=-1000000000;
	}

}

void __stdcall clearLine_FilterGraphDouble(double *_dataGraph, int _numberFrames, float _slope)
{

	int i;
	float left;
	float gradientLeft;
	float gradientRight;

	left = _dataGraph[0];


	for(i=1;i<_numberFrames-1;i++)
	{
//		if(i==1748)
//		{
//			__asm nop;
//		}
		if((left-_dataGraph[i])==0) {left = _dataGraph[i]; continue;}
		gradientLeft = 1 / (left-_dataGraph[i]);
		if(fabs(gradientLeft) > _slope) {left = _dataGraph[i]; continue;}

		if((_dataGraph[i]-_dataGraph[i+1])==0) {left = _dataGraph[i]; continue;}
		gradientRight = 1/(_dataGraph[i]-_dataGraph[i+1]);
		if(fabs(gradientRight) > _slope) {left = _dataGraph[i]; continue;}

		if((gradientLeft > 0) && (gradientRight > 0)) {left = _dataGraph[i]; continue;}
		if((gradientLeft < 0) && (gradientRight < 0)) {left = _dataGraph[i]; continue;}

		left = _dataGraph[i];

		_dataGraph[i]=-1000000000;
	}

}




void __stdcall clearline_RefShapeDistCalc(ReferenceShape_V10 *_Shape,
										  float _X, 
										  float _Y, 
										  double *_OrthoX, 
										  double *_OrthoY, 
										  double *_OrthoDistance, 
										  double _ShapeRadius,
										  double _ShapeCentreX,
										  double _ShapeCentreY,
										  double _ShapeRotation)
{

	Shapes shape(_Shape,
				 _ShapeRadius,
				 _ShapeCentreX,
				 _ShapeCentreY,
				 _ShapeRotation);
	shape.ProfileRefShapeDistCalc(_X,_Y,_OrthoX,_OrthoY,_OrthoDistance);

}

void __stdcall clearline_AutoRotate(ReferenceShape_V10 *_Shape,
									double _ShapeRadius,
									double _ShapeCentreX,
									double _ShapeCentreY,
									double _ShapeRotation,
									float *_PVDataX,
									float *_PVDataY,
									float *_PVCentreX,
									float *_PVCentreY,
									int _FromFrame,
									int _ToFrame,
									int *_EgnoreList,
									HWND _Hwnd,
									float _ScreenWidth,
									float _ScreenHeight,
									double _ScreenRatio)
{
 
	AutoRotate autorotate(_Shape,
						  _ShapeRadius,
						  _ShapeCentreX,
						  _ShapeCentreY,
						  _ShapeRotation,
						  _PVDataX,
						  _PVDataY,
						  _PVCentreX,
						  _PVCentreY,
						  _EgnoreList,
						  _Hwnd,
						  _ScreenWidth,
						  _ScreenHeight,
						  _ScreenRatio);
	autorotate.CalculateRotation(_FromFrame,_ToFrame);

}

//void __stdcall	clearline_CalculateFlat3d(float *pvDataX, 
//											float *pvDataY,
//											float *pvCentreX,
//											float *pvCentreY,
//											int *egnoreList,
//											int pvDataXYMultiplier, 
//											int *pvColourRed, 
//											int *pvColourGreen,
//											int *pvColourBlue,
//											double pvExpectedDiameter,
//											double pvRedLimit,
//											double pvBlueLimit,
//											int fromFrame,
//											int toFrame,
//											double offsetX,
//											double offsetY,
//											ReferenceShape_V10 *_Shape,
//											double _ShapeRotation)

void __stdcall	clearline_CalculateFlat3d(float *pvDataX, 
											float *pvDataY,
											float *pvCentreX,
											float *pvCentreY,
											int *egnoreList,
											int pvDataXYMultiplier, 
											int *pvColourRed, 
											int *pvColourGreen,
											int *pvColourBlue,
											double pvExpectedDiameter,
											double pvRedLimit,
											double pvBlueLimit,
											int fromFrame,
											int toFrame,
											float *offsetX,
											float *offsetY,
											ReferenceShape_V10 *_Shape,
											double _ShapeRotation,
											double *_graphData, //PCN4974
											int _shadingType) //PCN4974

{
	
	Flat3D flat3d(pvDataX,
				  pvDataY,
				  pvCentreX,
				  pvCentreY,
				  egnoreList,
				  pvDataXYMultiplier,
				  pvColourRed,
				  pvColourGreen,
				  pvColourBlue,
				  pvExpectedDiameter,
				  pvRedLimit,
				  pvBlueLimit,
				  fromFrame,
				  toFrame,
				  offsetX,
				  offsetY,
				  _Shape,
				  _ShapeRotation,
				  _graphData,  //PCN4974
				  _shadingType);  //PCN4974
	
	// All frames start at frame 1, and profile points are from 1 to 180 inclusive
	if(fromFrame<1) fromFrame=1; 

	flat3d.CalculateFlat3D(fromFrame, toFrame);

}

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

void __stdcall clearline_CalculateCapacity(float *pvDataX, 
										   float *pvDataY,
										   float *pvCentreX,
										   float *pvCentreY,
										   int *egnoreList,
										   float *pvCapacityFullData,
										   int pvDataXYMultiplier, 
										   double pvExpectedDiameter,
										   int fromFrame,
										   int toFrame)
{
	
	Capacity capacity(pvDataX,
					 pvDataY,
					 pvCentreX,
					 pvCentreY,
					 egnoreList,
					 pvCapacityFullData,
					 pvDataXYMultiplier,
					 pvExpectedDiameter,
					 fromFrame,
					 toFrame);
	// All frames start at frame 1, and profile points are from 1 to 180 inclusive
	if(fromFrame<1) fromFrame=1; 

	
	capacity.CalculateCapacity();
	
}

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

void __stdcall clearline_CalculateOvality(float *pvDataX, 
										  float *pvDataY,
										  float *pvCentreX,
										  float *pvCentreY,
										  int *egnoreList,
										  float *pvOvalityFullData,
										  int pvDataXYMultiplier, 
										  int fromFrame,
										  int toFrame)
{
//	int i;

	if(fromFrame<1) fromFrame=1; 
	Ovality ovality(pvDataX,
					 pvDataY,
					 pvCentreX,
					 pvCentreY,
					 egnoreList,
					 pvOvalityFullData,
					 pvDataXYMultiplier,
					 fromFrame,
					 toFrame,
					 0,
					 0,
					 0,
					 0);
	// All frames start at frame 1, and profile points are from 1 to 180 inclusive
	ovality.CalculateOvality();
	
//	if(fromFrame==1)
//		for(i=0;i<180;i++)
//			if(egnoreList[i]==1) 
//			{
//				FilterGraph FilterGraph(pvOvalityFullData, toFrame);
//				FilterGraph.Smooth();
//				return;
//			}

}

void __stdcall clearline_CalculateDebugOvality(float *pvDataX, 
											   float *pvDataY,
											   float *pvCentreX,
											   float *pvCentreY,
											   int *egnoreList,
											   float *pvOvalityFullData,
											   int pvDataXYMultiplier, 
											   int frame,
											   HWND hwnd,
											   float screenWidth,
											   float screenHeight,
											   double screenRatio)
{
	/*
	if(frame<1) frame = 1;
	Ovality ovality(pvDataX,
					 pvDataY,
					 pvCentreX,
					 pvCentreY,
					 egnoreList,
					 pvOvalityFullData,
					 pvDataXYMultiplier,
					 frame,
					 frame,
					 hwnd,
					 screenWidth,
					 screenHeight,
					 screenRatio);
	// All frames start at frame 1, and profile points are from 1 to 180 inclusive
	ovality.CalculateOvality();
	*/
}


////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

void __stdcall clearline_CalculateXYDiameter(float *pvDataX, 
										  float *pvDataY,
										  float *pvCentreX,
										  float *pvCentreY,
										  int *egnoreList,
										  double *pvXDiameterFullData,
										  double *pvYDiameterFullData,
										  int pvDataXYMultiplier, 
										  int fromFrame,
										  int toFrame)
{
	
	XYDiameter XYDiameter(pvDataX,
					 pvDataY,
					 pvCentreX,
					 pvCentreY,
					 egnoreList,
					 pvXDiameterFullData,
					 pvYDiameterFullData,
					 pvDataXYMultiplier,
					 fromFrame,
					 toFrame);
	// All frames start at frame 1, and profile points are from 1 to 180 inclusive
	if(fromFrame<1) fromFrame=1; 
	XYDiameter.CalculateXYDiameter();
	
}			

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

void __stdcall clearline_CalculateDiameterMaxMin(float *pvDataX, 
										  float *pvDataY,
										  float *pvCentreX,
										  float *pvCentreY,
										  int *egnoreList,
										  double *pvXYDiameterFullMax,
										  double *pvXYDiameterFullMin,
										  short *pvXYDiameterSegFullMax,
										  short *pvXYDiameterSegFullMin,
										  int pvDataXYMultiplier, 
										  int fromFrame,
										  int toFrame)
{
	
	XYDiameterMaxMin XYDiameterMaxMin(pvDataX,
					 pvDataY,
					 pvCentreX,
					 pvCentreY,
					 egnoreList,
					 pvXYDiameterFullMax,
					 pvXYDiameterFullMin,
					 pvXYDiameterSegFullMax,
					 pvXYDiameterSegFullMin,
					 pvDataXYMultiplier,
					 fromFrame,
					 toFrame);
	// All frames start at frame 1, and profile points are from 1 to 180 inclusive
	if(fromFrame<1) fromFrame=1; 

	XYDiameterMaxMin.CalculateXYDiameterMaxMin();
	
}		

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////




void __stdcall clearline_CalculateDeltaMaxMin(float *pvDataX, 
										  float *pvDataY,
										  float *pvCentreX,
										  float *pvCentreY,
										  int *egnoreList,
										  double *pvDeltaFullMax,
										  double *pvDeltaFullMin,
										  short *pvDeltaSegFullMax,
										  short *pvDeltaSegFullMin,
										  int pvDataXYMultiplier, 
										  int fromFrame,
										  int toFrame)
{
	
	DeltaMaxMin DeltaMaxMin(pvDataX,
					 pvDataY,
					 pvCentreX,
					 pvCentreY,
					 egnoreList,
					 pvDeltaFullMax,
					 pvDeltaFullMin,
					 pvDeltaSegFullMax,
					 pvDeltaSegFullMin,
					 pvDataXYMultiplier,
					 fromFrame,
					 toFrame);
	// All frames start at frame 1, and profile points are from 1 to 180 inclusive
	if(fromFrame<1) fromFrame=1; 
	DeltaMaxMin.CalculateDeltaMaxMin();
	
}		


////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

void __stdcall clearline_CalculateDiameterMedian(float *pvDataX, 
										  float *pvDataY,
										  float *pvCentreX,
										  float *pvCenterY,
										  int *egnoreList,
										  double *pvMedianFullData,
										  int pvDataXYMultiplier, 
										  int fromFrame,
										  int toFrame)
{

	if(fromFrame<1) fromFrame=1; 
	Median Median(pvDataX,
					 pvDataY,
					 pvCentreX,
					 pvCenterY,
					 egnoreList,
					 pvMedianFullData,
					 pvDataXYMultiplier,
					 fromFrame,
					 toFrame);
	// All frames start at frame 1, and profile points are from 1 to 180 inclusive

	Median.CalculateMedian();

	
	
}	

void __stdcall clearline_CalculateExceededLimits(float *_pvDataOne, 
												 float *_pvDataTwo,
												 double *_pvDistance,
												 int arraySize, 
                                                 float limitUp, 
												 float limitDown,
                                                 float *answer,
												 int twoGraphs)
{
	double totalDistance=0;
	double exceededDistance=0;
	int i;

	if(twoGraphs!=0) 
	{
		if(limitUp==limitDown)
		{
			for(i=0;i<arraySize-1;i++)
			{
				totalDistance+=fabs(_pvDistance[i+1]-_pvDistance[i]);
				if((_pvDataOne[i]>limitUp)  || (_pvDataTwo[i]>limitUp) ) exceededDistance+=fabs(_pvDistance[i+1]-_pvDistance[i]);
			}
		}
			
			

		else
		{
			
			for(i=0;i<arraySize-1;i++)
			{
				totalDistance+=fabs(_pvDistance[i+1]-_pvDistance[i]);
				if((_pvDataOne[i]>limitUp) || (_pvDataOne[i]<limitDown) || (_pvDataTwo[i]>limitUp) || (_pvDataTwo[i]<limitDown)) exceededDistance+=fabs(_pvDistance[i+1]-_pvDistance[i]);
			}
		}
	}
	else
	{
		if(limitUp==limitDown)
		{
			for(i=0;i<arraySize-1;i++)
			{
				totalDistance+=fabs(_pvDistance[i+1]-_pvDistance[i]);
				if(_pvDataOne[i]>limitUp) exceededDistance+=fabs(_pvDistance[i+1]-_pvDistance[i]);
			}
			
		}
		else
		{
			for(i=0;i<arraySize-1;i++)
			{
				totalDistance+=fabs(_pvDistance[i+1]-_pvDistance[i]);
				if((_pvDataOne[i]>limitUp) || (_pvDataOne[i]<limitDown)) exceededDistance+=fabs(_pvDistance[i+1]-_pvDistance[i]);
			}
		}
	}
	if(totalDistance==0) {*answer = 0; return;}
	*answer = exceededDistance/totalDistance*100; return;
}


////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

void __stdcall clearline_CalculateFractile(float *_pvData,
											int _ArraySize,
											float _Fractile,
											float &_Answer)
{
	
	if (_ArraySize<=2) {_Answer = 0; return;}

	float FractileIndex;
	float Ans;
	float *SortedData;

	SortedData = new float[_ArraySize];
	memcpy(SortedData,_pvData,sizeof(float)*(_ArraySize));
	//QuickSort(SortedData, (int) 1, _ArraySize);
	MySort(SortedData,_ArraySize);
	FractileIndex = (float) _ArraySize * _Fractile / 100;
	if(FractileIndex<1) FractileIndex = 1;
	if(FractileIndex>=_ArraySize) FractileIndex = _ArraySize;
	Ans = SortedData[(int) FractileIndex];
	delete[] SortedData;
	_Answer = Ans;
	
	
	//Fracts Fracts();
	//Fracts

	//Fracts Fracts(float *_pvData,
	//			  int _ArraySize,
	//			  float _Fractile);
	
	//Fracts();//
		
	//_Answer = Fracts.CalculateFractile();
	
}	



////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

void __stdcall clearline_CalculatePercentile(float *pvDataX, 
										  float *pvDataY,
										  float *pvCentreX,
										  float *pvCenterY,
										  int *egnoreList,
										  float *pvMedianFullData,
										  int pvDataXYMultiplier, 
										  int fromFrame,
										  int toFrame)
{
	
	if(fromFrame<1) fromFrame=1; 
	Percentile Percentile(pvDataX,
					 pvDataY,
					 pvCentreX,
					 pvCenterY,
					 egnoreList,
					 pvMedianFullData,
					 pvDataXYMultiplier,
					 fromFrame,
					 toFrame);
	// All frames start at frame 1, and profile points are from 1 to 180 inclusive

	Percentile.CalculatePercentile();
	
}	


////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

 void __stdcall clearline_LoadPVD_Data(char *pvFileName,
										   int pvDataStartAddress,
										   int pvDataBlockSize,
										   int xy,
										   float *pvDataX,
                                           float *pvDataY,
                                           double pvDataMultiplier,
                                           int fromFrame,
                                           int toFrame) //PCN3603		
 {
	 
	 LoadPVD LoadPVD(pvFileName,
				  pvDataStartAddress,
				  pvDataBlockSize,
				  xy,
				  pvDataX,
				  pvDataY,
				  pvDataMultiplier,
				  fromFrame,
				  toFrame);

	 LoadPVD.LoadPVDData();
	 
 }
		
 
 void __stdcall clearline_CalculateCentre(float *_pvDataX, 
										   float *_pvDataY,
										   float *_pvCentreX,
										   float *_pvCentreY,
										   int *_egnoreList,
										   int _fromFrame,
										   int _toFrame,
										   int _waterLevelCentre,
										   int _edgeCentre)
 {
	 
	 if(_fromFrame<1) _fromFrame=1; 
	 Centre Centre(_pvDataX,
					_pvDataY,
					_pvCentreX,
					_pvCentreY,
					_egnoreList,
					_fromFrame,
					_toFrame,
					_waterLevelCentre,
					_edgeCentre);
	 Centre.CalculateCentre();

 }


 void __stdcall clearline_CalculateShapeCentre(ReferenceShape_V10 *_Shape,
									double _ShapeRadius,
									double _ShapeCentreX,
									double _ShapeCentreY,
									double _ShapeRotation,
									float *_pvDataX,
									float *_pvDataY,
									float *_pvCentreX,
									float *_pvCentreY,
									int _fromFrame,
									int _toFrame,
									int *_egnoreList,
									HWND _hwnd,
									float _screenWidth,
									float _screenHeight,
									double _screenRatio)
 {
	 /*
	 if(_fromFrame<1) _fromFrame=1; 
	 Centre Centre(_Shape,
					_ShapeRadius,
					_ShapeCentreX,
					_ShapeCentreY,
					_ShapeRotation,
					_pvDataX,
					_pvDataY,
					_pvCentreX,
					_pvCentreY,
					_egnoreList,
					_fromFrame,
					_toFrame,
					_hwnd,
					_screenWidth,
					_screenHeight,
					_screenRatio);
	 Centre.CalculateCentre();

	 Centre.CalculateCentre();
*/
 }


void __stdcall clearline_MoveFileData(char *_pvFileName, int _FromFilePosition, int _ToFilePosition)
{
	
	FILE *f;
	fpos_t a;
	long totalFileLength;
	char *fileData;
	long dataBufferSize;


	int f_forFileLength;
   

   	_FromFilePosition--;
	_ToFilePosition--;

    f_forFileLength = _open( _pvFileName, _O_RDWR | _O_CREAT, _S_IREAD);
	totalFileLength = _filelength(f_forFileLength);
    _close(f_forFileLength);
 
	if(totalFileLength - _FromFilePosition <= 0) return; //If the data to be moved is allready at the end of the
														 //then it doesn't have to do anything.
	
	f = fopen(_pvFileName,"r+b");
	
	dataBufferSize = totalFileLength - _FromFilePosition;
	
	fileData = new char[dataBufferSize];
	
	a = (long) _FromFilePosition;
	fsetpos( f,  &a);
	fread(fileData,1,dataBufferSize,f);
	
	a = (long) _ToFilePosition;
	fsetpos(f, &a);
	fwrite(fileData,1,dataBufferSize,f);

	fclose(f);
	delete[] fileData;


//	EmbeddedFile embeddedFile();
//	embeddedFile.MoveFileData(_pvFileName,_FromFilePosition,_ToFilePosition);
	
}

void __stdcall clearline_EmbedFileData(char *_pvFileName, char *_EmbFilePosition, int _ToFilePosition)
{
	
	FILE *f;
	fpos_t a;
	long embFileLength;
	char *fileData;
	long dataBufferSize;

	int f_forFileLength;

	_ToFilePosition--;
   
    f_forFileLength = _open( _EmbFilePosition, _O_RDWR | _O_CREAT, _S_IREAD);
	embFileLength = _filelength(f_forFileLength);
    _close(f_forFileLength);
 
	f = fopen(_EmbFilePosition,"rb");
	dataBufferSize = embFileLength;
	fileData = new char[dataBufferSize];
	fread(fileData,1,dataBufferSize,f);
	fclose(f);

	f = fopen(_pvFileName,"r+b");
	
	a = (long) _ToFilePosition;
	fsetpos(f, &a);
	fwrite(fileData,1,dataBufferSize,f);
	fclose(f);

	delete[] fileData;
	
}

void __stdcall clearline_ExtractEmbedFile(char *_pvFileName, char *_EmbFile, int _FromFilePosition, int _FileLength)
{
	
	FILE *f;
	fpos_t a;
//	long embFileLength;
	char *fileData;
//	long dataBufferSize;

//	int f_forFileLength;

	_FromFilePosition--;
   
	fileData = new char[_FileLength];

	f = fopen(_pvFileName, "rb");
	a = (long) _FromFilePosition;
	fsetpos(f, &a);
	

	fread(fileData,1,_FileLength,f);
	fclose(f);

	f = fopen(_EmbFile,"wb");
	

	fwrite(fileData,1,_FileLength,f);
	fclose(f);

	delete[] fileData;
	
}

 /* Dont put back under the windows 2000 testing
 void __stdcall clearline_TestingDrawingLines(HWND whnd)
 {
	 vec2double point1(100,100);
	 vec2double point2(200,200);

	 vec2double point3(100,200);
	 vec2double point4(200,100);


	 vec2double intersect;
	 bool line1, line2;

	 HDC hdc;
	 hdc = GetDC(whnd);

	 Gdiplus::GdiplusStartupInput gdiplusStartupInput;
	 Gdiplus::GdiplusStartup(&m_gdiplusToken, &gdiplusStartupInput, NULL);

	 Graphics graphics(hdc);


	    // Create a Pen object.
   Pen blackPen(Color(255, 0, 0, 0), 1);
   Pen cirPen(Color::Red, 2);

   // Initialize the coordinates of the points that define the line.
   int x1 = 100;
   int y1 = 100;
   int x2 = 200;
   int y2 = 200;

   // Draw the line.
   ::Intersection().TwoLines(point1, point2, 
	 				       point3, point4,
						   intersect,
						   line1, line2);

	if(line1 && !line2) cirPen.SetColor(Color(255, 0, 0, 255));
	if(line1 && line2) cirPen.SetColor(Color(255,255,0,0));
	if(!line1 && line2) cirPen.SetColor(Color(255,0,255,0));
	if(!line1 && !line2) cirPen.SetColor(Color(225,0,0,0));

   graphics.DrawLine(&blackPen, (int) point1.x, (int) point1.y, (int) point2.x, (int) point2.y);
   graphics.DrawLine(&blackPen, (int) point3.x, (int) point3.y, (int) point4.x, (int) point4.y);
   graphics.DrawArc(&cirPen, (int) intersect.x-5, (int) intersect.y-5, 10, 10, 0,360);

   //Gdiplus::GdiplusShutdown(m_gdiplusToken);

 }
*/


vec2double GetProfileIntersection(vec2double point,vec2double *profilePoints)
{
	
	int i;
	vec2double a,b;
	vec2double intersect;
	bool section;
	bool orig;


	for(i=44;i<134;i++)
	{
		a = profilePoints[i];
		b = profilePoints[i+1];
		::Intersection().TwoLines(point,vec2double(0,0),a,b,intersect,section,orig);
		if(orig) return intersect;
	}
	return vec2double(0,0);
	
}

void __stdcall clearLine_EditProfile(float *pvDataX, 
									  float *pvDataY,
									  float *pvCentreX,
									  float *pvCentreY,
									  float *pvCentAdjX,
									  float *pvCentAdjY,
									  int *egnoreList,
									  int fromFrame,
									  int toFrame,
									  double _diameter,
									  float *_graphData)
{

	if(fromFrame<1) fromFrame=1; 
	EditProfile EditProfile(pvDataX,
							pvDataY,
							pvCentreX,
							pvCentreY,
							pvCentAdjX,
							pvCentAdjY,
							egnoreList,
							fromFrame,
							toFrame,
							_diameter,
							0);

	EditProfile.FilterCrap();
}	

void __stdcall clearline_CalculateDeflectionXY(double *diameterX, double *diameterY, 
											   double *medianDiameter, 
											   float *deflectionX, float *deflectionY,
											   int fromFrame, int toFrame)
{
	int i;
	float medDia;
	float defX;
	float defY;
	float diaX;
	float diaY;
	return;
	for(i=fromFrame;i<toFrame;i++)
	{

		medDia=medianDiameter[i];
		
		if (medDia>-100000000 && medDia!=0) 
		{

			diaX = diameterX[i];
			diaY = diameterY[i];

			if(diaX>-100000000)
			{
				diaX = fabs(diaX);
				defX = (diaX - medDia) /medDia * 100;
				//if(diaX<medDia) defX*=-1;
				deflectionX[i]=defX;
			}
			else
			{
				deflectionX[i]=-1000000000;
			}

			if(diaY>-100000000)
			{
				diaY = fabs(diaY);
				defY = (diaY - medDia) /medDia * 100;
				//if(diaY<medDia) defY*=-1;
				deflectionY[i]=defY;
			}
			else
			{
				deflectionY[i]=-1000000000;
			}
			

		}
		else
		{
			deflectionX[i]=-1000000000;
			deflectionY[i]=-1000000000;
		}
	}
}

void __stdcall clearline_CalculateDeflectionXYSmooth(float *diameterX, float *diameterY, 
											   float *medianDiameter, 
											   float *deflectionX, float *deflectionY,
											   int fromFrame, int toFrame)
{
	int i;
	float medDia;
	float defX;
	float defY;
	float diaX;
	float diaY;

	for(i=fromFrame;i<toFrame;i++)
	{
	
		medDia=medianDiameter[i];
		
		if (medDia>=-100000000 && medDia!=0) 
		{

			diaX = diameterX[i];
			diaY = diameterY[i];

			if(diaX>-100000000)
			{
				diaX = fabs(diaX);
				defX = (diaX - medDia) /medDia * 100; //PCN
				if(defX<-50 || defX>50) deflectionX[i] =  -1000000000;
				else deflectionX[i]=defX;
			}
			else
			{
				deflectionX[i]=-1000000000;
			}
					
			if(diaY>-100000000)
			{
				diaY = fabs(diaY);
				defY = (diaY - medDia) /medDia * 100;
				if(defY<-50 || defY>50) deflectionY[i]=-1000000000;
				else deflectionY[i]=defY;
			}
			else
			{
				deflectionY[i]=-1000000000;
			}
			

		}
		else
		{
			deflectionX[i]=-1000000000;
			deflectionY[i]=-1000000000;
		}
	}
}

void __stdcall clearline_calculateLayOfPipe_Richard(double startHeight, double endHeight,
													double *distance, 
													float *centre,
													float *inclination,
													double scale,
													int fromFrame, int toFrame,
													int totalFrames,
													double count)
{

	double drop;
	double distanceBetweenFrames;
	double heightBetweenFrames;
	double frameHeight;
	
	double height=0;
	double adjustedHeight=0;

	int adjustedFrameCount;

	double gradient;
	double currentGradient = 0;
	double nextGradient = 0;
	
	double fallBetweenRL;

	int frameNo;

	adjustedFrameCount = count;
	fallBetweenRL = (endHeight - startHeight) / (distance[toFrame-1] - distance[1]);


	currentGradient = 0;

	//startHeight = ((fallBetweenRL * (distance[toFrame] - distance[1])) - endHeight) * -1;

	for(frameNo = 1;frameNo<toFrame;frameNo++)
	{
		if ((frameNo+adjustedFrameCount)<=toFrame)
		{
			distanceBetweenFrames = fabs(distance[frameNo+adjustedFrameCount] - distance[frameNo])*1000;
			heightBetweenFrames = centre[frameNo+adjustedFrameCount] - centre[frameNo];
		}
		else
		{
			distanceBetweenFrames = fabs(distance[frameNo-adjustedFrameCount] - distance[frameNo])*1000;
			heightBetweenFrames = (centre[frameNo-adjustedFrameCount]-centre[frameNo]);// * -1;
			//distanceBetweenFrames = fabs(distance[toFrame] - distance[frameNo]) * 1000;
			//heightBetweenFrames = centre[toFrame] - centre[frameNo];
		}
		
		gradient = 0;

		if (distanceBetweenFrames != 0)
		{
			gradient = heightBetweenFrames / distanceBetweenFrames;
		}

		

		distanceBetweenFrames = fabs(distance[frameNo+1] - distance[frameNo]);

		nextGradient = currentGradient - gradient;
		if (abs(nextGradient - currentGradient) > 0.05)
		{
			currentGradient = currentGradient;
		}else{
			currentGradient = nextGradient;
		}
		
		//if (frameNo < (adjustedFrameCount)) {currentGradient = fallBetweenRL;}

		//if ((frameNo+(adjustedFrameCount/2))<=toFrame) inclination[frameNo+(adjustedFrameCount/2)] = 100 * gradient * -1; //for percentage inclination

		height = (currentGradient * distanceBetweenFrames) * 6;

		adjustedHeight = height; //(fallBetweenRL * ((distance[frameNo])-(distance[1]))) + startHeight + height;

		if(frameNo >= (toFrame - adjustedFrameCount-1))
		{
			//break;
			__asm nop
		}

		//if (frameNo<=(adjustedFrameCount/2))
		//{
		//	inclination[frameNo]=startHeight;
		//	inclination[frameNo+(adjustedFrameCount/2)]=adjustedHeight;
		//}
		//else if ((frameNo+(adjustedFrameCount/2))<=toFrame)
		//{
		//	inclination[frameNo+(adjustedFrameCount/2)]=adjustedHeight;
		//}
		inclination[frameNo]=adjustedHeight;
	}

	/*inclination[toFrame]=endHeight;*/
	


}

void __stdcall clearline_calcualteLayOfPipe(double startHeight, double endHeight,
											double *distance, 
											float *centre,
											float *inclination,
											double scale,
											int fromFrame, int toFrame,
											int totalFrames,
											double snapOnDistance)
{

	double drop;
	double distanceBetweenFrame;
	double frameHeight;

	vec2double slope;
	
	double height=0;

	double adjustStep=100;
	double adjust;
	int direction = 0;
	

	int frameNo;

	drop = endHeight-startHeight;
	adjust = 100;


	for(frameNo = 1;frameNo<toFrame-1;frameNo++)
	{
		if(frameNo==2400)
		{
			__asm nop
		}
		
		distanceBetweenFrame = fabs(distance[frameNo+1] - distance[frameNo]);
		frameHeight = centre[frameNo]-adjust;

		slope.x = snapOnDistance;
		slope.y = frameHeight;
		slope = slope.toVector();
		slope.y = distanceBetweenFrame;
		slope = slope.toCoordinate();

		height = height + slope.y;
		
		inclination[frameNo]=height;
	}


	while(fabs(adjustStep)>0.01)
	{
		if(inclination[totalFrames-2]>drop) 
		{
			if(direction==-1) adjustStep/=2;
			adjust += adjustStep;
			direction=1;
		}
		if(inclination[totalFrames-2]<drop)
		{
			if(direction==1) adjustStep/=2;
			adjust -= adjustStep;
			direction=-1;
		}

		height=0;
		for(frameNo = 1;frameNo<toFrame-1;frameNo++)
		{
			if(frameNo==2400)
			{
				__asm nop
			}
			
			distanceBetweenFrame = fabs(distance[frameNo+1] - distance[frameNo]);
			frameHeight = centre[frameNo]-adjust;

			slope.x = snapOnDistance;
			slope.y = frameHeight;
			slope = slope.toVector();
			slope.y = distanceBetweenFrame;
			slope = slope.toCoordinate();

			height = height + slope.y;
			
			inclination[frameNo]=height;
		}
	}
	for(frameNo = 1;frameNo<toFrame;frameNo++)
	{
		inclination[frameNo]=inclination[frameNo]+startHeight;
	}


}

//void __stdcall clearline_calcualteLayOfPipe(double startHeight, double endHeight,
//											double *distance, 
//											float *centre,
//											float *inclination,
//											double scale,
//											int fromFrame, int toFrame,
//											int totalFrames,
//											double adjust)
//{
//
//
//	double slope;
//	double slope2;
//	double offset;
//	double incl;
//	int i;
//	float height=0;
//	float angle;
//
//	float distBetweenFrames;
//	float distBetweenFrames2;
//	
//	float heightChangeBetweenFrames;
//	float heightChangeBetweenFrames2;
//	float currentSlope;
//
//	if(scale==0) scale = 1;
//
//	//slope = (endHeight-startHeight)/totalFrames;
//	
//	offset = startHeight;
//	currentSlope=adjust;
//
//
//
//	for(i=fromFrame;i<toFrame-1;i++)
//	{
//		
//		distBetweenFrames = fabs(distance[i+1]-distance[i])*1000;
//		heightChangeBetweenFrames = centre[i+1]-centre[i];
//
//		distBetweenFrames2 = fabs(distance[i+2]-distance[i+1])*1000;
//		heightChangeBetweenFrames2 = centre[i+2]-centre[i+1];
//		
//		if (distBetweenFrames == 0) slope = 0;
//		else slope = heightChangeBetweenFrames / distBetweenFrames;
//
//		if (distBetweenFrames == 0) slope2 = 0;
//		else slope2 = heightChangeBetweenFrames2 / distBetweenFrames2;
//
//
//		currentSlope = currentSlope + (slope2 - slope);
//
//		distBetweenFrames = fabs(distance[i+1]-distance[i]);
//
//		height = height + (distBetweenFrames * currentSlope);
//		//height = height + (heightChangeBetweenFrames);
//		inclination[i] = height;
//		
//
//		//angle = atan(heightChangeBetweenFrames/distBetweenFrames);
//		//if(centre[i+10]>centre[i]) angle*=-1;
//		//angle = angle/PI*180;
//		//inclination[i] = angle;
//	}
//
//	
//
//}

void __stdcall clearline_calcualteInclination(double startHeight, double endHeight, 
											  float *centre, float *inclination, 
											  double scale, 
											  int fromFrame, int toFrame, int totalFrames)
{
	float slope;
	float offset;
	float incl;
	int i;
	float height=0;
	float angle;

	float distBetweenFrames;

	if(scale==0) scale = 1;

	slope = (endHeight-startHeight)/totalFrames;
	
	offset = startHeight;

	for(i=fromFrame;i<=toFrame;i++)
	{
		
		incl = (i*slope) + offset + (centre[i]/scale);
		inclination[i] = incl;
	}

	
}



//void __stdcall clearline_calcualteInclination(double startHeight, double endHeight, 
//											  float *centre, float *inclination, 
//											  double scale, 
//											  int fromFrame, int toFrame, int totalFrames)
//{
//	float slope;
//	float offset;
//	float incl;
//	int i;
//	float height=0;
//
//	if(scale==0) scale = 1;
//
//	slope = (endHeight-startHeight)/totalFrames;
//	//if(startHeight<endHeight) offset = startHeight;
//	//if(startHeight>=endHeight) offset = endHeight;
//	
//	offset = startHeight;
//
//	for(i=fromFrame;i<=toFrame;i++)
//	{
//		
//		incl = (i*slope) + offset + (centre[i]/scale);
//		inclination[i] = incl;
//	}
//
//	
//}

//void __stdcall clearline_calcualteInclination(double startHeight, double endHeight, 
//											  float *centre, float *inclination, 
//											  double scale, 
//											  int fromFrame, int toFrame, int totalFrames)
//{
//	float slope;
//	float offset;
//	float incl;
//	int i;
//	float height=0;
//
//	if(scale==0) scale = 1;
//
//	slope = (endHeight-startHeight)/totalFrames;
//	//if(startHeight<endHeight) offset = startHeight;
//	//if(startHeight>=endHeight) offset = endHeight;
//	
//
//
//	//for(i=fromFrame;i<=toFrame;i++)
//	//{
//	//	
//	//	incl = (i*slope) + offset + (centre[i]/scale);
//	//	inclination[i] = incl;
//	//}
//	height = offset;
//	for(i=fromFrame;i<=toFrame;i++)
//	{
//		height=height+((centre[i]+offset)/scale/-100);
//		incl = (i*slope) + adjust + height;
//		inclination[i] = incl;
//	}
//	
//}

// _x1,_y1 :- First set of profile points
// _x2,_y2 :- Second set of profile points
// _xe,_yr :- Resulting zipped profile
// _cut1 :- Position to cut profiles , represents a angle to cut from, reference by centre
// _cut2 :- Position to cut profiles , represents a angle to cut to, refrence by centre
/*
void __stdcall ZipProfiles(float *_x1,float *_y1void,float *_x2,float *y2, float *_xr, float *_yr, float _cut1, float _cut2)
{
	int no1ProfilePoints;
	int no2ProfilePoints;
	float step;
	float range;
	int pointPos;
	vec2double destPoint;
	float RadToPI = 180*PI;
	int profilePoint;
	float angle;

		

	range = _cut2 - _cut1;
	no1ProfilePoints = (int) (180 / 360 * range);



	no2ProfilePoints = 180 - no1ProfilePoints;

	step = range / no1ProfilePoints;
	
	profilePoint = 0;

	for(angle = _cut1;angle<_cut2;angle+=step)
	{
		destPoint.x=angle;
		destPoint.y=1000;
		destPoint = destPoint.toCoordinate();
	}




}
*/