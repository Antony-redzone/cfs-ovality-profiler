#include "AutoRotate.h"

AutoRotate::AutoRotate(ReferenceShape_V10 *_Shape,
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
					    double _screenRatio)
{
	shapeRadius   = _ShapeRadius;
	shapeCentreX  =	_ShapeCentreX;
	shapeCentreY  =	_ShapeCentreY;
	shapeRotation =	_ShapeRotation;
	pvDataX		  = _PVDataX;
	pvDataY		  = _PVDataY;
	pvCentreX	  = _PVCentreX;
	pvCentreY	  = _PVCentreY;
	egnoreList = _EgnoreList;

	hwnd = _hwnd;
	screenWidth = _screenWidth;
	screenHeight = _screenHeight;
	screenRatio = _screenRatio;
	screenCentre.x = _screenWidth / 2;
	screenCentre.y = _screenHeight / 2;

	shape = new Shapes(_Shape,
					   _ShapeRadius,
					   _ShapeCentreX,
					   _ShapeCentreY,
					   _ShapeRotation);
}

AutoRotate::~AutoRotate(void)
{
	delete shape;
	delete[] rotationWeight;
}

void AutoRotate::CalculateRotation(int _FromFrame, int _ToFrame)
{

//	if(hwnd != 0) InitGraphics();
	rotationWeight = new double[_ToFrame+1];
	if(_FromFrame<1 || _ToFrame<1) return;
	for(currentFrame=_FromFrame;currentFrame<=_ToFrame;currentFrame++) AutoRotateFrame();
}

void AutoRotate::AutoRotateFrame(void)
{
	CreateFakePoints();

//	double lowestWeight;
	double lowestRot;
//	double rot;
//	double left,mid,right;


//
//	double currentWeight;
	
//	currentWeight = GetWeight(0);
//	lowestWeight = currentWeight;
//	lowestRot = 0;

	//for(rot=-PI/20;rot<PI*2;rot+=(PI/270))
//	for(rot=-PI/2;rot<PI/2;rot+=(PI/180))
//	{
//		currentWeight = GetWeight(rot);
//		if(currentWeight<lowestWeight) { lowestWeight=currentWeight; lowestRot = rot;}
//
//	}
	lowestRot = 0;
	FindBestRotation(lowestRot,PI/4);
	
	RotateProfile(lowestRot);
	AdjustTopIndex();
	CopyFakePointsBack();


	rotationWeight[currentFrame]=lowestRot;
}

void AutoRotate::FindBestRotation(double &curPoint, double size)
{
	if(size < PI/1800) return;

	double variance;
	double closestVar;
	bool iTBN=false; // is there better neighbour
	
	double closestRot;
	double lookingAt;

	closestRot=curPoint;
	closestVar=GetWeight(curPoint);

	lookingAt = curPoint - size;
	

	variance=GetWeight(lookingAt);
	if(variance<closestVar) {closestRot=lookingAt; closestVar=variance; iTBN=true;}

	lookingAt=curPoint-(size/2);
	variance=GetWeight(lookingAt);
	if(variance<closestVar) {closestRot=lookingAt; closestVar=variance; iTBN=true;}

	lookingAt = curPoint + (size/2);
	variance=GetWeight(lookingAt);
	if(variance<closestVar) {closestRot=lookingAt; closestVar=variance; iTBN=true;}

	lookingAt = curPoint + size;
	variance=GetWeight(lookingAt);
	if(variance<closestVar) {closestRot=lookingAt; closestVar=variance; iTBN=true;}

	FindBestRotation(closestRot, size/2);
	curPoint=closestRot;
}

void AutoRotate::RotateProfile(double rot)
{
	int i;
	for(i=0;i<180;i++)
	{
		if(fakePoints[i]!=0)
		{
			fakePoints[i]=fakePoints[i].rotateCoordinate(rot);
		}
	}
}

void AutoRotate::CopyFakePointsBack(void)
{
	int pointIndex;
	int point;
	int index;

	index = (currentFrame * 180)+1;

	for(point=0;point<180;point++)
	{
		pointIndex = point+index;
		if(fakePoints[point]==0) {pvDataX[pointIndex]=0; pvDataY[pointIndex]=0;}
		else 
		{
			pvDataX[pointIndex] = (float) fakePoints[point].x - pvCentreX[currentFrame];
			pvDataY[pointIndex] = (float) fakePoints[point].y - pvCentreY[currentFrame];
		}
		
	}

			
//	if(hwnd!=0)
//	{
//		for(point=0;point<180;point++)
//		{
//		DrawLine(fakePoints[point],fakePoints[(point+1)%180],Color::Green);
//		}
//	}



	
}

double AutoRotate::GetWeight(double rot)
{
	int i;
	int count=0;
	vec2double point;
	double ave=0;
	double variance=0;
	double weights[180];

	vec2double ortho;
	double orthoDistance;

	
	

	if(currentFrame == 3550)
	{
		__asm nop;
	}
	for(i = 0;i<180;i++)
	{
		if(fakePoints[i]==0) continue;
		point = fakePoints[i].rotateCoordinate(rot);
		//weights[count] = fabs((shape->ProfileRefShapeDistCalc(point)));
		shape->ProfileRefShapeDistCalc((float) point.x, (float) -point.y,&ortho.x,&ortho.y,&orthoDistance);
		
//		if(hwnd!=0)
//		{
//			ortho.y*=-1;
//			DrawLine(point,ortho,Color::Green);
//		}
			
			
			
			
		weights[count] = fabs(orthoDistance);
		
		ave += weights[count];
		count++;
	}
	if(count==0) return -1;
	ave /= count;

//	for(i=0;i<count;i++) variance+=fabs(ave-weights[i]);
	
//	variance/=(double) count;
//	return variance; // return the average variance
	return ave;

}

void AutoRotate::CreateFakePoints(void)
{

	int pointIndex;
	int point;
	int index;

	index = (currentFrame * 180)+1;

	for(point=0;point<180;point++)
	{
		pointIndex = point+index;
		if(pvDataX[pointIndex]==0 && pvDataY[pointIndex]==0) fakePoints[point]=0;
		else if(egnoreList[point]==1) fakePoints[point]=0;
		else fakePoints[point] = vec2double(pvDataX[pointIndex]+pvCentreX[currentFrame],pvDataY[pointIndex]+pvCentreY[currentFrame]);
		if((fakePoints[point].x>10000) || (fakePoints[point].x<-10000)) 
		{
			fakePoints[point].x = 0;
			fakePoints[point].y = 0;
		}
	}
}

/*
void AutoRotate::InitGraphics(void)
{
	 hdc = GetDC(hwnd);

	 Gdiplus::GdiplusStartupInput gdiplusStartupInput;
	 Gdiplus::GdiplusStartup(&m_gdiplusToken, &gdiplusStartupInput, NULL);

	 graphics = new Graphics(hdc);
}
*/

/*
void AutoRotate::DrawLine(vec2double a, vec2double b, Color colour)
{
	a.y*=-1;
	b.y*=-1;

	a = a / screenRatio; a = a + screenCentre;
	b = b / screenRatio; b = b + screenCentre;

	Pen linePen(colour, 1);
	
 
	
	graphics->DrawLine(&linePen, (int) a.x, (int) a.y, (int) b.x, (int) b.y);
}
*/

/*
void AutoRotate::DrawCircle(vec2double a,double size, Color colour)
{
	a.y*=-1;
	a = a / screenRatio; a = a + screenCentre;
	size = size / screenRatio;
	a.x -=(size/2);
	a.y -=(size/2);

	Pen cirPen(colour, 1);
	graphics->DrawArc(&cirPen, (int) a.x, (int) a.y, (int) size, (int) size, 0,360);
}
*/

int AutoRotate::GetMostVertical(void)
{
	int i;

	int mostVertIndex=0;
	vec2double vect;
	double mostVertAngle;

	vect = fakePoints[0];
	vect = vect.toVector();
	mostVertAngle = fabs(vect.x - 180);
		

	for(i=1;i<180;i++)
	{
		vect = fakePoints[i];
		vect = vect.toVector();
		vect.x = fabs(vect.x-180);
		if(mostVertAngle>vect.x) {mostVertAngle = vect.x; mostVertIndex = i;}
	}
	return mostVertIndex-90;

}

void AutoRotate::AdjustTopIndex(void)
{
	int i;
	int topIndex;
	vec2double tmpFakePoints[180];
	
	topIndex = GetMostVertical();
	for(i=0;i<180;i++)
	{
		tmpFakePoints[i] = fakePoints[(360+i+topIndex)%180];
	}
	for(i=0;i<180;i++)
	{
		fakePoints[i] = tmpFakePoints[i];
	}
}


