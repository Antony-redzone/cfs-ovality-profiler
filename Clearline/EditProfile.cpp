


#include "EditProfile.h"


EditProfile::EditProfile(float *_pvDataX, 
			   float *_pvDataY,
			   float *_pvCentreX,
			   float *_pvCentreY,
			   float *_pvCentAdjX,
			   float *_pvCentAdjY,
			   int *_egnoreList,
			   int _fromFrame,
			   int _toFrame,
			   double _diameter,
			   float *_graphData)
{
	pvDataX=_pvDataX;
	pvDataY=_pvDataY;
	pvCentreX=_pvCentreX;
	pvCentreY=_pvCentreY;
	pvCentAdjX=_pvCentAdjX;
	pvCentAdjY=_pvCentAdjY;
	fromFrame = _fromFrame;
	toFrame= _toFrame;
	egnoreList = _egnoreList;
	cutoffDifference = _diameter/10;
	cutoffHeight = _diameter*1.5;
	topDirection = 90; //57
	diameter = _diameter;
	graphData = _graphData;

}

EditProfile::~EditProfile(void)
{
}

void EditProfile::StampProfileData(void)
{
	currentFrame = toFrame;
	FrameStamp();
}

void EditProfile::FillInNoData(void)
{
	for(currentFrame=fromFrame-1;currentFrame<toFrame;currentFrame++) FrameFill();
}

void EditProfile::FilterCrap(void)
{
	for(currentFrame=fromFrame-1;currentFrame<toFrame;currentFrame++) FrameFilterCrap();
}
void EditProfile::CreatePerfectPipe(void)
{
	for(currentFrame=fromFrame-1;currentFrame<toFrame;currentFrame++) 	FrameCreatePerfectPipe();
}

void EditProfile::FrameStamp(void)
{
	int i;
//	long indexProfileOne;
	long pointIndexFrom;
	long pointIndexTo;

	pointIndexTo = ((fromFrame-1)*180)+1;
	pointIndexFrom = ((toFrame-1)*180)+1;


	for(i=0;i<180;i++)
	{
		pvDataX[pointIndexTo+i] = pvDataX[pointIndexFrom+i];
		pvDataY[pointIndexTo+i] = pvDataY[pointIndexFrom+i];
	}


//	indexProfileOne = (fromFrame*180)+1;
//	CreateFilteredPoints(indexProfileOne);
//	for(i=0;i<180;i++)
//	{
//		fakePointsOther[i] = fakePoints[i];
//	}
//	indexProfileOne = (currentFrame*180)+1;
//	for(i=0;i<180;i++)
//	{
//		fakePoints[i] = fakePointsOther[i];	
//	}
//	StoreProfile();
//
//
//				pvDataX[pointIndex] = pvDataX[pointIndex] - pvCentreX[currentFrame];
//			pvDataY[pointIndex] = pvDataY[pointIndex] - pvCentreY[currentFrame];

}

void EditProfile::FrameFilterCrap(void)
{
	long indexProfileOne;
	
	
	indexProfileOne = (currentFrame*180)+1;
	CreateFilteredPoints(indexProfileOne);
	FilterProfile();
	StoreProfile();
}

void EditProfile::FrameFill(void)
{
	long indexProfileOne;
	
	indexProfileOne = (currentFrame*180)+1;
	CreateFilteredPoints(indexProfileOne);
	FillPerfectPoints();
	StoreProfile();
}

void EditProfile::FrameCreatePerfectPipe(void)
{
	long indexProfileOne;

	indexProfileOne = (currentFrame*180)+1;
	CreateFilteredPoints(indexProfileOne);
	CreateProfile();
	ReSpreadProfile();
	StoreProfile();
}

void EditProfile::CreateFilteredPoints(long index)
{
	long point;
	long pointIndex;
	vec2double coord;

	for(point=0;point<180;point++)
	{
		pointIndex = point+index;
		if(pvDataX[pointIndex]==0 && pvDataY[pointIndex]==0) 
		{
			fakePoints[point]=0;
		}
		//else if(egnoreList[point]==1) fakePoints[point]=0;
		else fakePoints[point] = vec2double(pvDataX[pointIndex]+pvCentreX[currentFrame],pvDataY[pointIndex]+pvCentreY[currentFrame]);
		if((fakePoints[point].x>10000) || (fakePoints[point].x<-10000)) 
		{
			fakePoints[point].x = 0;
			fakePoints[point].y = 0;
		}
		currentCentre=vec2double(-pvCentAdjX[currentFrame],pvCentAdjY[currentFrame]);
	}
}

void EditProfile::StoreProfile(void)
{
	long index;
	long pointIndex;
	long point;

	index=(currentFrame*180)+1;
	for(point=0;point<180;point++)
	{
		pointIndex = point+index;
		if(fakePoints[point].x!=0 || fakePoints[point].y!=0) 
		{
			//if(fakePoints[point].x >=0) pvDataX[pointIndex] = (float) fakePoints[point].x + 20000;	//
			//else pvDataX[pointIndex] = (float) fakePoints[point].x - 20000;
			pvDataX[pointIndex] = (float) fakePoints[point].x;
			pvDataY[pointIndex] = (float) fakePoints[point].y;
			pvDataX[pointIndex] = pvDataX[pointIndex] - pvCentreX[currentFrame];
			pvDataY[pointIndex] = pvDataY[pointIndex] - pvCentreY[currentFrame];
		}
		else
		{
			pvDataX[pointIndex]= 0;
			pvDataY[pointIndex]= 0;
		}
	}
/*

	for(point=0;point<180;point++)										//
	{		
	

												//
		if(((pvDataX[pointIndex]==0) && pvDataY[pointIndex]==0) || egnoreList[point]==1)			//
		{	
			if((fakePoints[point].x!=0) || (fakePoints[point].y!=0))
			{
				if(fakePoints[point].x >=0) pvDataX[pointIndex] = (float) fakePoints[point].x + 20000;	//
				else pvDataX[pointIndex] = (float) fakePoints[point].x - 20000;
				pvDataY[pointIndex] = (float) fakePoints[point].y;
				pvDataX[pointIndex] = pvDataX[pointIndex] - pvCentreX[currentFrame];
				pvDataY[pointIndex] = pvDataY[pointIndex] - pvCentreY[currentFrame];

			}
			else
			{
				pvDataX[pointIndex] = 0;
				pvDataY[pointIndex] = 0;
			}

		}																//	
	}	

*/

}

void EditProfile::FilterProfile(void)
{
	int startPoint;
	int i;

	startPoint = topDirection;

	vec2double topPoint;
	topPoint = FindMedianTopPoint(startPoint);
	for(i=0;i<180;i++) FilterPoint(i,topPoint);
}

void EditProfile::FillPerfectPoints(void)
{
	double Gradiant;
//	int shape;
	double angle;
	int i;
//	int p;
	vec2double tempPoint;



	Gradiant = (PI*2)/180;
	
	for(i=0;i<180;i++)
	{
		angle = i * Gradiant;
		radialPoints[i] = radialPoints[i].toVector();
		radialPoints[i].x = diameter/2;
		radialPoints[i].y = angle;
		radialPoints[i] = radialPoints[i].toCoordinate();
	}



	
	for(i=0;i<180;i++)
	{
		if(fakePoints[i]==0) fakePoints[i] = radialPoints[i];
	}



	 ReOrderProfile();
}




void EditProfile::CreateProfile(void)
{
	int i;
	double Gradiant;

	Gradiant = (PI*2)/180;

	vec2double point;

	for(i=0;i<180;i++)
	{
		point = fakePoints[i];

		point = point.toVector();
		point.y = diameter/2;
		point.x = ((double) i * Gradiant)+PI;

		point = point.toCoordinate();
		
		if(point.y<-(diameter/2)+graphData[currentFrame]) 
		{
			point.y = -(diameter/2)+graphData[currentFrame];
		}
		fakePoints[i] = point;


	}
}

void EditProfile::ReSpreadProfile(void)
{

	vec2double TempPoints[180];
	int i;
	double Gradiant;
	vec2double point;

	Gradiant = (PI*2)/180;
	
	
	for(i=0;i<180;i++)
	{
		point = point.toVector();
		point.x = (Gradiant * i) + PI;
		if((point.x)> (2*PI)) point.x -=(2*PI);
		point.y = diameter;
		point = point.toCoordinate();

		TempPoints[i]=GetProfileIntersection(point);
	}

	for(i=0;i<180;i++)
	{
		fakePoints[i]=TempPoints[i];
	}


}

void EditProfile::FilterPoint(int i,vec2double topPoint)
{
	if(fakePoints[i]==0) return;
	
	int left, right;
	double difference;
	
	
	vec2double pointa, pointb, pointc;
	left = i - 1;
	right = i + 1;
	if(right>179) right -=180;
	if(left<0) left +=180;
	
	pointb = fakePoints[i];
	pointa = fakePoints[left];
	pointc = fakePoints[right];

	if(pointa==0) pointa = pointb;
	if(pointc==0) pointc = pointb;
	

	pointa = pointa - pointb;
	pointc = pointc - pointb;
	pointa = pointa.toVector();
	pointc = pointc.toVector();

	difference = fabs(pointa.y)+fabs(pointc.y);
	if(difference > cutoffDifference) 
	{
		fakePoints[i]=0;
	}


	pointb = pointb - topPoint;
	pointb = pointb.toVector();
	if(fabs(pointb.y)> cutoffHeight) fakePoints[i]=0;
}

vec2double EditProfile::FindMedianTopPoint(int startPoint)
{
	int left;
	int right;
	double yPoints[180];
	double xPoints[180];
	double t;
	int i;
	int index=0;
	int p;
	bool swap;
	vec2double top;

	left=startPoint-2;
	right = startPoint+2;
	for(i=0;i<180;i++,index++)
	{
		p=i;
		if(p<0  ) p +=180;
		if(p>179) p -=180;
		xPoints[index] = fakePoints[p].x;
		yPoints[index] = fakePoints[p].y;
	}
	
	swap=true;
	while(swap)
	{
		swap=false;
		for(i=0;i<179;i++)
		{
			if(xPoints[i]<xPoints[i+1]) {t=xPoints[i]; xPoints[i]=xPoints[i+1]; xPoints[i+1] = t;swap=true;}
			
		}
	}
	swap=true;
	while(swap)
	{
		swap=false;
		for(i=0;i<179;i++)
		{
			if(yPoints[i]<yPoints[i+1]) {t=yPoints[i]; yPoints[i]=yPoints[i+1]; yPoints[i+1] = t;swap=true;}
			
		}
	}

	top.x=0; //xPoints[2];
	top.y=yPoints[0];
	return top;
}

vec2double EditProfile::GetProfileIntersection(vec2double point)
{
	int i;
	vec2double a,b;
	vec2double intersect;
	bool section;
	bool orig;


	for(i=0;i<179;i++)
	{
		a = fakePoints[i];
		b = fakePoints[i+1];
		::Intersection().TwoLines(point,vec2double(0,0),a,b,intersect,section,orig);
		if(section && orig) return intersect;
	}
	a = fakePoints[179];
	b = fakePoints[0];
	::Intersection().TwoLines(point,vec2double(0,0),a,b,intersect,section,orig);
	if(section && orig) return intersect;

	return vec2double(0,0);
}

void EditProfile::ReOrderProfile(void)
{
	vec2double tempPoints[180];
	bool swap;
	int i;

	vec2double point;
	
	
	for(i=0;i<180;i++)
	{



		tempPoints[i] = fakePoints[i].toVector();
	}

	swap=true;
	while(swap==true)
	{
		swap = false;
		for(i=0;i<179;i++)
		{
			if(tempPoints[i].x>tempPoints[i+1].x)
			{
				point = tempPoints[i];
				tempPoints[i] = tempPoints[i+1];
				tempPoints[i+1]=point;
				swap=true;
			}
		}
	}

	for(i=0;i<180;i++)
	{
		fakePoints[i]=tempPoints[i].toCoordinate();
	}

}


