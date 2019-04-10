#include <stdio.h>
#include <math.h>
#include "Video.h"
#include "FishEyeTransform.h"
#include <time.h>

LensTransform::LensTransform(){
	buffer=NULL;
	TFactor		= 1.025;
	xCentre		= 0;
	yCentre		= 0;
	Scale		= 0;
	ImageWidth	= 0;
	ImageHeight	= 0;
	MaskCreated = false;
	Initialised = false;
	Active		= OFF;
	y_scale		= 1; // Setting no y_scale transformation PCN3303
	y_position	= 0; // Setting no y position adjustment  PCN3303
}

//PCN3085 Destructor Added to remove memory leaks (12 October 2004, Antony)
LensTransform::~LensTransform(){
	int i;
	
	for(i=0;i<ImageHeight;i++){
		delete[] buffer[i];
		delete[] Mask[i];
	}
	delete[] buffer;
	delete[] Mask;

	//if(buffer!=0) delete[] buffer;
	//if(Mask!=0) delete[] Mask;
}


void LensTransform::Flush(void){
	pixel empty;
	empty.red = 0;
	empty.blue= 0;
	empty.green=0;
	for(int i=0;i<ImageWidth;i++){
		for(int t=0;t<ImageHeight;t++){
			buffer[t][i]=empty;
//		    Mask[t][i].x = 0;
//			Mask[t][i].y = 0;
		}
	}
}

void LensTransform::Initialize(int CurrentWidth, int CurrentHeight, int BaseWidth, int BaseHeight){
	int i;
	ImageHeight		= CurrentHeight;
	ImageWidth		= CurrentWidth;
	OriginalHeight	= BaseHeight;
	OriginalWidth	= BaseWidth;

	if(OriginalHeight == 0) OriginalHeight = ImageHeight;
	if(OriginalWidth == 0) OriginalWidth = ImageWidth;

//	Msg("Image %dx%d Original %dx%d",CurrentHeight,CurrentWidth,BaseHeight,BaseWidth);
	y_ratio = (double)ImageHeight / OriginalHeight;
	x_ratio = (double)ImageWidth  / OriginalWidth;
	ratio34 = (OriginalWidth * 0.75) / OriginalHeight;
	Divider = (sqrt((double) (OriginalHeight * OriginalHeight) + (double) (OriginalWidth * OriginalWidth)) /2 -1)+10;
	buffer = new pixel * [ImageHeight];
	Mask = new vec2int * [ImageHeight];
	for(i=0;i<ImageHeight;i++){
		buffer[i] = new pixel[ImageWidth];
		Mask[i] = new vec2int[ImageWidth];
	}
	MaskCreated = false;
	Initialised = true;
}

void LensTransform::TurnFishEyeOn(void){
	Active = ON;
	//if((ImageHeight == OriginalHeight && ImageWidth == OriginalWidth) || Scale == 0)	SetDisplayScale();
	//if(MaskCreated == false) CreateMask();
}

void LensTransform::TurnFishEyeOff(void){
	Active = OFF;
}

bool LensTransform::FishEyeStatus(void){
	return Active == ON ? ON : OFF;
}

void LensTransform::SetOffsets(int x,int y){
	xCentre = OriginalWidth/2 + x;
	yCentre = OriginalHeight/2 + y;
}

void LensTransform::SetTFactor(double TF){
	//if( TF < 1.025 || TF > 1.65) { TurnFishEyeOff(); return; } //PCN3595 when TF set below 1.025 then turn off
	if(TF < 1.0 || TF > 1.65) { TurnFishEyeOff(); return; } //PCN3595 when TF set below 1.025 then turn off 'PCN3687
	TurnFishEyeOn();										//otherwise turn on.
	TFactor = TF;
}

void LensTransform::SetImageSize(void){
	OriginalHeight = ImageHeight;
	OriginalWidth = ImageWidth;
	y_ratio = (double)ImageHeight / OriginalHeight;
	x_ratio = (double)ImageWidth  / OriginalWidth;
	ratio34 = (OriginalWidth * 0.75) / OriginalHeight;
	Divider = (sqrt((double) (OriginalHeight *OriginalHeight) + (double) (OriginalWidth * OriginalWidth)) /2 -1)+10;
}

void LensTransform::SetOriginalSize(int width,int height){
	OriginalHeight = height;
	OriginalWidth = width;
	y_ratio = (double)ImageHeight / OriginalHeight;
	x_ratio = (double)ImageWidth  / OriginalWidth;
	ratio34 = (OriginalWidth * 0.75) / OriginalHeight;
	Divider = (sqrt((double) (OriginalHeight *OriginalHeight) + (double) (OriginalWidth * OriginalWidth)) /2 -1)+10;
}

void LensTransform::SetScale(double S){
	Scale = S;
}

double LensTransform::GetScale(void){
	return Scale;
}

void LensTransform::SetDisplayScale(){
	int iteration = 0;

	if(!Initialised) return;
	double inc;
	vec2double heightA, heightB;
	vec2double widthA,widthB;
	Scale = 100;
	
	//if((TFactor < 1.025 || TFactor > 1.65 || xCentre == 0 || yCentre == 0)) return;
	if((TFactor < 1.0 || TFactor > 1.65 || xCentre == 0 || yCentre == 0)) return; //PCN3687

	// PCN3687 replaced xtop ybottom etc, with p, a vec2double
	heightA.x = xCentre / x_ratio;	//xtop = xCentre / y_ratio; // Added this dont forget to comment
	heightA.y = ImageHeight;		//	ytop = ImageHeight;
	heightB.x = xCentre / x_ratio;	//xbottom = xCentre / x_ratio; // Added this dont forget to comment;
	heightB.y = 0;					//ybottom = 0;

	//Added width testing for scaling the image
	widthA.x = ImageWidth;			//xtop = xCentre / y_ratio; // Added this dont forget to comment
	widthA.y = yCentre / y_ratio;	//	ytop = ImageHeight;
	widthB.x = 0;					//xbottom = xCentre / x_ratio; // Added this dont forget to comment;
	widthB.y = yCentre / y_ratio;	//ybottom = 0;

	

	ConvertPoint(heightA); //ConvertPoint(xbottom,ybottom);
	ConvertPoint(heightB); //ConvertPoint(xtop,ytop);
	
	ConvertPoint(widthA);
	ConvertPoint(widthB);

	inc = 0.1; // Decressed the step from 2 to 0.1, this enabled no distortion masks to be closer to proper size and not to big only by a fraction
	//inc = ((int)(ytop - ybottom) < ImageHeight) ? 1.0 : -1.0;
	//while( ( (int)(ytop - ybottom) < ImageHeight))// || ((int)(xbottom-) < ImageWidth) ) // PCN3687
	  while( ( (int)(heightA.y-heightB.y) < ImageHeight) || ((int)(widthA.x-widthB.x) < ImageWidth))
	{
		iteration++; if(iteration>20000)
		{
			break;
		}
		Scale+=inc;
//		if(Scale > 250.0) break;

		heightA.x = xCentre / x_ratio;	// xtop = xCentre / y_ratio;
		heightA.y = ImageHeight;// / y_ratio;		// ytop = ImageHeight;
		heightB.x = xCentre / x_ratio;	// xbottom = xCentre / x_ratio;
		heightB.y = 0;					// ybottom = 0;

		widthA.x = ImageWidth;// / x_ratio;			//xtop = xCentre / y_ratio; // Added this dont forget to comment
		widthA.y = yCentre / y_ratio;	//	ytop = ImageHeight;
		widthB.x = 0;					//xbottom = xCentre / x_ratio; // Added this dont forget to comment;
		widthB.y = yCentre / y_ratio;	//ybottom = 0;

		ConvertPoint(heightA);			// ConvertPoint(xbottom,ybottom);
		ConvertPoint(heightB);			// ConvertPoint(xtop,ytop);

		ConvertPoint(widthA);
		ConvertPoint(widthB);

	}
	//ScaleOffset = ImageHeight - (int)ytop;
}

void LensTransform::ConvertPoint(double &x, double &y){
	double Radius,MidCal,Angle = 0,FloatX,FloatY,OffsetX,OffsetY,X,Y;



	x = x / x_ratio;
	y = y / y_ratio;



	OffsetX = x - (xCentre);///x_ratio); //PCN3687 added x_ration distortion to centre as with below
	OffsetY = y - (yCentre);///y_ratio);

	vec2double p(OffsetX,OffsetY);	

//	Radius = sqrt(OffsetX * OffsetX + OffsetY * OffsetY);
	p = p.toVector();
	Radius = p.y;
	Angle = p.x;
	
	
	MidCal = atan(Radius / Divider);

//	if(OffsetY > 0)							Angle = asin(OffsetX / Radius); else
//	if(OffsetY <=0 && OffsetX < xCentre)	Angle = acos(OffsetX / Radius) + PI/2; 

	FloatX = sin(Angle) * tan(MidCal * TFactor);
	FloatY = cos(Angle) * tan(MidCal * TFactor);
	X = (Scale * FloatX + 0.5) + (xCentre);///x_ratio);
	Y = (Scale * FloatY* y_scale + 0.5) + (yCentre);///y_ratio);

	x=X;
	y=Y;

	
	
	x = x * x_ratio;
	y = y * y_ratio;
	y = y + y_position; // PCN3303, after scaling the y axis, may need to recentre video with y position
	
}

void LensTransform::ConvertPoint(vec2double &point){
	ConvertPoint(point.x,point.y);
}

void LensTransform::Convert (double x, double y){
	int X,Y;
	double tempx,tempy;
	tempx = x;
	tempy = y;

	y = y * ratio34;

	ConvertPoint(tempx,tempy);
	X = (int) tempx;
	Y = (int) tempy;

	y = y / ratio34;


	if(X < ImageWidth && X >=0 && Y <ImageHeight && Y >=0 && y >= 0 && y < ImageHeight && x >= 0 && x < ImageWidth) { 
		Mask[Y][X].x = (int) x;
		Mask[Y][X].y = (int) y;
	}
}

void LensTransform::CreateMask(void){
	double dx,dy,inc;

	inc = 0.4;
	for(dy = 0; dy < ImageHeight; dy += inc){
		for(dx = 0; dx < ImageWidth; dx += inc){
			Convert(dx,dy);
		}

	}
	MaskCreated = true;
}

void LensTransform::Transform(pixel **image){
	int x,y;
	if(Active == OFF) return;
//	Msg("Centre (%d,%d) Scale = %f Original (%d,%d) Image (%d,%d)",xCentre,yCentre,Scale,OriginalHeight,OriginalWidth,ImageHeight,ImageWidth);
	if(MaskCreated == false) CreateMask();
//	Flush();
	for(y = 0; y < ImageHeight; y++){
		for(x = 0; x < ImageWidth; x++){
			if(Mask[y][x].x > 0 && Mask[y][x].y > 0 && Mask[y][x].y < ImageHeight && Mask[y][x].x < ImageWidth){
				buffer[y][x] = image[ Mask[y][x].y ][ Mask[y][x].x ];
			}
		}
	}
}

void LensTransform::CopyToVideo(pixel ** Destination,int xBound,int yBound){
	int i,t;

	if(xBound > ImageWidth) xBound = ImageWidth;
	if(yBound > ImageHeight)yBound = ImageHeight;

	for(i=0;i<xBound;i++){
		for(t=0;t<yBound;t++){
			Destination[t][i] = buffer[t][i];
		}
	}
}

void LensTransform::CopySinglePixel(int x,int y,pixel **Destination){
	Destination[y][x]=buffer[y][x];
}

void LensTransform::ConvertPoint(double &x,double &y,int cx,int cy, double Factor){
	double Radius,MidCal,Angle = 0,FloatX,FloatY,OffsetX,OffsetY,X,Y;


	OffsetX = x - cx;
	OffsetY = y - cy;

	Radius = sqrt(OffsetX * OffsetX + OffsetY * OffsetY);
	MidCal = atan(Radius / Divider);

	if(OffsetY > 0)							Angle = asin(OffsetX / Radius); else
	if(OffsetY <=0 && OffsetX < cx)			Angle = acos(OffsetX / Radius) + PI/2; 

	FloatX = sin(Angle) * tan(MidCal * Factor);
	FloatY = cos(Angle) * tan(MidCal * Factor);
	X = (Scale * FloatX ) + cx;
	Y = (Scale * FloatY ) + cy;

	x=X;
	y=Y;
}

void LensTransform::Find_Centre(int x,int y,double Factor){
	double Sum = 0;
	
	Factor = Find_TFactor(x,y,Factor);
	Sum = Asses_Grid(x,y,Factor);
	if(Sum < best) {
		best = Sum;
		bestx = x;
		besty = y;
		bestTFactor = Factor;
		Find_Centre(x-1,y-1,Factor);
		Find_Centre(x,y-1,Factor);
		Find_Centre(x+1,y-1,Factor);
		Find_Centre(x-1,y,Factor);
		Find_Centre(x+1,y,Factor);
		Find_Centre(x+1,y+1,Factor);
		Find_Centre(x,y+1,Factor);
		Find_Centre(x+1,y+1,Factor);
	}
}

double LensTransform::Find_TFactor(int x, int y,double Factor){
	double current,currentbest = 99999.9,inc;
	
	currentbest = Asses_Grid(x,y,Factor);
	inc = 0.01;
	current = Asses_Grid(x,y,Factor + inc);
	if(current == currentbest) return Factor;
	if(current > currentbest) {
		inc = -0.01;
		current = Asses_Grid(x,y,Factor + inc);
	}

	while(current < currentbest){
		current = Asses_Grid(x,y,Factor += inc);
		currentbest = current;
	}

	
/*
	for(temp = 1.05; temp < 1.65; temp += 0.01){
		Current = Asses_Grid(x,y,temp);
		if(Current < currentbest)	{
			bestTFactor = temp;
			currentbest = Current;
		}
	}
*/
	return Factor;
}

void LensTransform::Load_Grid(void){
	int i,t;
	FILE *f;
	char buffer[10];

	f = fopen("c:\\clearlineprofilerv5.4.2\\grid.dat","r");
	if(!f){
		Msg("Falied to open grid.dat");
		return;
	}

	for(i=0;i<3;i++){
		for(t=0;t<3;t++){
			fgets(buffer,10,f);
			Basegrid[t][i].x = atof(buffer);
			fgets(buffer,10,f);
			Basegrid[t][i].y = atof(buffer);

		}
	}
}

void LensTransform::copy_grid(void){
	int i,t;
	for(i=0;i<3;i++){
		for(t=0;t<3;t++){
			grid[i][t].x = Basegrid[i][t].x;
			grid[i][t].y = Basegrid[i][t].y;
		}
	}
}

double LensTransform::Asses_Grid(int x,int y,double f){
	int i,t;
	double val = 0;

	copy_grid();
	FEGrid(x,y,f);
	for(i=0;i<2;i++){
		for(t=0;t<2;t++){
			val += Square(grid[i][t],grid[i+1][t],grid[i][t+1],grid[i+1][t+1]);
		}
	}
	return val;
}

double LensTransform::Square(vec2double ul,vec2double ur,vec2double bl, vec2double br){
	double val = 0;
	val += fabs(fabs(ul.y - ur.y) - fabs(bl.y - br.y));
	val += fabs(fabs(ul.x - bl.x) - fabs(ur.x - br.x));
	return val;
}

void LensTransform::FEGrid(int x,int y,double f){
	int i,t;
	Scale = 120;
	for(i=0;i<3;i++){
		for(t=0;t<3;t++){
			ConvertPoint(grid[i][t].x,grid[i][t].y,x,y,f);
		}
	}
}

void LensTransform::AutoCalibrate(void){
	best = 99999.9;

	clock_t start,finish;
	start = clock();

	Load_Grid();
	bestx = ImageWidth/2;
	besty = ImageHeight/2;
	
	Find_Centre(bestx,besty,1.35);

	xCentre = bestx;
	yCentre = besty;
	TFactor = bestTFactor;
	
	finish = clock();

	Msg("Time taken = %d clock ticks",finish - start);
	Msg("Centre = (%d,%d) TFactor = %f and with best sum = %f",ImageWidth/2 - bestx,ImageHeight/2 - besty,(bestTFactor - 1) * 40,best);

	CreateMask();
}


void LensTransform::CalculateOldFishEyeScale(void){

	double dScale = (double)(4 * ImageHeight) / (3 * ImageWidth);
	int iPicWidth = (int)(ImageWidth * dScale + 0.5);
	double iMaxRadius = (double)(int)sqrt((double) (ImageHeight * ImageHeight) + (double) (iPicWidth * iPicWidth)) / 2 - 1; // PCN2433
	double iDivider = iMaxRadius + 10;
	double dCalResult = PI / iDivider;
	double dMidCal = atan(iMaxRadius / iDivider);

	double dMinX,dMinY,dMaxX,dMaxY;
	double dIncrement,dMaxVal,dTheta;
	double dI,dJ,dI2,dJ2,dX2,dY2;
	double dPropx,dPropy;


	dMinX = 99999;
	dMinY = 99999;
	dMaxX = -99999;
	dMaxY = -99999;

	dIncrement = PI / 2;
	dMaxVal = PI * 2;

	for ( dTheta = dIncrement ; dTheta <= dMaxVal; dTheta = dTheta + dIncrement )
	{
		dI = sin(dTheta);// * 2.0 / dTransParam4X; PCN2421
		dJ = cos(dTheta);// * 2.0 / dTransParam4Y; PCN2421
		
		dI2 = dI * tan(dMidCal * TFactor) / dCalResult;
		dJ2 = dJ * tan(dMidCal * TFactor) / dCalResult;

		dX2 = dI2 + (iPicWidth/2) + (ImageWidth/2 - xCentre);
		dY2 = dJ2 + yCentre;
		
		if (dX2 < dMinX)
			dMinX = dX2;
		if (dX2 > dMaxX)
			dMaxX = dX2;
		if (dY2 < dMinY)
			dMinY = dY2;
		if (dY2 > dMaxY)
			dMaxY = dY2;
	}

	dPropx = (double)iPicWidth / (dMaxX - dMinX) * sqrt((double) 2);// * (2.0 / dTransParam4X); PCN2421 PCN2433(sqrt(2))
	dPropy = (double)ImageHeight / (dMaxY - dMinY) * sqrt((double) 2);// * (2.0 / dTransParam4Y); PCN2421 PCN2433(sqrt(2))


	dPropx *= OldFishEyeLookupTable(TFactor);
	dPropy *= OldFishEyeLookupTable(TFactor);

	if (dPropx < dPropy) // PCN2433 > ==> <
		dPropy = dPropx;
	else
		dPropx = dPropy;
	Scale = dPropx;
	Msg("Scale = %f",Scale);
}

double LensTransform::OldFishEyeLookupTable(double Factor){

	if(Factor <= 1.05) return 1.21;
    if(Factor <= 1.075) return 1.21;
    if(Factor <= 1.10) return 1.22;
    if(Factor <= 1.125) return 1.23;
    if(Factor <= 1.15) return 1.24;
    if(Factor <= 1.175) return 1.26;
    if(Factor <= 1.2) return 1.28;
    if(Factor <= 1.225) return 1.29;
    if(Factor <= 1.25) return 1.31;
    if(Factor <= 1.275) return 1.34;
    if(Factor <= 1.3) return 1.36;
    if(Factor <= 1.325) return 1.38;
    if(Factor <= 1.35) return 1.4;
    if(Factor <= 1.375) return 1.4;
    if(Factor <= 1.4) return 1.42;
    if(Factor <= 1.425) return 1.44;
    if(Factor <= 1.45) return 1.47;
    if(Factor <= 1.475) return 1.5;
    if(Factor <= 1.5) return 1.54;
    if(Factor <= 1.525) return 1.57;
    if(Factor <= 1.55) return 1.62;
    if(Factor <= 1.575) return 1.66;
    if(Factor <= 1.6) return 1.72;
    if(Factor <= 1.625) return 1.78;
    if(Factor <= 1.65) return 1.84;
	return 0;
}