// SonarProfiler.cpp : Defines the entry point for the DLL application.
//
#include <stdio.h>
#include <windows.h>
#include "..\houghlibv2.0\cbsalgebra.h"
#include "SonarProfiler.h"

#include <math.h>




double VERSION = 1.1; // Different centre type selectable. (Edge, water or normal)

void __stdcall sonar_getversion(double *ver)
{
	*ver=VERSION;
}

void __stdcall InitialiseDLL(){

	pos = 0;
	bytesPerPixel = 3;
	homeAngle = 0 - PI/2;
	homeAngle = 0;
	scanAngle = homeAngle;
	lowPass = 110;
	highPass = 245;
	slope = (float)0.7;
	halfWay = 127;
	syncCount = 0;
	cableOffset = 0;
	scale = 1.0;
	//scale = 2.55;
	xCentre = 0;
	yCentre = 0;

}

int __stdcall StartSonarSweep(int width, int height, int preview, int arcSize, int centreAngle, double sampleRate, int overSamples, int samples, double sSize, int pSize){
//int __stdcall InitialiseDLL(int width, int height, int preview){

	imgW = width;
	imgH = height;
	previewImage = preview;

	
	if(overSamples == 2){
		copyWidth = 60;
	} else if (overSamples == 4) {
		copyWidth = 60;
	} else if (overSamples == 8) {
		copyWidth = 21;
	} else {
		copyWidth = 10;
	}

	drawLogo = true;

	logoOffsetX = (imgW/2) - (logoWidth/2);
	logoOffsetY = (imgH/2) - (logoHeight/2);

// initialise the m_Config class

	SAMPLES = samples;
	STEPS = 400;
	MAXDATASIZE = STEPS * SAMPLES;

	CConfiguration *CFig;
	CConfiguration localCFG;

	localCFG.m_Oversamples = overSamples;


	switch(arcSize){
		case 30:	as = eum30Degrees; break;
		case 60:	as = eum60Degrees; break;
		case 90:	as = eum90Degrees; break;
		case 120:	as = eum120Degrees; break;
		case 150:	as = eum150Degrees; break;
		case 180:	as = eum180Degrees; break;
		case 210:	as = eum210Degrees; break;
		case 240:	as = eum240Degrees; break;
		case 270:	as = eum270Degrees; break;
		case 360:	as = eum360Degrees; break;
	}
	//m_Config.m_ArcSize = as;
	
	switch(centreAngle){
		case 30:	ca = eumCentre30; break;
		case 60:	ca = eumCentre60; break;
		case 90:	ca = eumCentre90; break;
		case 120:	ca = eumCentre120; break;
		case 150:	ca = eumCentre150; break;
		case 180:	ca = eumCentre180; break;
		case 210:	ca = eumCentre210; break;
		case 240:	ca = eumCentre240; break;
		case 270:	ca = eumCentre270; break;
		case 300:	ca = eumCentre300; break;
		case 330:	ca = eumCentre330; break;
	}
	//m_Config.m_CentreAngle = ca;
	//m_Config.m_SampleRate = (float)sampleRate;

	
	if(sSize == 0.9){
		ss = eum09Degree;
	} else if(sSize == 1.8){
		ss = eum18Degree;
	} else if(sSize == 2.7){
		ss = eum27Degree;
	} else if(sSize == 3.6){
		ss = eum36Degree;
	}


	// Setup Initial Sonar Configuration
	// And Send the same configuration details to the outline processor 

	//if(usbComm->SetConfiguration(m_Config) == false)	return 0;
	//usbComm->SetConfiguration(m_Config);
	//outline->SetConfiguration(m_Config);

	// Create the 1512USB Interface
	usbComm = Sonar::Create1512USBCommObject();
	// Create Outline Processor Interface
	outline = CreateOutlineProcessorObject();
	// Initialise communication between PC and 1512 USB Box
	usbComm->InitialiseComms();

	usbComm->SetConfiguration(localCFG);
	outline->SetConfiguration(localCFG);


	//usbComm->SetConfiguration(m_Config);
	//outline->SetConfiguration(m_Config);

	CFig = usbComm->GetConfiguration();

	// Setup Shaft Encoder
	usbComm->SetShaftEncoder(SetQuadEncoder);



	// Setup Shaft Encoder
	//if(usbComm->SetShaftEncoder(SetQuadEncoder) == false)	return 0;
	//usbComm->SetShaftEncoder(SetQuadEncoder);

	usbComm->SetArcSize(as);
	usbComm->SetCentreAngle(ca);
	usbComm->SetSampleRate((float)sampleRate);
	usbComm->SetOversamples(overSamples);
	usbComm->SetSamples(samples);

	//if(usbComm->SetArcSize(as) == false)	return 0;
	//if(usbComm->SetCentreAngle(ca) == false)	return 0;
	//if(usbComm->SetSampleRate((float)sampleRate) == false)	return 0;
	//if(usbComm->SetOversamples(overSamples) == false)	return 0;
	//if(usbComm->SetSamples(samples) == false)	return 0;
	//if(usbComm->SetStepSize(ss) == false){
	//	return 0;

	//}
	//if(usbComm->SetPulseWidth(pSize) == false)	return 0;

	usbComm->SetStepSize(ss);
	usbComm->SetPulseWidth(pSize);


	frameCount = 0;

	start = 0;

	CFig = usbComm->GetConfiguration();

	//for(int i = 0; i < 1000000; i++) ProfileData[i] = Pallette[188];

	// Register Callback function
	usbComm->RegisterCallback(OnSonarUpdate);

	CFig = usbComm->GetConfiguration();

	return 1;

}


/*
int __stdcall InitialiseDLL(int width, int height, int preview, int arcSize, int centreAngle, double sampleRate, int overSamples, int samples, double sSize, int pSize ){
//int __stdcall InitialiseDLL(int width, int height, int preview){

	pos = 0;
	bytesPerPixel = 3;
	scanAngle = 0;
	imgW = width;
	imgH = height;
	previewImage = preview;
	lowPass = 110;
	highPass = 245;
	slope = (float)0.7;
	halfWay = 127;


// initialise the m_Config class

	SAMPLES = samples;
	STEPS = 400;
	MAXDATASIZE = STEPS * SAMPLES;

	EnumArcSize as;
	switch(arcSize){
		case 30:	as = eum30Degrees; break;
		case 60:	as = eum60Degrees; break;
		case 90:	as = eum90Degrees; break;
		case 120:	as = eum120Degrees; break;
		case 150:	as = eum150Degrees; break;
		case 180:	as = eum180Degrees; break;
		case 210:	as = eum210Degrees; break;
		case 240:	as = eum240Degrees; break;
		case 270:	as = eum270Degrees; break;
		case 360:	as = eum360Degrees; break;
	}
	//m_Config.m_ArcSize = as;
	EnumCentreAngle ca;
	switch(centreAngle){
		case 30:	ca = eumCentre30; break;
		case 60:	ca = eumCentre60; break;
		case 90:	ca = eumCentre90; break;
		case 120:	ca = eumCentre120; break;
		case 150:	ca = eumCentre150; break;
		case 180:	ca = eumCentre180; break;
		case 210:	ca = eumCentre210; break;
		case 240:	ca = eumCentre240; break;
		case 270:	ca = eumCentre270; break;
		case 300:	ca = eumCentre300; break;
		case 330:	ca = eumCentre330; break;
	}
	//m_Config.m_CentreAngle = ca;
	//m_Config.m_SampleRate = (float)sampleRate;

	EnumStepSize ss;
	if(sSize == 0.9){
		ss = eum09Degree;
	} else if(sSize == 1.8){
		ss = eum18Degree;
	} else if(sSize == 2.7){
		ss = eum27Degree;
	} else if(sSize == 3.6){
		ss = eum36Degree;
	}


	// Create the 1512USB Interface
	usbComm = Sonar::Create1512USBCommObject();


	// Create Outline Processor Interface
	outline = CreateOutlineProcessorObject();


	// Initialise communication between PC and 1512 USB Box
	if(usbComm->InitialiseComms() == false)		return 0;
	//usbComm->InitialiseComms();

	// Setup Initial Sonar Configuration
	// And Send the same configuration details to the outline processor 



	//if(usbComm->SetConfiguration(m_Config) == false)	return 0;
	usbComm->SetConfiguration(m_Config);
	outline->SetConfiguration(m_Config);


	// Setup Shaft Encoder
	if(usbComm->SetShaftEncoder(SetQuadEncoder) == false)	return 0;
	//usbComm->SetShaftEncoder(SetQuadEncoder);

	if(usbComm->SetCentreAngle(ca) == false)	return 0;
	if(usbComm->SetSampleRate((float)sampleRate) == false)	return 0;
	if(usbComm->SetOversamples(overSamples) == false)	return 0;
	if(usbComm->SetSamples(samples) == false)	return 0;
	//if(usbComm->SetStepSize(ss) == false){
	//	return 0;
	//}
	if(usbComm->SetPulseWidth(pSize) == false)	return 0;


	frameCount = 0;

	start = 0;

	//for(int i = 0; i < 1000000; i++) ProfileData[i] = Pallette[188];

	// Register Callback function
	usbComm->RegisterCallback(OnSonarUpdate);

	return 1;

}

*/

// Callback function
void OnSonarUpdate(void)
{
	
	if (usbComm->IsScanning()){


		//CConfiguration *CFig;
		//CFig = usbComm->GetConfiguration();

		// Retrieve data from sonar
		usbComm->GetScanData(ProfileData);

		// Process outline data
		outline->ProcessOutline(ProfileData);

		GetSystemTime(&currentSystemTime);

		SaveProfileToFile();
		if(frameCount <= CALIBRATE_SWEEPS){
			calibrateBuffer[frameCount].wHour = currentSystemTime.wHour;
			calibrateBuffer[frameCount].wMinute = currentSystemTime.wMinute;
			calibrateBuffer[frameCount].wSecond = currentSystemTime.wSecond;
			calibrateBuffer[frameCount].wMilliseconds = currentSystemTime.wMilliseconds;
		}
		frameCount ++;

		if(syncCount > 0){
			syncCount--;
			if(syncCount < 1) frameCount = 1;
		}

	}
}

void drawAverageRay(double angle,int i,unsigned char *Img)
{
	float rayPosition;
	int dataIndex;
	int colour;
	float xc = (imgW/2)+ (float) xCentre;
	float yc = (imgH/2)+ (float) yCentre;
	int pointX, pointY;
	int pointSize;


	angle+=homeAngle;

	for(rayPosition = 0;rayPosition<301;rayPosition++)
	{
		pointSize = (int) ((rayPosition /50)+2);

		dataIndex=(int) rayPosition+i;
		colour = ProfileData[dataIndex]*3;

		if((colour<lowPass) || (colour>highPass)) continue;
		pointX = (int) (sin(angle)*rayPosition*scale);
		pointY = (int) (cos(angle)*rayPosition*scale);
		pointX+=(int) xc;
		pointY+=(int) yc;
		fillCircle(pointX,pointY, pointSize,Img,colour);
//		fillCircle(pointX,pointY-1, pointSize,Img,colour);
//		fillCircle(pointX,pointY+1, pointSize,Img,colour);
//		fillCircle(pointX-1,pointY, pointSize,Img,colour);
//		fillCircle(pointX+1,pointY, pointSize,Img,colour);
		
		
	}
}



void __stdcall drawframe(unsigned char *Img){
//	int i;
	double angle;
	bytesPerPixel = 3;
	//angle = PI / 2;
	angle = homeAngle;

	maxSample = 0;
	maxRad = 0;
	previousMaxRad = 0;

	startX = 0;
	startX = 0;
	filterData();
	drawIMGframe((float) imgW/2,(float) imgH/2,Img);


	
//	for(i = 0; i < (120400); i+= 301){
//		
//
//		drawAverageRay(angle,i,Img);
//
//		//drawray(angle,i,i + 301,Img);
//		//drawrayOLD(angle,i,i+301,Img);
//		
//
//		//angle = angle + 0.45;
//		angle += ((PI*2) / 400);
//	}

//ANT
//	VBRay(Img);

}



void __stdcall VBRay(unsigned char *Img){
	int i;
	

	for(i=0;i<20;i++){
		drawrayOLD(scanAngle,pos,pos + SAMPLES,Img);
		//for(t=0;t<6;t++){
		//	drawray(scanAngle + t *(((PI*2)/STEPS) / 20) ,pos,pos + SAMPLES,Img);
		//}

		scanAngle += (PI*2)/STEPS;
		pos += SAMPLES;
		if(pos >= MAXDATASIZE) {
			drawLogo = drawLogo == true ? false : true;
			pos = 0;
			scanAngle = homeAngle;
		}
		if(scanAngle > (homeAngle + (PI*2))){
			scanAngle = homeAngle;
			pos = 0;
		}
	}

	

}


void drawArcOLD(int radius, double ang1, double ang2,unsigned char *Img, BYTE col){
	double x1, y1, x2, y2;
	

	x1 = (imgW/2) + (radius * cos(ang1));
	y1 = (imgH/2) + (radius * sin(ang1));
		
	x2 = (imgW/2) + (radius * cos(ang2));
	y2 = (imgH/2) + (radius * sin(ang2));
			

	int xSize = abs((int)x2 - (int)x1);
	int ySize = abs((int)y2 - (int)y1);
	double step;

//	if((ang1 >= PI/4 && ang1 < (3*PI/4)) || (ang1 >= (PI + (3*PI/4)) && ang1 < (PI + PI/4))){
	if(xSize > ySize){
		step = (fabs(ang2 - ang1) / xSize) / 2;
		for(int i = 0; i < xSize * 2; i++){
			setRadarColour(Img,col,(int)x1,(int)y1,radius,false);
			//setPixel((int)x1,(int)y1,Img,col);
			x1 = (imgW/2) + (radius * cos(ang1));
			y1 = (imgH/2) + (radius * sin(ang1));
			ang1 += step;
		}
		
	} 
	if(ySize > 0){
		step = (fabs(ang2 - ang1) / ySize) /2;
		for(int i = 0; i < ySize * 2; i++){
			setRadarColour(Img,col,(int)x1,(int)y1,radius,false);
			//setPixel((int)x1,(int)y1,Img,col);
			x1 = (imgW/2) + (radius * cos(ang1));
			y1 = (imgH/2) + (radius * sin(ang1));
			ang1 += step;
		}

	}

}

void drawrayOLD(double ang, int start, int finish,unsigned char *Img){
	int sample;
	int rad = 1;

	for(sample = start; sample < finish; sample ++){
		
		drawArcOLD(rad,ang,ang+(PI*2)/STEPS,Img,ProfileData[sample]);
		rad++;


		//x2 = x1 + (1 * cos(ang));
		//y2 = y1 + (1 * sin(ang));
		//drawlineBMP((int)x1,(int)y1,(int)x2,(int)y2,ProfileData[sample],Img);
		//x1 = x2;
		//y1 = y2;
	}

	
}


void drawray(double ang, int start, int finish,unsigned char *Img){
	int sample;
	double rad = 1.0;
	int x1,y1,x2,y2;

	for(sample = start; sample < finish; sample ++){

		//if(rad > 120.0)	drawArc(rad,ang,ang+(PI*2)/400,Img,ProfileData[sample],15,186);


		
		if(ProfileData[sample] > 50){ 
			if(sample > maxSample){
				
				maxSample = sample;
				maxRad = rad;
				previousMaxRad = maxRad;
			}
		}
		rad+= scale;

		//x2 = x1 + (1 * cos(ang));
		//y2 = y1 + (1 * sin(ang));
		//drawlineBMP((int)x1,(int)y1,(int)x2,(int)y2,ProfileData[sample],Img);
		//x1 = x2;
		//y1 = y2;
	}

	if(maxRad < 150.0/scale) return;

	x1 = (int)( (imgW/2 + xCentre) + (maxRad * cos(ang)) );
	y1 = (int)( (imgH/2 - yCentre)+ (maxRad * sin(ang)) );

	x2 = (int)( (imgW/2 + xCentre) + (maxRad * cos(ang+(PI*2)/400)) );
	y2 = (int)( (imgH/2 - yCentre)+ (maxRad * sin(ang+(PI*2)/400)) );



	if(startX != 0 && startY != 0) {
		drawlineBMP(startX,startY,x1,y1,(ProfileData[maxSample] < 155 ? ProfileData[maxSample] + 100 : 255),Img);
		//drawlineBMP(currentX,currentY,x2,y2,(ProfileData[maxSample] < 155 ? ProfileData[maxSample] + 100 : 255),Img);
		//drawlineBMP(currentX,currentY,x2,y2,ProfileData[maxSample],Img);
	}

	//drawArc(maxRad,ang,ang+(PI*2)/400,Img,ProfileData[maxSample],15,186);
	//drawArc(maxRad,ang,ang+(PI*2)/400,Img,(ProfileData[maxSample] < 155 ? ProfileData[maxSample] + 100 : 255),15,186);
	drawlineBMP(x1,y1,x2,y2,(ProfileData[maxSample] < 155 ? ProfileData[maxSample] + 100 : 255),Img);

	startX = x2;
	startY = y2;	
}


void __stdcall drawlineBMP(int x1, int y1,int x2, int y2, BYTE col, unsigned char *Img){

	double angle = fabs((double)(y2 - y1) / (double)(x2 - x1));	// angle 
	int a = abs(x2 - x1);		// adjacent
	int o = abs(y2 - y1);		// opposite
	int lx = x1 > x2 ? x2 : x1;
	int ly = y1 > y2 ? y2 : y1;
	double v;

	int len = (int)sqrt(a * a + o * o);
	if(len > 100) return;

	int pos;

	int i;

	if(angle > 0.5){
		angle = (1.0 / angle);


		//ly = (x1 <= x2) ? y1 : y2;
		
		
		v = (double)lx;

		for(i=0; i<o; i ++){
			//pos = ((ly + i) * (imgW * bytesPerPixel)) + ((int)v * bytesPerPixel);
			//pos = ((((y1 <= y2 && x1 <= x2) ? ly + i : ly - i)) * (imgW * bytesPerPixel)) + ((int)v * bytesPerPixel);

			if(ly == y2 && lx == x1) pos = ((y1 - i) * (imgW * bytesPerPixel)) + ((int)v * bytesPerPixel);
			if(ly == y2 && lx == x2) pos = ((y2 + i) * (imgW * bytesPerPixel)) + ((int)v * bytesPerPixel);
			if(ly == y1 && lx == x1) pos = ((y1 + i) * (imgW * bytesPerPixel)) + ((int)v * bytesPerPixel);
			if(ly == y1 && lx == x2) pos = ((y2 - i) * (imgW * bytesPerPixel)) + ((int)v * bytesPerPixel);

			if(pos >=0 && pos < (imgW * bytesPerPixel) * imgH){
				if(bytesPerPixel == 3){
					DrawSquare(Img,pos + 1,col);
					//DrawSquare(Img,pos,255);

				} else {
					Img[pos] = (unsigned char)col;
				}
			}
			v += angle;
			//v -= angle;
		}
	} else {

		
		ly = (x1 < x2) ? y1 : y2;
		if(y2 < y1) angle = 0 - angle;

		v = (double)ly;
		for(i=0; i<a; i++){
			pos = ((int)v * (imgW * bytesPerPixel)) + ((lx + i) * bytesPerPixel);

			if(pos >=0 && pos < (imgW * bytesPerPixel) * imgH){
				if(bytesPerPixel == 3){
					DrawSquare(Img,pos + 1,col);
					//DrawSquare(Img,pos,col);
				} else {
					Img[pos] = (unsigned char)col;
				}
			}
			v += angle;
		}
	}
}

void __stdcall ReadProfileFromFile(unsigned char *Img, int bytes, char *filename, int W, int H, int *Hours, int *Mins, int *mSec, int *Dist)
{


	FILE *input;
	short temp;

	bytesPerPixel = bytes;
	Image = Img;
	input = fopen(filename,"rb");

	imgW = W;
	imgH = H;
//	lowPass = 50;
//	highPass = 255;

	int i;
	for(i=0;i<((imgW) * bytes) * (imgH);i++) Img[i] = 0;


	//fillCircle(imgH/2,imgW/2 ,480,Img,0);
	//fillCircle(imgH/2,imgW/2 ,imgW,Img,0);


	if(input != NULL) {

		//fread(ProfileData,1,MAXDATASIZE,input);
		fread(ProfileData,1,120400,input);

		fread(&temp,1,2,input); //dummy?
		fread(&temp,1,2,input); //dummy?

		fread(Hours,1,2,input); //hour
		fread(Mins,1,2,input); //minute
		fread(&temp,1,2,input); //second
		*mSec = ((int)temp * 1000);
		fread(&temp,1,2,input); //millisecond
		*mSec = *mSec + (int)temp;

		fread(&temp,1,2,input);
		*Dist = (int)temp;

		fclose(input);

		drawframe(Img);
	}


}

void SaveProfileToFile(void)
{

	FILE *output;
	char filename[300];
	clock_t temp;
	CSensorData *SData;

	sprintf(filename,"%s%d.3S3",outputFileName,frameCount);

	output = fopen(filename,"wb");

	if( start == 0 ){
		fwrite(&start,4,1,output);
		//fprintf(output,"%d",0);
		start = clock();
	} else {
		temp = clock() - start;
		fwrite(&temp,4,1,output);
		//fprintf(output,"%d",temp - start);
	}

	SData = usbComm->GetSensorData();
	//system time
	if(output != NULL) {

		fwrite(ProfileData,1,MAXDATASIZE,output);
		//write the time stamp at the end of the frame
		fwrite(&currentSystemTime.wHour,1,sizeof(WORD),output);
		fwrite(&currentSystemTime.wMinute,1,sizeof(WORD),output);
		fwrite(&currentSystemTime.wSecond,1,sizeof(WORD),output);
		fwrite(&currentSystemTime.wMilliseconds,1,sizeof(WORD),output);
		fwrite(&SData->m_CablePayout,1,sizeof(WORD),output); // distance

	}
	fclose(output);
}

void __stdcall SetMaxAndMin(int Max, int Min){

	if(Min < 0 || Min > 200 || Max > 255 || Max < 50) return;
	highPass = Max;
	lowPass = Min;

}


void __stdcall StopScanning(){
	if(usbComm != NULL){
		usbComm->StopScan();
	}
}

void __stdcall StartScanning(char *directory){
	if(usbComm != NULL){
		strcpy(outputFileName,directory);
		frameCount = 1;
		usbComm->StartScan();
	}
}


void __stdcall CheckForScanning(int *status){
	if(usbComm != NULL){
		if (usbComm->IsScanning()){
			if (frameCount <= CALIBRATE_SWEEPS){
				*status = 2;
			} else {
				*status = 1;
			}
		} else {
			*status = 0;
		}

	} else {
		*status = 0;
	}
}



void __stdcall SynchronisedStart(int *secs, int *millisecs, _SYSTEMTIME *startTime, int countDown){
	int i, total, status = 0;

	for(i=1; i<=CALIBRATE_SWEEPS; i++){
		status += (calibrateBuffer[i].wSecond * 1000) + calibrateBuffer[i].wMilliseconds;
	}
	status = (int) ((double)status / (double)CALIBRATE_SWEEPS);

	total = status * 3;
	*secs = (int)(total / 1000);
	*millisecs = total - ((*secs) * 1000);

	startTime->wDay = currentSystemTime.wDay;
	startTime->wDayOfWeek = currentSystemTime.wDayOfWeek;
	startTime->wHour = currentSystemTime.wHour;
	startTime->wMilliseconds = currentSystemTime.wMilliseconds;
	startTime->wMinute = currentSystemTime.wMinute;
	startTime->wMonth = currentSystemTime.wMonth;
	startTime->wSecond = currentSystemTime.wSecond;
	startTime->wYear = currentSystemTime.wYear;

	syncCount = countDown;
}



void setPixel(int x, int y, unsigned char *Img,BYTE col){
	if(x>=0 && y>=0 && x<imgW && y<imgH){
		Img[y*imgW*3 + (x*3)]   = col;
		Img[y*imgW*3 + (x*3)+1] = col;
		Img[y*imgW*3 + (x*3)+2] = col;
	}
}

void drawArc(double radius, double ang1, double ang2,unsigned char *Img, BYTE col, int XOffset, int YOffset){
	double x1, y1, x2, y2;
	

	x1 = (imgW/2 + XOffset) + (radius * cos(ang1));
	y1 = (imgH/2 - YOffset) + (radius * sin(ang1));
		
	x2 = (imgW/2 + XOffset) + (radius * cos(ang2));
	y2 = (imgH/2 - YOffset) + (radius * sin(ang2));
			

	int xSize = abs((int)x2 - (int)x1);
	int ySize = abs((int)y2 - (int)y1);
	double step;

//	if((ang1 >= PI/4 && ang1 < (3*PI/4)) || (ang1 >= (PI + (3*PI/4)) && ang1 < (PI + PI/4))){
	if(xSize > ySize){
		step = (fabs(ang2 - ang1) / xSize) / 2;
		for(int i = 0; i < xSize * 2; i++){
			setRadarColour(Img,col,(int)x1,(int)y1,radius,true);
			//setPixel((int)x1,(int)y1,Img,col);
			x1 = (imgW/2 + XOffset) + (radius * cos(ang1));
			y1 = (imgH/2 - YOffset) + (radius * sin(ang1));
			ang1 += step;
		}
		
	} 
	if(ySize > 0){
		step = (fabs(ang2 - ang1) / ySize) /2;
		for(int i = 0; i < ySize * 2; i++){
			setRadarColour(Img,col,(int)x1,(int)y1,radius,true);
			//setPixel((int)x1,(int)y1,Img,col);
			x1 = (imgW/2 + XOffset) + (radius * cos(ang1));
			y1 = (imgH/2 - YOffset) + (radius * sin(ang1));
			ang1 += step;
		}

	}

}

void fillCircle(int xCentre, int yCentre, int width, unsigned char *Img, BYTE col){
	int i,t,x,y,pos;
	double dist;
	double boundaryCheck = (imgW-1)*3*imgH;
	
	pos = (yCentre*imgW*3)+(xCentre*3);
	

	for(i = xCentre - (width / 2); i < xCentre + (width / 2);i++){
		for(t = yCentre - (width / 2); t < yCentre + (width / 2);t++){
			
			x = abs(xCentre - i);
			y = abs(yCentre - t);

			dist = sqrt(x * x + y * y);
			if((int)dist < (width/2)){
				pos = i*imgW*3 + (t*3);
				if((pos<0) || (pos > boundaryCheck)) return;

				Img[pos]   = (unsigned char) (col*0.8);
				Img[pos+1] = (unsigned char) col;
				Img[pos+2] = (unsigned char) (col*0.8);
			}
		}
	}
}




void setRadarColour(unsigned char *Img, BYTE col,int x, int y, double radius, bool large){

	if((int)radius < copyWidth){
		int pos = y*imgW*3 + (x*3);
		int lX,lY;

		lX = x - logoOffsetX;
		lY = y - logoOffsetY;

		int lPos = (lY*(logoWidth)*3 + (lX *3)) +1; //((lY*(logoWidth+1)*3) + (lX*3));

		if(drawLogo == true){
			Img[pos]   = 15;
			Img[pos+1] = 200;
			Img[pos+2] = 90;
		} else {
			//Img[pos]   = 0;
			//Img[pos+1] = 0;
			//Img[pos+2] = 0;
			Img[pos]   = LogoData[lPos];
			Img[pos+1] = LogoData[lPos+1];
			Img[pos+2] = LogoData[lPos+2];
		}
		return;
	}

	if(x>=0 && y>=0 && x<imgW && y<imgH){

		int pos = y*imgW*3 + (x*3);

		//if(col < lowPass) return;
		//col = col + 100;
		//if(col > 255) col = 255;

		//Img[pos]   = col;
		//Img[pos+1] = col;
		//Img[pos+2] = col;



		if(large == true) {
			DrawSquare(Img,pos,0);
			DrawSquare(Img,pos+1,col);
			DrawSquare(Img,pos+2,0);
			return;
		} 



		if(col < lowPass) {
			Img[pos]   = Pallette[0];
			Img[pos+1] = Pallette[1];
			Img[pos+2] = Pallette[2];
			return;
		}
		
		if(col > highPass) {
			Img[pos]   = Pallette[765];
			Img[pos+1] = Pallette[766];
			Img[pos+2] = Pallette[767];
			return;
		}
		
		float ratio = ((float)(col - lowPass) / (float)(highPass - lowPass));
		int pal = lowPass + (int)(ratio * 256.0);
		if(pos>=1) Img[pos-1]   = Pallette[pal]; //ANT
		Img[pos+1] = Pallette[pal+1];
		Img[pos] = Pallette[pal+2];
	
	}
}


void __stdcall LoadPallette(unsigned char *Img){
	for(int i=0;i<768;i++) Pallette[i] = Img[i];
}


void __stdcall SetScanAngle(double Angle){

	homeAngle = (Angle * (PI/180));
	scanAngle = homeAngle;

}


void __stdcall LoadLogo(unsigned char *Img,int w, int h){
			
	logoWidth = w;
	logoHeight = h;
	//logoOffsetX = (imgW/2) - (logoWidth/2);
	//logoOffsetY = (imgH/2) - (logoHeight/2);

	for(int i = 0;i < w*h*3; i++){
		LogoData[i+1] = Img[i];
	}

}

void __stdcall DrawCircle(unsigned char *Img, int width, int height, double radius){

	if(imgW == 0 || imgH ==0){
		imgW = width;
		imgH = height;
	}

	drawArc(radius,0,PI,Img,255,0,0);
	drawArc(radius,PI,PI*2,Img,255,0,0);

}

void __stdcall SetCablePayoutStart(int cable){
	CSensorData *SData;

	SData = usbComm->GetSensorData();
	cableOffset = SData->m_CablePayout - cable;
	
}


void __stdcall GetCablePayout(int *cable){
	CSensorData *SData;

	SData = usbComm->GetSensorData();
	*cable = SData->m_CablePayout - cableOffset;

}


void __stdcall SetScale(double scaleVal){
	scale = scaleVal;	
}

void __stdcall SetSonarCentre(int x, int y){
	xCentre = x;
	yCentre = y;
}



/*
void setRadarColour(unsigned char *Img, BYTE col,int x, int y){
	if(x>=0 && y>=0 && x<imgW && y<imgH){

		int pos = y*imgW*3 + (x*3);
		if(col < lowPass) {
			Img[pos]   = 0;
			Img[pos+1] = 0;
			Img[pos+2] = 0;
			return;
		} 
		if(col > highPass) {
			Img[pos]   = 255;
			Img[pos+1] = 0;
			Img[pos+2] = 0;
			return;
		}

		if(col > halfWay){
			Img[pos] = (unsigned char)(((float)(col - halfWay) * slope * 2));
		} else {
			Img[pos] = 0; 
		}

		if(col < halfWay){
			Img[pos+2] = (unsigned char)(255 - ((float)col * slope * 2));
		} else {
			Img[pos + 2] = 0;
		}


		if(col == halfWay) {
			Img[pos+1] = 255;
		} else if(col < halfWay){
			Img[pos+1] = (unsigned char)((float)col * slope);
		} else {
			Img[pos+1] = (unsigned char)(255 - (float)col * slope);
		}
	}
}
*/

void DrawSquare(unsigned char *Img,int pos, BYTE col){

	if(pos > 6 && pos < (imgW*3*imgH)){
		Img[pos-6] = col;
		Img[pos-3] = col;
		Img[pos]   = col;
		Img[pos+3] = col;
		Img[pos+6] = col;
	}

	if(pos > (imgW*6)){
		//Img[pos-(imgW*6)-6] = col;
		Img[pos-(imgW*6)-3] = col;
		Img[pos-(imgW*6)]   = col;
		Img[pos-(imgW*6)+3] = col;
		//Img[pos-(imgW*6)+6] = col;
	}
	
	if(pos > (imgW*3) + 6){
		Img[pos-(imgW*3)-6] = col;
		Img[pos-(imgW*3)-3] = col;
		Img[pos-(imgW*3)]   = col;
		Img[pos-(imgW*3)+3] = col;
		Img[pos-(imgW*3)+6] = col;
	}

	if(pos < ((imgW*3)*imgH) - (imgW*3)){
		Img[pos+(imgW*3)-6] = col;
		Img[pos+(imgW*3)-3] = col;
		Img[pos+(imgW*3)]   = col;
		Img[pos+(imgW*3)+3] = col;
		Img[pos+(imgW*3)+6] = col;
	}

	if(pos < ((imgW*3)*imgH) - (imgW*6)){
		//Img[pos+(imgW*6)-6] = col;
		Img[pos+(imgW*6)-3] = col;
		Img[pos+(imgW*6)]   = col;
		Img[pos+(imgW*6)+3] = col;
		//Img[pos+(imgW*6)+6] = col;
	}
}


/*

 void ClearImage(unsigned char * Img){
	for(int i=0;i<imgH*imgW*3;i++) Img[i]=0;
}


*/

void drawIMGframe(float centreX, float centreY, unsigned char *Img)
{


	vec2double centre;

	int rayNo, rayPos;
	char colour;
	int rayIndex;

	int x,y;

	centre.x = centreX; centre.y=centreY;

	for(y=0;y<imgH;y++)
		for(x=0;x<imgW;x++)
		{
			getRayAddress(rayNo,rayPos,(int) (x-centreX), (int) (y-centreY));
			if((rayPos>=301) || (rayPos<0)) continue;
			if((rayNo<0) || (rayNo>=400)) continue;
			rayIndex = (301*rayNo)+rayPos;
			colour = ProfileData[rayIndex];
			setPixel(x,y,Img,(char) ((double) colour*1.5));
		}

}

void getRayAddress(int &rayNo, int &rayPos, int x, int y)
{
	int sonarStep = 400;
	int sonarSamples = 301;
	double stepSize;
	double sampleSize;
	
	vec2double point;

	stepSize = (2*PI)/sonarStep;
	sampleSize = 301/sonarSamples;


	point.x=x; point.y=y;
	point = point.toVector();
	point.y*=sampleSize;

	rayNo = (int) (point.x/stepSize);
	rayPos = (int) point.y;
	
}

void filterData(void)
{

	int rayNo;
	int raySample;
	int rayIndex;
	int leftDiff;
	int rightDiff;


	for(rayNo=0;rayNo<400;rayNo++)
		for(raySample=1;raySample<299;raySample++)
		{
			rayIndex = (301*rayNo)+raySample;
			leftDiff = abs(ProfileData[rayIndex]-ProfileData[rayIndex-1]);
			rightDiff = abs(ProfileData[rayIndex]-ProfileData[rayIndex+1]);
			
			if((leftDiff>35) && (rightDiff>35)) ProfileData[rayIndex]=0;
		}

}


