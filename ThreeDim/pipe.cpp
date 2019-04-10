#include <windows.h>
#include <stdlib.h>
#include <atlbase.h>
#include <d3dx9.h>

#include <stdio.h>
#include <conio.h>

#include <gl/gl.h>


#include "pipe.h"
#include "CBSAlgebra.h"

#define PROFILE_SIZE 180 // Number of Points per Profile
#define SECTION_SIZE 129 // Number Frames per Pipe Section, needs to be a power of two + 1, also defined in window.cpp as PIPE_SECTION_SIZE
#define D3DFVF_CUSTOMVERTEX_PIPE  (D3DFVF_XYZ|D3DFVF_NORMAL|D3DFVF_DIFFUSE|D3DFVF_TEX2)

#define NUMBER_PATCHES_DOWN_PIPE ((frames/8)+1)
#define NUMBER_PATCHES_AROUND_PIPE 10




void MsgPipe(TCHAR *szFormat, ...)
{
    TCHAR szBuffer[512];

    va_list pArgs;
    va_start(pArgs, szFormat);
    _vstprintf(szBuffer, szFormat, pArgs);
    va_end(pArgs);
    MessageBox(NULL, szBuffer, TEXT("LaserProfiler Message"), MB_OK | MB_ICONERROR);
}

// This is the constructor of a class that has been exported.
// see viewpipe.h for the class definition
CViewpipe::CViewpipe(bool _OpenGL, //To render in OpenGL or D3D
					 float *datax, //PCN2988 pass the x coordinate data
					 float *datay, //PCN2988 pass the y coordinate data
					 float *centrex,
					 float *centrey,
					 long dataSize, //PCN2453 (loading 3d Many times Crash) 
					 LPDIRECT3DDEVICE9 pd3dDevice, char *path,
					 int *pvColourRed,
					 int *pvColourGreen,
					 int *pvColourBlue,
					 double expectedRad, //PCN2693 Needed for colour calculations (Antony van Iersel, 15 March 2004)
					 languageText *language,   //PCN2473 Language pointer (Antony, 11 May 2004)
					 int pVCalcMult, //PCN2693 (9 August 2004, Antony) needed divide by multiplier to get real values
					 int pVDataXYMult, //PCN2988 13 Sept 2004 its a devision needed for the XY data
					 int unitsPassed) //PCN3111 0 = metric, 1 = imperial
{ 
	int i;
	test1=0;

	OpenGL = _OpenGL;

	languagePipe = language; //PCN2693 Needed for colour calculations (Antony van Iersel, 15 March 2004)
	pVCalculationsMultiplier = pVCalcMult; //PCN2693 (9 August 2004, Antony) needed divide by multiplier to get real values
	pVDataXYMultiplier = pVDataXYMult; //PCN2988 add a multiplier for the 3d XY data;
	units = unitsPassed;
	//MsgPipe("Expect radius passed = %f",expectedRad);
	exportingPipe=false;
	exportingPanel=NULL; //PCN2720 (Antony van Iersel, 1 April 2004)

	colRed		=D3DCOLOR_RGBA(255,0  ,0  ,  255);	// Red
	colGreen	=D3DCOLOR_RGBA(0  ,255,0  ,  255);	// Green
	colBlue		=D3DCOLOR_RGBA(0  ,0  ,255,  255);	// Blue
	
	colOrange	=D3DCOLOR_RGBA(255,128,0  ,  255);	// Orange
	colPurple	=D3DCOLOR_RGBA(64 ,0  ,128,  255);	// Purple
	colYellow	=D3DCOLOR_RGBA(255,255,0  ,  255);	// Yellow
	colAqua		=D3DCOLOR_RGBA(100,180,180,  255);  // Aqua
	colWhite	=D3DCOLOR_RGBA(255,255,255,	 255);  // White
	colClearBlack = D3DCOLOR_RGBA(0,0,0,0);			// Clear Black
	colClearWhite = D3DCOLOR_RGBA(255,255,255,0);	// Clear White

	colDeltaBlue	= D3DCOLOR_RGBA(40 ,70 ,181,	 255); // Blue was 111
	colDeltaAqua	= D3DCOLOR_RGBA(90 ,155,204,	 255);
	colDeltaGreen	= D3DCOLOR_RGBA(181,224,238,	 255);
	colDeltaOrange	= D3DCOLOR_RGBA(255,100,  0,	 255);
	colDeltaWhite	= D3DCOLOR_RGBA(255,255,255,	 255);

	
	pTexturePipe=NULL;
	pTextureLaser=NULL;
	pTextureWhite=NULL;

	pd3dDevicePipe=NULL;
	startSection=NULL;
	endSection=NULL; 
	nodeBuffer = NULL; //PCN2465 (8 December 2003, Antony)
	preCalcSin = NULL; //PCN2465 (8 December 2003, Antony)
	preCalcCos = NULL; //PCN2465 (8 December 2003, Antony)

	pIBDetail0_1PipeSection=NULL;
	pIBDetail0_2PipeSection=NULL; 
	pIBDetail0_3PipeSection=NULL; 
	nodeBuffer=NULL;
	preCalcSin=NULL;
	preCalcCos=NULL;
	
	sectionIndex=NULL;

	rawData=NULL; 

/////////////////////////////////////////////////////////
	frames = dataSize/PROFILE_SIZE;
	noNodes = (frames*PROFILE_SIZE)-1;	// Total number Profile Points if Passed Data

	rawDatax = datax; //PCN2988
	rawDatay = datay; //PCN2988
	rawCentrex = centrex;
	rawCentrey = centrey;
	if(!OpenGL)	{rawData = new double[noNodes]; ConvertXYtoRadiusRawData();} 
	rawColourRed = pvColourRed;
	rawColourGreen = pvColourGreen;
	rawColourBlue = pvColourBlue;

	pd3dDevicePipe=pd3dDevice;
	expRad		  = expectedRad;	//this needs to be passed //expected radius
	strcpy(textureDirectory,path);

	waterLevelLeft=15;
	waterLevelRight=165;
	pipePosition=D3DXVECTOR3(0,0,0);

/////////////////////////////////////////////////////////

	shadeOn = false;	//PCN2337 , remembers the colour limits setting
	//PCN2950 change the default type from 0 to three for both shadeType and shadeTypePrevious
	shadeType     = 3;	//PCN2337 , 0 for None, 1 for Capacity, 2 for Ovality, 3 for Delta
	shadeTypePrev = 3;  //PCN2337 , Turning of the colour, when turned back on remeber what it was
	shadeChange   = 0;  //PCN2337 , keep track if it needs to be repainted
	shadeCapacityMinLimit=-1; shadeCapacityMaxLimit=1; // PCN2337 Default Limits, Makes not Different
	shadeDeltaMinLimit   =-1; shadeDeltaMaxLimit  =1;  // 
	shadeOvalityMinLimit =-1; shadeOvalityMaxLimit=1;// 	
	
	if(!OpenGL)
	{
		numberSections=(frames/(SECTION_SIZE-1))+1; // Have to add one, C rounds down.
		sectionIndex = new pipeSectionStruct *[numberSections];
		for(i=0;i<numberSections;i++) sectionIndex[i]=NULL;
		avgRad  = CalcAvgRadius();		// Average Radius of the Pipe
		if(expRad==0) expRad=avgRad;  // If no Expected Radius is passed, default to average Radius

		// Really suppose to be SECTION SIZE, but 2 added for saftey net, its really
		// A buffer + 1 on either size to calculate the shading of the pipe.
		CreateIndexBufferLOD();
	//	FillWaterSectionIB(); //PCN2347 removed (Antony van Iersel, 19 March 2004)
		CreateExportingPanel();


		
		nodeBuffer = new nodeStruct[(SECTION_SIZE+2)*PROFILE_SIZE]; 
	}

	
												 
	percentLoaded  =0;
	pipeDisplayType=0;	// Display type is Colour Range Vertices or Textured
	
	PreCalcCosSin();				// Pre-calculate the Trig for creating the Pipe Data
	LoadPipeTextures();				// Pre-Load all the Pipe Textures

	depthScale=(float) 2;	// Default DepthScale
	
	laserFocus = 1;
	laserSpeed = 0;
	laserWidth = 2; // PCN2365
	if(!OpenGL)	TogglePipeShade();	//PCN2950 Toggle on the colour shading so its on by default.
	
}

void CViewpipe::ConvertXYtoRadiusRawData()
{
	vec2double point;
	int i;
	int frameNo;
	if(units==0)
		for(i=0;i<noNodes;i++)
		{
			if((rawDatax[i] == 0) && (rawDatay[i] == 0)) {rawData[i]=0; continue;}
			frameNo = (i / 180);
			point = vec2double((double) (rawDatax[i]+rawCentrex[frameNo])/pVDataXYMultiplier,
										 (double) (rawDatay[i]+rawCentrey[frameNo])/pVDataXYMultiplier);
			if(point.x>10000) point.x = point.x-20000;
			if(point.x<-10000) point.x = point.x+20000;

			rawData[i]=(DistOfTwoPoints(0, point));
			
		}
	else
		for(i=0;i<noNodes;i++)
		{
			if((rawDatax[i] == 0) && (rawDatay[i] == 0)) {rawData[i]=0; continue;}
			frameNo = (i / 180);
			rawData[i]=(DistOfTwoPoints(0,vec2double((double) (rawDatax[i]+rawCentrex[frameNo])/pVDataXYMultiplier*25.40,
													 (double) (rawDatay[i]+rawCentrey[frameNo])/pVDataXYMultiplier*25.40)));
		}
//	FILE *f;
//	f=fopen("c:\\Test.dat","w");
//	for(i=0;i<noNodes;i++)
//	{
//		fprintf(f,"X = %i \tY = %i, \ttradius = %f\n",rawDatax[i], rawDatay[i], rawData[i]); 
//	}
//	fclose(f);
}

CViewpipe::~CViewpipe()
{
	UnloadPipe();
	if(exportingPanel!=NULL) { delete exportingPanel; exportingPanel = NULL; } //PCN2465 (8 December 2003, Antony)
	if(pTexturePipe!=NULL)    pTexturePipe->Release();  //PCN2465 (8 December 2003, Antony)
	if(pTextureLaser!=NULL)   pTextureLaser->Release(); //PCN2465 (8 December 2003, Antony)
	if(pTextureWhite!=NULL)	  pTextureWhite->Release(); //PCN2465 (8 December 2003, Antony)
	if(pIBDetail0_1PipeSection!=NULL) pIBDetail0_1PipeSection->Release(); //PCN2465 (8 December 2003, Antony)
	if(pIBDetail0_2PipeSection!=NULL) pIBDetail0_2PipeSection->Release(); //PCN2465 (8 December 2003, Antony)
	if(pIBDetail0_3PipeSection!=NULL) pIBDetail0_3PipeSection->Release(); //PCN2465 (8 December 2003, Antony)
	if(nodeBuffer!=NULL) { delete[] nodeBuffer; nodeBuffer = NULL; } //PCN2465 (8 December 2003, Antony) PCN3085 [] added
	if(preCalcSin!=NULL) { delete[] preCalcSin; preCalcSin = NULL; } //PCN2465 (8 December 2003, Antony) PCN3085 [] added
	if(preCalcCos!=NULL) { delete[] preCalcCos;	preCalcCos = NULL; } //PCN2465 (8 December 2003, Antony) PCN3085 [] added
	for(int i=0;i<numberSections;i++) { delete sectionIndex[i]; sectionIndex[i] = NULL; } //PCN2465 (8 December 2003, Antony)
	if(sectionIndex!=NULL)   { delete[] sectionIndex; sectionIndex = NULL; } //PCN2465 (8 December 2003, Antony) PCN3085 [] added
	if(startSection!=NULL)   { delete startSection; startSection = NULL; } //PCN2465 (8 December 2003, Antony)
	if(endSection!=NULL)     { delete endSection; endSection = NULL; } //PCN2465 (8 December 2003, Antony)
	if(rawData!=NULL) { delete[] rawData; rawData = NULL; }
}

void CViewpipe::PreCalcCosSin(void)
{
	int i;
	double rad=0,step;

	preCalcSin = new double[PROFILE_SIZE+1]; // one is added,
	preCalcCos = new double[PROFILE_SIZE+1]; // 0 to 180 back to 0
	step=(2*D3DX_PI)/PROFILE_SIZE;

	for(i=0;i<(PROFILE_SIZE);i++)
		{
		preCalcSin[i]=sin(rad);
		preCalcCos[i]=cos(rad);
		rad+=step;
		}
	preCalcSin[PROFILE_SIZE]=sin(0);
	preCalcCos[PROFILE_SIZE]=cos(0);
}

void CViewpipe::CreateIndexBufferLOD(void)
{
	long strip, around, count;

	WORD *p_index =NULL;
	WORD *p_wIndex=NULL;

	long aroundDetail0_1;
	long aroundDetail0_2;
	long aroundDetail0_3;
	long secOffset1;
	long secOffset2;
	long secOffset3;
	long secOffset4;

	numberPointsSection=(PROFILE_SIZE+1) * SECTION_SIZE;			// Note: Not used here, but its a good place to initialise
	numberTrianglesSection=(PROFILE_SIZE+1) * (SECTION_SIZE-1) * 2; // Note: Not used here, but its a good place to initialise

	aroundDetail0_1=(PROFILE_SIZE+1)/2; // Level of Detail Half   , Around
	aroundDetail0_2=(PROFILE_SIZE+1)/4; // Level of Detail Quater , Around
	aroundDetail0_3=(PROFILE_SIZE+1)/8; // Level of Detail Eighth , Around
	if((aroundDetail0_1*2)<(PROFILE_SIZE+1)) aroundDetail0_1++; // If not quite a clean
	if((aroundDetail0_2*4)<(PROFILE_SIZE+1)) aroundDetail0_2++; // devision, add one to
	if((aroundDetail0_3*8)<(PROFILE_SIZE+1)) aroundDetail0_3++; // complete the profile
	
	numTriSecDetail0_1 = aroundDetail0_1 * ((SECTION_SIZE-1) / 2); // Level of Detail Half,   up
	numTriSecDetail0_2 = aroundDetail0_2 * ((SECTION_SIZE-1) / 4); // Level of Detail Quater, up
	numTriSecDetail0_3 = aroundDetail0_3 * ((SECTION_SIZE-1) / 8); // Level of Detail Eighth, up

	numTriSecDetail0_1*= 2; // Two Triangles per Squared
	numTriSecDetail0_2*= 2; // Two Triangles per Squared
	numTriSecDetail0_3*= 2; // Two Triangles per Squared

	numPoiSecDetail0_1 = aroundDetail0_1 * ( SECTION_SIZE/2);
	numPoiSecDetail0_2 = aroundDetail0_2 * ( SECTION_SIZE/4);
	numPoiSecDetail0_3 = aroundDetail0_3 * ( SECTION_SIZE/8);

//////// PCN2347 removed (Antony van Iersel, 19 March 2004)///////
//	numTriWaterSecDet0_0 = (SECTION_SIZE-1) * 2 /1; // PCN2347 Two Triangles per Frame
//	numTriWaterSecDet0_1 = (SECTION_SIZE-1) * 2 /2; // PCN2347 Two Triangles per 2nd Frame
//	numTriWaterSecDet0_2 = (SECTION_SIZE-1) * 2 /4; // PCN2347 Two Traingles per 4th Frame
//	numTriWaterSecDet0_3 = (SECTION_SIZE-1) * 2 /8; // PCN2347 Two Triangles per 8th Frame
//
//	numPoiWaterSecDed0_0 = (((SECTION_SIZE) /1))*2; // PCN2347 Two Points per Frame + 1 For OverLap Frame in Section
//	numPoiWaterSecDed0_1 = (((SECTION_SIZE) /2))*2; // PCN2347 ""
//	numPoiWaterSecDed0_2 = (((SECTION_SIZE) /3))*2; // PCN2347 ""
//	numPoiWaterSecDed0_3 = (((SECTION_SIZE) /4))*2; // PCN2347 ""
//////////////////////////////////////////////////////////////////

	pd3dDevicePipe->CreateIndexBuffer(numTriSecDetail0_1*3*sizeof(WORD),
								  0,
								  D3DFMT_INDEX16,
								  D3DPOOL_DEFAULT,
								  &pIBDetail0_1PipeSection,
								  NULL);
	pd3dDevicePipe->CreateIndexBuffer(numTriSecDetail0_2*3*sizeof(WORD),
								  0,
								  D3DFMT_INDEX16,
								  D3DPOOL_DEFAULT,
								  &pIBDetail0_2PipeSection,
								  NULL);
	pd3dDevicePipe->CreateIndexBuffer(numTriSecDetail0_3*3*sizeof(WORD),
								  0,
								  D3DFMT_INDEX16,
								  D3DPOOL_DEFAULT,
								  &pIBDetail0_3PipeSection,
								  NULL);
	////////////////////////////////////////////////////////////
	// PCN2347 Create Water Index Buffer,					  //
	// The water takes its vertices from the edge of the Pipe //
	// For diferent Detail Levels //////////////////////////////
/* PCN2461 (8 December 2003, Antony) Removed will be put back when water is needed.
	pd3dDevicePipe->CreateIndexBuffer(numTriWaterSecDet0_0*sizeof(WORD),
									  0,
									  D3DFMT_INDEX16,
									  D3DPOOL_DEFAULT,
									  &pIBDetail0_0WaterSection,
									  NULL);
	pd3dDevicePipe->CreateIndexBuffer(numTriWaterSecDet0_1*sizeof(WORD),
									  0,
									  D3DFMT_INDEX16,
									  D3DPOOL_DEFAULT,
									  &pIBDetail0_1WaterSection,
									  NULL);
	pd3dDevicePipe->CreateIndexBuffer(numTriWaterSecDet0_2*sizeof(WORD),
									  0,
									  D3DFMT_INDEX16,
									  D3DPOOL_DEFAULT,
									  &pIBDetail0_2WaterSection,
									  NULL);
	pd3dDevicePipe->CreateIndexBuffer(numTriWaterSecDet0_3*sizeof(WORD),
									  0,
									  D3DFMT_INDEX16,
									  D3DPOOL_DEFAULT,
									  &pIBDetail0_3WaterSection,
									  NULL);
*/
/////////////////////////////////////////////////////////////////
	int a;
	count=0;  

	pIBDetail0_1PipeSection->Lock(0,0,(void **)  &p_index,  0);
	for(strip=0;strip<SECTION_SIZE-1;strip+=2)
		{
		for(around=0;around<PROFILE_SIZE;around+=2)
			{
			a=around+2;
			if(a>PROFILE_SIZE) a=PROFILE_SIZE;
			secOffset1 = (around  ) + ((strip  ) * (PROFILE_SIZE+1)); // *----*
			secOffset2 = (around  ) + ((strip+2) * (PROFILE_SIZE+1)); // |\   |
			secOffset3 = (a       ) + ((strip  ) * (PROFILE_SIZE+1)); // |  \ |
			secOffset4 = (a       ) + ((strip+2) * (PROFILE_SIZE+1)); // *----*

			// First Triangle in Square
			p_index[count++]=(WORD) secOffset1; // *	3    
			p_index[count++]=(WORD) secOffset2; // | \ 
			p_index[count++]=(WORD) secOffset3; // *--* 1  2

			// Secound Triangle in Square
			p_index[count++]=(WORD) secOffset3; // *--* 3  4
			p_index[count++]=(WORD) secOffset2; //  \ |
			p_index[count++]=(WORD) secOffset4; //    *    2
			}
		}
	pIBDetail0_1PipeSection->Unlock();


///////////////////////////////////////////////////////////////////

	count=0;
	pIBDetail0_2PipeSection->Lock(0,0,(void **) &p_index, 0);
	for(strip=0;strip<SECTION_SIZE-1;strip+=4)
		{
		for(around=0;around<PROFILE_SIZE;around+=4)
			{
			a=around+4;
			if(a>PROFILE_SIZE) a=PROFILE_SIZE;
			secOffset1 = (around  ) + ((strip  ) * (PROFILE_SIZE+1)); // *----*
			secOffset2 = (around  ) + ((strip+4) * (PROFILE_SIZE+1)); // |\   |
			secOffset3 = (a       ) + ((strip  ) * (PROFILE_SIZE+1)); // |  \ |
			secOffset4 = (a       ) + ((strip+4) * (PROFILE_SIZE+1)); // *----*

			// First Triangle in Square
			p_index[count++]=(WORD) secOffset1; // *	3    
			p_index[count++]=(WORD) secOffset2; // | \ 
			p_index[count++]=(WORD) secOffset3; // *--* 1  2

			// Secound Triangle in Square
			p_index[count++]=(WORD) secOffset3; // *--* 3  4
			p_index[count++]=(WORD) secOffset2; //  \ |
			p_index[count++]=(WORD) secOffset4; //    *    2
			}
		}
	pIBDetail0_2PipeSection->Unlock();

///////////////////////////////////////////////////////////////////

	count=0;
	pIBDetail0_3PipeSection->Lock(0,0,(void **) &p_index, 0);

	for(strip=0;strip<SECTION_SIZE-1;strip+=8)
		{
		for(around=0;around<PROFILE_SIZE;around+=8)
			{
			a=around+8;
			if(a>PROFILE_SIZE) a=PROFILE_SIZE;
			secOffset1 = (around  ) + ((strip  ) * (PROFILE_SIZE+1)); // *----*
			secOffset2 = (around  ) + ((strip+8) * (PROFILE_SIZE+1)); // |\   |
			secOffset3 = (a       ) + ((strip  ) * (PROFILE_SIZE+1)); // |  \ |
			secOffset4 = (a       ) + ((strip+8) * (PROFILE_SIZE+1)); // *----*

			// First Triangle in Square
			p_index[count++]=(WORD) secOffset1; // *	3    
			p_index[count++]=(WORD) secOffset2; // | \ 
			p_index[count++]=(WORD) secOffset3; // *--* 1  2

			// Secound Triangle in Square
			p_index[count++]=(WORD) secOffset3; // *--* 3  4
			p_index[count++]=(WORD) secOffset2; //  \ |
			p_index[count++]=(WORD) secOffset4; //    *    2
			}
		}
	pIBDetail0_3PipeSection->Unlock();

///////////////////////////////////////////////////////////////////
}

//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Name: FillWaterSectionIB PCN2347 5 November 2003
// Created By: Antony van Iersel
// 
// Description - Fills in the Index buffer if allready filled it just changes
//				 the data giving the ability to move the water level around.
//				 The Same index buffer is used for all the Pipe Sections.
//				 Four, One for each Detail (x1, x2, x4, x8)
//			
// Input - left, right. How far around the Pipe should the vertex points
//		   be for the water Index Points. The Vertex Points From the Pipe
//		   are used to connect water Index Points 
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

/*	//2461 (8 December 2003, Antony) Put back when Water is needed
void CViewpipe::FillWaterSectionIB()
{

	long strip;
	int count0=0;
	int count1=0;
	int count2=0;
	int count3=0;

	WORD *p_index0;	//Level Zero Detail,  every Frame
	WORD *p_index1; //Level One Detail,   every 2nd Frame
	WORD *p_index2; //Level Two Detail,   every 3rd Frame
	WORD *p_index3; //Level Three Detail, every 4th Frame

	pIBDetail0_0WaterSection->Lock(0,0,(void **) &p_index0, 0);
	pIBDetail0_1WaterSection->Lock(0,0,(void **) &p_index1, 0);	
	pIBDetail0_2WaterSection->Lock(0,0,(void **) &p_index2, 0);
	pIBDetail0_3WaterSection->Lock(0,0,(void **) &p_index3, 0);
	
	int offsetLeft, offsetRight;
	for(strip=0;strip<(SECTION_SIZE);strip++)
		{
		offsetLeft=count0; offsetRight=count0+1;
		p_index0[count0++]  = offsetLeft; 
		if(!(strip%2)) {p_index1[count1++] = offsetLeft; p_index1[count1++] = offsetRight;}
		if(!(strip%4)) {p_index2[count2++] = offsetLeft; p_index2[count2++] = offsetRight;}
		if(!(strip%8)) {p_index2[count3++] = offsetLeft; p_index3[count3++] = offsetRight;}
		p_index0[count0++] = offsetRight;
		}

	pIBDetail0_3WaterSection->Unlock();
	pIBDetail0_2WaterSection->Unlock();
	pIBDetail0_1WaterSection->Unlock();
	pIBDetail0_0WaterSection->Unlock();

}
*/

void CViewpipe::LoadNewSection(long start)
{
	if(start>frames) return;
	if(sectionIndex[(start/(SECTION_SIZE-1))]!=NULL) return;

	pipeSectionStruct *newPipeSection;
	newPipeSection = new pipeSectionStruct;
	newPipeSection->frameSectionStart = start;
	newPipeSection->frameSectionEnd   = start+SECTION_SIZE;

	sectionIndex[(start/(SECTION_SIZE-1))] = newPipeSection;
	

// For Number Vertex Points, one is Added to Profile size to sew up the top of the Profile.
// For Number Triangles, one is removed from Pipe Section Size because the triangles are
// inbetween the frames, two triangles per square.
// When calculating the number of points for the Index Buffer, three points per Triangle.
	
	ProcessPipeBuffer(start);

	pd3dDevicePipe->CreateVertexBuffer(numberPointsSection*sizeof(nodeStruct), 
		                           0, 
								   D3DFVF_CUSTOMVERTEX_PIPE, 
								   D3DPOOL_DEFAULT,
								   &newPipeSection->pVBPipeSection,
								   NULL);
/////////////// PCN2347 removed (Antony van Iersel, 19 March 2004)
//	pd3dDevicePipe->CreateVertexBuffer(numPoiWaterSecDed0_0*sizeof(nodeStruct),
//									0,
//									D3DFVF_CUSTOMVERTEX_PIPE,
//									D3DPOOL_DEFAULT,
//									&newPipeSection->pVBWaterSection,
//									NULL);
///////////////////////////////////////////////////////////////////

	pd3dDevicePipe->CreateIndexBuffer(numberTrianglesSection*3*sizeof(WORD),
								  0,
								  D3DFMT_INDEX16,
								  D3DPOOL_DEFAULT,
								  &newPipeSection->pIBPipeSection,
								  NULL);

	FillPipeSectionVB(start, newPipeSection->pVBPipeSection);
//	FillWaterSectionVB(start, newPipeSection->pVBWaterSection,waterLevelLeft,waterLevelRight);
	FillPipeSectionIB(start, newPipeSection->pIBPipeSection);
	newPipeSection->shade=shadeChange;	//PCN2337 keep track to see if Pipe needs to be painted
										//If shade is different to shadeChage then repaint

// Setup the Linked List Pointers, This has to be done last otherwise the new section
// would be drawn before it is ready.

	if(startSection==NULL)
		{
		startSection = newPipeSection;
		endSection   = newPipeSection;
		newPipeSection->prev=NULL;
		newPipeSection->next=NULL;

		}
	else {
		newPipeSection->prev=endSection;
		newPipeSection->next=NULL;

		endSection->next=newPipeSection;
		endSection=newPipeSection;
		}
    if((laserFocus>=start) && (laserFocus<=start+(SECTION_SIZE-1) )) MarkPipe(laserFocus);
}

void CViewpipe::SetVertexColours(nodeStruct *p_vertex, long secOffset, long frame, long profile)
{
	long profilePoint;

	if(((frame*PROFILE_SIZE)+profile)>noNodes) { test1++; return;}//MsgPipe("Buffer over run");
	profilePoint=(frame*(PROFILE_SIZE+1))+profile;
	
	
	//if(rawData[profilePoint] == 0) return;
	p_vertex[secOffset].colour=  D3DCOLOR_RGBA(rawColourRed[profilePoint] ,
											   rawColourGreen[profilePoint],
											   rawColourBlue[profilePoint],	 255);
}

//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Product Change Number: PCN2337 Antony van Iersel 4 November 2003
//						 
// Name: ShadeSection
// Created By: Antony van Iersel
//
// Description:
//		Changes to Vertex Colours of a given section according to VB set limits
//		*section - pointer to a Loaded section to change
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

void CViewpipe::ShadeSection(pipeSectionStruct *section)
{
	int up, around;	//frame count (up) a single section, profile point (around) the frame 
	nodeStruct *p_vertex; // pointer to the vertex buffer of this particullar section
	long secOffset; // pre-calculate the offset the particullar vertex curretly being looked at
	long frame; // frame that is matched with the PVD data that is passed

	section->pVBPipeSection->Lock(0,0,(void **) &p_vertex, 0); // unlock the vertex buffer for this section
	for(up=0;up<SECTION_SIZE;up++) //
		{
		frame = up + section->frameSectionStart;
		for(around=0;around<(PROFILE_SIZE+1);around++)
			{
			secOffset = (up * (PROFILE_SIZE+1) + around);
			if(frame<frames)
				SetVertexColours(p_vertex, secOffset, frame, around%PROFILE_SIZE); //PCN2692 Function repeated code
			else p_vertex->colour=D3DCOLOR_RGBA(0, 0  ,0  ,  0);
			
			//													 was repeated in Fillpipe Section
			/////////////////////////////////////////////////////now its in Function SetVertexColours
			}
		}
	section->pVBPipeSection->Unlock();
	section->shade=shadeChange;
}

void CViewpipe::FillPipeSectionVB(long start, LPDIRECT3DVERTEXBUFFER9 &lpVB)
{
	int up , around; // up pipe, around pipe
	float z=0;
	long frame;
	long secOffset;    //Pipe Section offset is the pipe VB offset;
	long bufferOffset; //Buffer offset is the temp buffer with pre-calc values eg vertex, normal etc
	nodeStruct *p_vertex;
	p_vertex=NULL;
	lpVB->Lock(0,0,(void **) &p_vertex, 0);

	for(up=0;up<SECTION_SIZE;up++) // Start counting frames up the pipe;
		{
		frame=up+start;
		for(around=0;around<(PROFILE_SIZE+1);around++) // Start counting points around the pipe
			{
			secOffset    = (up     *(PROFILE_SIZE+1)) +  around;
			bufferOffset = ((up+1) * PROFILE_SIZE   ) + (around%(PROFILE_SIZE));
		
			p_vertex[secOffset].vertice.x=nodeBuffer[bufferOffset].vertice.x;
			p_vertex[secOffset].vertice.y=nodeBuffer[bufferOffset].vertice.y;
			p_vertex[secOffset].vertice.z=nodeBuffer[bufferOffset].vertice.z;
			p_vertex[secOffset].normal   = nodeBuffer[bufferOffset].normal;
			
			p_vertex[secOffset].t1u      =(float) up*8/SECTION_SIZE;
			p_vertex[secOffset].t2u		 =(float) 0;
			p_vertex[secOffset].t1v		 =(float) around*8/PROFILE_SIZE+1;
			p_vertex[secOffset].t2v		 =(float) around/PROFILE_SIZE+1;;

			/////////////////////////////////////////////////////////////
			// If Shading is set then paint the pipe when Loading PCN2337
			
			if(frame<frames)
			
				SetVertexColours(p_vertex, secOffset, frame, around%PROFILE_SIZE); //PCN2692 Function repeated code
			else p_vertex->colour = D3DCOLOR_RGBA(0,0,0, 0);
			//											  //was repeated in shade section
			////////////////////////////////////////////////now its in Function SetVertexColours
			}
		z=z-1;
		}
	lpVB->Unlock();
}

/*  PCN2347 removed (Antony van Iersel, 19 March 2004)
void CViewpipe::FillWaterSectionVB(long start, LPDIRECT3DVERTEXBUFFER9 &lpVB, long left, long right)
{
	long count;
	int up;				// up pipe, around pipe
	float z=0;
	long frame;
//	long secOffsetLeft;		//Water Level Left  Section offset is the water VB offset;
//	long secOffsetRight;	//Water Level Right Section offset is the water VB offset;
	long bufferOffsetLeft;	//Buffer offset is the temp buffer with pre-calc values eg vertex, normal etc
	long bufferOffsetRight;	//Buffer offset is the temp buffer with pre-calc values eg vertex, normal etc
	nodeStruct *p_vertex;
	p_vertex=NULL;
	
	count=0;
	lpVB->Lock(0,0,(void **) &p_vertex, 0);
	for(up=0;up<SECTION_SIZE;up++) // Start counting frames up the pipe;
		{
		frame=up+start;
			
		bufferOffsetLeft = ((up+1) * PROFILE_SIZE   ) + (left %(PROFILE_SIZE));
		bufferOffsetRight= ((up+1) * PROFILE_SIZE   ) + (right%(PROFILE_SIZE));
	
		/// Vertex Values for the Left water Level
		p_vertex[count].vertice.x=nodeBuffer[bufferOffsetLeft].vertice.x;
		p_vertex[count].vertice.y=nodeBuffer[bufferOffsetLeft].vertice.y;
		p_vertex[count].vertice.z=(float) z;
		p_vertex[count].normal   = nodeBuffer[bufferOffsetLeft].normal;
		p_vertex[count].colour= D3DCOLOR_RGBA(255,255,255,180);
		p_vertex[count].t1u      =(float) up*8/SECTION_SIZE;
		p_vertex[count].t2u		 =(float) 0;
		p_vertex[count].t1v		 =(float) 0;
		p_vertex[count++].t2v	 =(float) 0;

		/// Vertex Values for the Right water Level
		p_vertex[count].vertice.x=nodeBuffer[bufferOffsetRight].vertice.x;
		p_vertex[count].vertice.y=nodeBuffer[bufferOffsetRight].vertice.y;
		p_vertex[count].vertice.z=(float) z;
		p_vertex[count].normal   = nodeBuffer[bufferOffsetRight].normal;
		p_vertex[count].colour   = D3DCOLOR_RGBA(255,255,255,220);
		p_vertex[count].t1u      =(float) up*8/SECTION_SIZE;
		p_vertex[count].t2u		 =(float) 0;
		p_vertex[count].t1v		 =(float) 1;
		p_vertex[count++].t2v	 =(float) 1;

		/////////////////////////////////////////////////////////////
		
		z=z-1;
		}

	lpVB->Unlock();
}
*/

void CViewpipe::FillPipeSectionIB(long start, LPDIRECT3DINDEXBUFFER9 &lpIB)
{
	long strip, around;
	long count;
	long secOffset1, secOffset2, secOffset3, secOffset4; // Offset for the four courners on the pipe section
	long datOffset1, datOffset2, datOffset3, datOffset4; // Offset for the four courners from the raw data and
	long frame;											 // is only used for bounds checking, if out of bounds then
	WORD *p_index;										 //	trinagle data is set to zero.

	p_index=NULL;
	
	lpIB->Lock(0,0,(void **) &p_index, 0);
	count=0;	// Reset the Count to start indexing from the start of the Index Buffer
	for(strip=0;strip<SECTION_SIZE-1;strip++)
		{
		frame=start+strip;
		for(around=0;around<PROFILE_SIZE;around++) //ANTTesting -1
			{

			datOffset1 = (around  ) + ((frame  ) * PROFILE_SIZE);
			datOffset2 = (around  ) + ((frame+1) * PROFILE_SIZE);
			if(!(around%(PROFILE_SIZE-1)) && (around>0))
				{
				datOffset3 = (0       ) + ((frame  ) * PROFILE_SIZE);
				datOffset4 = (0       ) + ((frame+1) * PROFILE_SIZE);  
				}
			else{
				datOffset3 = (around+1) + ((frame  ) * PROFILE_SIZE);
				datOffset4 = (around+1) + ((frame+1) * PROFILE_SIZE);
				}

			secOffset1 = (around  ) + ((strip  ) * (PROFILE_SIZE+1) ); // *----*
			secOffset2 = (around  ) + ((strip+1) * (PROFILE_SIZE+1) ); // |\   |
			secOffset3 = (around+1) + ((strip  ) * (PROFILE_SIZE+1) ); // |  \ |
			secOffset4 = (around+1) + ((strip+1) * (PROFILE_SIZE+1) ); // *----*
			
			if(((datOffset1>noNodes) ||
				(datOffset2>noNodes) ||
				(datOffset3>noNodes) ||
				(datOffset4>noNodes)))
				{
				p_index[count++]=(WORD) 0; // *    
				p_index[count++]=(WORD) 0; // | \ 
				p_index[count++]=(WORD) 0; // *--* 

				// Secound Triangle in Square
				p_index[count++]=(WORD) 0; // *--* 
				p_index[count++]=(WORD) 0; //  \ |
				p_index[count++]=(WORD) 0; //    * 
				}
			else if(
			    (rawData[datOffset1]==0) ||
				(rawData[datOffset2]==0) ||
				(rawData[datOffset3]==0) ||
				(rawData[datOffset4]==0))
				{
				p_index[count++]=(WORD) 0; // *    
				p_index[count++]=(WORD) 0; // | \ 
				p_index[count++]=(WORD) 0; // *--* 

				// Secound Triangle in Square
				p_index[count++]=(WORD) 0; // *--* 
				p_index[count++]=(WORD) 0; //  \ |
				p_index[count++]=(WORD) 0; //    * 
				}
			else{
				// First Triangle in Square
				p_index[count++]=(WORD) secOffset1; // *	3    
				p_index[count++]=(WORD) secOffset2; // | \ 
				p_index[count++]=(WORD) secOffset3; // *--* 1  2

				// Secound Triangle in Square
				p_index[count++]=(WORD) secOffset3; // *--* 3  4
				p_index[count++]=(WORD) secOffset2; //  \ |
				p_index[count++]=(WORD) secOffset4; //    *    2
				}
			}
		}
	lpIB->Unlock();
	
}

void CViewpipe::UnloadPipe(void)
{

	while(startSection!=NULL) UnloadPipeSection(startSection);
}

void CViewpipe::UnloadPipeSection(pipeSectionStruct *section)
{
	pipeSectionStruct *newNext;
	pipeSectionStruct *newPrev;


	newNext=section->next;
	newPrev=section->prev;
	if(newNext!=NULL) newNext->prev=newPrev;
	else endSection=section->prev;	// Check to see if Ending Pointer has to be shifted

	if(newPrev!=NULL) newPrev->next=newNext;
	else startSection=section->next;// Check to see if the Starting Pointer has to be shifted
	
	if(section->pIBPipeSection!=NULL) section->pIBPipeSection->Release();
	if(section->pVBPipeSection!=NULL) section->pVBPipeSection->Release();

	//////PCN2347 removed (Antony van Iersel, 19 March 2004)
	//	if(section->pVBWaterSection!=NULL) section->pVBWaterSection->Release(); //PCN2465 (8 December 2003, Antony)
	////////////////////////////////////////////////////////
	
	//delete section->pd3dDevicePipe;
	
	sectionIndex[(section->frameSectionStart/(SECTION_SIZE-1))]=NULL;
	delete section;
}

void CViewpipe::RenderD3D(void)
	{

	if(NULL == pd3dDevicePipe) {MsgPipe("d3d102"); exit(1);} //"pd3dDevice NULL"

	// Clear the backbuffer to D3DCOLOR_XRGB color

    pd3dDevicePipe->Clear( 0, NULL, D3DCLEAR_TARGET|D3DCLEAR_ZBUFFER,
                         D3DCOLOR_XRGB(250,250,255), 1.0f, 0 );

	if(exportingPanel!=NULL) exportingPanel->Draw();
	// Present the backbuffer contects to the display
	pd3dDevicePipe->Present( NULL, NULL, NULL, NULL);
	}

void CViewpipe::ChangePipeTexture(char *pTex)
{
	char texFile[800];

	if(pTexturePipe!=NULL) pTexturePipe->Release();
	strcpy(texFile,textureDirectory);
	strcat(texFile, pTex);
	if(FAILED( D3DXCreateTextureFromFile(pd3dDevicePipe, texFile, &pTexturePipe) ));	
	//PCN4240 MsgPipe("%s %s\n%s", languagePipe[15].text, // PCN2473 Language pointer (Antony, 11 May 2004)
	//					 languagePipe[16].text, // PCN2473 Language pointer (Antony, 11 May 2004)
	//					 texFile); //PCN2467 Texture file and directory added to error, (texFile)
}													//(9 December 2003, Antony van Iersel)

void CViewpipe::SetVBLaserTexture(pipeSectionStruct *section, long fraction, float texPos)
{
	long up, around;
	long offset;
	long ringCount=0;

	nodeStruct *p_vertex;
	p_vertex=NULL;
	up=fraction+1;

	section->pVBPipeSection->Lock(0,0,(void **) &p_vertex, 0);
	for(around=0;around<PROFILE_SIZE+1;around++)
		{
		offset=around + (up*(PROFILE_SIZE+1));
		p_vertex[offset].t2u=(float) texPos;
		laserRing[ringCount]=p_vertex[offset].vertice;
		laserRing[ringCount].z=(section->frameSectionStart-laserRing[ringCount].z)*-depthScale;
		ringCount++;
		}
	section->pVBPipeSection->Unlock();
	offset=up*2; // PCN2347 The Two vertex Points that need to change for laser on Water

////////////PCN2347 removed (Antony van Iersel, 19 March 2004)
//	section->pVBWaterSection->Lock(0,0,(void **) &p_vertex, 0);
//		p_vertex[offset].t2u  =(float) texPos;
//		p_vertex[offset+1].t2u=(float) texPos;
//	section->pVBWaterSection->Unlock();
//////////////////////////////////////////////////////////////

}


void CViewpipe::NextPipeTexture(void)
{
	currentTexture++; currentTexture%=4;
	ChangePipeTexture(textureList[currentTexture]);
}


void CViewpipe::ClearLaser(long fr)
	{

	long section;
	long fraction;

	if(OpenGL) return;

	if(fr>frames) return;
	if(fr<1) return;
	section  = fr/(SECTION_SIZE-1);
	fraction = fr-(section*(SECTION_SIZE-1));
	if(sectionIndex[section]==NULL) LoadNewSection(section*(SECTION_SIZE-1));
	SetVBLaserTexture(sectionIndex[section],fraction, 0.0);
	}

void CViewpipe::MarkPipe(long fr)
	{
	long section;
	long fraction;

	if(OpenGL) return;

	if(fr>frames) return;
	if(fr<1) return;
	section  = fr/(SECTION_SIZE-1);
	fraction = fr-(section*(SECTION_SIZE-1));
	if(sectionIndex[section]==NULL) LoadNewSection(section*(SECTION_SIZE-1));
	SetVBLaserTexture(sectionIndex[section],fraction, 0.5);
	}

//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Name: ChangeLaserWidth PCN2365 7 November 2003
// Created By: Antony van Iersel
// 
// Description - Changes the Width of the Laser, this is controlled by the distance
//				 from the laser line to the current cammera. 
//				 It is called just before
//               the drawing loop of the sections.
// Input - width , should be a division of 2, but if not then it will change it
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

void CViewpipe::ChangeLaserWidth(int width)
{
	int i;
	
	if(width<2) return;			  // Minimum Laser Width it 2
	width=(width/2)*2;			  // Make sure width is a division of 2;
	if(width>50) width=50;
	if(width==laserWidth) return; // If there is no change then egnore.
	

	if(width<laserWidth) // If reducing the width of the laser then erase it first
		{
		laserFocus-=(laserWidth/2); 
		for(i=0;i<laserWidth;i++) { ClearLaser(laserFocus); laserFocus++; }
		laserFocus-=(laserWidth/2);
		}
	
	laserWidth=width;

	laserFocus-=(laserWidth/2);
	for(i=0;i<laserWidth;i++) { MarkPipe(laserFocus); laserFocus++; }
	laserFocus-=(laserWidth/2);

}

void CViewpipe::MoveLaserTo(int f)
{
	int i;
	if(f<1) { f=1; laserSpeed=0; }
	if(f>frames-8) { f=frames-8; laserSpeed=0; }
								// PCN2365 adjusted to make the laser frame to be
	laserFocus-=(laserWidth/2);	// the centre of the drawn laser.
	for(i=0;i<laserWidth;i++) { ClearLaser(laserFocus); laserFocus++; }
	laserFocus-=(laserWidth/2); //

	laserFocus=f;				//
	laserFocus-=(laserWidth/2); //
	for(i=0;i<laserWidth;i++) { MarkPipe(laserFocus); laserFocus++; }
	laserFocus-=(laserWidth/2); //
}


void CViewpipe::MoveLaser(int dis)
{
	MoveLaserTo(laserFocus+dis);
}

void CViewpipe::FindHighestPoint(void)
	{
	int i;
		
	D3DXVECTOR3 dest;
	D3DVIEWPORT9 viewport;
	D3DXMATRIX projection;
	D3DXMATRIX view;
	D3DXMATRIX world;

	highestPoint=0;

	pd3dDevicePipe->GetViewport(&viewport);
	pd3dDevicePipe->GetTransform(D3DTS_PROJECTION , &projection);
    pd3dDevicePipe->GetTransform(D3DTS_VIEW, &view);

	topRing = laserRing[0];
	for(i=0;i<PROFILE_SIZE;i++)
		{
		D3DXVec3Project(
			&laserRingProj[i],
			&laserRing[i],
			&viewport,
			&projection,
			&view,
			NULL);
		}
	for(i=1;i<PROFILE_SIZE;i++)
		{
		if(laserRingProj[i].y<laserRingProj[highestPoint].y) highestPoint=i;
		}
}

//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Product Change Number: PCNAVI110603-1
// Name: PipeDraw
// Created By: Antony van Iersel
// Description:
//		Actuall Drawing of Pipe Vertices, Done with Triangles , Ring at a time
//		Section at a time
// Changed PCN2337, Pipe Colour limited Shading added 5 November 2003
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

void CViewpipe::PipeDraw(void)
{
	if(OpenGL) {OpenGLDraw(); return;}
	int lod=0;
	float distanceToCamera;
	float f;
	float lod0=1000*((float) expRad/100);
	float lod1=2000*((float) expRad/100);
	float lod2=4000*((float) expRad/100);
	float lod3=8000*((float) expRad/100);
	long i,start,end;

	D3DXVECTOR3 temp;
////////////////////////////
//	Update Section Buffer //
///////////////////////////////////////////////////////////////////////////////

	long lf;	// lf - what frame to focus the detail on.
	long camPosFrame;
	long noOfChanges;

	// if camera type 0 or 1, Pipe or Rail cam then the laser position in centre of detail
	// else closest frame to the camera's z position.
	if((cameraType==0) || (cameraType==1)) lf=laserFocus/(SECTION_SIZE-1); // What section to start with
	else { 
		camPosFrame=(long) -cameraPosition.z;		
		camPosFrame=(long) ((float) camPosFrame/depthScale);
		if(camPosFrame<0) camPosFrame=0;
		if(camPosFrame>frames) camPosFrame=frames;	
		lf=camPosFrame/(SECTION_SIZE-1);
		}

	start=lf-50;
	end=lf+50;
	if(start<0) start=0;
	if(end>(frames/(SECTION_SIZE-1))+1) end=(frames/(SECTION_SIZE-1)+1);

	loadingPipe=false;
	repaintingPipe=false; // PCN2337
	noOfChanges=0;		  // PCN2337
	for(i=0;i<start;i++) // Remove sections from Buffer are before the start;
		{
		
		if(sectionIndex[i]!=NULL) 
			{ 
			UnloadPipeSection(sectionIndex[i]); 
			loadingPipe=true;
			noOfChanges++; 
			if(noOfChanges>1) break;
			}
		}
	noOfChanges=0;
	for(i=end;i<(frames/(SECTION_SIZE-1)+1);i++) // Remove sections from Buffer are after the end 
		{
		if(sectionIndex[i]!=NULL) 
			{ 
			UnloadPipeSection(sectionIndex[i]); 
			noOfChanges++; 
			if(noOfChanges>1) break;	
			}
		
		}
	noOfChanges=0;
	for(i=start;i<end;i++) // Any section not in buffer, that is supposed to be, add.
		{
		if(sectionIndex[i]==NULL) 
			{

			LoadNewSection(i*(SECTION_SIZE-1));
			loadingPipe=true;
			noOfChanges++;
			if(noOfChanges>1) break;
			}
		else if(sectionIndex[i]->shade!=shadeChange) // If the shadeCange is not same as current Section
		{										     // Then repaint before drawing
			ShadeSection(sectionIndex[i]);
			repaintingPipe=true;
			noOfChanges++;
			if(noOfChanges>1) break;	// Only do one change per screen refresh.
			}
		}

//////////////////////////////////////////////////////////////
	
	if(pipeDisplayType==0) pd3dDevicePipe->SetTexture(0,pTexturePipe); // PCN2337 
	if(pipeDisplayType==1) pd3dDevicePipe->SetTexture(0,pTextureWhite);// PCN2337 Make pipe white for Colour Shading
	pd3dDevicePipe->SetTextureStageState( 0, D3DTSS_COLOROP, D3DTOP_MODULATE);
	pd3dDevicePipe->SetTextureStageState( 0, D3DTSS_COLORARG1, D3DTA_TEXTURE);
	pd3dDevicePipe->SetTextureStageState( 0, D3DTSS_COLORARG2, D3DTA_CURRENT);
	pd3dDevicePipe->SetTextureStageState( 0, D3DTSS_ALPHAOP, D3DTOP_DISABLE);
	
	pd3dDevicePipe->SetTexture(1,pTextureLaser);
	pd3dDevicePipe->SetTextureStageState( 1, D3DTSS_COLOROP, D3DTOP_ADD);
	pd3dDevicePipe->SetTextureStageState( 1, D3DTSS_COLORARG1, D3DTA_TEXTURE );
	pd3dDevicePipe->SetTextureStageState( 1, D3DTSS_COLORARG2, D3DTA_CURRENT);
	pd3dDevicePipe->SetTextureStageState( 1, D3DTSS_ALPHAOP, D3DTOP_DISABLE);

	pd3dDevicePipe->SetFVF( D3DFVF_CUSTOMVERTEX_PIPE );
	D3DXMATRIX scale,rot,translation,world;

	
	D3DXMatrixScaling(&scale,1,1,1); // PCN2410 last perameter was depthScale but now scaling
									 // is done in the buffer section so that shading is effected.
									 // (18 November 2003 Antony).
	long count=0;
//	float length=0;

	pipeSectionStruct *currentSection;
//	pipeSectionStruct *laserSection;
	
	currentSection=startSection;
//	laserSection=sectionIndex[laserFocus/(SECTION_SIZE-1)];

//	Don't know why I did this... if(laserSection!=NULL)					
//	Don't know why I did this...	{
//	Don't know why I did this...	f=(float) laserSection->frameSectionStart+(SECTION_SIZE/2);
//	Don't know why I did this...	f*=depthScale;
//	Don't know why I did this...	centerLaserRing=D3DXVECTOR3(0,0,f);
//	Don't know why I did this...	}
//	Don't know why I did this...temp = cameraPosition-centerLaserRing;
//	Don't know why I did this...distanceToCamera=D3DXVec3Length(&temp);
//////// So Out it comes // 7 November 2003 //////////////////////////////////

	// PCN2365 Finds the centre of the Laser Ring, then calculates the distance to
	// the current Cammera and divides the distance by 200 to get the laser width
	centerLaserRing=D3DXVECTOR3(0,0,(laserFocus*-depthScale));
	temp=cameraPosition-centerLaserRing;		
	distanceToCamera=D3DXVec3Length(&temp);		
	ChangeLaserWidth((int) distanceToCamera/200); //
	////////////////////////////////////////////////
	//D3DXMatrixTranslation(&translation,pipePosition.x,pipePosition.y,pipePosition.z);
	//pd3dDevicePipe->SetTransform(D3DTS_WORLD,&translation);
	while(true)
		{
		if(currentSection==NULL) break;
		// Moves current section to appropriate start position in 3D coordinates //
		D3DXMatrixTranslation(&translation,pipePosition.x,0,(float) -currentSection->frameSectionStart*depthScale); // PCN2410 (18 November 2003 Antony). 
		D3DXMatrixMultiply(&world, &translation, &scale);						 // 
		pd3dDevicePipe->SetTransform( D3DTS_WORLD, &world);						 //
		///////////////////////////////////////////////////////////////////////////

		f=(float) currentSection->frameSectionStart+(SECTION_SIZE/2); // With the centre
		f*=-depthScale;				// frame find the 3D coordingate of centre from section
		
//		temp = cameraPosition-D3DXVECTOR3(0,0,f);	// Find the distance from the camera
//		distanceToCamera=D3DXVec3Length(&temp);     // the the centre of the section, this is used
//													// to determine the LOD when drawing the pipe
		temp = D3DXVECTOR3(0,0,f);
		temp = cameraPosition-temp;
		distanceToCamera=D3DXVec3Length(&temp);
		


		if((laserFocus>=currentSection->frameSectionStart) && (laserFocus<=currentSection->frameSectionEnd)) lod=0; 

	
//		else if(distanceToCamera<lod0*depthScale) lod=0;
//		else if((distanceToCamera>=lod0*depthScale) && (distanceToCamera<lod1*depthScale)) lod=1;
//		else if((distanceToCamera>=lod1*depthScale) && (distanceToCamera<lod2*depthScale)) lod=2;
//		else lod=3;
	
	
		else if(distanceToCamera<lod0) lod=0;
		else if((distanceToCamera>=lod0) && (distanceToCamera<lod1)) lod=1;
		else if((distanceToCamera>=lod1) && (distanceToCamera<lod2)) lod=2;
		else lod=3;



		if(pipeDisplayType==0) pd3dDevicePipe->SetTexture(0,pTexturePipe); // PCN2337 
		if(pipeDisplayType==1) pd3dDevicePipe->SetTexture(0,pTextureWhite);// PCN2337 Make pipe white for Colour Shading
		pd3dDevicePipe->SetTextureStageState( 0, D3DTSS_COLOROP, D3DTOP_MODULATE);
		pd3dDevicePipe->SetTextureStageState( 0, D3DTSS_COLORARG1, D3DTA_TEXTURE);
		pd3dDevicePipe->SetTextureStageState( 0, D3DTSS_COLORARG2, D3DTA_CURRENT);
		pd3dDevicePipe->SetTextureStageState( 0, D3DTSS_ALPHAOP, D3DTOP_DISABLE);
		
		pd3dDevicePipe->SetStreamSource(0, currentSection->pVBPipeSection, 0, sizeof(nodeStruct));

		// PCN2347 Been tidied so compliment the new water section
		long numTri;
//		lod=2;
		
		if(lod==0) { pd3dDevicePipe->SetIndices(currentSection->pIBPipeSection); numTri=numberTrianglesSection;}
		if(lod==1) { pd3dDevicePipe->SetIndices(pIBDetail0_1PipeSection); numTri=numTriSecDetail0_1;}
		if(lod==2) { pd3dDevicePipe->SetIndices(pIBDetail0_2PipeSection); numTri=numTriSecDetail0_2;}
		if(lod==3) { pd3dDevicePipe->SetIndices(pIBDetail0_3PipeSection); numTri=numTriSecDetail0_3;}

		test1=numTri;
		pd3dDevicePipe->DrawIndexedPrimitive(D3DPT_TRIANGLELIST,
											 0,
											 0,
											 numberPointsSection,
											 0,
											 numTri);

	

//PCN2461 pd3dDevicePipe->SetTexture(0,pTextureWater);// PCN2337 Make pipe white for Colour Shading
//PCN2461 pd3dDevicePipe->SetTextureStageState( 0, D3DTSS_COLOROP, D3DTOP_MODULATE);
//PCN2461 pd3dDevicePipe->SetTextureStageState( 0, D3DTSS_COLORARG1, D3DTA_TEXTURE);
//PCN2461 pd3dDevicePipe->SetTextureStageState( 0, D3DTSS_COLORARG2, D3DTA_CURRENT);
//		pd3dDevicePipe->SetTextureStageState( 0, D3DTSS_ALPHAOP, D3DTOP_DISABLE);

		// PCN2347 Drawing to water in the Pipe. 6 Nomvember 2003, Antony van Iersel //
//PCN2461 pd3dDevicePipe->SetStreamSource(0, currentSection->pVBWaterSection, 0, sizeof(nodeStruct));
//		if(lod==0) { pd3dDevicePipe->SetIndices(pIBDetail0_0WaterSection); numTri=numTriWaterSecDet0_0-2; }
//		if(lod==1) { pd3dDevicePipe->SetIndices(pIBDetail0_1WaterSection); numTri=numTriWaterSecDet0_1-2; }
//		if(lod==2) { pd3dDevicePipe->SetIndices(pIBDetail0_2WaterSection); numTri=numTriWaterSecDet0_2-2; }
//		if(lod==3) { pd3dDevicePipe->SetIndices(pIBDetail0_3WaterSection); numTri=numTriWaterSecDet0_3-2; }
	
//		pd3dDevicePipe->DrawIndexedPrimitive(D3DPT_TRIANGLESTRIP,					 //
//											 0,										 //
//											 0,										 //
//											 numPoiWaterSecDed0_0,					 //
//											 0,										 //
//											 numTri);								 //
		///////////////////////////////////////////////////////////////////////////////
		
		if(currentSection==endSection) break;
		currentSection=currentSection->next;
//		length-=100;
		}

	if(sectionIndex[laserFocus/(SECTION_SIZE-1)]!=NULL) FindHighestPoint();

	D3DXMatrixTranslation(&translation,0,0,0);
	D3DXMatrixScaling(&scale,1,1,1);
	D3DXMatrixMultiply(&world, &translation, &scale);
    pd3dDevicePipe->SetTransform( D3DTS_WORLD, &world);
}

double  CViewpipe::CalcAvgRadius(void)
{
	long i;
	double total=0;
	for(i=0;i<noNodes;i++) total=total+rawData[i];
	return(total/noNodes);
}





//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Name: CViewpipe::ProcessPipeBuffer
// Created By: Antony van Iersel
// Change	  : PCNAVI 16 October, 
//              Input now has out Vector 'n' for Normal,
//				Once nomal calulated its put back into D3DXVECTOR3 variable
// Reason	  : New Pipe data directly loaded into Vertex Buffer insted of into 
//				a node buffer. Now surrounding normals have to be calculated temporary,
//				then the result placed strait into Verbtex Buffer.
// Description:
//		Calculating Vertice Normal, by getting Avg of The Six surrounding Trangles
//		Input Vertice - (2nd) seg: segment, (3rd) ring: how deap into pipe. 				   
//					  - destination Vertex (1st)
//
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


void CViewpipe::ProcessPipeBuffer(long start)
{
	long around;
	long i,b;
	long offsetFrame;
	long section=0;
	
//	nodeBuffer->sectionLength=0;

	for(section=0;section<SECTION_SIZE+2;section++)
		{
		offsetFrame=(start-1+section);
		if(offsetFrame<0)      offsetFrame=0;
		if(offsetFrame>frames) offsetFrame=frames;
		for(around=0;around<PROFILE_SIZE;around++)
			{
			i = around + (offsetFrame*PROFILE_SIZE); // i Index for the raw data of the Pipe
			b = around + (section    *PROFILE_SIZE); // b Index for the Buffer data of the Pipe
			if(i>noNodes)
				{
				nodeBuffer[b].vertice.x=(float) (float) preCalcSin[around]*-(float) (expRad); //PCN3188 negate the X axis to flip the 3D from left to right. around; //
				nodeBuffer[b].vertice.y=(float) (float) preCalcCos[around]*-(float) (expRad); //0 ;
				nodeBuffer[b].colour=D3DCOLOR_RGBA(0,0,0,0);
				}
			else if(rawData[i]==0)
				{
				nodeBuffer[b].vertice.x=(float) preCalcSin[around]*-(float) (expRad); //PCN3188 negate the X axis to flip the 3D from left to right.(around*depthScale); //
				nodeBuffer[b].vertice.y=(float) preCalcCos[around]*-(float) (expRad); //0; //
				nodeBuffer[b].colour=colClearWhite;
				}
			else {
				nodeBuffer[b].vertice.x=(float) preCalcSin[around]*-(float) (rawData[i]); //PCN3188 negate the X axis to flip the 3D from left to right. (around*depthScale);//
				nodeBuffer[b].vertice.y=(float) preCalcCos[around]*-(float) (rawData[i]); //(rawData[i]); //
				nodeBuffer[b].colour= colWhite;	
				}
			nodeBuffer[b].vertice.z=(float) (-(section-1)*depthScale); // Was times scale, but that is now at PipeDraw
			//  sectionLength+=rawLength[i];						   // PCN2410, now its back (x by depthScael, needs to
																	   // done here so that the shading is adjusted for scale.
			}
		}
	for(section=1;section<(SECTION_SIZE+1);section++)
		for(around=0;around<PROFILE_SIZE;around++)
			{
			AveSixTriangles(nodeBuffer[around+(section*PROFILE_SIZE)].normal, section,around);
			}
}



//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Product Change Number: PCNAVI110603-1
//						 
// Name: AveSixTriangles
// Created By: Antony van Iersel
// Change	  : PCNAVI 16 October, 
//              Input now has out Vector 'n' for Normal,
//				Once nomal calulated its put back into D3DXVECTOR3 variable
// Reason	  : New Pipe data directly loaded into Vertex Buffer insted of into 
//				a node buffer. Now surrounding normals have to be calculated temporary,
//				then the result placed strait into Verbtex Buffer.
// Description:
//		Calculating Vertice Normal, by getting Avg of The Six surrounding Trangles
//		Input Vertice - (2nd) seg: segment, (3rd) ring: how deap into pipe. 				   
//					  - destination Vertex (1st)
//
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#define POne		((frameNo  )*offset)+(seg  ) 
#define PTwo		((frameNo-1)*offset)+(seg  )
#define PThree		((frameNo-1)*offset)+(seg+1)
#define PThree2	    ((frameNo-1)*offset)
#define PFour  	    ((frameNo  )*offset)+(seg+1)
#define PFour2		((frameNo  )*offset)
#define PFive  	    ((frameNo+1)*offset)+(seg  )
#define PSix		((frameNo+1)*offset)+(seg-1)
#define PSix2		((frameNo+1)*offset)+(PROFILE_SIZE-1)
#define PSeven		((frameNo  )*offset)+(seg-1)
#define PSeven2     ((frameNo  )*offset)+(PROFILE_SIZE-1)

void CViewpipe::AveSixTriangles(D3DXVECTOR3 &n, long frameNo, long seg)
{
	int i;
	long offset=PROFILE_SIZE;
	if((frameNo<1) || (frameNo>=SECTION_SIZE+1)) { n=D3DXVECTOR3(0,0,0); return;}

	D3DXVECTOR3 tOne, tTwo, tThree, tFour, tFive, tSix;
	D3DXVECTOR3 *vOne, *vTwo, *vThree, *vFour, *vFive, *vSix, *vSeven;
//	MsgPipe("AveSixTriangles - frameNo %i, seg %i",frameNo, seg);
		
								vOne	=&nodeBuffer[POne].vertice;
								vTwo	=&nodeBuffer[PTwo].vertice;
		if(seg==PROFILE_SIZE-1) vThree	=&nodeBuffer[PThree2].vertice; else vThree =&nodeBuffer[PThree].vertice;
		if(seg==PROFILE_SIZE-1) vFour	=&nodeBuffer[PFour2].vertice;  else vFour  =&nodeBuffer[PFour].vertice;
								vFive	=&nodeBuffer[PFive].vertice;
		if(seg==0)				vSix	=&nodeBuffer[PSix2].vertice;   else vSix   =&nodeBuffer[PSix].vertice;
		if(seg==0)				vSeven	=&nodeBuffer[PSeven2].vertice; else vSeven =&nodeBuffer[PSeven].vertice;
		

	MakeNormal(tOne, 	*vOne, 	*vTwo, 	 *vThree);
	MakeNormal(tTwo, 	*vOne, 	*vThree, *vFour);
	MakeNormal(tThree,	*vOne,	*vFour,	 *vFive);
	MakeNormal(tFour,	*vOne,	*vFive,	 *vSix);
	MakeNormal(tFive, 	*vOne,	*vSix,	 *vSeven);
	MakeNormal(tSix,	*vOne,	*vSeven, *vTwo);

	
	// Next is the Average of the Six Normals for the Triangles around the Vertice.
	for(i=0;i<3;i++) n[i]=(tOne[i]+tTwo[i]+tThree[i]+tFour[i]+tFive[i]+tSix[i])/6;
}


void CViewpipe::MakeNormal(D3DXVECTOR3 &dst, const D3DXVECTOR3 a, const D3DXVECTOR3 b, const D3DXVECTOR3 c )
{
  D3DXVECTOR3 ab=b-a;
  D3DXVECTOR3 ac=c-a;
  D3DXVec3Cross( &dst,&ab, &ac);
  D3DXVec3Normalize(&dst,&dst);
}

void CViewpipe::LoadPipeTextures(void)
{
	strcpy(textureList[0],"\\Pipe\\Blue_Lining.jpg");
	strcpy(textureList[1],"\\Pipe\\TextureConcrete2.jpg");
	strcpy(textureList[2],"\\Pipe\\TextureClay8in2.jpg");
	strcpy(textureList[3],"\\Pipe\\TextureClayRed8in2.jpg");

	char texFile[800];
	currentTexture=1;	// Default Texture to 1
	prevTexture=1;
//////////////////
// Load Texture //
//////////////////
	if(!OpenGL)
	{
		/////////////////////////////////// Loading Pipe Texture ///////////////////////
		strcpy(texFile,textureDirectory);
		strcat(texFile, textureList[currentTexture]);
		if(FAILED( D3DXCreateTextureFromFile(pd3dDevicePipe, texFile, &pTexturePipe) ))	;
//		PCN4240 MsgPipe("%s %s\n%s", languagePipe[15], // PCN2473 Language pointer (Antony, 11 May 2004)
//							 languagePipe[16], // PCN2473 Language pointer (Antony, 11 May 2004)
//							 texFile); //PCN2467 texFile Added (9 Dec 2003, Antony)

		/////////////////////////////////// Loading Plain White Texture ////////////////
		strcpy(texFile,textureDirectory);
		strcat(texFile,"\\white.jpg");
		if(FAILED( D3DXCreateTextureFromFile(pd3dDevicePipe, texFile, &pTextureWhite) ));	
//		PCN4240 MsgPipe("%s %s\n%s", languagePipe[15], // PCN2473 Language pointer (Antony, 11 May 2004)
//							 languagePipe[16], // PCN2473 Language pointer (Antony, 11 May 2004)
//							 texFile); //PCN2467 texFile Added (9 Dec 2003, Antony)

		/////////////////////////////////// Loading Laser Texture //////////////////////
		strcpy(texFile,textureDirectory);
		strcat(texFile,"\\laser.jpg");
		if(FAILED( D3DXCreateTextureFromFile(pd3dDevicePipe, texFile, &pTextureLaser) )) ;
//		PCN4240 MsgPipe("%s %s\n%s", languagePipe[15], // PCN2473 Language pointer (Antony, 11 May 2004)
//							 languagePipe[16], // PCN2473 Language pointer (Antony, 11 May 2004)
//							 texFile); //PCN2467 texFile Added (9 Dec 2003, Antony)

		/////////////////////////////////// Loading Water Texture PCN2347 //////////////
	//	strcpy(texFile,textureDirectory);
	//	strcat(texFile,"\\water.jpg");
	//	if(FAILED( D3DXCreateTextureFromFile(pd3dDevicePipe, texFile, &pTextureWater) )) 
	//	MsgPipe("Can't find Water Texture\n%s",texFile); //PCN2467 texFile Added (9 Dec 2003, Antony)
	}
}


//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Name: ShadePipeType PCN2337 4 November 2003
// Created By: Antony van Iersel
// 
// Description:
//		Input- type, set the shade type 0 None, 1 Capactiy, 2 Ovality, 3 Delta
//		
//		Set the new shade type
//
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


void CViewpipe::ShadePipeType(int type)
{
	if(!shadeOn) 
		{
		currentTexture=prevTexture; // Make the texture what it was before colour limiting was turned on
		ChangePipeTexture(textureList[currentTexture]);
		shadeChange++;	// There is a change in the shadeType, this is to make sure the
		shadeType=0;    // the pipe is repainted, and set shade type to none
		}
	else {
		prevTexture=currentTexture;		// Remeber what the current texture is to put it back when
		ChangePipeTexture("\\white.jpg"); // shade was off, and make the pipe white to see colours 
		//PCN2510 was  , ("white.jpg") the \\ was missing. There was a change in directory structure, this one got missed.(Antony van Iersel, 23 December 2003) 

		shadeType=type; // There is a change in the shadeType, this is to make sure the
		shadeChange++;  // the pipe is repainted, and set shade type input type
		}
}

//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Name: SetShade PCN2337 4 November 2003
// Created By: Antony van Iersel
// 
// Description:
//		Sets the new shade Limits, these values come from VB PVGraphtype
//		
// Input: 
//		type, and uper and lower limts, Same limits as VB PVGraph
//		
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


void CViewpipe::SetShade(int type, double minLimit, double maxLimit)
{
	if(type==1) 
		{
		shadeCapacityMinLimit= minLimit;
		shadeCapacityMaxLimit= maxLimit;
		}
	if(type==2) 
		{
		shadeOvalityMinLimit= maxLimit;
		shadeOvalityMaxLimit= maxLimit;
		}
	if(type==3) 
		{
		shadeDeltaMinLimit= minLimit;
		shadeDeltaMaxLimit= maxLimit;
//		MsgPipe(" shadeDeltaMinLimit = %f, shadeDeltaMaxLimit = %f ",minLimit, maxLimit);
		}

	// PCN3825 If shade is off. Then just set the limits (Above) and trick it into thinking the
	// previous type was new type. This way when you turn the colour on it will  take on the type
	// of the last PVGraph type clicked. (Antony van Iersel 14 June 2004)
	if(!shadeOn) 
		{
		shadeTypePrev=type;
		return;
		}
	ShadePipeType(type);
}

//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Name: TogglePipeShade PCN2337 4 November 2003
// Created By: Antony van Iersel
// 
// Description:
//		Toggles the Pipe Limits Colourisation, form of to on and back agian
//		
// Input: 
//		None
//		
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


void CViewpipe::TogglePipeShade(void)
{
	if(shadeOn)
		{
		shadeOn=false;
		shadeTypePrev=shadeType;	// If Turning off, remember what type was being viewed
		ShadePipeType(0);
		}
	else{
		shadeOn=true;
		ShadePipeType(shadeTypePrev); // If Turning on, place back the last type being viewed
		}
}

//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Name: ExportPipeSTL PCN2376 14 November 2003
// Created By: Antony van Iersel
// 
// Description:	Exports Pipe 3D Data as a .stl format file that 
//				Cad and Modeling Programs can import.
// Input: None
// Output: Exported File. No Function Return.
//		
//	File Format:
//		................Header.........................................
//		80 bytes - Any Text Such as the Creator's Name
//		4  bytes - int equal to the number of facets (triagles) in file
//		................Data   ........................................
//		................Facet 1........................................
//		4  bytes - float normal  x (facet normal, not vertice)
//		4  bytes - float normal  y ("")
//      4  bytes - float normal  z ("")
//		4  bytes - float vertex1 x (first vertice of facet x cord)
//		4  bytes - float vertex1 y (first vertice of facet y cord)
//		4  bytes - float vertex1 z (first vertice of facet z cord)
//		4  bytes - float vertex2 x (secound vertice of facet x cord)
//		4  bytes - float vertex2 y (secound vertice of facet y cord)
//		4  bytes - float vertex2 z (secound vertice of facet z cord)
//		4  bytes - float vertex3 x (third vertice of facet x cord)
//		4  bytes - float vertex3 y (third vertice of facet y cord)
//		4  bytes - float vertex3 z (third vertice of facet z cord)
//		2  bytes - unused (padding to bake 50-bytes)
//
//		................Data   ........................................
//		................Facet 2........................................
//		The same as above for Every Facet
//	
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


void CViewpipe::ExportPipeSTL(char *filename)
{
	long total=0;
	long fraction;
	long sec;	// Run thru all the sections of the pipe.
	char headText[800]="Exported from ClearLine Profiler - www.cleanflowsystems.com"; // This goes in the any text
	nodeStruct *p_vbIndex;		// Tempory pointer to access Vertex Buffer Data from pipe sections
	WORD *p_ibIndex;			// Tempory pointer to access Index Buffer Data from pipe sections
	DWORD numberTriangles;		// Number of triangles per section to export
	facetDataStruct facetData;	// A structure to fill out to make file writing easier
	DWORD i, count;				
	D3DXVECTOR3 norm;			// normal that is going to be calculated for 
								// the facet (triangle) normal
	insideOut=false;	// Choose which way the facets are exported, clockwise or anti-clockwise
		
	HANDLE hFile = CreateFile(filename,
							  GENERIC_WRITE,
							  FILE_SHARE_WRITE,
							  NULL,
							  CREATE_ALWAYS,
							  FILE_ATTRIBUTE_NORMAL,
							  NULL);
	if (hFile == INVALID_HANDLE_VALUE) {MsgPipe("%s",languagePipe[23].text); return;}


	// Number of traingles to export is numberTriangles Note: not counting the last pipe section.
	// Then add however many frames flow over to the last pipe section.

	numberTriangles=numberTrianglesSection*(numberSections-1); // Total number of triangles to export (not last section)
	fraction = frames-((SECTION_SIZE-1)*(numberSections-1));   // Number of frames in last Section
	numberTriangles = numberTriangles + (fraction*360);	// Add the two together to get real total

	pipeSectionStruct *currentSection;	// Current section of pipe to export

	DWORD dwCnt; // Number of bytes writen to file
	WriteFile(hFile,(char*) headText,		  80, &dwCnt,NULL); // 80 bytes in the header
	WriteFile(hFile,(char*) &numberTriangles, 4,  &dwCnt,NULL); // 4 bytes to store number of triangles exported

	if(exportingPanel!=NULL) exportingPanel->UpdateText(2,(int) frames-1);	// Display how many frames are to be exported
	UnloadPipe();	// Unload the pipe as not to run out of memory as the pipe is being exported
					// just load and unload a section at a time

	/////////////////////////////////////////////////////////////////////////////////////////////
	
	for(sec=0;sec<numberSections;sec++) // Run through all the pipe sections
		{
		if(sectionIndex[sec]==NULL) // If current section not loaded then load it into memory
			{						// and if not first section unload previous section from memory
			LoadNewSection(sec*(SECTION_SIZE-1));
			if(sec>0) UnloadPipeSection(sectionIndex[sec-1]);
			}
		currentSection=sectionIndex[sec]; // Point the current section to newly loaded section
		count=0;	//
		
		if(currentSection==NULL) { MsgPipe("%s",languagePipe[24].text); break; }

		// ..... unlock both the vertex and index buffer to gain access to the data ....//
		currentSection->pIBPipeSection->Lock(0,0,(void **) &p_ibIndex, 0);			    //
		currentSection->pVBPipeSection->Lock(0,0,(void **) &p_vbIndex, 0);				//
		//////////////////////////////////////////////////////////////////////////////////

		// Start the loop to export the pipe section..........
		for(i=0;i<(DWORD) numberTrianglesSection;i++)
			{											  // PCN2414 (20 Nov 2003)
			if(GetAsyncKeyState (VK_ESCAPE) & 0x8000)     // If the Escape is pressed
				{										  // unlock the Buffers,
				currentSection->pVBPipeSection->Unlock(); // close the file and
				currentSection->pIBPipeSection->Unlock(); // exit the export function
				CloseHandle(hFile);						  //
				return;									  //
				}

			// If the last frame is exported and is midway up the section, finnish the export
			if(((i/360)+currentSection->frameSectionStart)>(DWORD) frames-1) break; 

			if((i%360)==0)	// Every 360th triangle, one frame, update the exporting panel
				if(exportingPanel!=NULL)
					{			//	
					exportingPanel->UpdateText(1,(int) ((i/360)+currentSection->frameSectionStart)); 
					RenderD3D();// 
					}			//
			
			// Fill the the facet Structure with the first triangle vertices //
			facetData.vertex1x=p_vbIndex[p_ibIndex[count]].vertice.x;		 //
			facetData.vertex1y=p_vbIndex[p_ibIndex[count]].vertice.y;	     //
			facetData.vertex1z=p_vbIndex[p_ibIndex[count]].vertice.z-(currentSection->frameSectionStart*depthScale);
			norm = p_vbIndex[p_ibIndex[count++]].normal;					 //
			///////////////////////////////////////////////////////////////////
						  
			if(insideOut) // Writes triangle vertices in 1 - 2 - 3 order
				{		  // If insideOut, then triangles go Clockwise for culling,
						  // this stops the outer walls from being drawn when culled.
				// Secound Vertice /////////////////////////////////////////
				facetData.vertex2x=p_vbIndex[p_ibIndex[count]].vertice.x; //
				facetData.vertex2y=p_vbIndex[p_ibIndex[count]].vertice.y; // Below, the frame depth scale is exported
				facetData.vertex2z=p_vbIndex[p_ibIndex[count]].vertice.z-(currentSection->frameSectionStart*depthScale);
				norm += p_vbIndex[p_ibIndex[count++]].normal;			  //
				////////////////////////////////////////////////////////////

				// Third and last Vertice in the facet (Triangle) //////////
				facetData.vertex3x=p_vbIndex[p_ibIndex[count]].vertice.x; //
				facetData.vertex3y=p_vbIndex[p_ibIndex[count]].vertice.y; // Below, the frame depth scale is exported
				facetData.vertex3z=p_vbIndex[p_ibIndex[count]].vertice.z-(currentSection->frameSectionStart*depthScale);
				////////////////////////////////////////////////////////////

				norm += p_vbIndex[p_ibIndex[count++]].normal; // Get the average from the
				norm/=3;									  // three normals form the pipes
				D3DXVec3Normalize(&norm, &norm);			  // triangle then use that for the face normal
				}											  // for the export facet
						
			else		// Writes triangles vertices in 1 - 3 - 2 order
				{		// else not insideOut, then triangles go Counter Clockwise for culling
						// this stops the inside walls from being drawn when culled.
				// Secound Vertice in the facet (Triangle) /////////////////
				facetData.vertex3x=p_vbIndex[p_ibIndex[count]].vertice.x; //
				facetData.vertex3y=p_vbIndex[p_ibIndex[count]].vertice.y; // Below, the frame depth scale is exported
				facetData.vertex3z=p_vbIndex[p_ibIndex[count]].vertice.z-(currentSection->frameSectionStart*depthScale);
				norm += p_vbIndex[p_ibIndex[count++]].normal;			  //
				////////////////////////////////////////////////////////////

				// Third and last Vertice in the facet (Triangle) //////////
				facetData.vertex2x=p_vbIndex[p_ibIndex[count]].vertice.x; //
				facetData.vertex2y=p_vbIndex[p_ibIndex[count]].vertice.y; // Below, the frame depth scale is exported
				facetData.vertex2z=p_vbIndex[p_ibIndex[count]].vertice.z-(currentSection->frameSectionStart*depthScale);
				norm += p_vbIndex[p_ibIndex[count++]].normal;			  //
				////////////////////////////////////////////////////////////

				norm/=-3;	// -ve so to invert the normals, Same average as above but invert
				
				D3DXVec3Normalize(&norm, &norm); // the normal for inside out view			
				}
		
			facetData.normalx=norm.x; // Fill the facet face Normae
			facetData.normaly=norm.y; // for export
			facetData.normalz=norm.z; //
			facetData.padding=0; // Padding to make sure there is 50 bytes in the Structure
			
			// Writes the facet data structure to file
			total++;
			WriteFile(hFile,(char*) &facetData, 50, &dwCnt, NULL);
			}
		currentSection->pVBPipeSection->Unlock();	// Unlock the buffers ready for the
		currentSection->pIBPipeSection->Unlock();   // the next section of pipe
		}
	////////////////////////////////////////////////////////////////////////////////////////
	
	CloseHandle(hFile);

	// And we are all done, Note: Exported files are rather large, 30 thousand frames was
	// half a gigabyte
}

//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Name: CreateExportingPanel PCN2376 14 November 2003
// Created By: Antony van Iersel
// 
// Description:	Creates the panel to display while exporting
//
// 1 entry is "Exorting frames......"
// 2 entry is number of frames curently exorted
// 3 entry is number of frames out of
// 4 entry is "out of"
// 5 entry is "Esc to cancel"
//
// Input: None
// Output: None
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

void CViewpipe::CreateExportingPanel(void)
{
	exportingPanel = new D3DFont(pd3dDevicePipe, 
								 textureDirectory,
								 "\\Panel.jpg",
								 "\\Boarder.jpg",
								 (D3DFont::languageText *) languagePipe); //PCN2473 Language pointer (Antony, 11 May 2004)
	exportingPanel->panelDim.x			   = 275; exportingPanel->panelDim.y			  = 155;
	exportingPanel->panelDim.width		   = 145; exportingPanel->panelDim.height		  = 65;
	exportingPanel->panelDim.boarderHeight = 1;	  exportingPanel->panelDim.boarderWidth = 1;
	exportingPanel->panelDim.colour        = D3DCOLOR_RGBA(255,255,255,90);
	exportingPanel->panelDim.boarderColour = D3DCOLOR_RGBA(255,255,255,255);
	exportingPanel->fontPosition.top    = (long) exportingPanel->panelDim.y;
	exportingPanel->fontPosition.left   = (long) exportingPanel->panelDim.x;
	exportingPanel->fontPosition.right  = (long) exportingPanel->panelDim.x + (long) exportingPanel->panelDim.width-3;
	exportingPanel->fontPosition.bottom = (long) exportingPanel->panelDim.y + (long) exportingPanel->panelDim.height;

	exportingPanel->SetPanel();
	exportingPanel->NewText("Exporting frames.....",D3DXVECTOR2((float) exportingPanel->fontPosition.left+4 ,(float) exportingPanel->fontPosition.top));
	exportingPanel->NewText(0				  ,D3DXVECTOR2((float) exportingPanel->fontPosition.left+4 ,	 (float) exportingPanel->fontPosition.top+20),DT_LEFT);
	exportingPanel->NewText(0				  ,D3DXVECTOR2((float) exportingPanel->fontPosition.left+50 ,	 (float) exportingPanel->fontPosition.top+20));
	exportingPanel->NewText("out of"		  ,D3DXVECTOR2((float) exportingPanel->fontPosition.left+50,	 (float) exportingPanel->fontPosition.top+20));
	exportingPanel->NewText("ESC to cancel"   ,D3DXVECTOR2((float) exportingPanel->fontPosition.left+25,     (float) exportingPanel->fontPosition.top+40));	
	exportingPanel->InitGeometry();
}

void CViewpipe::OpenGLDraw(void)
{
	long startFrame, endFrame;
	long currentFrame;

	


	glColor3f(1.0f,1.0f,1.0f);		
//	glBegin(GL_QUAD_STRIP);					// Start Drawing The Pyramid
	startFrame = laserFocus - 180; if(startFrame<1) startFrame = 1;
	endFrame = laserFocus + 180; if(endFrame>frames-1) endFrame = frames-1;

	threeDFrameIndex = 0;
	GetThreeDFrame(startFrame,0);
	GetThreeDFrame(startFrame+1,1);
	currentFrame=startFrame+2;
	
	for(;currentFrame<endFrame-1;currentFrame++)
	{
		RenderRing();
		threeDFrameIndex=(threeDFrameIndex+1)%2;
		GetThreeDFrame(currentFrame,threeDFrameIndex);
	}

	

/*
	GetThreeDFrame(2,threeDFrameIndex);
	threeDFrameIndex=(threeDFrameIndex+1)%2;
	GetThreeDFrame(3,threeDFrameIndex);
	
	for(currentFrame = 4;currentFrame<20;currentFrame++)
	{
		RenderRing();
		threeDFrameIndex=(threeDFrameIndex+1)%2;
		GetThreeDFrame(currentFrame,threeDFrameIndex);
	}
*/	



}
void CViewpipe::RenderRing(void)
{
	int firstFrame;
	int secondFrame;
	int i;

	firstFrame = threeDFrameIndex;
	secondFrame = (threeDFrameIndex+1)%2;

	glBegin(GL_QUADS);


	if(threeDRing[firstFrame][179].vertice.x!=0 && threeDRing[firstFrame][179].vertice.y!=0 &&
		   threeDRing[firstFrame][0].vertice.x!=0 && threeDRing[firstFrame][0].vertice.y!=0 &&
		   threeDRing[secondFrame][179].vertice.x!=0 && threeDRing[secondFrame][179].vertice.y!=0 &&
		   threeDRing[secondFrame][0].vertice.x!=0 && threeDRing[secondFrame][0].vertice.y!=0)
		{
			DrawQuad(threeDRing[firstFrame][179],
					 threeDRing[firstFrame][0],
					 threeDRing[secondFrame][0],
					 threeDRing[secondFrame][179]);
		}


		for(i=0;i<179;i++)
		{
			
			if(threeDRing[firstFrame][i].vertice.x!=0 && threeDRing[firstFrame][i].vertice.y!=0 &&
			   threeDRing[firstFrame][i+1].vertice.x!=0 && threeDRing[firstFrame][i+1].vertice.y!=0 &&
			   threeDRing[secondFrame][i].vertice.x!=0 && threeDRing[secondFrame][i].vertice.y!=0 &&
			   threeDRing[secondFrame][i+1].vertice.x!=0 && threeDRing[secondFrame][i+1].vertice.y!=0)
			{
				DrawQuad(threeDRing[firstFrame][i],
						 threeDRing[firstFrame][i+1],
						 threeDRing[secondFrame][i+1],
						 threeDRing[secondFrame][i]);
			}
		}

	glEnd();

	

	

}

void CViewpipe::DrawQuad(nodeStruct A, nodeStruct B, nodeStruct C, nodeStruct D)
{
		glColor4ubv((unsigned char *) &A.colour); glVertex3fv((float *) &A.vertice);
		glColor4ubv((unsigned char *) &A.colour); glVertex3fv((float *) &B.vertice);
		glColor4ubv((unsigned char *) &A.colour); glVertex3fv((float *) &C.vertice);
		glColor4ubv((unsigned char *) &A.colour); glVertex3fv((float *) &D.vertice);

}

void CViewpipe::Unit(float &x, float &y, float &z)
{
	float length;

	length = (float) sqrt((x * x) +
						  (y * y) +
						  (z * z));
	if(length==0) length=1;
	x/=length;
	y/=length;
	z/=length;
}


//	struct nodeStruct {
//		D3DXVECTOR3 vertice;
//		D3DXVECTOR3 normal;
//		DWORD colour;
//		FLOAT t1u,t1v;
//		FLOAT t2u,t2v;
//		};

void CViewpipe::GetThreeDFrame(long frameNumber, int ring)
{
	long profileNode;
	long colourNode;
	int i;

	float x, y;
	float centreX,centreY;

	profileNode = 180 * frameNumber;
	colourNode =  181 * frameNumber;
	centreX = rawCentrex[frameNumber];
	centreY = rawCentrey[frameNumber];
	//centreX = 0;
	//centreY = 0;

	for(i=0;i<180;i++)
	{
		x = rawDatax[profileNode];
		y = rawDatay[profileNode];

		//x = sin((double) i/90*PI)*5;
		//y = cos((double) i/90*PI)*5;

		if(x!=0 || y!=0)
		{
			if(x>10000) x-=20000;
			else if(x<-10000) x+=20000;
			x+=centreX;
			y+=centreY;
			threeDRing[ring][i].colour = D3DCOLOR_RGBA(rawColourBlue[colourNode],
													   rawColourGreen[colourNode],
													   rawColourRed[colourNode],1);
			threeDRing[ring][i].vertice = D3DXVECTOR3(x*20,y*20,(float) frameNumber*2);
		}
		else
		{
			threeDRing[ring][i].colour = 0;
			threeDRing[ring][i].vertice = D3DXVECTOR3(0,0,frameNumber*2);
		}
		profileNode++;
		colourNode++;
	}
//				glColor3f( (float) (rawColourRed[colourNode])/255, 
//						   (float) (rawColourGreen[colourNode])/255, 
//						   (float) (rawColourBlue[colourNode])/255);
//				glNormal3f(norX,norY,norZ);
//
//				glVertex3f( x1, y1, (float) (frameNo*-2));			// Top Of Triangle (Front)
//				glVertex3f( x2, y2, (float) ((frameNo+1)*-2));

}
