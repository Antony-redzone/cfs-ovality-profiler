#include <d3dx9.h>
#include "D3DFont.h"


#define PROFILE_SIZE 180 // Number of Points per Profile

 // PCN2473 , Antony 11 May 2004, copy of language pointer from window to acces language array.

class CViewpipe {
	
public:
	struct languageText{char text[101];}; //PCN2473 Antony van Iesel
	CViewpipe(bool _OpenGL, //to render in opengl or d3d
			  float *datax, //PCN2988 Pass the x coordinate data
			  float *datay, //PCN2988 Pass the y coordinate data
			  float *centrex,
			  float *centrey,
			  long dataSize, // PCN2453 startf, endf changed to datasize (loading many times crash)
			  LPDIRECT3DDEVICE9 pd3dDevice,char *path,
			  int *pvColourRed,
			  int *pvColourGreen,
			  int *pvColourBlue,
			  double expectedRad, //PCN2693 Needed for colour calculations (Antony van Iersel, 15 March 2004)
			  languageText *language,//PCN2473 Language pointer (Antony, 11 May 2004)
			  int  pVCalcMult, //PCN2693 (9 August 2004, Antony) needed divide by multiplier to get real values
			  int pVDataXYMult, //PCN2988 13 Sept 2004 its a devision needed for the XY data
			  int unitsPassed); //PCN3111 0 = metric , 1 = imperial
				
	~CViewpipe();						 //destructor, deallocate the memory
	bool OpenGL; //To render in open gl or d3d
	languageText *languagePipe;  //PCN2473 Language pointer (Antony, 11 May 2004)
	int test1, test2;

//Antony's stuff
	bool scankeyTable[128]; // There are 128 max scancodes on a keyboard

	struct nodeStruct {
		D3DXVECTOR3 vertice;
		D3DXVECTOR3 normal;
		DWORD colour;
		FLOAT t1u,t1v;
		FLOAT t2u,t2v;
		};

	struct pipeSectionStruct 
		{
		long shade;
		pipeSectionStruct *next;
		pipeSectionStruct *prev;
		LPDIRECT3DDEVICE9 pd3dDevicePipe;
		LPDIRECT3DVERTEXBUFFER9 pVBPipeSection;
		LPDIRECT3DINDEXBUFFER9  pIBPipeSection;
//		LPDIRECT3DVERTEXBUFFER9 pVBWaterSection; //PCN2347 removed (Antony van Iersel, 19 March 2004)
		long frameSectionStart;
		long frameSectionEnd;
		};

	struct facetDataStruct 
		{
		float normalx;
		float normaly;
		float normalz;
		float vertex2x, vertex2y, vertex2z;
		float vertex1x, vertex1y, vertex1z;
		float vertex3x, vertex3y, vertex3z;
		WORD padding;
		};

	char insideOut;
	double pVCalculationsMultiplier; //PCN2693 (9 August 2004, Antony) needed divide by multiplier to get real values
	double pVDataXYMultiplier; //PCN2988 add a multiplier for the XY data.
	int units; //PCN3111 units of mesearments 0 = metric, 1 = imperial. Inches will be converted to mm when
			   //        when calculating the radius values
	pipeSectionStruct *startSection;
	pipeSectionStruct *endSection;
	pipeSectionStruct **sectionIndex;

	LPDIRECT3DDEVICE9		pd3dDevicePipe;
	LPDIRECT3DTEXTURE9		pTexturePipe;
	LPDIRECT3DTEXTURE9		pTextureLaser;
	LPDIRECT3DTEXTURE9		pTextureWhite; 
//	LPDIRECT3DTEXTURE9		pTextureLaserLightMap;  //PCN2461
//	LPDIRECT3DTEXTURE9		pTextureWater;			//PCN2461 To be put back when using water
	LPDIRECT3DINDEXBUFFER9	pIBDetail0_1PipeSection;	// Every Secound Frame x Secound Vertex
	LPDIRECT3DINDEXBUFFER9	pIBDetail0_2PipeSection;	// Every Forth Frame   x Forth Vertex
	LPDIRECT3DINDEXBUFFER9	pIBDetail0_3PipeSection;	// Every Eigth Frame   x Eigth Vertex

	//////////////////////////////////////////////////// PCN2347
//PCN2461 will be put back when using water	// LPDIRECT3DINDEXBUFFER9	pIBDetail0_0WaterSection; // Every Frame
//PCN2461 will be put back when using water	// LPDIRECT3DINDEXBUFFER9	pIBDetail0_1WaterSection; // Every Secound Frame
//PCN2461 will be put back when using water	// LPDIRECT3DINDEXBUFFER9	pIBDetail0_2WaterSection; // Every Forth Frame
//PCN2461 will be put back when using water // LPDIRECT3DINDEXBUFFER9	pIBDetail0_3WaterSection; // Every Eigth Frame

	long numTriSecDetail0_1; 
	long numTriSecDetail0_2; 
	long numTriSecDetail0_3;

	long numPoiSecDetail0_1;
	long numPoiSecDetail0_2;
	long numPoiSecDetail0_3;

	long numberPointsSection;
	long numberTrianglesSection;

//	long numTriWaterSecDet0_0; // PCN2347 
//	long numTriWaterSecDet0_1; // PCN2347
//	long numTriWaterSecDet0_2; // PCN2347
//	long numTriWaterSecDet0_3; // PCN2347

//	long numPoiWaterSecDed0_0; // PCN2347
//	long numPoiWaterSecDed0_1; // PCN2347
//	long numPoiWaterSecDed0_2; // PCN2347
//	long numPoiWaterSecDed0_3; // PCN2347

	nodeStruct *nodeBuffer;
	double *rawData;
	float *rawDatax; //PCN2988 raw data x coordinate passed from VB
	float *rawDatay; //PCN2988 raw data y coordiante passed from VB
	float *rawCentrex;
	float *rawCentrey;
	int *rawColourRed;
	int *rawColourGreen;
	int *rawColourBlue;
	long numberSections;
	D3DFont *exportingPanel;
	
	bool exportingPipe;
	int percentLoaded;
	bool loadingPipe;
	bool repaintingPipe;
	long laserFocus;
	long highestPoint;
	D3DXVECTOR3 pipePosition;
	D3DXVECTOR3 laserRing[PROFILE_SIZE+1];
	D3DXVECTOR3 laserRingProj[PROFILE_SIZE+1];
	D3DXVECTOR3 topRing;
	D3DXVECTOR3 cameraPosition;
	int cameraType;
	D3DXVECTOR3 centerLaserRing;

	double *preCalcSin;
	double *preCalcCos;

	int rings;
	int frames;
	long noNodes;
	float twist;
	int waterTwist;
	long waterLevelLeft, waterLevelRight;

	long shadeChange;
	int shadeType;
	int shadeTypePrev;
	bool shadeOn;
	double shadeCapacityMinLimit, shadeCapacityMaxLimit;
	double shadeDeltaMinLimit,	 shadeDeltaMaxLimit;	 
	double shadeOvalityMinLimit,  shadeOvalityMaxLimit;

	DWORD colRed;
	DWORD colGreen;
	DWORD colBlue;
	DWORD colOrange;
	DWORD colPurple;
	DWORD colYellow;
	DWORD colWhite;
	DWORD colAqua;
	DWORD colClearBlack;
	DWORD colClearWhite;

	DWORD colDeltaBlue;
	DWORD colDeltaAqua;
	DWORD colDeltaGreen;
	DWORD colDeltaOrange;
	DWORD colDeltaWhite;

	
	
	double expRad;  //expected radius
	double avgRad;
	int laserSpeed;
	int laserWidth;

	float depthScale;
	bool flat;
	char pipeDisplayType;
	char pipeTexture[800];
	char textureDirectory[800];
	char textureList[10][800];
	int currentTexture;
	int prevTexture;

	double radius;

	void PreCalcCosSin(void);
	void LoadNewSection(long start);
	void FillPipeSectionVB(long start, LPDIRECT3DVERTEXBUFFER9 &lpVB);
	void FillWaterSectionVB(long start, LPDIRECT3DVERTEXBUFFER9 &lpVB, long left, long right);
	void FillPipeSectionIB(long start, LPDIRECT3DINDEXBUFFER9 &lpIB);
//	void FillWaterSectionIB(); //PCN2347 removed (Antony van Iersel, 19 March 2004)
	void ShadeSection(pipeSectionStruct *section);
	void SetVertexColours(nodeStruct *p_vertex, long secOffset, long frame, long profile );
	void RenderD3D(void);
	void ProcessPipeBuffer(long start);
	double CalcAvgRadius(void);
	void CreateIndexBufferLOD(void);
	void ChangePipeTexture(char *pTex);
	void SetVBLaserTexture(pipeSectionStruct *section, long fraction, float texPos);
	void NextPipeTexture(void);
	void InitIBDetail(char det123);
	void ClearLaser(long fr);
	void MarkPipe(long fr);
	void MoveLaserTo(int f);
	void MoveLaser(int dis);
	void ChangeLaserWidth(int width);
	void UnloadPipe(void);
	void UnloadPipeSection(pipeSectionStruct *section);
	void FindHighestPoint(void);
	void SetupMatrices(void);
	void MakeNormal(D3DXVECTOR3 &dst, const D3DXVECTOR3 a, const D3DXVECTOR3 b, const D3DXVECTOR3 c );
	void AveSixTriangles(D3DXVECTOR3 &n, long ring, long seg);
	void PipeDraw(void);
	void LoadPipeTextures(void);
	void ShadePipeType(int type);
	void SetShade(int type, double minLimit, double maxLimit);
	void TogglePipeShade(void);
	void ExportPipeSTL(char *filename); // PCN2376
	void CreateExportingPanel(void); // PCN2376
private:
	nodeStruct threeDRing[2][180];
	int threeDFrameIndex;
	
	void ConvertXYtoRadiusRawData(void);
	void OpenGLDraw(void);
	void Unit(float &x, float &y, float &z);
	void GetThreeDFrame(long frameNumber, int ring);
	void RenderRing(void);
	void DrawQuad(nodeStruct A, nodeStruct B, nodeStruct C, nodeStruct D);
};