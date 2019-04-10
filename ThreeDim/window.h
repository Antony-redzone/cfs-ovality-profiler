

//#include "D3DFont.h"
#include "Camera.h"
#include "pipe.h"
#include "scape.h"
#include "D3DObject.h" //PCN2488 (Antony van Iersel, 9 Jan 2003)

#include "Arrows.h"
#include "Lights.h"
#include <d3dx9.h>

struct languageText{char text[101];}; //PCN2473 Antony van Iesel

class window{
public:
	HWND  hWnd;
	HDC hDC;
	HGLRC hglrc;
	HPALETTE hPalette;

	void show();
	void hide();
	void destroy();
};



class D3Dwindow:public window{
public:
	PIXELFORMATDESCRIPTOR pfd;


	D3Dwindow(HWND passedHWnd, int vertexMode, 
		float *datax, //PCN2988 pass x coordiante data
		float *datay, //PCN2988 pass y coordiante data
		float *centrex,
		float *centrey,
		long numdata,
		char *path,
		int *pvColourRed,
		int *pvColourGreen,
		int *pvColourBlue,
		double expectedRad,
		int pVCalcMult, //PCN2860 Multiplier Passed to allow greater presision when displaying capacity, ovality and delta
		int pVDataXYMult, //PCN2988 13 Sept 2004 its a devision needed for the XY data
		int units); //PCN3111 0 = metric, 1 = imperial
	~D3Dwindow();

	LPDIRECT3D9		  g_pD3D;
	LPDIRECT3DDEVICE9 g_pd3dDevice;

	float aspectRatio;
	bool deviceCreated;
	bool pipeLoaded;
	int pVCalculationsMultiplier;
	bool OpenGL; //deciding to render in opengl or d3d.

////////////////////////////
// PCN2473
	languageText *languageArray;
////////////////////////////

	CViewpipe *p;
	CViewpipe *p2;

	Scape *landScape;
	Camera *railCam;
	Camera *pipeCam;
	Camera *freeCam;
	Camera *mapCam;
	Camera *cam;

	D3DFont *headsUp;
	D3DFont *loadingPanel;
	D3DFont *repaintingPanel;
	D3DFont *debugPanel;

	D3DObject *dirArrow; //PCN2488 (Antony van Iersel, 9 Jan 2003)
//	D3DObject *dirX, *dirY, *dirZ;
	
//	Arrow *targetArrow; ANT22march2004
	Lights light1;
	Lights light2;

	char textureDir[800];
	char installationDir[800];

	int camSelect;

	int leftX, leftY;
	int rightX, rightY;
	bool rightBut;
	bool leftBut;
	int speed;
	bool drawPipeTrueFalse;
	bool drawScapeTrueFalse;
	bool drawWaterTrueFalse;
	bool drawArrowTrueFalse;
	bool lineView;
	int focus;
	bool roll;
	int laserDistance;
	long watl,watr;

	HRESULT InitD3D(int vertexMode);
	void SetupLights();
	void SetupMaterial(void);
	void SetupMatrices(void);
	void ResetCameras(void);
	
	void RenderD3D(void);
	void DrawDirectionArrow(void);
	void MouseMotion(int x, int y);
	void AllStop(void);
	void Keyboard(int value);
	void SelectCamera(int c);
	void UpdateCameraObjects();
	void Motion(void);
	void MoveLaserTo(int frame);
	
	void UnloadPipe(void);
	void WindowToBmp (char *name, HWND hWndCapture);
	void InitialiseLanguageArray(void);
	void CreateHeadsUpPanel(void);
	void CreateLoadingUnloadingPanel(void);
	void CreateRepaintingPanel(void);
	void CreateDebugPanel(void);
	void UpdatePanelText(void);
	void LoadDirectionArrow(void);
	void DebugVariables(void);
	void ZoomLaserDistance(int dis);
	void PipeDepthScale(float scale);
	void Zoom(int speed);
	void ExportPipe(char *filename);	// PCN2376

	// OpenGL mostly from here //
	void OpenGLWindowInit(void);
	void GLInit(void);
	void ResiseGLWindow(int width, int height);
	void RenderOpenGL(void);



    friend LRESULT WINAPI MainWndProc(HWND  hWnd,UINT  message, WPARAM  wParam, LPARAM  lParam);
private:

};