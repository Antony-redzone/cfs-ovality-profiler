

#include "window.h"
#include <atlbase.h>
#include <stdio.h>
#include <iostream.h>
#include <gl\gl.h>
#include <gl\glu.h>
#include <gl\glaux.h>

#include "CBSalgebra.h"

float  whiteLight[] = 	{ 0, 0, 0, 1.0 };
float  sourceLight[] = { 1.0, 1.0, 1.0, 1.0 };
float lightpos[]= {0.1f, -0.1f, 1.0f, 0.0f};

#define PIPE_SECTION_SIZE 129 // Number Frames per Pipe Section, needs to be a power of two + 1, also defined in pipe.cpp as SECTION_SIZE
int *testDataX;
int *testDataY;

void MsgWindow(TCHAR *szFormat, ...)
{
    TCHAR szBuffer[512];

    va_list pArgs;
    va_start(pArgs, szFormat);
    _vstprintf(szBuffer, szFormat, pArgs);
    va_end(pArgs);

    MessageBox(NULL, szBuffer, TEXT("LaserProfiler Message"), MB_OK | MB_ICONERROR);
}

//constructor, creates all objects within the window and the window itself
D3Dwindow::D3Dwindow(HWND passedHWnd, int vertexMode, //Bla
					 float *datax, //PCN2988 passing x coordinate data
					 float *datay, //PCN2988 passing y coordinate data
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
					 int units) //PCN3111 0 = metric, 1 = imperial
{
	OpenGL = false; //PCN OPENGL
//	numdata=450*180;
//	int count=0;
//	double radStep = (2*PI)/180;
//	double rad=0;
//	testDataX = new int[numdata];
//	testDataY = new int[numdata];
//	for(count=0;count<numdata;count++)
//	{
//		testDataX[count] = (int) (sin(rad) * 500);
//		testDataY[count] = (int) (cos(rad) * 500);
//		rad+=radStep;
//	}

//	datax=testDataX;
//	datay=testDataY;


	RECT viewWindow; int height, width;
	hWnd = passedHWnd;
	hDC = GetDC(hWnd);

	::GetClientRect(hWnd, &viewWindow);
	height=viewWindow.bottom-viewWindow.top;
	width=viewWindow.right-viewWindow.left;

	pVCalculationsMultiplier = pVCalcMult;
	if(height==0) height=1;
	aspectRatio = (float) width / (float) height;

	strcpy(textureDir, path);
	strcat(textureDir,"Textures");

	strcpy(installationDir, path);
	

	p			= NULL;// p2 = NULL;
	landScape	= NULL;
	headsUp		= NULL;
	railCam		= NULL;
	pipeCam		= NULL;
	freeCam		= NULL;
	mapCam		= NULL;
	loadingPanel	= NULL; //PCN2461 (8 December 2003, Antony)
	repaintingPanel	= NULL; //PCN2461 (8 December 2003, Antony)
	debugPanel		= NULL; //PCN2461 (8 December 2003, Antony)
	dirArrow	= NULL; //PCN2367 (Antony van Iersel, 9 Jan 2003)
	languageArray = NULL;

	languageArray = new languageText[30]; //PCN2473 Language Text Array (Antony 26 Feb 2004)
	InitialiseLanguageArray();
	
	watl=45;watr=135;

 	focus=1;
	laserDistance=200;
	g_pD3D = NULL;
	g_pd3dDevice = NULL;

	//PCN3122, used to check if it InitD3D return E_FIAL then completly fail the
	// initialising of the D3D, now we check if it succeded if not then move on
	// to the next vertexMode type, ps 0 = hardware, 1 = mixed, 2 = software
	if(!OpenGL)
	{
		if(vertexMode==0)
		{
			if(InitD3D(0)!=E_FAIL) deviceCreated=true;      // Try Hardware Vertex Processing if not then
			else if(InitD3D(1)!=E_FAIL) deviceCreated=true; // Try Mixed Vertex Processing if not then
			else if(InitD3D(2)!=E_FAIL) deviceCreated=true; // Try Software Vertex Processing in not then
			else {deviceCreated=false; return;}             // return at error, can't display 3D Pipe
		}
		else{
			if(InitD3D(2)!=E_FAIL) deviceCreated=true; // Try Software Vertex Processing in not then
			else {deviceCreated=false; return;}             // return at error, can't display 3D Pipe
		}
	}
	else
	{
		OpenGLWindowInit();
	}

	pipeLoaded        =false;
	drawPipeTrueFalse =true;
	drawScapeTrueFalse=true;
	drawWaterTrueFalse=true;
	drawArrowTrueFalse=true;
	lineView          =false;

	
	p = new CViewpipe(OpenGL,
					  datax, //PCN2988 pass x coordinate data
					  datay, //PCN2988 pass y coordinate data
					  centrex,
					  centrey,
					  numdata, g_pd3dDevice,textureDir,
					  pvColourRed,
					  pvColourGreen,
					  pvColourBlue,
					  expectedRad,
					  (CViewpipe::languageText *) languageArray,
					  pVCalculationsMultiplier,
					  pVDataXYMult,  //PCN2988 13 Sept 2004 its a devision needed for the XY data
					  units); //PCN3111 0 = metric, 1 = imperial
/*	p2 = new CViewpipe(data, numdata, g_pd3dDevice,textureDir,
					 capacity,
					 ovality,
					 deltaMax,
					 deltaMin,
					 expectedRad);
	p2->pipePosition.x=200;
*/
//	MsgWindow("about to create pipe");
	if(!OpenGL)
	{

	//////////////////////////////////////////////////////////PCN2367 (Antony van Iersel)
		dirArrow = new D3DObject(g_pd3dDevice,			// Device to be rendered to
								 textureDir,			// 3D Object directory, texture, model etc
								 "fat_arrow_rgb.x",//thin_arrow_xyz.x",    // 3D Object File Name
								 D3DXVECTOR3((float) -13,   (float) 112,   (float) -168),	// Rotation x,y,z
								 D3DXVECTOR3((float)-6.70, (float)-4.70,(float)10), // Position x,y,z
								 //D3DXVECTOR3((float)0, (float)0,(float)10), // Position x,y,z
								 D3DXVECTOR3((float)0.02,(float)0.02, (float)0.02), //  Scale x,y,z
								 0,						// Align roation to world (0)
								 1,                     // Align position to Camera (1) 
								 (D3DObject::languageText *) languageArray);  //PCN2473 Language pointer (Antony, 11 May 2004)

		//////////////////////////////////////////////

		landScape    = new Scape(g_pd3dDevice, textureDir, (Scape::languageText*) languageArray);

		CreateHeadsUpPanel();
		CreateLoadingUnloadingPanel();  // Create and fill in the Unloading and Loading Display Panel	
		CreateRepaintingPanel(); // Creates and fill Repainting Panel PCN2337
		CreateDebugPanel(); 
	}
	
	camSelect=1;
	railCam = new Camera();
	pipeCam = new Camera(); pipeCam->ZoomTarget(850); pipeCam->TiltTarget(45); pipeCam->RollTargetY(135);
	freeCam = new Camera(); freeCam->MoveTo(0,1440,1440); freeCam->Tilt(45);
	mapCam  = new Camera(); mapCam->MoveTo(0,25000,0); mapCam->Pan(180); mapCam->Tilt(90);
	cam=freeCam;

	light1.Tilt(D3DXToRadian(45));
	light2.Tilt(D3DXToRadian(-45));

//	MsgWindow("ProcessPipeData Function");

	if(landScape!=NULL)		landScape->InitGeometry();	//PCN2461 (8 December 2003, Antony)
	if(headsUp!=NULL)		headsUp->InitGeometry();	//PCN2461 (8 December 2003, Antony)

//	RenderD3D();
//	p->ProcessPipeData();
//Test	p->InitGeometry();

	if(p!=NULL) pipeLoaded=true; // PCN2461 (22 March 2004 , Antony van Iersel)
	if(p!=NULL && landScape!=NULL) landScape->offset.y=((float) p->avgRad)*10; //PCN2461 (8 December 2003, Antony)
	if(p!=NULL) pipeCam->ZoomTarget((float) p->avgRad*3); //PCN2461 (8 December 2003, Antony)
	
////////////////////////////

	int fr;
	int direction=-1;
	if(OpenGL) direction=1;


	if(p!=NULL) railCam->InitRail(p->frames-direction); //PCN2461 (8 December 2003, Antony)
	if(p!=NULL) pipeCam->InitRail(p->frames-direction); //PCN2461 (8 December 2003, Antony)
	
	if(p!=NULL) for(fr=0;fr<p->frames;fr++) //PCN2461 p!=NULL condition added (8 December 2003, Antony)
					{ 
					railCam->rail[fr] = D3DXVECTOR3( 0,0, (fr*direction)*p->depthScale);
					pipeCam->rail[fr] = D3DXVECTOR3( 0,0, (fr*direction)*p->depthScale);
					}
	
	railCam->MoveToRail(0); 
	pipeCam->MoveToRail(0);

///////////////////////////////
	SelectCamera(camSelect+1); // This replaces the following 8 lines.
	
	//	if(landScape!=NULL) //PCN2461 (8 December 2003, Antony)
	//	{
	//	if(camSelect==0) { cam=railCam; landScape->SetCam(cam); if(targetArrow!=NULL) targetArrow->SetCamera(cam);} //PCN2461
	//	if(camSelect==1) { cam=pipeCam; landScape->SetCam(cam); if(targetArrow!=NULL) targetArrow->SetCamera(cam);} //PCN2461
	//	if(camSelect==2) { cam=freeCam; landScape->SetCam(cam); if(targetArrow!=NULL) targetArrow->SetCamera(cam);} //PCN2461
	//	if(camSelect==3) { cam=mapCam;  landScape->SetCam(cam); if(targetArrow!=NULL) targetArrow->SetCamera(cam);} //PCN2461
	//	}
	speed=0;
}

D3Dwindow::~D3Dwindow()
{
	if(railCam!=NULL) { delete railCam; railCam = NULL; } //PCN3085
	if(mapCam!=NULL)  { delete mapCam; mapCam = NULL; } //PCN3085
	if(freeCam!=NULL) { delete freeCam; freeCam = NULL; } //PCN3085
	if(pipeCam!=NULL) { delete pipeCam; pipeCam = NULL; } //PCN3085
	if(p!=NULL) { delete p;	p = NULL; } //PCN3085

//	if(targetArrow!=NULL) delete targetArrow; ANT22march2004
	if(landScape  !=NULL) { delete landScape; landScape = NULL; } //PCN3085
	if(headsUp    !=NULL) { delete headsUp; headsUp = NULL; } //PCN3085
	if(loadingPanel!= NULL)		{ delete loadingPanel; loadingPanel = NULL; }  //PCN2461 (8 December 2003, Antony)
	if(repaintingPanel!= NULL)	{ delete repaintingPanel; repaintingPanel = NULL; }//PCN2461 (8 December 2003, Antony)
	if(debugPanel!= NULL)	{ delete debugPanel; debugPanel = NULL; }	   //PCN2461 (8 December 2003, Antony)
	if(dirArrow!=NULL) { delete  dirArrow; dirArrow = NULL; }  //PCN2563 (26 Feb 2004, Antony)
	if(languageArray!=NULL) { delete[] languageArray; languageArray = NULL; } //PCN2473 Language Text Array (Antony 26 Feb 2004) PCN3085 [] added
	
	if(g_pd3dDevice != NULL) g_pd3dDevice->Release();
	if(g_pD3D != NULL) g_pD3D->Release();
	
}

void D3Dwindow::UnloadPipe(void)
{
	if(p!=NULL) { delete p; p = NULL; }
}

// vertexMode -- 0 = Hardware Vertexing
//            -- 1 = Mixed, Hardware and Software Vertexing
//			  -- 2 = Software Vertexing

HRESULT D3Dwindow::InitD3D(int vertexMode)
	{
	D3DDISPLAYMODE display_mode;
	HRESULT hr;

	//Create the D3D object, which is needed to create the D3DDevice.
    g_pD3D = Direct3DCreate9( D3D_SDK_VERSION );
    g_pD3D->GetAdapterDisplayMode(D3DADAPTER_DEFAULT,&display_mode);
    
	// Set up the structure used to create the D3DDevice. Most parameters are
	// zeroed out. We set Windowed to TRUE, since we want to do D3D in a
	// window, and then set the SwapEffect to "discard", which is the most
	// efficient method of presenting the back buffer to the display. And
	// we request a back buffer format that matches the current desktop display
	// format.
	

	D3DPRESENT_PARAMETERS d3dpp;
	
	ZeroMemory( &d3dpp, sizeof(d3dpp));
	d3dpp.Windowed = TRUE;
	d3dpp.SwapEffect = D3DSWAPEFFECT_DISCARD;
	d3dpp.BackBufferFormat = D3DFMT_UNKNOWN;
    d3dpp.EnableAutoDepthStencil = TRUE;
    d3dpp.AutoDepthStencilFormat = D3DFMT_D16;

	// Create the Direct3D Device. Here we are using the default adapter (most
	// systems only have one, unless they have multiple graphics hardware cards
	// installed) and requesting the HAL (which is saying we want the hardware
	// device rather than a software one). Software vertex processing is
	// specified since we know it will work on all cards. On cards that support
	// hardware vertex processing, though, we would see a big performance gain
	// by specifying hardware vertex processing.

	
	// Set up Graphics Device to Run in Hardware Vertexing Mode.
//	vertexMode=0;

	if(vertexMode==0) 
		{
		hr=g_pD3D->CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, 
											  (HWND) hWnd, D3DCREATE_HARDWARE_VERTEXPROCESSING,
											  &d3dpp, &g_pd3dDevice);
		
		if((hr==D3DERR_INVALIDCALL) || (hr==D3DERR_NOTAVAILABLE)) 
			{
			//PCN3112, no longer need to inform user, will just try the next vertex type
			//MsgWindow("%s\n%s\n%s",
			//	languageArray[7].text,
			//	languageArray[8].text,
			//	languageArray[9].text);

			if(g_pd3dDevice != NULL) g_pd3dDevice->Release(); //PCN3112 need to release to reuse on next type
			if(g_pD3D != NULL) g_pD3D->Release();             //
			return E_FAIL;
			}
		}

	// Set up Graphics Device to Run in Mixed - Hardware and Software Vertexing Mode.
	if(vertexMode==1) 
		{
		hr=g_pD3D->CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, 
											  (HWND) hWnd, D3DCREATE_MIXED_VERTEXPROCESSING, 
											  &d3dpp, &g_pd3dDevice);
		if((hr==D3DERR_INVALIDCALL) || (hr==D3DERR_NOTAVAILABLE))
			{
			//PCN3112, no longer need to inform user, will just try the next vertex type
			//MsgWindow("%s\n%s\n%s",
			//	languageArray[10].text,
			//	languageArray[11].text,
			//	languageArray[12].text);
			
			if(g_pd3dDevice != NULL) g_pd3dDevice->Release(); //PCN3112 need to release to reuse on next type
			if(g_pD3D != NULL) g_pD3D->Release();			  //	
			return E_FAIL;
			}
		}

	// Set up Grahics Device to Run in Software Veretexing Mode.
	if(vertexMode==2) 
		{
		hr=g_pD3D->CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, 
											  (HWND) hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING,
											  &d3dpp, &g_pd3dDevice);
		if((hr==D3DERR_INVALIDCALL) || (hr==D3DERR_NOTAVAILABLE))
			{
			MsgWindow("%s\n%s",
				languageArray[13].text,
				languageArray[14].text);
			if(g_pd3dDevice != NULL) g_pd3dDevice->Release();
			if(g_pD3D != NULL) g_pD3D->Release();
			return E_FAIL;
			}
		}
		
	g_pd3dDevice->SetRenderState( D3DRS_CULLMODE, D3DCULL_NONE );
    g_pd3dDevice->SetRenderState(D3DRS_ZENABLE, D3DZB_TRUE );
	
	g_pd3dDevice->SetRenderState(D3DRS_ZFUNC, D3DCMP_LESS );
	g_pd3dDevice->SetRenderState(D3DRS_ALPHABLENDENABLE,TRUE);
	g_pd3dDevice->SetRenderState(D3DRS_SRCBLEND,D3DBLEND_SRCALPHA);
	g_pd3dDevice->SetRenderState(D3DRS_DESTBLEND,D3DBLEND_INVSRCALPHA);
	g_pd3dDevice->SetRenderState( D3DRS_LIGHTING, TRUE );
	g_pd3dDevice->SetRenderState( D3DRS_AMBIENT, 0x00202020 );
//	g_pd3dDevice->SetRenderState( D3DRS_SHADEMODE, D3DSHADE_FLAT);
	

	SetupMaterial();
	return S_OK;
}

void D3Dwindow::SetupMatrices(void)
	{
	D3DXMATRIX matView;
	D3DXMatrixLookAtLH( &matView, 
		&D3DXVECTOR3(cam->position[0], cam->position[1], cam->position[2]),
		&D3DXVECTOR3(cam->target[0], cam->target[1], cam->target[2]),
		&D3DXVECTOR3(cam->up[0], cam->up[1], cam->up[2])
		);
	g_pd3dDevice->SetTransform(D3DTS_VIEW, &matView);
	D3DXMATRIX matProj;
	D3DXMatrixPerspectiveFovLH( &matProj, D3DX_PI/3, aspectRatio, 1.0, 1000000.0);
	g_pd3dDevice->SetTransform(D3DTS_PROJECTION, &matProj);
	}

void D3Dwindow::SetupMaterial(void)
{
    // Set up a material. The material here just has the diffuse and ambient
    // colors set to yellow. Note that only one material can be used at a time.
    D3DMATERIAL9 mtrl;
    ZeroMemory( &mtrl, sizeof(D3DMATERIAL9) );
    mtrl.Diffuse.r = mtrl.Ambient.r = 1.0f;
    mtrl.Diffuse.g = mtrl.Ambient.g = 1.0f;
    mtrl.Diffuse.b = mtrl.Ambient.b = 1.0f;
    mtrl.Diffuse.a = mtrl.Ambient.a = 1.0f;
	mtrl.Emissive.r = 0.1f;
	mtrl.Emissive.g = 0.1f;
	mtrl.Emissive.b = 0.1f;
	mtrl.Emissive.a = 1.0f;
    g_pd3dDevice->SetMaterial( &mtrl );
}


void D3Dwindow::SetupLights()
{

    g_pd3dDevice->SetLight( 0, &light1.light );
    g_pd3dDevice->LightEnable( 0, TRUE );
    g_pd3dDevice->SetLight( 1, &light2.light );
    g_pd3dDevice->LightEnable( 1, TRUE );

    // Finally, turn on some ambient light.
}

void D3Dwindow::RenderD3D(void)
	{

	if(p!=NULL) p->cameraPosition=cam->position; //PCN2461 (8 December 2003,Antony)
	if(p!=NULL) p->cameraType=camSelect;		 //PCN2461 (8 December 2003,Antony)

	if(OpenGL) {RenderOpenGL(); return;}

	if(NULL == g_pd3dDevice) {MsgWindow("d3d101"); exit(1);}//g_pd3dDevice NULL");

	// Clear the backbuffer to D3DCOLOR_XRGB color

    g_pd3dDevice->Clear( 0, NULL, D3DCLEAR_TARGET|D3DCLEAR_ZBUFFER,
                         D3DCOLOR_XRGB(250,250,255), 1.0f, 0 );

	// Begin the scene



	g_pd3dDevice->BeginScene();
        SetupLights();
		SetupMatrices();

		if(pipeLoaded && p!=NULL)
			{
			if(drawPipeTrueFalse) p->PipeDraw(); //p2->PipeDraw(); //PCN2461 (8 December 2003, Antony)
			if(p->loadingPipe && loadingPanel!=NULL) loadingPanel->Draw();  //PCN2461 (8 December 2003, Antony)
			if(p->repaintingPipe && repaintingPanel!=NULL) repaintingPanel->Draw();  //PCN2461 (8 December 2003, Antony)
			
			}
		else {
			}
		if(drawScapeTrueFalse && landScape!=NULL) landScape->DrawScape(); //PCN2461 (8 December 2003, Antony)

		if(dirArrow!=NULL) dirArrow->Draw(); //PCN2653 (26 Feb 2004 Antony) only draw if dirArrow is initialised
//		dirX->Draw();
//		dirY->Draw();
//		dirZ->Draw();


//	DebugVariables();
//	if(pipeLoaded && p!=NULL && headsUp!=NULL) //PCN2461 (8 December 2003, Antony)
//		{
//		lf=p->laserFocus;
//		headsUp->UpdateText(1,lf);
//		headsUp->UpdateText(3,(float) (p->pipeCapacity[lf])/pVCalculationsMultiplier); //PCN2829 PCN2860 div by Mulitplier (Antony, 2 June 2004)
//		headsUp->UpdateText(5,(float) (p->pipeOvality[lf]) /pVCalculationsMultiplier); //PCN2829 PCN2860 div by Mulitplier (Antony, 2 June 2004)	
//		headsUp->UpdateText(7,(float) (p->pipeDeltaMax[lf])/pVCalculationsMultiplier); //PCN2829 PCN2860 div by Mulitplier (Antony, 2 June 2004)
//		headsUp->UpdateText(9,(float) (p->pipeDeltaMin[lf])/pVCalculationsMultiplier); //PCN2829 PCN2860 div by Mulitplier (Antony, 2 June 2004)
//		}
//	if(headsUp!=NULL) headsUp->Draw();//PCN
//	DebugVariables();
	g_pd3dDevice->EndScene(); // Moved below the last thing to draw, but funny thing
							  // it worked fine until recently, its new SDK, not 9b
							  // Don't know if this has anything to do with it
							  // Antony van Iersel, 19 March 2004

	// Present the backbuffer contects to the display
	g_pd3dDevice->Present( NULL, NULL, NULL, NULL);
	}


void D3Dwindow::DebugVariables(void)
{
//DF	if(debugPanel!=NULL && p!=NULL && dirArrow!=NULL) debugPanel->UpdateText(2,dirArrow->rotation.x); //PCN2461 (8 December 2003, Antony)
//DF	if(debugPanel!=NULL && p!=NULL && dirArrow!=NULL) debugPanel->UpdateText(4,dirArrow->rotation.y);
//DF	if(debugPanel!=NULL && p!=NULL && dirArrow!=NULL) debugPanel->UpdateText(6,dirArrow->rotation.z);
//	if(p!=NULL)	debugPanel->UpdateText(2,p->test1);

//	if(debugPanel!=NULL) debugPanel->Draw(); //PCN2461 (8 December 2003, Antony)
}


//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Name: InitialiseLanguageArray PCN2473
// Created By: Antony van Iersel
// Date: 26 Febuary 2004
// 
// Description:	Initilaises the language array with default English
//
// Input: None
// Output: None
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

void D3Dwindow::InitialiseLanguageArray(void)
{          
	strcpy(languageArray[0].text  ,"Frame:");
	strcpy(languageArray[1].text  ,"Capacity:");
	strcpy(languageArray[2].text  ,"Ovality:");
	strcpy(languageArray[3].text  ,"Delta max:");
	strcpy(languageArray[4].text  ,"Delta min:");
	strcpy(languageArray[5].text  ,"Loading");
	strcpy(languageArray[6].text  ,"Re-Painting");

	strcpy(languageArray[7].text  ,"Sorry, Hardware Vertexing is not available on your Graphics Processing Unit");
	strcpy(languageArray[8].text  ,"Cannot Initialise Hardware Vertexing");
	strcpy(languageArray[9].text  ,"Please Try Mixed or Software Vertexing");

	strcpy(languageArray[10].text ,"Sorry, Mixed Vertexing is not available on your Graphics Processing Unit");
	strcpy(languageArray[11].text ,"Cannot Initialise Mixed Vertexing");
	strcpy(languageArray[12].text ,"Please Try Software Vertexing");

	strcpy(languageArray[13].text ,"Sorry, Software Vertexing is not available on your Graphics Processing Unit");
	strcpy(languageArray[14].text ,"This hardware cannot run the 3D Application");

	strcpy(languageArray[15].text ,"Could not find");
	strcpy(languageArray[16].text ,"texture");
	strcpy(languageArray[17].text ,"pipe");

	strcpy(languageArray[18].text ,"white");
	strcpy(languageArray[19].text ,"laser");
	strcpy(languageArray[20].text ,"panel");

	strcpy(languageArray[21].text ,"boarder");
	strcpy(languageArray[22].text ,"D3D model");
	strcpy(languageArray[23].text ,"Invalid export file name");
	strcpy(languageArray[24].text ,"No data in this section to export");
}

void D3Dwindow::CreateHeadsUpPanel(void)
{
	//if(headsUp!=NULL) //PCN2461 (8 December 2003, Antony)
	headsUp      = new D3DFont(g_pd3dDevice, textureDir,
							   "\\Panel_Loading.jpg",
							   "\\Boarder.jpg",
							   (D3DFont::languageText *) languageArray); //PCN2473 Language pointer (Antony, 11 May 2004) 
	headsUp->SetPanel();
	headsUp->NewText(languageArray[0].text,D3DXVECTOR2(12,70)); headsUp->NewText(0,D3DXVECTOR2(70,70));
	headsUp->NewText(languageArray[1].text,D3DXVECTOR2(12,10)); headsUp->NewText(0,D3DXVECTOR2(65,10));
	headsUp->NewText(languageArray[2].text,D3DXVECTOR2(12,25)); headsUp->NewText(0,D3DXVECTOR2(65,25));
	headsUp->NewText(languageArray[3].text,D3DXVECTOR2(12,40)); headsUp->NewText(0,D3DXVECTOR2(65,40));
	headsUp->NewText(languageArray[4].text,D3DXVECTOR2(12,55)); headsUp->NewText(0,D3DXVECTOR2(65,55));
	headsUp->NewText("%",D3DXVECTOR2(128,10));
	headsUp->NewText("%",D3DXVECTOR2(128,25));
	headsUp->NewText("%",D3DXVECTOR2(128,40));
	headsUp->NewText("%",D3DXVECTOR2(128,55));
}

void D3Dwindow::CreateLoadingUnloadingPanel(void)
{
	loadingPanel = new D3DFont(g_pd3dDevice, 
							   textureDir,
							   "\\Panel_Loading.jpg", 
							   "\\Boarder.jpg",
							   (D3DFont::languageText *) languageArray);
	loadingPanel->panelDim.x			   = 580; loadingPanel->panelDim.y			  = 5;
	loadingPanel->panelDim.width		   = 95; loadingPanel->panelDim.height		  = 20;
	loadingPanel->panelDim.boarderHeight = 1;	  loadingPanel->panelDim.boarderWidth = 1;
	loadingPanel->panelDim.colour        = D3DCOLOR_RGBA(255,255,255,90);
	loadingPanel->panelDim.boarderColour = D3DCOLOR_RGBA(255,255,255,255);
	loadingPanel->fontPosition.top    = (long) loadingPanel->panelDim.y;
	loadingPanel->fontPosition.left   = (long) loadingPanel->panelDim.x;
	loadingPanel->fontPosition.right  = (long) loadingPanel->panelDim.x + (long) loadingPanel->panelDim.width;
	loadingPanel->fontPosition.bottom = (long) loadingPanel->panelDim.y + (long) loadingPanel->panelDim.height;

	loadingPanel->SetPanel();
	loadingPanel->NewText(languageArray[5].text,D3DXVECTOR2((float) loadingPanel->fontPosition.left+2
		                                       ,(float) loadingPanel->fontPosition.top));
	loadingPanel->InitGeometry();
}

void D3Dwindow::CreateRepaintingPanel()
{
	repaintingPanel = new D3DFont(g_pd3dDevice, 
								  textureDir,
								  "\\Panel_Loading.jpg", 
								  "\\Boarder.jpg",
								  (D3DFont::languageText *) languageArray);
	repaintingPanel->panelDim.x			   = 410; repaintingPanel->panelDim.y			  = 5;
	repaintingPanel->panelDim.width		   = 95; repaintingPanel->panelDim.height		  = 20;
	repaintingPanel->panelDim.boarderHeight = 1;	  repaintingPanel->panelDim.boarderWidth = 1;
	repaintingPanel->panelDim.colour        = D3DCOLOR_RGBA(255,255,255,90);
	repaintingPanel->panelDim.boarderColour = D3DCOLOR_RGBA(255,255,255,255);
	repaintingPanel->fontPosition.top    = (long) repaintingPanel->panelDim.y;
	repaintingPanel->fontPosition.left   = (long) repaintingPanel->panelDim.x;
	repaintingPanel->fontPosition.right  = (long) repaintingPanel->panelDim.x + (long) repaintingPanel->panelDim.width;
	repaintingPanel->fontPosition.bottom = (long) repaintingPanel->panelDim.y + (long) repaintingPanel->panelDim.height;

	repaintingPanel->SetPanel();
	repaintingPanel->NewText(languageArray[6].text,D3DXVECTOR2((float) repaintingPanel->fontPosition.left+2
		                                       ,(float) repaintingPanel->fontPosition.top));
	repaintingPanel->InitGeometry();
}

void D3Dwindow::UpdatePanelText(void)
{
if(headsUp!=NULL)
		{
		headsUp->UpdateText(0,languageArray[0].text);	//Frame
		headsUp->UpdateText(2,languageArray[1].text);   //Capacity
		headsUp->UpdateText(4,languageArray[2].text);	//Ovality
		headsUp->UpdateText(6,languageArray[3].text);	//Delta max
		headsUp->UpdateText(8,languageArray[4].text);	//Delta min
		}
if(loadingPanel!=NULL) loadingPanel->UpdateText(0,languageArray[5].text);		//Loading
if(repaintingPanel!=NULL) repaintingPanel->UpdateText(0,languageArray[6].text); //Re-Painting

}
	

void D3Dwindow::CreateDebugPanel(void)
{
	debugPanel = new D3DFont(g_pd3dDevice, 
							 textureDir,
							 "\\Panel_Loading.jpg", 
							 "\\Boarder.jpg",
							 (D3DFont::languageText *) languageArray);  //PCN2473 Language pointer (Antony, 11 May 2004)
	debugPanel->panelDim.x			   = 10; debugPanel->panelDim.y				= 100;
	debugPanel->panelDim.width		   = 185; debugPanel->panelDim.height		= 220;
	debugPanel->panelDim.boarderHeight = 2;	  debugPanel->panelDim.boarderWidth = 2;
	debugPanel->panelDim.colour        = D3DCOLOR_RGBA(255,255,255,90);
	debugPanel->panelDim.boarderColour = D3DCOLOR_RGBA(255,255,255,255);
	debugPanel->fontPosition.top    = (long) debugPanel->panelDim.y;
	debugPanel->fontPosition.left   = (long) debugPanel->panelDim.x;
	debugPanel->fontPosition.right  = (long) debugPanel->panelDim.x + (long) debugPanel->panelDim.width;
	debugPanel->fontPosition.bottom = (long) debugPanel->panelDim.y + (long) debugPanel->panelDim.height;

	debugPanel->SetPanel();
	debugPanel->NewText("Debug"	  ,	D3DXVECTOR2((float) debugPanel->fontPosition.left+2,   (float) debugPanel->fontPosition.top));
	debugPanel->NewText("Left:"   ,	D3DXVECTOR2((float) debugPanel->fontPosition.left+2,   (float) debugPanel->fontPosition.top+15));
	debugPanel->NewText(""        ,	D3DXVECTOR2((float) debugPanel->fontPosition.left+120, (float) debugPanel->fontPosition.top+15));
	debugPanel->NewText("Right:"  ,	D3DXVECTOR2((float) debugPanel->fontPosition.left+2,   (float) debugPanel->fontPosition.top+30));
	debugPanel->NewText(""        , D3DXVECTOR2((float) debugPanel->fontPosition.left+120, (float) debugPanel->fontPosition.top+30));
	debugPanel->NewText(""		  ,	D3DXVECTOR2((float) debugPanel->fontPosition.left+2,   (float) debugPanel->fontPosition.top+45));
	debugPanel->NewText(""        , D3DXVECTOR2((float) debugPanel->fontPosition.left+120, (float) debugPanel->fontPosition.top+45));
	debugPanel->InitGeometry();
}

void D3Dwindow::MouseMotion(int x, int y)
	{
	float r;
	if(leftBut)
		{
		if(camSelect==0) { railCam->Pan((float) (x-leftX)/3); railCam->Tilt((float) (y-leftY)/6); }
		if(camSelect==1) { pipeCam->RollTargetY((float) (x-leftX)/3); pipeCam->TiltTarget((float) (y-leftY)/6); }
		if(camSelect==2) { freeCam->Pan((float) (x-leftX)/3); freeCam->Tilt((float) (y-leftY)/6); }
		if(camSelect==3) 
			{ 
			r=mapCam->position[1]/5000;
			if(r<0.25) r=0.25;
			mapCam->MoveStrafe((leftX-x)*r*12); 
			mapCam->MoveFoward((y-leftY)*r*12);
			}
		}
	if(rightBut)
		{
		if(camSelect==0)
			{
			ZoomLaserDistance((y-rightY)*5);
			}
		if(camSelect==1) 
			{
			r=freeCam->position[1]/5000;
			if(r<0.25) r=0.25;
			pipeCam->Zoom((y-rightY)*40*r);
			}

		if(camSelect==2) 
			{
			r=freeCam->position[1]/5000;
			if(r<0.25) r=0.25;
			freeCam->MoveHeight((rightY-y)*r*50);
			}
		if(camSelect==3)
			{
			r=mapCam->position[1]/5000;
			if(r<0.25) r=0.25;
			mapCam->MoveHeight((rightY-y)*r*50);
			}
		}
	rightX=x; rightY=y;
	leftX=x; leftY=y;
	}

void D3Dwindow::AllStop(void)
{
	railCam->speed=0;
	pipeCam->speed=0;
	freeCam->speed=0;
	if(p!=NULL) p->laserSpeed=0; //PCN2461 (8 December 2003, Antony)
}

void D3Dwindow::Zoom(int speed)
{
	float r;
	if(camSelect==0) ZoomLaserDistance(speed);
	if(camSelect==1) 
		{
		r=freeCam->position[1]/5000;
		if(r<0.25) r=0.25;
		pipeCam->Zoom((speed)*8*r);
		}
	if(camSelect==2) 
		{
		r=freeCam->position[1]/5000;
		if(r<0.25) r=0.25;
		freeCam->MoveHeight((speed)*r*50);
		}
	if(camSelect==3)
		{
		r=mapCam->position[1]/5000;
		if(r<0.25) r=0.25;
		mapCam->MoveHeight((speed)*r*50);
		}
}

void D3Dwindow::ResetCameras(void)
{
	pipeCam->Reset(); freeCam->Reset(); mapCam->Reset(); railCam->Reset();
	pipeCam->ZoomTarget(850); pipeCam->TiltTarget(45); pipeCam->RollTargetY(-45);
	if(p!=NULL) pipeCam->ZoomTarget((float) p->avgRad*3); //PCN2461 (8 December 2003, Antony)

	mapCam->MoveTo(0,25000,0); mapCam->Pan(180); mapCam->Tilt(90);
	freeCam->Reset();
	freeCam->MoveTo(0,1440,1440); freeCam->Tilt(45);
	laserDistance=200;
}

void D3Dwindow::Keyboard(int value)
{
	float r;


	if(value=='1') SelectCamera(1);
	if(value=='2') SelectCamera(2);
	if(value=='3') SelectCamera(3);
	if(value=='4') SelectCamera(4);

	if(value=='0') g_pd3dDevice->SetRenderState( D3DRS_FILLMODE, D3DFILL_WIREFRAME);
	if(value=='9') g_pd3dDevice->SetRenderState( D3DRS_FILLMODE, D3DFILL_SOLID);

	if(camSelect==0)
		{

		}
	if(camSelect==1)
		{
		if(value==',') pipeCam->RollTargetZ(5);
		if(value=='.') pipeCam->RollTargetZ(-5);
		}
	if(camSelect==2)
		{
		r=freeCam->position[1]/5000;
		if(r<0.25) r=0.25;
		if(value==VK_UP)    {freeCam->speed+=20; if(freeCam->speed>40) freeCam->speed=60;}
		if(value==VK_DOWN)  {freeCam->speed-=20; if(freeCam->speed<-40) freeCam->speed=-60;}
		if(value==VK_LEFT)  freeCam->MoveStrafe(+100*r);
		if(value==VK_RIGHT) freeCam->MoveStrafe(-100*r);
		}
	if(camSelect==3)
		{
		if(value==VK_UP) landScape->offset[2]+=100;
		if(value==VK_DOWN) landScape->offset[2]-=100;
		if(value==VK_LEFT) landScape->offset[0]+=100;
		if(value==VK_RIGHT) landScape->offset[0]-=100;
		if(value==VK_HOME) landScape->rotateLand++;
		if(value==VK_END) landScape->rotateLand--;
		if(value=='+') landScape->scale*=2;
		if(value=='-') landScape->scale/=2;
		}
	if(value=='n' && p!=NULL) p->ChangeLaserWidth(p->laserWidth-2); // Added p!=NULL PCN2461 (8 December 2003, Antony)
	if(value=='m' && p!=NULL) p->ChangeLaserWidth(p->laserWidth+2); // Added p!=NULL PCN2461 (8 December 2003, Antony)
	
	if(value==' ') 
		{
		AllStop();
		}

	if(value=='5') if(drawPipeTrueFalse==false) drawPipeTrueFalse=true; else drawPipeTrueFalse=false;
	if(value=='6') if(drawScapeTrueFalse==false) drawScapeTrueFalse=true; else drawScapeTrueFalse=false;
	if(value=='7') if(drawWaterTrueFalse==false) drawWaterTrueFalse=true; else drawWaterTrueFalse=false;


//	if(value=='e') { p->ExportPipeSTL("c:\\test.stl"); }

//	if(value=='u') if(dirArrow!=NULL) dirArrow->rotation.x++;
//	if(value=='j') if(dirArrow!=NULL) dirArrow->rotation.x--;
//	if(value=='i') if(dirArrow!=NULL) dirArrow->rotation.y++;
//	if(value=='k') if(dirArrow!=NULL) dirArrow->rotation.y--;
//	if(value=='o') if(dirArrow!=NULL) dirArrow->rotation.z++;
//	if(value=='l') if(dirArrow!=NULL) dirArrow->rotation.z--;

}

void D3Dwindow::SelectCamera(int c)
{
//	if(landScape==NULL) return; //PCN2461 (8 december 2003, Antony)
	if(c==1) {AllStop(); camSelect=0; cam=railCam; }//landScape->SetCam(cam); if(dirArrow!=NULL) dirArrow->SetCam(cam);} //PCN2461 (8 December 2003, Antony)
	if(c==2) {AllStop(); camSelect=1; cam=pipeCam; }//landScape->SetCam(cam); if(dirArrow!=NULL) dirArrow->SetCam(cam);} //PCN2461 (8 December 2003, Antony)
	if(c==3) {AllStop(); camSelect=2; cam=freeCam; }//landScape->SetCam(cam); if(dirArrow!=NULL) dirArrow->SetCam(cam);} //PCN2461 (8 December 2003, Antony)
	if(c==4) {AllStop(); camSelect=3; cam=mapCam;  }//landScape->SetCam(cam); if(dirArrow!=NULL) dirArrow->SetCam(cam);} //PCN2461 (8 December 2003, Antony)
	UpdateCameraObjects();
}

void D3Dwindow::UpdateCameraObjects()
{
	if(landScape!=NULL) landScape->SetCam(cam);
	if(dirArrow!=NULL)  dirArrow->SetCam(cam);
//	if(dirX!=NULL)		dirX->SetCam(cam);
//	if(dirY!=NULL)		dirY->SetCam(cam);
//	if(dirZ!=NULL)		dirZ->SetCam(cam);
}

void D3Dwindow::PipeDepthScale(float scale)
{
	if(p==NULL) return; //PCN2461 (8 December 2003, Antony)
	int fr;
	p->depthScale=scale;
	p->MarkPipe(p->laserFocus);
	for(fr=0;fr<p->frames;fr++) 
		{ 
		railCam->rail[fr] = D3DXVECTOR3( 0,0, (-fr)*p->depthScale);
		pipeCam->rail[fr] = D3DXVECTOR3( 0,0, (-fr)*p->depthScale);
		}
	p->UnloadPipe(); //PCN2410 unload Pipe to redraw with new scale. (18 November 2003 Antony).
}

void D3Dwindow::Motion(void)
{
	float r=freeCam->position[1]/5000;
	if(r<0.25) r=0.25;

//	p->MoveLaser(p->laserSpeed);
	
	if(camSelect==0) railCam->MoveToRail(p->laserFocus-laserDistance);
	if(p==NULL) return; //PCN2461 (8 December 2003, Antony)
	if(camSelect==1) pipeCam->MoveToRailTarget(p->laserFocus);
	if(camSelect==2) freeCam->MoveDirection(freeCam->speed*r);
}

void D3Dwindow::MoveLaserTo(int frame)
{
	if(p==NULL) return; //PCN2461 (8 December 2003, Antony)
	p->MoveLaserTo(frame);
	p->laserSpeed=0;
}

void D3Dwindow::ZoomLaserDistance(int dis)
{
	if(p==NULL) return; //PCN2461 (8 December 2003, Antony)
	if((laserDistance+dis)>p->laserFocus) return;
	dis=dis/5;
	laserDistance+=dis;
	if(laserDistance<-500) laserDistance=-500;
	if(laserDistance>500) laserDistance=500;
}

void D3Dwindow::WindowToBmp (char *fileName, HWND hWndCapture)
	{

	//////////Richard - 030310//////////////////////////////////////////////
	HRESULT hr;

	LPDIRECT3DTEXTURE9 pRenderTexture = NULL;

	LPDIRECT3DSURFACE9 pRenderSurface = NULL,pBackBuffer = NULL;
	D3DXMATRIX matProjection,matOldProjection;


	hr = g_pd3dDevice->CreateTexture(512,512,
							   1,
											   D3DUSAGE_RENDERTARGET,
											   D3DFMT_X8R8G8B8,
											   D3DPOOL_DEFAULT,
											   &pRenderTexture,
											   NULL);

	hr = pRenderTexture->GetSurfaceLevel(0,&pRenderSurface);

	hr = g_pd3dDevice->GetTransform(D3DTS_PROJECTION,&matOldProjection);
	hr = g_pd3dDevice->GetRenderTarget(0,&pBackBuffer);

	hr = g_pd3dDevice->SetRenderTarget(0,pRenderSurface);

	RenderD3D();

	hr = D3DXSaveSurfaceToFile(fileName,D3DXIFF_JPG,pRenderSurface,NULL,NULL);


	hr = g_pd3dDevice->SetRenderTarget(0,pBackBuffer);
	////////////////////////////////////////////////////////////////////////

/*	HDC hdcScreen = ::GetDC(hWndCapture);
	RECT rc;
	::GetClientRect(hWndCapture,&rc);
	int iScrWidth=rc.right;
	int iScrHeight=rc.bottom;



	MsgWindow("Width = %i, Height = %i",iScrWidth, iScrHeight);
	BITMAPINFO bmpInfo;
	bmpInfo.bmiHeader.biSize			   = (DWORD) sizeof(BITMAPINFOHEADER);
	bmpInfo.bmiHeader.biWidth              = (LONG) iScrWidth;
	bmpInfo.bmiHeader.biHeight             = (LONG) iScrHeight;
	bmpInfo.bmiHeader.biPlanes             = (WORD) 1;
	bmpInfo.bmiHeader.biBitCount           = (WORD) 24;
	bmpInfo.bmiHeader.biCompression        = (DWORD) BI_RGB;
	bmpInfo.bmiHeader.biSizeImage          = (DWORD) 0;//0x000ff800; //iScrWidth*iScrHeight*3;
	bmpInfo.bmiHeader.biXPelsPerMeter      = (LONG) 0;// 0x0ec4;
	bmpInfo.bmiHeader.biYPelsPerMeter      = (LONG) 0;// 0x0ec4;
	bmpInfo.bmiHeader.biClrUsed            = (DWORD) 0;
	bmpInfo.bmiHeader.biClrImportant       = (DWORD) 0;

	void *pvBits = NULL;
	HBITMAP hBitmap = CreateDIBSection(hdcScreen,&bmpInfo,DIB_RGB_COLORS,&pvBits,NULL,0);
	HDC hdcCompatible = CreateCompatibleDC(hdcScreen);
	HBITMAP hOldBitmap = (HBITMAP)SelectObject(hdcCompatible,hBitmap);
	BOOL b = BitBlt(hdcCompatible,0,0,iScrWidth,iScrHeight,hdcScreen,0,0,SRCCOPY);

	int size;
	size = iScrWidth*iScrHeight*3;
	HANDLE hFile = CreateFile(fileName,GENERIC_WRITE,FILE_SHARE_WRITE,NULL,
        CREATE_ALWAYS,FILE_ATTRIBUTE_NORMAL,NULL);

	if (hFile != INVALID_HANDLE_VALUE) 
		{
		DWORD dwCnt;
		BITMAPFILEHEADER BM_Header;
		BM_Header.bfType = ((WORD) ('M' << 8) | 'B');
		BM_Header.bfSize = (DWORD) ((iScrWidth*(iScrHeight)*3) + sizeof(BM_Header) +sizeof(bmpInfo.bmiHeader));
		BM_Header.bfReserved1 = 0;
		BM_Header.bfReserved2 = 0;
		BM_Header.bfOffBits = (DWORD) ((sizeof(BM_Header)) + sizeof(bmpInfo.bmiHeader));
		WriteFile(hFile,(char*) &BM_Header,         sizeof(BM_Header),         &dwCnt,NULL);
		WriteFile(hFile,(char*) &bmpInfo.bmiHeader, sizeof(bmpInfo.bmiHeader), &dwCnt,NULL);
		WriteFile(hFile,(char*) pvBits,             ((iScrWidth)*(iScrHeight+1)*3), &dwCnt,NULL);//+1022,            &dwCnt,NULL);
		CloseHandle(hFile);
		}     

////// Antony - 9 October ver 1.1 ////////////////////////////////
// Note: 1 is Added to the iScrHeight							//
// (works if one is aded to iScrWidth aswell , but not both)    //
// without it the File is 1022byts short. And wont load.        //
//////////////////////////////////////////////////////////////////

	SelectObject(hdcCompatible,hOldBitmap);
	DeleteDC(hdcCompatible);
	DeleteObject(hBitmap);
	::ReleaseDC(NULL,hdcScreen);*/
}

//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Name: ExportPipe PCN2376 14 November 2003
// Created By: Antony van Iersel
// 
// Description:	Calls the ExportPipe Function in Pipe class passing it
// where the pipe exported data has to be saved.
// Input: File Name on where to export pipe
// Output: None
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

void D3Dwindow::ExportPipe(char *filename)
{
	if(p==NULL) return; //PCN2461 (8 December 2003, Antony)
	p->ExportPipeSTL(filename);
}

///////////////////////////////////////////////////////////////////////////////////////////////////
//								   ////////////////////////////////////////////////////////////////	
// Here onward is opengl rendering ////////////////////////////////////////////////////////////////
//                                 ////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////

void D3Dwindow::OpenGLWindowInit(void)
{

        ZeroMemory( &pfd, sizeof( pfd ) );
        pfd.nSize = sizeof( pfd );
        pfd.nVersion = 1;
        pfd.dwFlags = PFD_DRAW_TO_WINDOW | PFD_SUPPORT_OPENGL |
                      PFD_DOUBLEBUFFER;
        pfd.iPixelType = PFD_TYPE_RGBA;
        pfd.cColorBits = 24;
        pfd.cDepthBits = 16;
        pfd.iLayerType = PFD_MAIN_PLANE;
        int format = ChoosePixelFormat( hDC, &pfd );
        SetPixelFormat( hDC, format, &pfd );

        // create the render context (RC)
        hglrc = wglCreateContext( hDC );

        // make it the current render context
        wglMakeCurrent( hDC, hglrc );


	
	RECT rc;

//	int PixelFormat;
//
//	static PIXELFORMATDESCRIPTOR pfd=
//	{
//		sizeof(PIXELFORMATDESCRIPTOR),
//		1,
//		PFD_DRAW_TO_WINDOW |						// Format Must Support Window
//		PFD_SUPPORT_OPENGL |						// Format Must Support OpenGL
//		PFD_DOUBLEBUFFER,							// Must Support Double Buffering
//		PFD_TYPE_RGBA,
//		16,	//Bits
//		0,0,0,0,0,0,
//		0,
//		0,
//		0,
//		0,0,0,0,
//		16,
//		0,
//		0,
//		PFD_MAIN_PLANE,
//		0,
//		0,0,0
//	};

//	if(!(PixelFormat=ChoosePixelFormat(hDC,&pfd)))
//		MessageBox(NULL,"Can't find a suitable PixelFormat,","ERROR",MB_OK|MB_ICONEXCLAMATION);
//	if(!SetPixelFormat(hDC,PixelFormat,&pfd))
//		MessageBox(NULL,"Can't set the PixelFormat,","ERROR",MB_OK|MB_ICONEXCLAMATION);

	
//	hglrc = wglCreateContext(hDC);
//	wglMakeCurrent(hDC,hglrc);
    GetClientRect(hWnd, &rc);
	ResiseGLWindow(rc.right-rc.left,rc.bottom-rc.top);	

}

void D3Dwindow::GLInit(void)
{







	glClearColor((float) 1.0, (float) 1.0, (float) 1.0, (float) 0.0);
	
	glPolygonMode(GL_FRONT_AND_BACK,GL_FILL);
	//glShadeModel(GL_SMOOTH);							// Enable Smooth Shading
	//glClearColor(1.0f, 1.0f, 1.0f, 0.5f);				// Black Background
	//glClearDepth(1.0f);									// Depth Buffer Setup
	//glEnable(GL_DEPTH_TEST);							// Enables Depth Testing
	//glDepthFunc(GL_LEQUAL);								// The Type Of Depth Testing To Do
	//glHint(GL_PERSPECTIVE_CORRECTION_HINT, GL_NICEST);	// Really Nice Perspective Calculations



//	glShadeModel(GL_SMOOTH);						// Enables Smooth Shading
//	glDepthFunc(GL_LEQUAL);							// The Type Of Depth Test To Do
//	glHint(GL_PERSPECTIVE_CORRECTION_HINT, GL_NICEST);			// Really Nice Perspective Calculations
//	glPolygonMode(GL_FRONT_AND_BACK,GL_FILL);
//	glEnable(GL_CULL_FACE);
//	glEnable(GL_DEPTH_TEST);

//	glEnable(GL_LIGHTING);
//	glEnable(GL_COLOR_MATERIAL);
//	glLightModelfv(GL_LIGHT_MODEL_AMBIENT,whiteLight);
//	glLightfv(GL_LIGHT0,GL_DIFFUSE,sourceLight);

//	
//	glEnable(GL_LIGHT0);
//	glColorMaterial(GL_FRONT, GL_AMBIENT_AND_DIFFUSE);

	//////////////////////
//	glPolygonMode(GL_FRONT_AND_BACK,GL_FILL);
//	glEnable(GL_CULL_FACE);
//	glEnable(GL_LIGHTING);
//	glEnable(GL_COLOR_MATERIAL);
//	glLightModelfv(GL_LIGHT_MODEL_AMBIENT,whiteLight);
//	glLightfv(GL_LIGHT0,GL_DIFFUSE,sourceLight);


//	glEnable(GL_MODULATE);
//	glEnable(GL_TEXTURE_2D);
//	glEnable(GL_LIGHT0);
//	glColorMaterial(GL_FRONT, GL_AMBIENT_AND_DIFFUSE);
//	glLightfv(GL_LIGHT0, GL_POSITION, lightpos );
}

void D3Dwindow::ResiseGLWindow(int width, int height)
{
	if(height==0) height = 1;
	glViewport(0,0,width,height);
	glMatrixMode(GL_PROJECTION);
	glLoadIdentity();
	gluPerspective(45.0,(float) width / (float) height,0.1,10000.0);
	glMatrixMode(GL_MODELVIEW);
	glLoadIdentity();
}

void D3Dwindow::RenderOpenGL(void)
{


	glClear(GL_COLOR_BUFFER_BIT |GL_DEPTH_BUFFER_BIT);
	
	glLoadIdentity();					// Reset The View
	gluLookAt(cam->position.x,cam->position.y,cam->position.z,
			  cam->target.x,cam->target.y,cam->target.z,
			  cam->up.x, cam->up.y, cam->up.z); // camera (from where x,y,z lootat x,y,z cameras up, x,y,z)
//	glTranslatef(-1.5f,0.0f,-6.0f);						// Move Left 1.5 Units And Into The Screen 6.0
//	glRotatef(rtri,0.0f,1.0f,0.0f);						// Rotate The Triangle On The Y axis ( NEW )
//	glLightfv(GL_LIGHT0, GL_POSITION, lightpos );

	p->PipeDraw();

	SwapBuffers(hDC);					// Swap Buffers (Double Buffering)
}


