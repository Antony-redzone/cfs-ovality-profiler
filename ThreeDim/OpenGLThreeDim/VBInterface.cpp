#include <windows.h>
#include <gl\gl.h>
#include <gl\glu.h>
#include <gl\glaux.h>

void OpenGLInit(HWND hwnd);
void ResiseGLWindow(int width, int height);
void GLInit(void);

void __stdcall ThreeDim_Initialise(HWND _hwnd)
{
	OpenGLInit(_hwnd);
}

void __stdcall ThreeDim_RenderScene(void)
{
	glClear(GL_COLOR_BUFFER_BIT | GL_DEPTH_BUFFER_BIT);
	glLoadIdentity();
}

void OpenGLInit(HWND hwnd)
{
	HDC hdc;
	HGLRC hglrc;
	RECT rc;
	int PixelFormat;

	static PIXELFORMATDESCRIPTOR pfd=
	{
		sizeof(PIXELFORMATDESCRIPTOR),
		1,
		PFD_DRAW_TO_WINDOW |
		PFD_SUPPORT_OPENGL |
		PFD_TYPE_RGBA,
		16,	//Bits
		0,0,0,0,0,0,
		0,
		0,
		0,
		0,0,0,0,
		16,
		0,
		0,
		PFD_MAIN_PLANE,
		0,
		0,0,0
	};

	hdc = GetDC(hwnd);

	if(!(PixelFormat=ChoosePixelFormat(hdc,&pfd)))
		MessageBox(NULL,"Can't find a suitable PixelFormat,","ERROR",MB_OK|MB_ICONEXCLAMATION);
	if(!SetPixelFormat(hdc,PixelFormat,&pfd))
		MessageBox(NULL,"Can't set the PixelFormat,","ERROR",MB_OK|MB_ICONEXCLAMATION);

	
	hglrc = wglCreateContext(hdc);
	wglMakeCurrent(hdc,hglrc);
    GetClientRect(hwnd, &rc);
	ResiseGLWindow(rc.right-rc.left,rc.bottom-rc.top);	

}

void GLInit(void)
{
	glClearColor(0.0,0.0,0.0,0.0);
	glClearDepth(1.0);
	glEnable(GL_DEPTH_TEST);
	glDepthFunc(GL_EQUAL);
	glHint(GL_PERSPECTIVE_CORRECTION_HINT,GL_NICEST);

}

void ResiseGLWindow(int width, int height)
{
	if(height==0) height = 1;
	glViewport(0,0,width,height);
	glMatrixMode(GL_PROJECTION);
	glLoadIdentity();
	gluPerspective(45.0,(float) width / (float) height,0.1,100.0);
	glMatrixMode(GL_MODELVIEW);
	glLoadIdentity();
}

