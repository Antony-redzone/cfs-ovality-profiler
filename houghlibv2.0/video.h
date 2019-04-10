//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Video 24 November 2003
// Created By: Louise Shrimpton
// 
// Description:	This is not a class, just a file containing all functions that 
//	interact with the Visual Basic code.
//
// Functionality:
//
// Inherited From:  None
//	 	
// 
//
// 
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#ifndef video
#define video 

//------------------------------------------------------------------------------
#include <windows.h>

#include "qedit.h"
//#include <dshow.h>
//#include <fstream.h>


//
// Function prototypes
//
extern HWND hwnd;

enum PLAYSTATE {Stopped, Paused, Running, Init};

//
// Macros
//
#define SAFE_RELEASE(x) if(x !=NULL) { x->Release(); x = NULL; }
// convenient macro for releasing interfaces
//#define HELPER_RELEASE(x)   if (x != NULL) \
//                            { \
//                                x->Release(); \
//                                x = NULL; \
//                            }

#define JIF(x) if (FAILED(hr=(x))) \
    {Msg(TEXT("FAILED(hr=0x%x) in ") TEXT(#x) TEXT("\n"), hr); return hr;}

//
// Constants
//
// Application-defined message to notify app of filtergraph events
#define WM_GRAPHNOTIFY  WM_APP+1

//
// Resource constants
//
#define IDI_VIDPREVIEW          100

void Msg(TCHAR *szFormat, ...);

extern IBaseFilter *pSrcFilter;
// extern IBaseFilter *pF; //PCN2289



//extern int deviceType; // PCN2289

struct pixel { // PCN2516
	unsigned char blue,green,red;
};

class Video {
public:
	int seekByTime; //PCN3289 to be able to switch between seek by time and frame. Default time
//data (at some point make private...)
	char *vbfname;

	int width;
	int height;
	double frameRate; //PCN3533
	long lastSeek; //PCN3533
	double lastRecordedTime;
//	int videoDeviceType;	//PCN 2289, 20 Oct 2003, (Antony) to select device type eg default or nvidia

	LONGLONG videoLengthTime;
//	LONGLONG videoLengthFrames; PCN3289 we removing all references to frames (3 Feb 2005)
//	LONGLONG framesPerTime; PCN3289 we removing all references to frames (3 Feb 2005)

//functions
	Video(void);/* {width=300;height = 240;}  For cues AVI*/
	~Video(void);
	HRESULT showPropertyPage(IUnknown* pIU, const WCHAR* name);
//	HRESULT showPropertyPage(void);
	IBaseFilter *createsamp(void);
	HRESULT GetPin(IBaseFilter *pFilter, PIN_DIRECTION PinDir, IPin **ppPin);
	HRESULT ConnectFilters(IGraphBuilder *pGraph, IBaseFilter *pFirst, IBaseFilter *pSecond);
	IBaseFilter *LoadFile(char *FileName);
	HRESULT CaptureVideo(int live,int captureDevice,bool NoSync); //PCN2395 select capture device (21 Sept, Ant)
	HRESULT FindCaptureDevice(IBaseFilter ** ppSrcFilter, int captureDevice); //PCN2395 select caputure device (21 Sept, Ant)
	HRESULT GetInterfaces(void);
	void CloseInterfaces(int live);
	HRESULT SetupVideoWindow(void);
	void ResizeVideoWindow(void);
	double GetFrameRate(void);		//PCN3533
	void CalculateFrameRate(void);	//PCN3533
	HRESULT AddGraphToRot(IUnknown *pUnkGraph, DWORD *pdwRegister);
	void RemoveGraphFromRot(DWORD pdwRegister);
	HRESULT HandleGraphEvent(void);
	void specifyfile(const char *name);
	LONGLONG getTotalTime(void);

	double getRate(void);
	void setRate(double r);
//	void seekFrame(int frame);
	void approxSeekFrame(int frame);
	void pause(void);
	void run(void);
	void step(void);
	void FrameAdvance(void);
	void FrameRewind(void);
	bool Refresh(void);
	bool seekTime(double t);
	void grab(char *name, int registered, char *watermark, int fishEyeOn);
	//void getlastframe(void); PCN3289 not used and never would have worked properly
	int matend(void);
	void ffwd();
	void rwnd();
//	int getframe(); PCN3284
	bool recordprofileinfo;
	bool window;
	LONGLONG gettime();
	void GetDimensions(int *w, int *h);
	int getnumdisframes();

	void SetDeviceInput(void);
	void CheckCapabilities(void);
	bool IsVideoRunning(void);
	void VideoRun(bool state);

	void grablive(); // PCN2476
	//LONGLONG ConvertTimeToFrames(LONGLONG time); PCN3289

	HRESULT ConnectPins(IBaseFilter *pOne, int pinOneCount, IBaseFilter *pTwo, int pinTwoCount);
	bool AudioSample(short int &distance);
	void ResetLastRecordedTime(void);
	void FindFilter(void);
};


#endif
