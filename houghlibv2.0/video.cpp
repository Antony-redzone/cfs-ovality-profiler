//------------------------------------------------------------------------------
// File: video.cpp
//
// Desc: directshow capture stuff
//
//------------------------------------------------------------------------------

// <atlbase.h>
#include <atlbase.h>
#include <windows.h>

#include <stdio.h>
#include <iostream> // #include <iostream.h> PCN4896

#include "video.h"
#include "laserprofiler.h"
//#include "ovtool.h"
#include <comdef.h>
#include <fstream> //PCN4896
#include <dshow.h>





HANDLE DeadLockThreedHandle=0;

//#include <mtype.h>
using namespace std;
//const m_fShowCaptureProperties=0;
//char *FName="c:\\pipe5.avi";
//_COM_SMARTPTR_TYPEDEF(IOvMgr, __uuidof(IOvMgr));


// An application can advertise the existence of its filter graph
// by registering the graph with a global Running Object Table (ROT).
// The GraphEdit application can detect and remotely view the running
// filter graph, allowing you to 'spy' on the graph with GraphEdit.
//
// To enable registration in this sample, define REGISTER_FILTERGRAPH.
//
#define REGISTER_FILTERGRAPH

#define DEVICE_DEFAULT 0	// PCN2289
#define DEVICE_NVIDIA 1		// PCN2289

extern Laserprofiler *hw;


#define semaphore HANDLE
semaphore lockSB; //Lock the sample buffer

void signal(semaphore h) 
{

	if(h == lockSB) DeadLockThreedHandle = 0;


	ReleaseSemaphore(h,1,NULL);
}


unsigned long __stdcall DeadLockDetection(void *s) 
{
	int i;
	for(i=0;i<1000;i++)
	{
		Sleep(1);
		if (DeadLockThreedHandle == 0) 
		{
			return 0;
		}
	}
   if(DeadLockThreedHandle!=0) 
   {
		ReleaseSemaphore(lockSB,1,NULL);//signal(lockSB);
   }
   DeadLockThreedHandle = 0;
   return 0;
}


void wait(semaphore h) 
{
	unsigned long id;
	int i=0;

	WaitForSingleObject( h, MAXLONG);
  
	if(h == lockSB)
	{
    DeadLockThreedHandle = CreateThread(NULL,0, // security, stack size
							DeadLockDetection, // start
							(void *)i,0,&id); // param, creation flags, id
	}
}



semaphore create(int v) {
	return CreateSemaphore(NULL,(long)v, MAXLONG, NULL);
}



//PCN3258
//DEFINE_GUID(CLSID_SampleGrabber, 0x083863F1, 0x70DE, 0x11d0, 0xbd, 0x40, 0x00, 0xa0, 0xc9, 0x11, 0xce, 0x86);
//#define INITGUID	// have to define this so next line actually creates GUID structure
//DEFINE_GUID(CLSID_EMPGDMX, 0x136DCBF5, 0x3874, 0x4B70, 0xAE, 0x3E, 0x15, 0x99, 0x7D, 0x63, 0x34, 0xF7);
static const CLSID CLSID_EMPGDMX  = { 0x136DCBF5, 0x3874, 0x4B70, { 0xAE, 0x3E, 0x15, 0x99, 0x7D, 0x63, 0x34, 0xF7 } };
static const CLSID CLSID_FFDSHOW  = { 0x04FE9017, 0xF873, 0x410E, { 0x87, 0x1E, 0xAB, 0x91, 0x66, 0x1A, 0x4E, 0xF7 } };
static const CLSID CLSID_EM2VD    = { 0xBC4EB321, 0x771F, 0x4E9F, { 0xAF, 0x67, 0x37, 0xC6, 0x31, 0xEC, 0xA1, 0x06 } };

//			CLSID_MPEG2Demultiplexer						  
//
// Global data
//
DWORD g_dwGraphRegister=0;
IVideoFrameStep *g_pFS = NULL; // PCN2668 (3 March 2004, Antony van Iersel) Method for Frame Advance (Step)
IVideoWindow  * g_pVW = NULL;
IMediaControl * g_pMC = NULL;   //provides methods for controlling the flow of data through the filter graph
IMediaSeeking * g_pMS = NULL;   //
IMediaEventEx * g_pME = NULL;   //contains methods for retrieving event notifications and for overriding the Filter Graph Manager's default handling of events
								//adds methods that enable an application window to receive messages when events occur

IQualProp * iqprop = NULL;
ISampleGrabber *pGrabber = NULL;
IBaseFilter	   *pSampleGrabberFilter = 0;
IBaseFilter    *pMPEG2Demultiplexer;
IBaseFilter	   *pMPEG2Decoder;




IGraphBuilder * g_pGraph = NULL;
ICaptureGraphBuilder2 * g_pCapture = NULL;
PLAYSTATE g_psCurrent = Stopped;
//int pauseCount = 0;
//int videorunCount = 0;
//bool wasRunning;	// Flag to decide if graph needs to run after pausing to process image
bool amRefreshing = false;

bool videoRun = false;
bool amSeeking = false;


IBaseFilter *pSrcFilter=NULL;
IBaseFilter *pF=NULL;// PCN2289



class CGrabCB : public ISampleGrabberCB 
{
	
public:

    // These will get set by the main thread below. We need to
    // know this in order to write out the bmp
    long Width;
    long Height;

    // Fake out any COM ref counting
    //
    STDMETHODIMP_(ULONG) AddRef() { return 2; }
    STDMETHODIMP_(ULONG) Release() { return 1; }

    // Fake out any COM QI'ing
    //
    STDMETHODIMP QueryInterface(REFIID riid, void ** ppv)
    {
        if( riid == IID_ISampleGrabberCB || riid == IID_IUnknown ) 
        {
            *ppv = (void *) static_cast<ISampleGrabberCB*> ( this );
            return NOERROR;
        }    

        return E_NOINTERFACE;
    }


    // We don't implement this one
    //
 	STDMETHODIMP SampleCB(double SampleTime, IMediaSample *pSample)
		{
		return S_OK;
		}

	STDMETHODIMP BufferCB(double SampleTime, BYTE *pBuffer, long BufferLen)
		{
		hw->currentFrameGrabeTime = SampleTime;
		if(hw->mediaType!=0) return S_OK;
		if(!amSeeking) 
		{
			hw->movie->lastSeek=(long) (SampleTime*1000);
			hw->movie->pause(); // Allways Pause to process video.
		}
		
		hw->framecb(SampleTime,pBuffer);

		
		if(!amSeeking) if(videoRun==true) hw->movie->run();
		signal(lockSB);
		return S_OK;
	}

	
	
};

CGrabCB *cb;

///////////////////////////////////////////////////////////////////////
// showPropertyPage: Display the property page of a COM object
// Auguments:
//   pIU:    A interface of the COM object
//   name:   the name of the dialog box of the property page
// return value: error code
///////////////////////////////////////////////////////////////////////
HRESULT Video::showPropertyPage(IUnknown* pIU, const WCHAR* name)
{
    HRESULT hr=0;
    if (pIU) 
		{
        ISpecifyPropertyPages* pispp = 0;
		JIF(pIU->QueryInterface(IID_ISpecifyPropertyPages, (void **)&pispp));
        CAUUID caGUID;
        if (SUCCEEDED(pispp->GetPages(&caGUID))) 
			{
            OleCreatePropertyFrame(0, 0, 0,
                L"Setup",     // Caption for the dialog box
                1,        // Number of filters
                (IUnknown**)&pIU,
                caGUID.cElems,
                caGUID.pElems,
                0, 0, 0);
            // Release the memory
            CoTaskMemFree(caGUID.pElems);
			}
        SAFE_RELEASE(pispp);
		}
    return hr;
}


IBaseFilter *Video::createsamp(void) 
	{
	HRESULT hr;
//	IBaseFilter *pF = NULL; 
	
	//CoCreateInstance to create an instance of the Filter Graph Manager. 
	//If this call succeeds, DirectShow is installed on the machine
	hr = CoCreateInstance(CLSID_SampleGrabber, NULL, CLSCTX_INPROC_SERVER,
		IID_IBaseFilter, reinterpret_cast<void**>(&pF));
  if (FAILED(hr))
		{
        Msg(TEXT("1001 hr=0x%x"), hr);
        return NULL;
		}	

	//
	hr = pF->QueryInterface(IID_ISampleGrabber,reinterpret_cast<void**>(&pGrabber));
	hr = g_pGraph->AddFilter(pF, L"SampleGrabber");
	
	// Find the current bit depth.
	HDC hdc = GetDC(NULL);
//	int iBitDepth = GetDeviceCaps(hdc, BITSPIXEL);
	ReleaseDC(NULL, hdc);
	// Set the media type.
	AM_MEDIA_TYPE mt;
	ZeroMemory(&mt, sizeof(AM_MEDIA_TYPE));
	mt.majortype = MEDIATYPE_Video;
	mt.subtype = MEDIASUBTYPE_RGB24;
	hr = pGrabber->SetMediaType(&mt);
	hr = pGrabber->SetBufferSamples(TRUE);
	//CGrabCB *
	cb = new CGrabCB();  //create the ISampleGrabberCB object
	
	//PCN3533/////////////////////// Set to 1 is the samples into a buffer instead of getting
	pGrabber->SetCallback(cb, 1); // the data directly from the sample grabber array. This
								  // was originally set to 0, direct data (27th May 2005, Antony)
								  
	
	return pF;
	}

HRESULT Video::GetPin(IBaseFilter *pFilter, PIN_DIRECTION PinDir, IPin **ppPin)
	{
    IEnumPins  *pEnum;
    IPin       *pPin;
    pFilter->EnumPins(&pEnum);
    while(pEnum->Next(1, &pPin, 0) == S_OK)
		{
        PIN_DIRECTION PinDirThis;
        pPin->QueryDirection(&PinDirThis);
        if (PinDir == PinDirThis)
			{
            pEnum->Release();
            *ppPin = pPin;
            return S_OK;
			}
        pPin->Release();
		}
    pEnum->Release();
    return E_FAIL;  
	}

HRESULT Video::ConnectFilters(IGraphBuilder *pGraph, IBaseFilter *pFirst, IBaseFilter *pSecond)
	{
    IPin *pOut = NULL, *pIn = NULL;
    HRESULT hr = GetPin(pFirst, PINDIR_OUTPUT, &pOut);
    if (FAILED(hr)) return hr;
    hr = GetPin(pSecond, PINDIR_INPUT, &pIn);
    if (FAILED(hr)) 
		{
        pOut->Release();
        return E_FAIL;
		}
    hr = pGraph->Connect(pOut, pIn);
    pIn->Release();
    pOut->Release();
    return hr;
	}

IBaseFilter *Video::LoadFile(char *FileName) 
{
	HRESULT hr;
	IBaseFilter    *m_pIFileSource;

	// get the interface to the file source filter
	hr = CoCreateInstance((REFCLSID)CLSID_AsyncReader,
						   NULL, 
						   CLSCTX_INPROC_SERVER, 
						   IID_IBaseFilter,
						   (void**)&m_pIFileSource);
	if(FAILED(hr) || NULL == m_pIFileSource)
		{
		Msg("1002");
		return FALSE;
		};

	// add the filter to the graph
	hr = g_pGraph->AddFilter(m_pIFileSource, NULL);
	if(FAILED(hr))
		{
		Msg("1003");
		return FALSE;
		};
	IFileSourceFilter *pIFileSource;


	hr = m_pIFileSource->QueryInterface((REFIID)IID_IFileSourceFilter, 
										(void **) &pIFileSource);
	if (FAILED(hr))
		{
		return FALSE;
		};

	// Get the file name...if no file then exit
	if (FileName[0] == 0)
		{
		Msg("1004");
	//  HELPER_RELEASE(pIFileSource);
		return FALSE;
		};

	WCHAR wPath[MAX_PATH];
	MultiByteToWideChar( CP_ACP, 0, FileName, -1, wPath, MAX_PATH );

	// load the File Source (Async) with the filename.  If this step is not done
	// the filter does not create an output pin so we won't be able to connect
	hr = pIFileSource->Load(wPath, NULL);
	if( FAILED( hr ) )
		{
		Msg("1005" );
		return FALSE;
		}
	SAFE_RELEASE(pIFileSource);

	return m_pIFileSource;
}

HRESULT Video::CaptureVideo(int live, int captureDevice,bool NoSync) //PCN2395 select Capture Device (21 Sept Ant)
{
    HRESULT hr=0;
	//  IBaseFilter *pSrcFilter=NULL;
 //   IBaseFilter *pF=NULL; 

    // Get DirectShow interfaces
    hr = GetInterfaces();
    if (FAILED(hr))
		{
        Msg(TEXT("1006  hr=0x%x"), hr);
        return hr;
		}
    // Attach the filter graph to the capture graph
    
	hr = g_pCapture->SetFiltergraph(g_pGraph);
    if (FAILED(hr))
		{
        Msg(TEXT("1007  hr=0x%x"), hr);
        return hr;
		}

    // Use the system device enumerator and class enumerator to find
    // a video capture/preview device, such as a desktop USB video camera.
	if(live)
		{
		hw->movie->width=352;
		hw->movie->height=240;
	    hr = FindCaptureDevice(&pSrcFilter, captureDevice);
		if (FAILED(hr)) 
			{
	        return hr;
			}

		// Add Capture filter to our graph.
		hr = g_pGraph->AddFilter(pSrcFilter, L"Video Capture");
		if (FAILED(hr)) 
			{
			Msg(TEXT("1008  hr=0x%x"), hr);
			pSrcFilter->Release();
			return hr;
			}
		
		// Now decide the capturing config by the property page or set a 
		// default one. (i.e. Frame rate, resolution.....)
		CComPtr<IAMStreamConfig> pSC;  // Media Stream config interface
		// check capture device capabilities
		JIF(g_pCapture->FindInterface(&PIN_CATEGORY_CAPTURE,
									  &MEDIATYPE_Video, pSrcFilter, 
									  IID_IAMStreamConfig, 
									  (void **)&pSC));
	
		AM_MEDIA_TYPE *pmt;
		int iCount, iSize, ind;//,maxx=0,imax=0;
		// if (!m_fShowCaptureProperties) { // default capture setup
		// default capture: 320 * 240 resolution, 15f/s
		VIDEO_STREAM_CONFIG_CAPS caps;
		pSC->GetNumberOfCapabilities(&iCount, &iSize);
		for (ind = 0; ind < iCount; ind++) 
			{
			pSC->GetStreamCaps(ind, &pmt, (BYTE *)&caps);
// PCN2289 NVidia condition removed, Not needed, it was to change the default capture resolution
// from 640 x 480, to 320 to 240. But the Default was wrong. It was suppose to be 320 x 240. Now
// that is corrected condition no longer needed. (14 November 2003, Antony van Iersel)
//			if(deviceType==DEVICE_NVIDIA)	// PCN2289 , if nvidia chipset force the Capture Properties
//				{
//				VIDEOINFOHEADER *pvi = (VIDEOINFOHEADER *)pmt->pbFormat;
//				pvi->bmiHeader.biWidth=320;
//				pvi->bmiHeader.biHeight=240;
//				hw->movie->width=capturewidth;
//				hw->movie->height=captureheight;
//				pvi->AvgTimePerFrame = (LONGLONG)(10000000 / captureframerate);
//				pSC->SetFormat(pmt);
//				}
//			else // End removed Condition.
			if (pmt->formattype == FORMAT_VideoInfo && pmt->subtype == MEDIASUBTYPE_RGB24)
				{ 
				// maxx=caps.InputSize.cx;
				// imax=ind;
				// Set the capturing frame rate
				VIDEOINFOHEADER *pvi = (VIDEOINFOHEADER *)pmt->pbFormat;
				pvi->bmiHeader.biWidth=capturewidth;
				pvi->bmiHeader.biHeight=captureheight;
				hw->movie->width=capturewidth;
				hw->movie->height=captureheight;
				pvi->AvgTimePerFrame = (LONGLONG)(10000000 / captureframerate);
		
				///////////////////////////////////////////////////////////////////////////////////
				// PCN1967.  .... Antony van Iersel, 1 July 2003 ....
				// The SetFormat was returning a exception caught by JIF Macro because it was
				// trying to reset the Framerate if needed.
				// But this cant be set when using a USB capture Device ("Belkin, USB Videobus II")
				// The exception is not needed. 
				// Solution: Try to set it, if can't just egnore it.
				// 
				//			JIF(pSC->SetFormat(pmt));
				pSC->SetFormat(pmt);
				break; // found default setup, quit the loop				}
				}
			// if (m_fShowCaptureProperties) { 
			// showPropertyPage(pSC, L"Setup the capture");
			}
		}
	//start here	

	IEnumPins	*EnumPins=NULL;
	IPin		*Pin=NULL;
//	IBaseFilter *overlaymixerF=NULL;


	//test
	//	char temp[200];
	//	strcpy(temp, "z:\\Louise_Testing_19_12_02\\houghlib\\test3.avi");

	//end test
	
	

	if(!live)
		{

		char ext[]="ext";
		int stringLength;

		stringLength = strlen(vbfname);


		IBaseFilter    *FileSource=LoadFile((char *)vbfname);//temp);//FName);
		//FindFilter();
		
		if(stringLength>3) 
		{
			memcpy(ext,&vbfname[stringLength-3],sizeof(char)*4);
		
			if(_stricmp(ext,"mpg")==0)
			{
				if(pMPEG2Demultiplexer!=0) hr = g_pGraph->AddFilter(pMPEG2Demultiplexer, L"The Elecard MPEG Demultiplexer");
				if(pMPEG2Decoder!=0) 
				{
					hr = g_pGraph->AddFilter(pMPEG2Decoder, L"Ele Decoder");

//					HKEY hKey=0;				// Declare a key to store the result
//					DWORD dwValue = 0;

//					RegOpenKeyEx (HKEY_CURRENT_USER,"Software\\Elecard\\Elecard MPEG-2 Video Decoder HD\\Profiler.exe",NULL,KEY_WRITE,&hKey);
//					RegSetValueEx(hKey,            // subkey handle 
//								  "VMR maintain aspect ratio",         // value name 
//								  0,                       // must be zero 
//								 REG_DWORD,               // value type 
//								 (LPBYTE) &dwValue, // pointer to value data 
//								 sizeof(DWORD)) ;         // length of value data 
//					RegCloseKey(hKey);

//					RegOpenKeyEx (HKEY_CURRENT_USER,"Software\\Elecard\\Elecard MPEG-2 Video Decoder HD\\Default",NULL,KEY_WRITE,&hKey);
//					RegSetValueEx(hKey,            // subkey handle 
//								  "VMR maintain aspect ratio",         // value name 
//								  0,                       // must be zero 
//								 REG_DWORD,               // value type 
//								 (LPBYTE) &dwValue, // pointer to value data 
//								 sizeof(DWORD)) ;         // length of value data 
//					RegCloseKey(hKey);



				}

				
				HKEY hKey=0;				// Declare a key to store the result
				DWORD dwValue = 0;
				PHKEY phkResult=0;
				DWORD dwDisposition;


				//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
				// Adding elecard reg entry for ascpect ratio

				RegCreateKeyEx (HKEY_CURRENT_USER, "Software\\Elecard\\Elecard MPEG-2 Video Decoder HD\\Default", 0 , 0, REG_OPTION_NON_VOLATILE, KEY_SET_VALUE,NULL, &hKey, &dwDisposition);
				RegCloseKey(hKey);


				
				//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
				// Setting elecard aspect ratio

				RegOpenKeyEx (HKEY_CURRENT_USER,"Software\\Elecard\\Elecard MPEG-2 Video Decoder HD\\Profiler.exe",NULL,KEY_WRITE,&hKey);
				RegSetValueEx(hKey,            // subkey handle 
							  "VMR maintain aspect ratio",         // value name 
							  0,                       // must be zero 
							 REG_DWORD,               // value type 
							 (LPBYTE) &dwValue, // pointer to value data 
							 sizeof(DWORD)) ;         // length of value data 
				RegCloseKey(hKey);

				RegOpenKeyEx (HKEY_CURRENT_USER,"Software\\Elecard\\Elecard MPEG-2 Video Decoder HD\\Default",NULL,KEY_WRITE,&hKey);
				RegSetValueEx(hKey,            // subkey handle 
							  "VMR maintain aspect ratio",         // value name 
							  0,                       // must be zero 
							 REG_DWORD,               // value type 
							 (LPBYTE) &dwValue, // pointer to value data 
							 sizeof(DWORD)) ;         // length of value data 
				RegCloseKey(hKey);

				///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
				// Setting multiplexor index mode

				dwValue = 3;
				RegOpenKeyEx (HKEY_CURRENT_USER,"Software\\Elecard\\Elecard MPEG Demultiplexer\\Profiler.exe",NULL,KEY_WRITE,&hKey);
				RegSetValueEx(hKey,            // subkey handle 
							  "Indexing mode",         // value name 
							  0,                       // must be zero 
							 REG_DWORD,               // value type 
							 (LPBYTE) &dwValue, // pointer to value data 
							 sizeof(DWORD)) ;         // length of value data 
				RegCloseKey(hKey);

				RegOpenKeyEx (HKEY_CURRENT_USER,"Software\\Elecard\\Elecard MPEG Demultiplexer\\Default",NULL,KEY_WRITE,&hKey);
				RegSetValueEx(hKey,            // subkey handle 
							  "Indexing mode",         // value name 
							  0,                       // must be zero 
							 REG_DWORD,               // value type 
							 (LPBYTE) &dwValue, // pointer to value data 
							 sizeof(DWORD)) ;         // length of value data 
				RegCloseKey(hKey);


			}
			
		}
		//hr = g_pGraph->AddFilter(pMPEG2Decoder, L"FFDSHOW Decoder");

		



		pF=createsamp();
		hr=ConnectFilters(g_pGraph,FileSource,pF);
		SAFE_RELEASE(FileSource);

		}
	else {
		pF=createsamp();
		hr=ConnectFilters(g_pGraph,pSrcFilter,pF);
// PCN2289, This condition for a NVidia chipset has been removed, This setup will work for all
// capture devices. (Needs to be tested propery to be made sure) 14 November 2003, Antony van Iersel
//
//		if(deviceType==DEVICE_NVIDIA) //PCN2289, changes of NVidia Chipset can be used for all
//			{						  // Capture Chipset, no special setting now needed in VB
			//////////////////////////////////////////////////////////////////////////////////
			// PCN2289																		//
			// This whole condition is for the nvidia capture chipset, it needs something,	//
			// anything connected to the preview pin to work, even if its not used			//
			//	if(deviceType==DEVICE_NVIDIA) 
			//		{
			//		hr=g_pGraph->AddFilter(overlaymixerF,L"Overlay Mixer");
			//		if (FAILED(hr)) { Msg(TEXT("1009aaa  hr=0x%x"), hr); pSrcFilter->Release(); return hr; }
			//  	hr=ConnectFilters(g_pGraph,pSrcFilter,overlaymixerF);
			//		if (FAILED(hr)) { Msg(TEXT("1009a  hr=0x%x"), hr);   pSrcFilter->Release(); return hr; }
			
//			const struct _GUID pintype=PIN_CATEGORY_PREVIEW; // With the RenderStream 1st parameter, when
//			hr=ConnectFilters(g_pGraph,pSrcFilter,pF);       // set to NULL, it finds the next pin. Was told PIN_CATEGORY_PRIVIEW
		hr=g_pCapture->RenderStream(NULL, NULL,pSrcFilter,NULL,NULL); // PCN 2289  Leadtek Error
		if (FAILED(hr)) { Msg(TEXT("1018  hr=0x%x"), hr); pSrcFilter->Release(); return hr; } //PCN2428 error code was 1009, but that already used
																							  //so now its 1018. (21 Nov 2003 , Antony van Iersel)
		IBaseFilter *pRender;				
		g_pGraph->FindFilterByName(L"Video Renderer" , &pRender); // PCN2289, this added Later
		g_pGraph->RemoveFilter(pRender); 						  // to rid of secound render Window
		SAFE_RELEASE(pRender);														  // 14 November 2003, Antony van Iersel
			////////////////////////////////////////////////////////////////////////////////////////////////
//			} End of removed condition.
		}
	//end here

	pF->EnumPins(&EnumPins);
	EnumPins->Reset();
	EnumPins->Skip(1);
	EnumPins->Next(1, &Pin, NULL);
	EnumPins->Release();

	g_pGraph->Render(Pin);	// graphbuilder completes graph



    // Render the preview pin on the video capture filter
    // Use this instead of g_pGraph->RenderFile
    if (FAILED(hr))
		{
        Msg(TEXT("Codec required - hr=0x%x"), hr);
        //pSrcFilter->Release();
		return hr;
		}
	//	hr=Pin->QueryInterface(IID_IAMDroppedFrames,(void **)&droppedFrames);
	//	if (FAILED(hr))
	//		{
	//       Msg(TEXT("Couldn't set up Dropped Frames Interface"));
	//		Pin->Release();
	//		}
	
	//	hr = g_pCapture->FindInterface(NULL,NULL,pSrcFilter,IID_IAMDroppedFrames,(void**)&droppedFrames);
	//	if (FAILED(hr))
	//		{
	//      Msg(TEXT("Failed to find dropped frames interface"));
	//		}
    // Now that the filter has been added to the graph and we have
    // rendered its stream, we can release this reference to the filter.
	//	if(!usefile)  pSrcFilter->Release();

	//Martin's addition 29/1/03
	long x=0,y=0;
    IBasicVideo *pBasicVideo = NULL; 
    if (g_pGraph) hr = g_pGraph->QueryInterface(IID_IBasicVideo, (void **) &pBasicVideo);
    if (SUCCEEDED(hr)) 
		{
        hr=pBasicVideo->get_SourceWidth(&x);
        hr=pBasicVideo->get_SourceHeight(&y);
        pBasicVideo->Release();
		}
    width=x;
    height=y;

	//end Martin's addition

	
	if(!live)
	{

		if(NoSync)
		{
 			//PCN3533 , (Moandy 23 May 2005, Antony), nothing to do with video deadlock
			//but took of the vidoe time syncing so if the computer is fast enough it will
			//process faster that realtime.
			///////////////////////////////////////////////////////////////////   
			IMediaFilter *pMediaFilter = 0;
			g_pGraph->QueryInterface(IID_IMediaFilter, (void**)&pMediaFilter);
			pMediaFilter->SetSyncSource(NULL);
			pMediaFilter->Release();
			///////////////////////////////////////////////////////////////////
		}
	}


    hr = SetupVideoWindow();
    if (FAILED(hr))
		{
        Msg(TEXT("1010  hr=0x%x"), hr);
        return hr;
		}
	
	// Add our graph to the running object table, which will allow
    // the GraphEdit application to "spy" on our graph


	
#ifdef REGISTER_FILTERGRAPH
    hr = AddGraphToRot(g_pGraph, &g_dwGraphRegister);
   if (FAILED(hr))
		{
        Msg(TEXT("1011  hr=0x%x"), hr);
        g_dwGraphRegister = 0;
		}
#endif
// PCN3289 It should never have started running the graph at this point
// (10 Feb 2005, Antony)
//   // Start previewing video data
//	hr = g_pMC->Run();
//	if (FAILED(hr))
//		{
//      Msg(TEXT("1012  hr=0x%x"), hr);
//    return hr;
//		}

//    // Remember current state
//    g_psCurrent = Stopped;
	//	window = true;
	//	if(window == true)
	//		{
	//		hr = hw->movie->showPropertyPage(pF, L"Adjust Camera");
	SAFE_RELEASE(Pin);
	videoLengthTime = getTotalTime(); //PCN3289
	CalculateFrameRate();
    return S_OK;
}

HRESULT Video::FindCaptureDevice(IBaseFilter ** ppSrcFilter,int captureDevice)
	{
    HRESULT hr;
    IBaseFilter * pSrc = NULL;
    IMoniker *pMoniker = 0; // PCN2395 was CComPtr <IMoniker> pMoniker = NULL
    //ULONG cFetched;
   
    // Create the system device enumerator
    CComPtr <ICreateDevEnum> pDevEnum =NULL;

	//Creates a single uninitialized object of the class associated with a specified CLSID
    hr = CoCreateInstance (CLSID_SystemDeviceEnum, 
						   NULL, 
						   CLSCTX_INPROC,
						   IID_ICreateDevEnum, 
						   (void ** ) &pDevEnum);
    if (FAILED(hr))
		{
        Msg(TEXT("1013  hr=0x%x"), hr);
        return hr;
		}

    // Create an enumerator for the video capture devices
    CComPtr <IEnumMoniker> pClassEnum = NULL;

    hr = pDevEnum->CreateClassEnumerator (CLSID_VideoInputDeviceCategory, &pClassEnum, 0);
    if (FAILED(hr))
		{
        Msg(TEXT("1013  hr=0x%x"), hr);
        return hr;
		}

    // If there are no enumerators for the requested type, then 
    // CreateClassEnumerator will succeed, but pClassEnum will be NULL.
    if (pClassEnum == NULL)
		{
        Msg(TEXT("1014"));
        return E_FAIL;
		}

    // Use the first video capture device on the device list.
    // Note that if the Next() call succeeds but there are no monikers,
    // it will return S_FALSE (which is not a failure).  Therefore, we
    // check that the return code is S_OK instead of using SUCCEEDED() macro.
	//pClassEnum->Next (1, &pMoniker, &cFetched);

	//PCN2395, above is wrong, we no longer find the first capture device, we find
	// the nth capture device decided by the VB pass of captureDevice

	int numberCaptureDevices=0;

	while (pClassEnum->Next(1, &pMoniker, NULL) == S_OK) //PCN2395
//    if (S_OK == (pClassEnum->Next (1, &pMoniker, &cFetched)))
		{
        // Bind Moniker to a filter object
        hr = pMoniker->BindToObject(0,0,IID_IBaseFilter, (void**)&pSrc);

		//PCN2395vvvvvvvvvvvvvvv
        if (FAILED(hr))
			{
			pMoniker->Release();
			continue;
			}
		if(numberCaptureDevices==captureDevice) break;
		numberCaptureDevices++;
		//^^^^^^^^^^^^^^^^^^^^^^
		pMoniker->Release();
//	}
//   else {
//		Msg(TEXT("1016"));   
 //     return E_FAIL;
		}

    // Copy the found filter pointer to the output parameter.
    // Do NOT Release() the reference, since it will still be used
    // by the calling function.
    *ppSrcFilter = pSrc;


    return hr;
}


HRESULT Video::GetInterfaces(void)
{
    HRESULT hr;

    // Create the filter graph
    hr = CoCreateInstance (CLSID_FilterGraph, 
						   NULL, 
						   CLSCTX_INPROC,
						   IID_IGraphBuilder, 
						   (void **) &g_pGraph);
    if (FAILED(hr)) 
		{
		return hr;
		}
    // Create the capture graph builder
    hr = CoCreateInstance (CLSID_CaptureGraphBuilder2 , 
						   NULL, 
						   CLSCTX_INPROC,
						   IID_ICaptureGraphBuilder2, 
						   (void **) &g_pCapture);
    if (FAILED(hr)) 
		{
		return hr;
		}

	hr = CoCreateInstance(CLSID_EMPGDMX, 
						  NULL, 
						  CLSCTX_INPROC_SERVER, 
						  IID_IBaseFilter, (void**)&pMPEG2Demultiplexer);
	if (FAILED(hr)) __asm nop;

	hr = CoCreateInstance(CLSID_EM2VD,
						  NULL,
						  CLSCTX_INPROC_SERVER,
						  IID_IBaseFilter, (void**)&pMPEG2Decoder);
	if (FAILED(hr)) __asm nop;
    


	// Obtain interfaces for media control and Video Window
    hr = g_pGraph->QueryInterface(IID_IMediaControl, (LPVOID *) &g_pMC);
    if (FAILED(hr)) 
		{
		return hr;
		}

	hr = g_pGraph->QueryInterface(IID_IVideoWindow, (LPVOID *) &g_pVW);
	if (FAILED(hr)) 
		{
		return hr;
		}

	hr = g_pGraph->QueryInterface(IID_IMediaEvent, (LPVOID *) &g_pME);
	if (FAILED(hr)) 
		{
		return hr;
		}
    
	hr = g_pGraph->QueryInterface(IID_IMediaSeeking, (LPVOID *) &g_pMS);
    if (FAILED(hr))
		{
        return hr;
		}

	//PCN2668 (3 March 2004, Antony van Iersel) Method to allow frame advance (Step by 1 frame)
	hr = g_pGraph->QueryInterface(__uuidof(IVideoFrameStep), (PVOID *)&g_pFS); //
	if (FAILED(hr)) /////////////////////////////////////////////////////////////
		{			//
		return hr;	//
		}			//
	//////////////////
	
	
	//	hr = g_pGraph->QueryInterface(IID_IQualProp, (LPVOID *) &iqprop);
	//  if (FAILED(hr))
	//		{
	//      Msg(TEXT("Failed to get Quality interface"));
	//		return hr;
	//		}

    // Set the window handle used to process graph events
    hr = g_pME->SetNotifyWindow((OAHWND)hwnd, WM_GRAPHNOTIFY, 0);
    return hr;
}

// Close the graph and tidy up.
void Video::CloseInterfaces(int live)
{
    // Stop previewing data
    if (g_pMC) g_pMC->StopWhenReady();

    g_psCurrent = Stopped;

    // Stop receiving events
    if (g_pME) g_pME->SetNotifyWindow(NULL, WM_GRAPHNOTIFY, 0);

    // Relinquish ownership (IMPORTANT!) of the video window.
    // Failing to call put_Owner can lead to assert failures within
    // the video renderer, as it still assumes that it has a valid
    // parent window.
    if(g_pVW)
		{
        g_pVW->put_Visible(OAFALSE);
        g_pVW->put_Owner(NULL);
		}

#ifdef REGISTER_FILTERGRAPH
    // Remove filter graph from the running object table   
    if (g_dwGraphRegister)
		{
        RemoveGraphFromRot(g_dwGraphRegister);
		}
#endif
    // Release DirectShow interfaces
	if(live){ SAFE_RELEASE(pSrcFilter); }

	//---------------------------------------------
	
g_pMC->Stop();
g_psCurrent = Stopped;

//g_pGraph->Release();


// Enumerate the filters in the graph.

IEnumFilters *pEnum = NULL;
HRESULT hr = g_pGraph->EnumFilters(&pEnum);
if (SUCCEEDED(hr))
{
    IBaseFilter *pFilter = NULL;
    while (S_OK == pEnum->Next(1, &pFilter, NULL))
     {
         // Remove the filter.
         g_pGraph->RemoveFilter(pFilter);
         // Reset the enumerator.
         pEnum->Reset();
         pFilter->Release();
    }
    pEnum->Release();
}

	SAFE_RELEASE(g_pGraph);
	SAFE_RELEASE(iqprop);
	SAFE_RELEASE(g_pMS);
	SAFE_RELEASE(g_pMC);
	SAFE_RELEASE(g_pME);
	SAFE_RELEASE(g_pVW);

	SAFE_RELEASE(g_pCapture);
	SAFE_RELEASE(g_pFS);
	SAFE_RELEASE(pSrcFilter);
	SAFE_RELEASE(pF);
	SAFE_RELEASE(pGrabber);
	delete cb;
}


// Sets the video overlay parent
HRESULT Video::SetupVideoWindow(void)
{
    HRESULT hr;

    // Set the video window to be a child of the main window
    hr = g_pVW->put_Owner((OAHWND)hwnd);
    if (FAILED(hr))
		{
        return hr;
		}
    
	// Set video window style
    hr = g_pVW->put_WindowStyle(WS_CHILD | WS_CLIPCHILDREN);
    if (FAILED(hr))
		{
        return hr;
		}
    // Use helper function to position video window in client rect 
    // of main application window
    ResizeVideoWindow();

    // Make the video window visible, now that it is properly positioned
    hr = g_pVW->put_Visible(OATRUE);
    if (FAILED(hr))
		{
        return hr; //Msg(TEXT("Failed! hr = %x"), hr);
		}
	hr = g_pVW->put_MessageDrain((OAHWND)hwnd);	////////// PCN4380
    if (FAILED(hr))										//
		{												//
        return hr; //Msg(TEXT("Failed! hr = %x"), hr);	//
		}												//
	//////////////////////////////////////////////////////
    return hr;															
}

// Returns the length of the video by the time in 100 nanosecounds
LONGLONG Video::getTotalTime(void) 
{
	LONGLONG g_rtTotalTime;

	g_pMS->SetTimeFormat(&TIME_FORMAT_MEDIA_TIME);
	g_pMS->GetDuration(&g_rtTotalTime);
	return g_rtTotalTime;
}

double Video::GetFrameRate(void)
{
	return frameRate;
}

void Video::CalculateFrameRate(void)
{
	LONGLONG totalTime = getTotalTime();
	LONGLONG totalFrames;

	g_pMS->SetTimeFormat(&TIME_FORMAT_FRAME);
	g_pMS->GetDuration(&totalFrames);
	g_pMS->SetTimeFormat(&TIME_FORMAT_MEDIA_TIME);

	frameRate = (double) totalFrames / ((double) totalTime / 10000000);
}



// Pause Graph
void Video::pause(void)
{
	g_pMC->Pause();
	g_psCurrent = Paused;
}

// Run Graph, if a video file,the video will start playing
void Video::run(void) 
{
	g_pMC->Run();
	g_psCurrent = Running; //PCN3289

}



// ( PCN2668 )------------------------------------------------------
//
// Name		: Vidoe::step
// Created	: 3 March 2004
// Updated	:
// Prg By	: Antony van Iersel

// Desc		: Steps and video one frame foward.
// Usage	: This is to replace playing the video when recording the PVD.
//			  Instead the graph set to play and trying to prosses as many frames
//			  as possible and missing the ones it can't keep up with. It will pause the
//            video and frame advance after each profile is processed, this giving every
//			  frame in the video a profile.
// Inputs	: none
// Output	: none

void Video::step()
{
	// The graph must be paused for frame stepping to work, if not paused, pause it.
   // if (g_psCurrent != Paused) pause();
    pause();
	// Step just 1 frame foward. Possible for than 1 frame or less the 1 frame but
	// not always supported by the graph filters. It is not needed for this application.
	// 1 Frame is Fine
    g_pFS->Step(1, NULL);
}

// ( PCN2865 )------------------------------------------------------
//
// Name		: Video::FrameAdvance
// Created	: 2 June 2004
// Updated	:
// Prg By	: Antony van Iersel

// Desc		: Steps and video one frame foward.
// Usage	  This is more stable than video step. Vidoe step is good for recording PVD
//			  but not very good for advance a single frame, it will cause the video to play
//			  sometimes.
// Inputs	: none
// Output	: none
void Video::FrameAdvance(void)
{
	seekTime(hw->currentFrameGrabeTime+(1/frameRate));
}

// ( PCN????)------------------------------------------------------
//
// Name		: Video::FrameRewind
// Created	: 4 September 2006
// Updated	:
// Prg By	: Antony van Iersel

// Desc		: Steps and video one frame backwards
// Usage	  This is more stable than video step. Vidoe step is good for recording PVD
//			  but not very good for advance a single frame, it will cause the video to play
//			  sometimes.
// Inputs	: none
// Output	: none
void Video::FrameRewind(void)
{
	seekTime(hw->currentFrameGrabeTime-(1/frameRate));
}

// ( PCN???? )------------------------------------------------------
//
// Name		: Video::Refresh
// Created	: 11 June 2004
// Updated	:
// Prg By	: Antony van Iersel

// Desc		: Refresh the Video FrameGrab.
// Usage	: Refreshes the video by calling the frame number and seeking to exactly the same  
//			  position. Replaces and bad refresh that called frameseek.
// Inputs	: none
// Output	: none

bool Video::Refresh(void)
{
	if(videoRun==true) return false; // If graph is running or was running and
									 // in the middle of procssing profile then there is no need to
									 // refresh the video is or will be in running mode and will
									 // refresh on the next frame.
	return seekTime(hw->currentFrameGrabeTime);
}


// Get number of droped frames, not used
int Video::getnumdisframes()
{
	int i=-1;
	HRESULT hr;
	hr = iqprop->get_FramesDroppedInRenderer(&i);
	if (FAILED(hr))
		{
        Msg(TEXT("1017 hr = %x"), hr);
		}
	return i;
}

// Seek to a place in the video index by nearest frame
void Video::approxSeekFrame(int frame) 
{
	LONGLONG newTime=frame;
	
	g_pMS->SetTimeFormat(&TIME_FORMAT_FRAME);
	g_pMS->SetPositions(&newTime, 
						AM_SEEKING_AbsolutePositioning,
						NULL,
						AM_SEEKING_SeekToKeyFrame);
}

// Seek to a place in the video index by time
bool Video::seekTime(double t)
{
	 //PCN3553 /////////////////////

	//	if(videoRun) return false;
//	FILE *logger; //PCN3251
	if(amSeeking) return false; //Trying to prevent lockup, if its allready seeking then dont try
//	logger=fopen("C:\\LookingForALockUp","a+");	fprintf(logger, "About to do the first seek\n"); fclose(logger);
	wait(lockSB);
//	logger=fopen("C:\\LookingForALockUp","a+");	fprintf(logger, "past the first wait\n"); fclose(logger);

	amSeeking=true;

	////////////////////////////////
	if(t<0) t=0;
	if(t>(double) videoLengthTime/10000000) t = (double) videoLengthTime/10000000;
	g_pMS->SetTimeFormat(&TIME_FORMAT_MEDIA_TIME);
	
	LONGLONG newTime= LONGLONG(t * 10000000); // convert s to nanoseconds.

	g_pMS->SetPositions(&newTime, 
						AM_SEEKING_AbsolutePositioning,
						NULL, 
						AM_SEEKING_NoPositioning);
	Sleep(100);

//	logger=fopen("C:\\LookingForALockUp","a+");	fprintf(logger, "About to do the second seek\n"); fclose(logger);
	wait(lockSB);
//	logger=fopen("C:\\LookingForALockUp","a+");	fprintf(logger, "past the second seek\n"); fclose(logger);
	amSeeking=false;
	signal(lockSB);

	return true;
}

// Sets the video playback overlay size
void Video::ResizeVideoWindow(void)
{
    RECT rc;

    // Make the preview video fill our window
    GetClientRect(hwnd, &rc);

    // Resize the video preview window to match owner window size
    if (g_pVW) g_pVW->SetWindowPosition(0, 0, rc.right, rc.bottom);
}



#ifdef REGISTER_FILTERGRAPH

// Add graph to GraphEdit
HRESULT Video::AddGraphToRot(IUnknown *pUnkGraph, DWORD *pdwRegister) 
{
    IMoniker * pMoniker;
    IRunningObjectTable *pROT;
    WCHAR wsz[128];
    HRESULT hr;

    if (FAILED(GetRunningObjectTable(0, &pROT))) 
		{
        return E_FAIL;
		}

    wsprintfW(wsz, 
			  L"FilterGraph %08x pid %08x", 
			  (DWORD_PTR)pUnkGraph, GetCurrentProcessId());
    hr = CreateItemMoniker(L"!", wsz, &pMoniker);
    if (SUCCEEDED(hr)) 
		{
        hr = pROT->Register(0, pUnkGraph, pMoniker, pdwRegister);
        pMoniker->Release();
		}
    pROT->Release();
    return hr;
}


// Remove the graph from GraphEdit for debugging
void Video::RemoveGraphFromRot(DWORD pdwRegister)
{
    IRunningObjectTable *pROT;

    if (SUCCEEDED(GetRunningObjectTable(0, &pROT))) 
		{
        pROT->Revoke(pdwRegister);
        pROT->Release();
		}
}

#endif

// Msg display
void Msg(TCHAR *szFormat, ...)
	{
    TCHAR szBuffer[512];

    va_list pArgs;
    va_start(pArgs, szFormat);
    _vstprintf(szBuffer, szFormat, pArgs);
    va_end(pArgs);

    MessageBox(NULL, szBuffer, TEXT("LaserProfiler Message"), MB_OK | MB_ICONERROR);
	}


// Event handler for FrameGrab
HRESULT Video::HandleGraphEvent(void)
{
    LONG evCode, evParam1, evParam2;
    HRESULT hr=S_OK;

    while(SUCCEEDED(g_pME->GetEvent(&evCode, 
									(LONG_PTR *) &evParam1, 
									(LONG_PTR *) &evParam2, 
									0)))
		{
        //
        // Free event parameters to prevent memory leaks associated with
        // event parameter data.  While this application is not interested
        // in the received events, applications should always process them.
        //
        hr = g_pME->FreeEventParams(evCode, evParam1, evParam2);
        
        // Insert event processing code here, if desired
		}

    return hr;
}

// Sets up video file to play
void Video::specifyfile(const char *name)
{
	vbfname = new char[256];
	strcpy(vbfname, name);//"z:\\Louise_Testing_19_12_02\\houghlib\\test3.avi");
}

// Grab a video frame
void Video::grab(char *name, int registered, char *watermark,int fishEyeOn) //PCN4596 0 = fishEye Off, 1 = fisheEye On
{

	pixel **fishEyeBuffer;
	int i,j;

	HRESULT hr;
	AM_MEDIA_TYPE MediaType; 
	ZeroMemory(&MediaType,sizeof(MediaType));
	hr = pGrabber->GetConnectedMediaType(&MediaType); 

	// Get a pointer to the video header. 
	VIDEOINFOHEADER *pVideoHeader = (VIDEOINFOHEADER*)MediaType.pbFormat; 

	// The video header contains the bitmap information. 
	// Copy it into a BITMAPINFO structure. 
	BITMAPINFO BitmapInfo; 
	ZeroMemory(&BitmapInfo, sizeof(BitmapInfo)); 
	CopyMemory(&BitmapInfo.bmiHeader, &(pVideoHeader->bmiHeader), 
	sizeof(BITMAPINFOHEADER));
		
	char *buffer = NULL; 
	// Copy the image into the buffer.
	long size = 0;
	hr = pGrabber->GetCurrentBuffer(&size,NULL);
	buffer=new char[size];
	hr = pGrabber->GetCurrentBuffer(&size,(long *)buffer);


	
	//if(hw->FEye.FishEyeStatus() == ON) PCN4596 Fisheye is now never ever off, only a flag for Snapshot
	//1 to snap with fishe eye  calculation on, 0 to egnore fisheye calculations
	if((fishEyeOn==1) && (hw->FEye.FishEyeStatus() == ON)) 
	{
		fishEyeBuffer = new pixel * [hw->movie->height];
		for(i=0;i<hw->movie->height;i++) fishEyeBuffer[i] = new pixel[hw->movie->width];
		
		for ( i = 0 ; i < hw->movie->width ; i++ ) 
			for ( j = 0 ; j < hw->movie->height ; j++ )
			{
				fishEyeBuffer[j][i].blue  = (unsigned char) buffer[(i + j * hw->movie->width) * 3];      // blue;
				fishEyeBuffer[j][i].green = (unsigned char) buffer[(i + j * hw->movie->width) * 3 + 1];  // green;
				fishEyeBuffer[j][i].red   = (unsigned char) buffer[(i + j * hw->movie->width) * 3 + 2];  // red;
			}

		hw->FEye.Transform(fishEyeBuffer);
		hw->FEye.CopyToVideo(fishEyeBuffer,hw->movie->width,hw->movie->height);

		for ( i = 0 ; i < hw->movie->width ; i++ ) 
			for ( j = 0 ; j < hw->movie->height ; j++ )
			{
				buffer[(i + j * hw->movie->width) * 3] =     (char) fishEyeBuffer[j][i].blue;  // blue;
				buffer[(i + j * hw->movie->width) * 3 + 1] = (char) fishEyeBuffer[j][i].green; // green;
				buffer[(i + j * hw->movie->width) * 3 + 2] = (char) fishEyeBuffer[j][i].red;   // red;
			}
		for(i=0;i<hw->movie->height;i++) delete[] fishEyeBuffer[i];
		delete[] fishEyeBuffer;
	}

	if(registered==0)
		{
		for(long y=(long)(((double)size/4.0)*2.5);y<(long)(size/4.0)*3;y++)
			{
			buffer[y]=(char)128;
			}
		}

	ofstream *BitMapFile=new ofstream(name,ios::binary); //E:\Documents and Settings\LouiseS\Desktop\Snapshot131.bmp",ios::binary);
	BITMAPFILEHEADER BM_Header;
	BM_Header.bfType = ((WORD) ('M' << 8) | 'B');
	BM_Header.bfSize = sizeof(BitmapInfo.bmiHeader) + size + sizeof( BM_Header );
	BM_Header.bfReserved1 = 0;
	BM_Header.bfReserved2 = 0;
	BM_Header.bfOffBits = (DWORD)(sizeof(BM_Header) + sizeof(BitmapInfo.bmiHeader));
	BitMapFile->write((char *)&BM_Header, sizeof(BM_Header));
	BitMapFile->write((char *)&BitmapInfo.bmiHeader, sizeof(BitmapInfo.bmiHeader));
	BitMapFile->write((char *)buffer, size );
	BitMapFile->close();

	if(buffer!=NULL) {delete[] buffer; buffer=NULL;} //PCN2561 Antony van Iersel 23 March 2004 (Memory cleanup)
	delete BitMapFile;
} 

// PCN3289 never used and would not have worked properly
// Gets the last frame, will not work. Uses the wrong function.
//void Video::getlastframe(void)
//{
//	LONGLONG stoptime;
//	g_pMS->GetStopPosition(&stoptime);
//	g_pMS->SetPositions(&stoptime, AM_SEEKING_AbsolutePositioning,NULL, AM_SEEKING_NoPositioning);
//}

int Video::matend(void)
{
	//returns 1 if at the end of the movie, 0 otherwise
	LONGLONG stoptime;
	g_pMS->GetStopPosition(&stoptime);
	LONGLONG currtime;
	g_pMS->GetCurrentPosition(&currtime);
	if(stoptime == currtime) return 1;
	return 0;
}

// Increae play speed
void Video::ffwd() 
{
	//increases speed
	double currRate;
	g_pMS->GetRate(&currRate);
	if(currRate == 1)
		{ 
		g_pMS->SetRate(4.0);
		}
	else{
		g_pMS->SetRate(1.0);
		}
}

// Get current playback speed
double Video::getRate(void)
{
	double currRate;
	g_pMS->GetRate(&currRate);
	return currRate;
}

// Set video speed, 0.5 half speed, 2 double speed, 1 normal speed etc
void Video::setRate(double r)
{
	g_pMS->SetRate(r);
}

// Play in reverese
void Video::rwnd() 
{
	//slows speed
	double currRate;
	g_pMS->GetRate(&currRate);
	if(currRate == 1) g_pMS->SetRate(-1.0);
	else g_pMS->SetRate(1.0);
}

// No longer to be used PCN3289 (3 Feb 2005)
// Get current Frame
//
//int Video::getframe() 
//{
//	HRESULT error = S_OK;
//	LONGLONG currTimePos;
//	LONGLONG convertedTimeToFrame;
//
//	g_pMS->SetTimeFormat(&TIME_FORMAT_MEDIA_TIME);
//	error = g_pMS->GetCurrentPosition(&currTimePos);
//	if(error == E_INVALIDARG) Msg("GetCurrentPosition, Invalid");
//	if(error == E_NOTIMPL) Msg("GetCurrentPosition, Method is not supported");
//	if(error == E_POINTER) Msg("GetCurrentPosition, NULL pointer argument");
//	if(error == S_OK) Msg("GetCurrentPosition Time, Success");
//
//
//
//	convertedTimeToFrame = ConvertTimeToFrames(currTimePos);
//	Msg("Current Frame  = %i",(int) convertedTimeToFrame);
//
//	return (int)convertedTimeToFrame;//currFramePos;
//}

// Get current Time
LONGLONG Video::gettime()
{


	LONGLONG time;

//	g_pMS->SetTimeFormat(&TIME_FORMAT_MEDIA_TIME);
	g_pMS->GetCurrentPosition(&time);
	return time;
}

// Get Video width and height
void Video::GetDimensions(int *w, int *h)
{
	*w = width;
	*h = height;
}

// Video Select( PCN2326 )------------------------------------------------------^
//
// Name		: SetDeviceInput
// Created	: 31 October 2003, PCNPCN2326
// Updated	:
// Prg By	: Antony van Iersel
// Param	: 
// Desc		: Calls the SetVideoInput Function in Video.CPP
// Usage	: When ever the input from the capture device incorect,
//			: this function is called to correct the input.


// Sets the caputre device input
void Video::SetDeviceInput(void)
{
	IBaseFilter *pXF;


	// Calling the Crossbar Interface to Set the Video in Properties
	// eg. Composite, Tuner, SVHS etc
	g_pCapture->FindInterface(&PIN_CATEGORY_CAPTURE,
							  &MEDIATYPE_Video,
							  pSrcFilter,
							  IID_IAMCrossbar, 
							  (void**) &pXF);
	showPropertyPage(pXF, L"Capture Input Select"); // Calling the Property diaglog
}

void Video::CheckCapabilities(void)
{

	HRESULT error;

	bool	canSeekAbsolute = false, 
			canSeekForwards = false,
			canSeekBackwards= false,
			canGetCurrentPos= false,
			canGetStopPos   = false,
			canGetDuration  = false,
			canPlayBackwards= false,
			canDoSegments   = false,
			source          = false; 

	DWORD capabilities =
		AM_SEEKING_CanSeekAbsolute |
		AM_SEEKING_CanSeekForwards |
		AM_SEEKING_CanSeekBackwards |
		AM_SEEKING_CanGetCurrentPos |
		AM_SEEKING_CanGetStopPos |
		AM_SEEKING_CanGetDuration |
		AM_SEEKING_CanPlayBackwards | 
		AM_SEEKING_CanDoSegments  |
		AM_SEEKING_Source ;

	g_pMS->SetTimeFormat(&TIME_FORMAT_MEDIA_TIME);
	error = g_pMS->CheckCapabilities(&capabilities);
	

	if(capabilities & AM_SEEKING_CanSeekAbsolute) canSeekAbsolute=true;
	if(capabilities & AM_SEEKING_CanSeekForwards) canSeekForwards=true;
	if(capabilities & AM_SEEKING_CanSeekBackwards) canSeekBackwards=true;
	if(capabilities & AM_SEEKING_CanGetCurrentPos) canGetCurrentPos=true;
	if(capabilities & AM_SEEKING_CanGetStopPos) canGetStopPos=true;
	if(capabilities & AM_SEEKING_CanGetDuration) canGetDuration=true;
	if(capabilities & AM_SEEKING_CanPlayBackwards) canPlayBackwards=true;
	if(capabilities & AM_SEEKING_CanDoSegments) canDoSegments=true;
	if(capabilities & AM_SEEKING_Source) source=true;


	if(error == S_FALSE) Msg("All video access ok\nAbsolute %i\nSeekFowards %i\nSeekBackwards %i\n GetCurrentPos %i\n GetStopPos %i\nGetDuration %i\nPlayBackwards %i\nDoSegments %i\nsource %i,"
		,canSeekAbsolute, 
		canSeekForwards,
		canSeekBackwards,
		canGetCurrentPos,
		canGetStopPos,
		canGetDuration,
		canPlayBackwards,
		canDoSegments,
		source);
	
	if(error == E_FAIL) Msg("No video access available");
	if(error == S_OK) Msg("Not all video access available");
}
// PCN3289 Short lived, not going to be used (3 Feb 2005, Antony)
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
//  PCN3289
//	Name	: Video::ConvertTimeToFrames
//	Created	: 03 Feb 2005
//	By		: Antony van Iersel
//	Desc    : Some codecs don't like retrieving the current frame number, this is been replaced
//			: with finding the current time and converting it to frames manually.
//	Usage   : Pass the time that is needed to be converted to frames, return this frame number
//  Paramm	: Time to convert in LONGLONG format
//	Returns : Frame number converted from time in LONGLONG format
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

//LONGLONG Video::ConvertTimeToFrames(LONGLONG time)
//{
//
//	LONGLONG frames=0;
//	if(framesPerTime!=0) frames = time / framesPerTime;
//	Msg("time %u\n framesPerTime %u\nvideoLengthTime %u\nvideoLengthFrames %uframes %u",
//		time,framesPerTime,videoLengthTime,videoLengthFrames,frames);
//	return frames;
//}

bool Video::IsVideoRunning(void)
{
	return (videoRun==true)? true:false;
}

void Video::VideoRun(bool state)
{
	videoRun=state;
}

//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
//  PCN2561
//	Name	: Video::Video
//	Created	: 23 March 2004
//	By		: Antony van Iersel
//	Desc    : Original Initilisation of Vidoe Class was in Header, Moved here to Tidy 
//			: up the code. Mirrors the Destroy Video Class which was created at the same time
//	Usage   : Initialised the Video Class (Movie)
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Video::Video(void)
{
	lockSB=create(1);
	lastSeek=0;
	lastRecordedTime=-1; //Initialise, this is to make sure the a pause or a step back is not recorded
	seekByTime = true; //PCN3251 default initialised to seek by time.
					   // When asked to seek to a certain frame in the video
	// it will be manually converted to a time seek. Also when asked for a frame
	// It will manually convert from time to frame.
				
	vbfname = 0;
	width = 768;
	height=576;
	recordprofileinfo=false;
	videoLengthTime=0;
}

void Video::ResetLastRecordedTime(void)
{
	lastRecordedTime = -1;
}


//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
//  PCN2561
//	Name	: Video::~Video
//	Created	: 23 March 2004
//	By		: Antony van Iersel
//	Desc    : Was no place to properly destroy the video class, there are memory leaks all
//			  over, this is to Tidy up the distruction of the Video Class and hopefully
//			  catch the severe memory leaks
//	Usage   : Destroys the Video Class, remove any object from memory
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Video::~Video(void)
{
// Stop the graph.

	if(vbfname!=0) delete[] vbfname;
}

HRESULT Video::ConnectPins(IBaseFilter *pOne, int pinOneCount, IBaseFilter *pTwo, int pinTwoCount)
{
	HRESULT hr;
	int i;
	IEnumPins	*pEnumPins = 0;
	IPin		*pPinOne = 0;
	IPin		*pPinTwo = 0;
//	PIN_INFO	pinfo;

	pOne->EnumPins(&pEnumPins); //Enumerate IBaseFilterOne pins
	pEnumPins->Reset();	// Reset the Pins
	for(i=0;i<=pinOneCount;i++)
	{
		hr = pEnumPins->Next(1, &pPinOne, NULL);
		if(hr!=S_OK)
		{
			if(pPinOne!=0) pPinOne->Release();
			if(pEnumPins!=0) pEnumPins->Release();
			return hr==S_FALSE;
		}
	}
	pEnumPins->Release();

	pTwo->EnumPins(&pEnumPins);
	pEnumPins->Reset();
	for(i=0;i<=pinTwoCount;i++)
	{
		hr = pEnumPins->Next(1, &pPinTwo, NULL);
		if(hr!=S_OK)
		{
			if(pPinTwo!=0) pPinTwo->Release();
			if(pPinOne!=0) pPinOne->Release();
			if(pEnumPins!=0) pEnumPins->Release();
			return hr==S_FALSE;
		}
	}
	pEnumPins->Release();

	hr=g_pGraph->Connect(pPinOne, pPinTwo);
	if(pPinOne!=0) pPinOne->Release();
	if(pPinTwo!=0) pPinTwo->Release();
	return hr;
}



void Video::FindFilter(void)
{
/*
// Create the System Device Enumerator.
HRESULT hr;
ICreateDevEnum *pSysDevEnum = NULL;
hr = CoCreateInstance(CLSID_SystemDeviceEnum, NULL, CLSCTX_INPROC_SERVER,
    IID_ICreateDevEnum, (void **)&pSysDevEnum);

// Obtain a class enumerator for the video compressor category.
IEnumMoniker *pEnumCat = NULL;
hr = pSysDevEnum->CreateClassEnumerator(CLSID_EMPGDMX, &pEnumCat, 0);
//CLSID_VideoCompressor
//CLSID_MPEG2Demultiplexer
if (hr == S_OK) 
{
    // Enumerate the monikers.
    IMoniker *pMoniker;
    ULONG cFetched;
    while(pEnumCat->Next(1, &pMoniker, &cFetched) == S_OK)
    {
        IPropertyBag *pPropBag;
        pMoniker->BindToStorage(0, 0, IID_IPropertyBag, (void **)&pPropBag);

        // To retrieve the friendly name of the filter, do the following:
        VARIANT varName;
        VariantInit(&varName);
        hr = pPropBag->Read(L"FriendlyName", &varName, 0);
        if (SUCCEEDED(hr))
        {
            // Display the name in your UI somehow.
        }
        VariantClear(&varName);

        // To create an instance of the filter, do the following:
        IBaseFilter *pFilter;
        pMoniker->BindToObject(NULL, NULL, IID_IBaseFilter, (void**)&pFilter);
        // Now add the filter to the graph. Remember to release pFilter later.
    
        // Clean up.
        pPropBag->Release();
        pMoniker->Release();
    }
    pEnumCat->Release();
}
pSysDevEnum->Release();

*/
}



