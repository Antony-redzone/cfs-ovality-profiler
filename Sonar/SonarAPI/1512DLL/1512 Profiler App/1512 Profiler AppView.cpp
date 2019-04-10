// 1512 Profiler AppView.cpp : implementation of the CProfilerAppView class
//

#include "stdafx.h"
#include "1512 Profiler App.h"

#include "1512 Profiler AppDoc.h"
#include "1512 Profiler AppView.h"
#include ".\1512 profiler appview.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

// Callback Function definition
void OnSonarUpdate();
CProfilerAppView *pView; // to enable use of the view in the callback function

// CProfilerAppView

IMPLEMENT_DYNCREATE(CProfilerAppView, CView)

BEGIN_MESSAGE_MAP(CProfilerAppView, CView)
	// Standard printing commands
	ON_COMMAND(ID_FILE_PRINT, CView::OnFilePrint)
	ON_COMMAND(ID_FILE_PRINT_DIRECT, CView::OnFilePrint)
	ON_COMMAND(ID_FILE_PRINT_PREVIEW, CView::OnFilePrintPreview)
END_MESSAGE_MAP()

// CProfilerAppView construction/destruction

CProfilerAppView::CProfilerAppView()
: m_ControlsCreated(false)
, m_USBConnected(false), m_UWUConnected(false), m_Scanning(false)
{
	// create objects and set callback
	m_pProfiler = Sonar::Create1512USBCommObject();
	m_pProfiler->RegisterCallback(&OnSonarUpdate);

	m_pOutlineProcessor1 = CreateOutlineProcessorObject();
	m_pOutlineProcessor2 = CreateOutlineProcessorObject();
}

CProfilerAppView::~CProfilerAppView()
{
}

BOOL CProfilerAppView::PreCreateWindow(CREATESTRUCT& cs)
{
	// TODO: Modify the Window class or styles here by modifying
	//  the CREATESTRUCT cs

	return CView::PreCreateWindow(cs);
}

// CProfilerAppView drawing

void CProfilerAppView::OnDraw(CDC* /*pDC*/)
{
	CProfilerAppDoc* pDoc = GetDocument();
	ASSERT_VALID(pDoc);
	if (!pDoc)
		return;

	// TODO: add draw code for native data here
}


// CProfilerAppView printing

BOOL CProfilerAppView::OnPreparePrinting(CPrintInfo* pInfo)
{
	// default preparation
	return DoPreparePrinting(pInfo);
}

void CProfilerAppView::OnBeginPrinting(CDC* /*pDC*/, CPrintInfo* /*pInfo*/)
{
	// TODO: add extra initialization before printing
}

void CProfilerAppView::OnEndPrinting(CDC* /*pDC*/, CPrintInfo* /*pInfo*/)
{
	// TODO: add cleanup after printing
}


// CProfilerAppView diagnostics

#ifdef _DEBUG
void CProfilerAppView::AssertValid() const
{
	CView::AssertValid();
}

void CProfilerAppView::Dump(CDumpContext& dc) const
{
	CView::Dump(dc);
}

CProfilerAppDoc* CProfilerAppView::GetDocument() const // non-debug version is inline
{
	ASSERT(m_pDocument->IsKindOf(RUNTIME_CLASS(CProfilerAppDoc)));
	return (CProfilerAppDoc*)m_pDocument;
}
#endif //_DEBUG


// CProfilerAppView message handlers

void CProfilerAppView::OnInitialUpdate()
{
	CView::OnInitialUpdate();
	pView = this;

	// Create various dialog windows
	m_DataView.Create(IDD_DATAVIEWDLG, this);
	m_OutlineDlg1.Create(IDD_OUTLINEDLG, this);
	m_OutlineDlg1.SetWindowPos(&wndTop, 550, 0, 0, 0, SWP_NOSIZE);
	m_OutlineDlg2.Create(IDD_OUTLINEDLG, this);
	m_OutlineDlg2.SetWindowPos(&wndTop, 550, 200, 0, 0, SWP_NOSIZE);
	if (m_Controls.Create(IDD_CONTROLSDLG,this)) m_ControlsCreated = true;
	
	// Initialise Comms, if true is returned, the USB Device is connected
	m_USBConnected = m_pProfiler->InitialiseComms();
	
	// If GetVersion Returns true, the under water unit is connected
	unsigned char str[80];
	if(m_USBConnected) m_UWUConnected = m_pProfiler->GetVersion(str);

	// Setup Outline Processors
	m_pOutlineProcessor1->ChooseAlgorithm(0, m_OutlineDlg1.m_AlgorithmName);
	m_pOutlineProcessor2->ChooseAlgorithm(1, m_OutlineDlg2.m_AlgorithmName);
	m_pOutlineProcessor1->SetConfiguration(GetDocument()->m_Config);
	m_pOutlineProcessor2->SetConfiguration(GetDocument()->m_Config);

}

LRESULT CProfilerAppView::WindowProc(UINT message, WPARAM wParam, LPARAM lParam)
{
	CRect rect, childRect;

	switch(message)
	{
	case WM_SIZE:
		if (m_ControlsCreated)
		{
			GetWindowRect(&rect);
			m_Controls.GetWindowRect(&childRect);
			m_Controls.SetWindowPos(&wndTop, 0, rect.Height() - childRect.Height(), 0, 0, SWP_NOSIZE);
		}
		return 0;
	}

	return CView::WindowProc(message, wParam, lParam);
}


// Callback Function
void OnSonarUpdate()
{
	CProfilerAppDoc *pDoc = pView->GetDocument();

	if (pView->m_Scanning)
	{
		memset(pDoc->Data, 0, USB_MAXDATABLOCKSIZE*6); // Reset local Sonar Data
		pView->m_pProfiler->GetScanData(pDoc->Data); // Get Sonar Data
		memcpy(&pDoc->m_SensorData, pView->m_pProfiler->GetSensorData(), sizeof(CSensorData)); // Get Sensor Data
	
		// Process Outlines and copy the data into the the Outline Dialog windows.
		pView->m_pOutlineProcessor1->ProcessOutline(pDoc->Data);
		pView->m_pOutlineProcessor2->ProcessOutline(pDoc->Data);
		memcpy(pView->m_OutlineDlg1.m_OutlineData, pView->m_pOutlineProcessor1->GetOutline(), sizeof(Outline)*400);
		memcpy(pView->m_OutlineDlg2.m_OutlineData, pView->m_pOutlineProcessor2->GetOutline(), sizeof(Outline)*400);

		// Refresh data and outline dialogs
		pView->m_DataView.InvalidateRect(0, true);
		pView->m_OutlineDlg1.InvalidateRect(0, true);
		pView->m_OutlineDlg2.InvalidateRect(0, true);
	}
}
