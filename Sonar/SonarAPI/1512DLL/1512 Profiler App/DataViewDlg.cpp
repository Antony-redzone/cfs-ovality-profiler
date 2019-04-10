// DataViewDlg.cpp : implementation file
//

#include "stdafx.h"
#include "1512 Profiler App.h"
#include "1512 Profiler AppDoc.h"
#include "1512 Profiler AppView.h"
#include "DataViewDlg.h"
#include ".\dataviewdlg.h"


// CDataViewDlg dialog

IMPLEMENT_DYNAMIC(CDataViewDlg, CDialog)
CDataViewDlg::CDataViewDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CDataViewDlg::IDD, pParent)
{
	m_Palette.Load("default.pal");
}

CDataViewDlg::~CDataViewDlg()
{
}

void CDataViewDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
}


BEGIN_MESSAGE_MAP(CDataViewDlg, CDialog)
END_MESSAGE_MAP()


// CDataViewDlg message handlers

LRESULT CDataViewDlg::WindowProc(UINT message, WPARAM wParam, LPARAM lParam)
{
	PAINTSTRUCT ps;
	CDC *pDC;
	CProfilerAppDoc *pDoc;
	char str[80];

	CBitmap Offscreen;
	CBitmap *pOld;
	CDC memDC;
	CRect rect;

	switch(message)
	{
	case WM_ERASEBKGND:
		return 0;
	case WM_PAINT:
		pDoc = ((CProfilerAppView *)GetParent())->GetDocument();
		pDC = BeginPaint(&ps);

			pDC->SetBkColor(RGB(255, 255, 255));

			// Display Sensor Data
			sprintf(str, "%f", pDoc->m_SensorData.m_Pitch);
			pDC->TextOut(120, 27, str);
			sprintf(str, "%f", pDoc->m_SensorData.m_Roll);
			pDC->TextOut(120, 46, str);
			sprintf(str, "%d    ", pDoc->m_SensorData.m_MotorPosition);
			pDC->TextOut(120, 76, str);
			sprintf(str, "%f", pDoc->m_SensorData.m_SupplyVoltage);
			pDC->TextOut(120, 96, str);
			sprintf(str, "%d    ", pDoc->m_SensorData.m_CablePayout);
			pDC->TextOut(120, 116, str);

			// For Double Buffering
			pDC->GetClipBox(&rect);
			memDC.CreateCompatibleDC(pDC);
			Offscreen.CreateCompatibleBitmap(pDC, rect.Width(), rect.Height());
			pOld = memDC.SelectObject(&Offscreen);
			
			// Offset needed to image for different step sizes
			int Offset = 1;
			switch(pDoc->m_Config.m_StepSize)
			{
			case eum18Degree:
				Offset = 2;
				break;
			case eum27Degree:
				Offset = 3;
				break;
			case eum36Degree:
				Offset = 4;
				break;
			}

			// Display sonar Data
			int nSamples = pDoc->m_Config.m_Samples;
			float sampleheight = 250.0f / (float)nSamples;
			for(int i = 0; i <  400; i++)
			{
				for (int j = 0; j < nSamples; j++)
				{
					// Give each pixel a colour from the palette based on its intensity.
					memDC.SetPixel(10 + (i), 395 - (j*sampleheight), *m_Palette.GetColor(pDoc->Data[(i/Offset)*(nSamples)+j]));
				}
			}
			// Blit Back buffer.
			
			pDC->TransparentBlt(10, 130, rect.Width()-15, rect.Height()-130, &memDC, 10, 130, rect.Width()- 10, rect.Height()-130, RGB(0,0,0));
			memDC.SelectObject(pOld);
			Offscreen.DeleteObject();
			memDC.DeleteDC();

		EndPaint(&ps);
		return 0;
	}

	return CDialog::WindowProc(message, wParam, lParam);
}
