// OutlineDlg.cpp : implementation file
//

#include "stdafx.h"
#include "1512 Profiler App.h"
#include "1512 Profiler AppDoc.h"
#include "1512 Profiler AppView.h"
#include "OutlineDlg.h"
#include ".\outlinedlg.h"


// COutlineDlg dialog

IMPLEMENT_DYNAMIC(COutlineDlg, CDialog)
COutlineDlg::COutlineDlg(CWnd* pParent /*=NULL*/)
	: CDialog(COutlineDlg::IDD, pParent)
{
	m_AlgorithmName[0] = '\0';
}

COutlineDlg::~COutlineDlg()
{
}

void COutlineDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
}


BEGIN_MESSAGE_MAP(COutlineDlg, CDialog)
END_MESSAGE_MAP()


// COutlineDlg message handlers

LRESULT COutlineDlg::WindowProc(UINT message, WPARAM wParam, LPARAM lParam)
{
	PAINTSTRUCT ps;
	CDC *pDC;

	switch(message)
	{
	//case WM_ERASEBKGND:
	//	return 0;
	case WM_PAINT:
		pDC = BeginPaint(&ps);
		pDC->TextOut(0, 0, m_AlgorithmName);

		// Offset needed to image for different step sizes
			int Offset = 1;
			switch(((CProfilerAppView *)GetParent())->GetDocument()->m_Config.m_StepSize)
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

		for (int i = 0; i < 400/Offset; i++)
		{
			if(m_OutlineData[i].intensity > 0)pDC->SetPixel(i*Offset, 180 - int(m_OutlineData[i].range*500), RGB(0, 0, 0));
		}


		EndPaint(&ps);
		return 0;
	}

	return CDialog::WindowProc(message, wParam, lParam);
}
