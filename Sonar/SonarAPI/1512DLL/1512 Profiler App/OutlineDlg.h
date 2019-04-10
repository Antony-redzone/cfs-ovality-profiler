#pragma once
#include "OutlineProcessor.h"

// COutlineDlg dialog

class COutlineDlg : public CDialog
{
	DECLARE_DYNAMIC(COutlineDlg)

public:
	COutlineDlg(CWnd* pParent = NULL);   // standard constructor
	virtual ~COutlineDlg();

	Outline m_OutlineData[400];
	char m_AlgorithmName[80];

// Dialog Data
	enum { IDD = IDD_OUTLINEDLG };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()
	virtual LRESULT WindowProc(UINT message, WPARAM wParam, LPARAM lParam);
};
