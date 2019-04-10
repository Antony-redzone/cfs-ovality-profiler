#pragma once
#include "Palette.h"

// CDataViewDlg dialog

class CDataViewDlg : public CDialog
{
	DECLARE_DYNAMIC(CDataViewDlg)

public:
	CDataViewDlg(CWnd* pParent = NULL);   // standard constructor
	virtual ~CDataViewDlg();
	Sonar::CPalette m_Palette;

// Dialog Data
	enum { IDD = IDD_DATAVIEWDLG };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()
	virtual LRESULT WindowProc(UINT message, WPARAM wParam, LPARAM lParam);
};
