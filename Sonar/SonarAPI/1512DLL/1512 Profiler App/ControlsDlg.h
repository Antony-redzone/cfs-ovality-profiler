#pragma once
#include "afxcmn.h"
#include "afxwin.h"


// CControlsDlg dialog

class CControlsDlg : public CDialog
{
	DECLARE_DYNAMIC(CControlsDlg)

public:
	CControlsDlg(CWnd* pParent = NULL);   // standard constructor
	virtual ~CControlsDlg();
	virtual BOOL OnInitDialog();

// Dialog Data
	enum { IDD = IDD_CONTROLSDLG };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	virtual LRESULT WindowProc(UINT message, WPARAM wParam, LPARAM lParam);

	DECLARE_MESSAGE_MAP()
public:
	CSliderCtrl m_StepSizeSlider;
	CSliderCtrl m_ArcSizeSlider;
	CSliderCtrl m_CentreAngleSlider;
	CSliderCtrl m_ThresholdSlider;
	CSliderCtrl m_Outline1Slider;
	CSliderCtrl m_Outline2Slider;
	CButton m_ApplyButton;
	CButton m_StartButton;
	CButton m_StopButton;

	unsigned int m_Position;

	afx_msg void OnEnChangeTextBox();
	afx_msg void OnNMReleasedcaptureSlider(NMHDR *pNMHDR, LRESULT *pResult);
};
