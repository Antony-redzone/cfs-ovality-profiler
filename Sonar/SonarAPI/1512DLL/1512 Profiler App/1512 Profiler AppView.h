// 1512 Profiler AppView.h : interface of the CProfilerAppView class
//


#pragma once
#include "ControlsDlg.h"
#include "DataViewDlg.h"
#include "OutlineDlg.h"
#include "OutlineProcessor.h"
#include "1512USBInterface.h"

class CProfilerAppView : public CView
{
protected: // create from serialization only
	CProfilerAppView();
	DECLARE_DYNCREATE(CProfilerAppView)

// Attributes
public:
	CProfilerAppDoc* GetDocument() const;

	CControlsDlg m_Controls; // Controls dialog Window
	CDataViewDlg m_DataView; // Data Dialog Window
	COutlineDlg m_OutlineDlg1, m_OutlineDlg2; // 2 Outline Dialog Windows
	bool m_ControlsCreated; // Used to test if the control dialog has been created (for positioning)
	Sonar::C1512USBComm *m_pProfiler; // Profiler USB Interface
	bool m_USBConnected; // Is USB Connected?
	bool m_UWUConnected; // Is UWU Connected?
	bool m_Scanning; // Is Profiler Scanning?

	COutlineProcessor *m_pOutlineProcessor1; // Outline Processor Object 1
	COutlineProcessor *m_pOutlineProcessor2; // Outline Processor Object 2

// Operations
public:

// Overrides
	public:
	virtual void OnDraw(CDC* pDC);  // overridden to draw this view
virtual BOOL PreCreateWindow(CREATESTRUCT& cs);
protected:
	virtual BOOL OnPreparePrinting(CPrintInfo* pInfo);
	virtual void OnBeginPrinting(CDC* pDC, CPrintInfo* pInfo);
	virtual void OnEndPrinting(CDC* pDC, CPrintInfo* pInfo);

// Implementation
public:
	virtual ~CProfilerAppView();
#ifdef _DEBUG
	virtual void AssertValid() const;
	virtual void Dump(CDumpContext& dc) const;
#endif

protected:

// Generated message map functions
protected:
	DECLARE_MESSAGE_MAP()
public:
	virtual void OnInitialUpdate();
protected:
	virtual LRESULT WindowProc(UINT message, WPARAM wParam, LPARAM lParam);
};

#ifndef _DEBUG  // debug version in 1512 Profiler AppView.cpp
inline CProfilerAppDoc* CProfilerAppView::GetDocument() const
   { return reinterpret_cast<CProfilerAppDoc*>(m_pDocument); }
#endif

