// 1512 Profiler App.h : main header file for the 1512 Profiler App application
//
#pragma once

#ifndef __AFXWIN_H__
	#error include 'stdafx.h' before including this file for PCH
#endif

#include "resource.h"       // main symbols


// CProfilerAppApp:
// See 1512 Profiler App.cpp for the implementation of this class
//

class CProfilerAppApp : public CWinApp
{
public:
	CProfilerAppApp();


// Overrides
public:
	virtual BOOL InitInstance();

// Implementation
	afx_msg void OnAppAbout();
	DECLARE_MESSAGE_MAP()
};

extern CProfilerAppApp theApp;