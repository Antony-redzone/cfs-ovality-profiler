// 1512 Profiler AppDoc.h : interface of the CProfilerAppDoc class
//


#pragma once
#include "Configuration.h"
#include "SensorData.h"

#define USB_MAXDATABLOCKSIZE 50176

class CProfilerAppDoc : public CDocument
{
protected: // create from serialization only
	CProfilerAppDoc();
	DECLARE_DYNCREATE(CProfilerAppDoc)

// Attributes
public:
	CConfiguration m_Config; // Sonar Configuration Data
	CSensorData m_SensorData; // Sonar Sensor Data
	int m_Blanking; // Blanking
	BYTE Data[USB_MAXDATABLOCKSIZE*1200]; // Raw Sonar Data

// Operations
public:

// Overrides
	public:
	virtual BOOL OnNewDocument();
	virtual void Serialize(CArchive& ar);

// Implementation
public:
	virtual ~CProfilerAppDoc();
#ifdef _DEBUG
	virtual void AssertValid() const;
	virtual void Dump(CDumpContext& dc) const;
#endif

protected:

// Generated message map functions
protected:
	DECLARE_MESSAGE_MAP()
};


