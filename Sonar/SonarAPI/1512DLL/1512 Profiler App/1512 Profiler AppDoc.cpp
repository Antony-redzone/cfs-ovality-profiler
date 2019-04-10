// 1512 Profiler AppDoc.cpp : implementation of the CProfilerAppDoc class
//

#include "stdafx.h"
#include "1512 Profiler App.h"

#include "1512 Profiler AppDoc.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CProfilerAppDoc

IMPLEMENT_DYNCREATE(CProfilerAppDoc, CDocument)

BEGIN_MESSAGE_MAP(CProfilerAppDoc, CDocument)
END_MESSAGE_MAP()


// CProfilerAppDoc construction/destruction

CProfilerAppDoc::CProfilerAppDoc()
: m_Blanking(0)
{
	memset(Data, 0, USB_MAXDATABLOCKSIZE*6);

}

CProfilerAppDoc::~CProfilerAppDoc()
{
}

BOOL CProfilerAppDoc::OnNewDocument()
{
	if (!CDocument::OnNewDocument())
		return FALSE;

	// TODO: add reinitialization code here
	// (SDI documents will reuse this document)

	return TRUE;
}




// CProfilerAppDoc serialization

void CProfilerAppDoc::Serialize(CArchive& ar)
{
	if (ar.IsStoring())
	{
		// TODO: add storing code here
	}
	else
	{
		// TODO: add loading code here
	}
}


// CProfilerAppDoc diagnostics

#ifdef _DEBUG
void CProfilerAppDoc::AssertValid() const
{
	CDocument::AssertValid();
}

void CProfilerAppDoc::Dump(CDumpContext& dc) const
{
	CDocument::Dump(dc);
}
#endif //_DEBUG


// CProfilerAppDoc commands
