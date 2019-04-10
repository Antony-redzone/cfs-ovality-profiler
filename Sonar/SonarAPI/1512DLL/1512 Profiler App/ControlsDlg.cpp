// ControlsDlg.cpp : implementation file
//

#include "stdafx.h"
#include "1512 Profiler App.h"
#include "1512 Profiler AppDoc.h"
#include "1512 Profiler AppView.h"
#include "ControlsDlg.h"
#include ".\controlsdlg.h"


// CControlsDlg dialog

IMPLEMENT_DYNAMIC(CControlsDlg, CDialog)
CControlsDlg::CControlsDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CControlsDlg::IDD, pParent)
	, m_Position(0)
{
}

CControlsDlg::~CControlsDlg()
{
}

void CControlsDlg::DoDataExchange(CDataExchange* pDX)
{
	CProfilerAppDoc *pDoc = ((CProfilerAppView *)GetParent())->GetDocument();

	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_STEPSIZE, m_StepSizeSlider);
	DDX_Control(pDX, IDC_ARCSIZE, m_ArcSizeSlider);
	DDX_Control(pDX, IDC_CENTREANGLE, m_CentreAngleSlider);
	DDX_Control(pDX, IDC_OUTTHRESHOLD, m_ThresholdSlider);
	DDX_Control(pDX, IDC_OUT1, m_Outline1Slider);
	DDX_Control(pDX, IDC_OUT2, m_Outline2Slider);
	DDX_Control(pDX, IDC_APPLYCONFIG, m_ApplyButton);
	DDX_Control(pDX, IDC_STARTSCAN, m_StartButton);
	DDX_Control(pDX, IDC_STOPSCAN, m_StopButton);
	DDX_Text(pDX, IDC_OVERSAMPLES, pDoc->m_Config.m_Oversamples);
	DDX_Text(pDX, IDC_PULSEWIDTH, pDoc->m_Config.m_TxPulse);
	DDX_Text(pDX, IDC_SAMPLES, pDoc->m_Config.m_Samples);
	DDX_Text(pDX, IDC_SAMPLERATE, pDoc->m_Config.m_SampleRate);
	DDX_Text(pDX, IDC_BLANKING, pDoc->m_Blanking);
}


BEGIN_MESSAGE_MAP(CControlsDlg, CDialog)
	ON_EN_CHANGE(IDC_OVERSAMPLES, OnEnChangeTextBox)
	ON_EN_CHANGE(IDC_PULSEWIDTH, OnEnChangeTextBox)
	ON_EN_CHANGE(IDC_SAMPLES, OnEnChangeTextBox)
	ON_EN_CHANGE(IDC_SAMPLERATE, OnEnChangeTextBox)
	ON_EN_CHANGE(IDC_BLANKING, OnEnChangeTextBox)
	ON_NOTIFY(NM_RELEASEDCAPTURE, IDC_STEPSIZE, OnNMReleasedcaptureSlider)
	ON_NOTIFY(NM_RELEASEDCAPTURE, IDC_ARCSIZE, OnNMReleasedcaptureSlider)
	ON_NOTIFY(NM_RELEASEDCAPTURE, IDC_CENTREANGLE, OnNMReleasedcaptureSlider)
	ON_NOTIFY(NM_RELEASEDCAPTURE, IDC_OUTTHRESHOLD, OnNMReleasedcaptureSlider)
	ON_NOTIFY(NM_RELEASEDCAPTURE, IDC_OUT1, OnNMReleasedcaptureSlider)
	ON_NOTIFY(NM_RELEASEDCAPTURE, IDC_OUT2, OnNMReleasedcaptureSlider)
END_MESSAGE_MAP()


// CControlsDlg message handlers

BOOL CControlsDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	// Set Up Controls
	int nOutlines = ((CProfilerAppView *)GetParent())->m_pOutlineProcessor1->GetNumberOfAlgorithms();
	m_Outline1Slider.SetRange(0, nOutlines-1, true);
	m_Outline2Slider.SetRange(0, nOutlines-1, true);
	m_ThresholdSlider.SetRange(0, 255, true);
	m_Outline2Slider.SetPos(1);
	m_ThresholdSlider.SetPos(50);

	m_StepSizeSlider.SetRange(0, 3, true);
	m_ArcSizeSlider.SetRange(0, 9, true);
	m_CentreAngleSlider.SetRange(0, 10, true);

	m_ArcSizeSlider.SetPos(9);
	m_CentreAngleSlider.SetPos(5);

	m_StopButton.EnableWindow(false);

	return TRUE;  // return TRUE unless you set the focus to a control
	// EXCEPTION: OCX Property Pages should return FALSE
}

LRESULT CControlsDlg::WindowProc(UINT message, WPARAM wParam, LPARAM lParam)
{
	PAINTSTRUCT ps;
	CDC *pDC;
	CProfilerAppView *pView;

	switch(message)
	{
	case WM_COMMAND:
		pView = (CProfilerAppView *)GetParent();

		switch(LOWORD(wParam))
		{
		case IDC_STARTSCAN:
			if (pView->m_UWUConnected)
			{
				m_ApplyButton.EnableWindow(false);
				m_StartButton.EnableWindow(false);
				m_StopButton.EnableWindow();
				pView->m_pProfiler->StartScan();
				pView->m_Scanning = true;
				InvalidateRect(CRect(800, 25, 900, 50), true);
			}
			return 0;
		case IDC_STOPSCAN:
			m_ApplyButton.EnableWindow();
			m_StartButton.EnableWindow();
			m_StopButton.EnableWindow(false);
			pView->m_pProfiler->StopScan();
			pView->m_Scanning = false;
			InvalidateRect(CRect(800, 25, 900, 50), true);
			return 0;
		case IDC_APPLYCONFIG:
			m_ApplyButton.EnableWindow(false);
			m_StartButton.EnableWindow(false);
			pView->m_pProfiler->SetConfiguration((pView->GetDocument()->m_Config));
			m_StartButton.EnableWindow();
			m_ApplyButton.EnableWindow();
			pView->m_pOutlineProcessor1->SetConfiguration(pView->GetDocument()->m_Config);
			pView->m_pOutlineProcessor2->SetConfiguration(pView->GetDocument()->m_Config);
			return 0;
		case IDC_SETSAMPLES:
			pView->m_pProfiler->SetSamples(pView->GetDocument()->m_Config.m_Samples);
			pView->m_pOutlineProcessor1->SetConfiguration(pView->GetDocument()->m_Config);
			pView->m_pOutlineProcessor2->SetConfiguration(pView->GetDocument()->m_Config);
			return 0;
		case IDC_SETSAMPLERATE:
			pView->m_pProfiler->SetSampleRate(pView->GetDocument()->m_Config.m_SampleRate);
			pView->m_pOutlineProcessor1->SetConfiguration(pView->GetDocument()->m_Config);
			pView->m_pOutlineProcessor2->SetConfiguration(pView->GetDocument()->m_Config);
			return 0;
		case IDC_SETBLANKING:
			pView->m_pProfiler->SetBlanking(pView->GetDocument()->m_Blanking);
			pView->m_pOutlineProcessor1->SetConfiguration(pView->GetDocument()->m_Config);
			pView->m_pOutlineProcessor2->SetConfiguration(pView->GetDocument()->m_Config);
			return 0;
		case IDC_STEPCLOCKWISE:
			if (pView->m_pProfiler->TransmitStepClockwise(&pView->GetDocument()->Data[pView->GetDocument()->m_Config.m_Samples*m_Position]))
			{
				// Fetch Sensor Data
				pView->m_pProfiler->CollectSensorDataFromSonar();
				memcpy(&pView->GetDocument()->m_SensorData, pView->m_pProfiler->GetSensorData(), sizeof(CSensorData));

				m_Position = pView->GetDocument()->m_SensorData.m_MotorPosition;
				pView->m_DataView.InvalidateRect(0, true);
			}
			return 0;
		case IDC_STEPANTICLOCKWISE:
			if (pView->m_pProfiler->TransmitStepAntiClockwise(&pView->GetDocument()->Data[pView->GetDocument()->m_Config.m_Samples*m_Position]))
			{
				// Fetch Sensor Data
				pView->m_pProfiler->CollectSensorDataFromSonar();
				memcpy(&pView->GetDocument()->m_SensorData, pView->m_pProfiler->GetSensorData(), sizeof(CSensorData));

				m_Position = pView->GetDocument()->m_SensorData.m_MotorPosition;
				pView->m_DataView.InvalidateRect(0, true);
			}
			return 0;
		case IDC_TRANSMIT:
			if (pView->m_pProfiler->Transmit(&pView->GetDocument()->Data[pView->GetDocument()->m_Config.m_Samples*m_Position]))
			{
				// Fetch Sensor Data
				pView->m_pProfiler->CollectSensorDataFromSonar();
				memcpy(&pView->GetDocument()->m_SensorData, pView->m_pProfiler->GetSensorData(), sizeof(CSensorData));

				m_Position = pView->GetDocument()->m_SensorData.m_MotorPosition;
				pView->m_DataView.InvalidateRect(0, true);
			}
			return 0;
		}
		break;
	case WM_PAINT:
		pView = (CProfilerAppView *)GetParent();
		pDC = BeginPaint(&ps);
			if(pView->m_USBConnected)
			{
				if (pView->m_UWUConnected)
					pDC->TextOut(650, 25, "UWU Connected");
				else pDC->TextOut(650, 25, "UWU Not Connected");
			}
			else pDC->TextOut(650, 25, "USB Not Connected");
			if (pView->m_Scanning) pDC->TextOut(800, 25, "Scanning");
		EndPaint(&ps);
	}

	return CDialog::WindowProc(message, wParam, lParam);
}

void CControlsDlg::OnEnChangeTextBox()
{
	// TODO:  If this is a RICHEDIT control, the control will not
	// send this notification unless you override the CDialog::OnInitDialog()
	// function and call CRichEditCtrl().SetEventMask()
	// with the ENM_CHANGE flag ORed into the mask.

	UpdateData();
}

void CControlsDlg::OnNMReleasedcaptureSlider(NMHDR *pNMHDR, LRESULT *pResult)
{
	int value;
	CProfilerAppDoc *pDoc = ((CProfilerAppView *)GetParent())->GetDocument();

	switch(pNMHDR->idFrom)
	{
	case IDC_STEPSIZE:
		value = m_StepSizeSlider.GetPos();
		switch(value)
		{
		case 0:
			pDoc->m_Config.m_StepSize = eum09Degree;
			break;
		case 1:
			pDoc->m_Config.m_StepSize = eum18Degree;
			break;
		case 2:
			pDoc->m_Config.m_StepSize = eum27Degree;
			break;
		case 3:
			pDoc->m_Config.m_StepSize = eum36Degree;
		}
		break;
	case IDC_ARCSIZE:
		value = m_ArcSizeSlider.GetPos();
		switch(value)
		{
		case 0:
			pDoc->m_Config.m_ArcSize = eum30Degrees;
			break;
		case 1:
			pDoc->m_Config.m_ArcSize = eum60Degrees;
			break;
		case 2:
			pDoc->m_Config.m_ArcSize = eum90Degrees;
			break;
		case 3:
			pDoc->m_Config.m_ArcSize = eum120Degrees;
			break;
		case 4:
			pDoc->m_Config.m_ArcSize = eum150Degrees;
			break;
		case 5:
			pDoc->m_Config.m_ArcSize = eum180Degrees;
			break;
		case 6:
			pDoc->m_Config.m_ArcSize = eum210Degrees;
			break;
		case 7:
			pDoc->m_Config.m_ArcSize = eum240Degrees;
			break;
		case 8:
			pDoc->m_Config.m_ArcSize = eum270Degrees;
			break;
		case 9:
			pDoc->m_Config.m_ArcSize = eum360Degrees;
			break;
		}
		break;
	case IDC_CENTREANGLE:
			value = m_CentreAngleSlider.GetPos();
		switch(value)
		{
		case 0:
			pDoc->m_Config.m_CentreAngle = eumCentre30;
			break;
		case 1:
			pDoc->m_Config.m_CentreAngle = eumCentre60;
			break;
		case 2:
			pDoc->m_Config.m_CentreAngle = eumCentre90;
			break;
		case 3:
			pDoc->m_Config.m_CentreAngle = eumCentre120;
			break;
		case 4:
			pDoc->m_Config.m_CentreAngle = eumCentre150;
			break;
		case 5:
			pDoc->m_Config.m_CentreAngle = eumCentre180;
			break;
		case 6:
			pDoc->m_Config.m_CentreAngle = eumCentre210;
			break;
		case 7:
			pDoc->m_Config.m_CentreAngle = eumCentre240;
			break;
		case 8:
			pDoc->m_Config.m_CentreAngle = eumCentre270;
			break;
		case 9:
			pDoc->m_Config.m_CentreAngle = eumCentre300;
			break;
		case 10:
			pDoc->m_Config.m_CentreAngle = eumCentre330;
			break;
		}
		break;
	case IDC_OUTTHRESHOLD:
		value = m_ThresholdSlider.GetPos();
		((CProfilerAppView *)GetParent())->m_pOutlineProcessor1->SetThreshold(value);
		((CProfilerAppView *)GetParent())->m_pOutlineProcessor2->SetThreshold(value);
		break;
	case IDC_OUT1:
		value = m_Outline1Slider.GetPos();
		((CProfilerAppView *)GetParent())->m_pOutlineProcessor1->ChooseAlgorithm(value, ((CProfilerAppView *)GetParent())->m_OutlineDlg1.m_AlgorithmName);
		break;
	case IDC_OUT2:
		value = m_Outline2Slider.GetPos();
		((CProfilerAppView *)GetParent())->m_pOutlineProcessor2->ChooseAlgorithm(value, ((CProfilerAppView *)GetParent())->m_OutlineDlg2.m_AlgorithmName);
		break;
	}

	*pResult = 0;
}
