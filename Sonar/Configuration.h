#pragma once

#include "Definitions.h"

class CConfiguration
{
public:
	EnumArcSize m_ArcSize; 
	EnumCentreAngle m_CentreAngle;
	float m_SampleRate; // sample rate 2.5 - 5MHz
	int m_Oversamples; // samples per cell
	int m_Samples; // samples per scan line
	EnumStepSize m_StepSize;
	int m_TxPulse; // Tx Pulse width (in us)

public:
	CConfiguration(void);
	//CConfiguration(EnumArcSize as, EnumCentreAngle ca, float sr, int oversamples, int samples, EnumStepSize sSize, int pSize);
	virtual ~CConfiguration(void);
};
