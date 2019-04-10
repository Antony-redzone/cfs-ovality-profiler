#include "StdAfx.h"
#include ".\configuration.h"
#include "definitions.h"

CConfiguration::CConfiguration(void)
: m_ArcSize(eum360Degrees)
, m_CentreAngle(eumCentre180)
, m_SampleRate(2.5)
, m_Oversamples(4)
, m_Samples(301)
, m_StepSize(eum09Degree)
, m_TxPulse(2)
{
}


//CConfiguration::CConfiguration(EnumArcSize as, EnumCentreAngle ca, float sr, int oversamples, int samples, EnumStepSize sSize, int pSize)
//: m_ArcSize(as)
//, m_CentreAngle(ca)
//, m_SampleRate(sr)
//, m_Oversamples(oversamples)
//, m_Samples(samples)
//, m_StepSize(sSize)
//, m_TxPulse(pSize)
//{
//}

CConfiguration::~CConfiguration(void)
{
}
