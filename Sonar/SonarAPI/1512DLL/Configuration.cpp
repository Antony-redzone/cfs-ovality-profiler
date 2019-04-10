#include "StdAfx.h"
#include ".\configuration.h"

CConfiguration::CConfiguration(void)
: m_ArcSize(eum360Degrees)
, m_CentreAngle(eumCentre180)
, m_SampleRate(2.5)
, m_Oversamples(2)
, m_Samples(250)
, m_StepSize(eum09Degree)
, m_TxPulse(2)
{
}

CConfiguration::~CConfiguration(void)
{
}
