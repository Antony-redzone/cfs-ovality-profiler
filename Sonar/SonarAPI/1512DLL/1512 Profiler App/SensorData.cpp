#include "StdAfx.h"
#include ".\sensordata.h"

CSensorData::CSensorData(void)
: m_Pitch(0), m_Roll(0), m_CablePayout(0)
, m_SupplyVoltage(0), m_MotorPosition(0)
{
}

CSensorData::~CSensorData(void)
{
}
