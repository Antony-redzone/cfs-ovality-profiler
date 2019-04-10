#pragma once

class CSensorData
{
public:

	short m_MotorPosition; // position in Gradians 000 to 399
	float m_Pitch; // pitch in degrees (+/-180deg, 999 = not fitted)
	float m_Roll; // roll in degrees (+/-180deg, 999 = not fitted)
	WORD m_CablePayout; // cable payout in metres
	float m_SupplyVoltage;

public:
	CSensorData(void);
	virtual ~CSensorData(void);

	
};
