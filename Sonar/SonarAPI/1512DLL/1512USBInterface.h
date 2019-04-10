#pragma once
#include "Definitions.h"
#include "SensorData.h"
#include "Configuration.h"

namespace Sonar
{

class C1512USBComm
	{
	public:
		__declspec(dllimport) C1512USBComm(void);
		__declspec(dllimport) virtual ~C1512USBComm(void);

		__declspec(dllimport) virtual bool InitialiseComms();
		__declspec(dllimport) virtual bool LocateToZeroPosition();
		__declspec(dllimport) virtual bool Transmit(BYTE *memptr);
		__declspec(dllimport) virtual bool TransmitStepClockwise(BYTE *memptr);
		__declspec(dllimport) virtual bool TransmitStepAntiClockwise(BYTE *memptr);
		__declspec(dllimport) virtual bool CollectSensorDataFromSonar();
		__declspec(dllimport) virtual void StartScan();
		__declspec(dllimport) virtual void StopScan();
		__declspec(dllimport) virtual bool IsScanning();

		__declspec(dllimport) virtual void GetScanData(BYTE *dataBuffer);
		__declspec(dllimport) virtual CConfiguration *GetConfiguration();	
		__declspec(dllimport) virtual CSensorData *GetSensorData();
		__declspec(dllimport) virtual bool GetVersion(unsigned char *version);
		__declspec(dllimport) virtual double GetStartTime();
		__declspec(dllimport) virtual double GetEndTime();
	
		__declspec(dllimport) virtual bool SetArcSize(EnumArcSize arcSize);
		__declspec(dllimport) virtual void SetBlanking(int blanking);
		__declspec(dllimport) virtual bool SetCentreAngle(EnumCentreAngle centreAngle);
		__declspec(dllimport) virtual bool SetSampleRate(float sampleRate);
		__declspec(dllimport) virtual bool SetOversamples(int oversamples);
		__declspec(dllimport) virtual bool SetPulseWidth(int txPulse);
		__declspec(dllimport) virtual bool SetSamples(int samples);
		__declspec(dllimport) virtual bool SetShaftEncoder(eEncoder encoder);
		__declspec(dllimport) virtual bool SetStepSize(EnumStepSize stepSize);
		__declspec(dllimport) virtual bool SetConfiguration(const CConfiguration &configuration);

		__declspec(dllimport) virtual void RegisterCallback( void (*Callback) (void));
		__declspec(dllimport) virtual void RegisterTimeCallback( void (*TimeCallback)(double *time));
	};

__declspec(dllimport) C1512USBComm* Create1512USBCommObject();

};