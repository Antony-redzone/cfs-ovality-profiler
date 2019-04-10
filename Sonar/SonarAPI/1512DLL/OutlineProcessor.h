#pragma once
#include "Configuration.h"

struct Outline
{
	float range;
	int intensity;
};

class COutlineProcessor
{
public:
	__declspec(dllimport) COutlineProcessor(void);
	__declspec(dllimport) virtual ~COutlineProcessor(void);

	__declspec(dllimport) virtual void ChooseAlgorithm(int algorithm, char *name);
	__declspec(dllimport) virtual int GetNumberOfAlgorithms();
	__declspec(dllimport) virtual Outline *GetOutline(void);
	__declspec(dllimport) virtual void ProcessOutline(BYTE *data);
	__declspec(dllimport) virtual void SetConfiguration(const CConfiguration &configuration);
	__declspec(dllimport) virtual void SetThreshold(int threshold);
	__declspec(dllimport) virtual void SetVelocityOfSound(int vos);
};

__declspec(dllimport) COutlineProcessor* CreateOutlineProcessorObject();