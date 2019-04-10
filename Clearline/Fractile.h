#include "..\houghlibv2.0\CBSAlgebra.h"

class Fracts
{
public:
	Fracts();
	Fracts(float *_pvData,
		   int _ArraySize,
		   float _Fractile);

	~Fracts(void);
	float CalculateFractile(void);
	void CalcIt(void);

private:
	float *pvData;
	float *SortedData;
	int	arraySize;
	float fractilePer;
};