#include "Fractile.h"
#include <math.h>


//#include "Common.h"

Fracts::Fracts()
{
}

Fracts::Fracts(float *_pvData,
				   int _ArraySize,
				   float _FractilePer)

{
	pvData = _pvData;
	arraySize = _ArraySize;
	fractilePer = _FractilePer;
}


//Fracts::~Fracts(void)
//{
//}

//float Fracts::CalculateFractile(void)
//{
//	float FractileIndex;
//	float Ans;

//	SortedData = new float[arraySize];
//	memcpy(pvData,SortedData,sizeof(float)*arraySize);
//	QuickSort(SortedData, 0, arraySize);
//	FractileIndex = (float) arraySize * fractilePer / 100;
//	Ans = SortedData[(int) FractileIndex];
//	delete[] SortedData;
//	return Ans;
//}

//void Fracts::CalcIt(void)
//{
//}





