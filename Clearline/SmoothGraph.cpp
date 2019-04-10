#include <math.h>
#include "FilterGraph.h"

FilterGraph::FilterGraph(float *_graphData, int _numberFrames)
{
	graphData = _graphData;
	numberFrames = _numberFrames;
}

void FilterGraph::Smooth(void)
{
	int i;
	float a,b,c,d,e;
	float answer;
	tempData = new float[numberFrames+1];

	for(i=2;i<numberFrames-2;i++) 
	{
		a = (float) fabs(graphData[i-2]); if(a>=100000) a= 0;
		b = (float) fabs(graphData[i-1]); if(b>=100000) b= 0;
		c = (float) fabs(graphData[i]);   if(c>=100000) c= 0;
		d = (float) fabs(graphData[i+1]); if(d>=100000) d= 0;
		e = (float) fabs(graphData[i+2]); if(e>=100000) e= 0;

		answer = (a + b + c + d + e)/5;
		if(graphData[i]<0 && graphData[i]>-100000) answer*=-1;
		tempData[i] = answer;
	}

	for(i=2;i<numberFrames-2;i++) graphData[i] = tempData[i];

	delete[] tempData;
}
