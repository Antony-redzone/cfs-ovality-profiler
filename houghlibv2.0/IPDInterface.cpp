#include "IPDInterface.h"
#include <stdlib.h>
#include <string.h>
#include <stdio.h>
#include  <math.h>

IPDInterface::IPDInterface(char *_fileName)
{
	distance = NULL;
	time = NULL;
	numberOfLines = 0;
	strcpy(fileName,_fileName);

}

IPDInterface::~IPDInterface(void)
{
	if(distance!=NULL) {delete distance; distance = NULL;}
	if(time != NULL)   {delete time;     time = NULL;}
}

int IPDInterface::Initialise(void)
{
	int i,j=0;
	char IPDExtension[] = ".IPD";
	int len;			//This was char, big big big oops
	len = strlen(fileName);
	if(len<5) return -1;
	
	for(i=len-4;i<len;i++) 
	{
		fileName[i]=IPDExtension[j++];
	}
	

	filePointer = fopen(fileName,"r");
	if(filePointer == NULL) return -1;
	fclose(filePointer);

	LoadData();
	return 0;
}

void IPDInterface::LoadData(void)
{
	char line[256];
	char left[256];
	char right[256];

	int dist=0;
	double t=0;
	int count=0;

	numberOfLines = GetNumberFileLines();
	distance = new int[numberOfLines];
	time = new double[numberOfLines];

	if ((filePointer = fopen(fileName,"r"))==NULL) return;

	while(fscanf(filePointer,"%255s",line)!=EOF)
	{
		SplitLine(line,left,right,';');
		dist=atoi(left);
		t=atof(right);

		distance[count]=dist;
		if(count<numberOfLines) time[count++]=t;
	}
}

int IPDInterface::GetNumberFileLines(void)
{
	int count=0;
	char dummy[256];

	filePointer = fopen(fileName,"r");
	while(fscanf(filePointer,"%255s",dummy)!=EOF)
	{
		count++;
	}
	fclose(filePointer);

	return count;
}

void IPDInterface::SplitLine(char *line, char *left, char *right, char c)
{
	int i=0;
	int j=0;

	while((line[i]!=c) && (line[i]!=0) && (i<255))
	{
		left[i] = line[i];
		i++;
	}

	left[i++]=0;
	while(line[i]!=0 && (i<255))
	{
		right[j++]=line[i++];
	}	
	right[j]=0;
}

int IPDInterface::GetDistance(double t)
{
	int dist  = 0;
	int step = numberOfLines/4;
	int index = numberOfLines/2;

	if (numberOfLines==0) return 0;

	if(index<0) index=0;
	if(index>(numberOfLines-1)) index = numberOfLines-1;
	
	dist = SearchTime(t,index,step);
	return dist;
}

int IPDInterface::SearchTime(double t,int index, int step)
{
	double leftDiff;
	double rightDiff;
	double centreDiff;

	int leftIndex;
	int rightIndex;
	
	leftIndex = (index-step); if(leftIndex<0) leftIndex=0;
	rightIndex = (index+step); if(rightIndex>(numberOfLines-1)) rightIndex=numberOfLines-1;

	
	leftDiff=fabs(time[leftIndex]-t);
	rightDiff=fabs(time[rightIndex]-t);
	centreDiff=fabs(time[index]-t);
	if(step==1)
	{
		if((leftDiff<rightDiff) && (leftDiff<centreDiff)) return distance[leftIndex];
		if((rightDiff<leftDiff) && (rightDiff<centreDiff)) return distance[rightIndex];
		return distance[index];
	}
	if((leftDiff<rightDiff) && (leftDiff<centreDiff)) return SearchTime(t,leftIndex,(step+1)/2);
	if((rightDiff<leftDiff) && (rightDiff<centreDiff)) return SearchTime(t,rightIndex,(step+1)/2);
	return SearchTime(t , index, step/2);
}
