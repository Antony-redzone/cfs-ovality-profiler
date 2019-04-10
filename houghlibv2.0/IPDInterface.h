#include <stdio.h>

class IPDInterface
{
public:
	IPDInterface(char *_fileName);
	~IPDInterface(void);
	int Initialise(void);
	int GetDistance(double t);

private:
	int *distance;
	double *time;
	char fileName[800]; //PCN3744 Make sure if the IPD file is really deep in the directory, that there is enough room
						//to store the file name, had this problem with th CBS_Video dll, in cleanflow. Two arrays
						//for the file string were to short (8 November 2005,Antony)
	FILE *filePointer;
	int numberOfLines;

	void LoadData(void);
	int GetNumberFileLines(void);
	void SplitLine(char *line, char *left, char *right, char c);
	int SearchTime(double t,int index,int step);
};