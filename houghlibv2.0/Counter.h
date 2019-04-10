
#define X 0
#define Y 1

#define METRIC 0
#define IMPERICAL 1

struct history
	{
	int x;
	int y;
	int nextDecimal;
	};

struct edge
	{
	int avgPixel;
	int edgeRight;
	int edgeLower;
	int edgeDiag;
	int display;
	int marked;
	};

struct point
	{
	int i,j;
	edge e;
	};

struct trace
	{
	int numberPoints;
	int i_max;
	int i_min;
	int j_max;
	int j_min;
	int i_width;
	int j_width;
	point p[600];
	int flagged;
	};

struct object
	{
	int i_min; 
	int i_max;
	int j_min; 
	int j_max;
	int i_width;
	int j_width;
	int flagged;
	};
	
class Counter
{
public:

	int isDecimalFound;
	int contrast;
	int neighContr;
	int count;
	int numPixDiff;
	int threshold;
	int bufferValue;
	int xLeft,xRight,yTop,yLower;		// Counter Mask Left, Right, Top, Lower
	int sxLeft,sxRight,syTop,syLower;	// Secound Mask, where the counter is looking for changes
	int imWidth,imHeight;
	int maskWidth, maskHeight;
	int decimal_x,decimal_y;
	pixel **imVideo; // 
	edge **imMask; // Mask image, holds a copy of data in the mask
	trace traceMask[100];
	history foundDecimals[100];
	int countDecimals;
	int **chBuff;
	int neg;
	int totalAverageChange;
	int head,tail;
	int glitch;
	int direction;
	int isSet;
	int sampleColour[100];
	int countTraces;
	int countObjects;
	int lengthOffLine;
	int textColour;
	int pointSize;
	int sxHistory;
	int units; // Metric = 0 or Imperical = 1;

	Counter(void);
	~Counter();
	void SetCounterMask(int xL, int xR, int yT, int yL);
	void SetCounterPointer(pixel **originalIm, int w, int h);
	void ResetMaskOverSingleDigit(void);
	int CountDiff(void);
	int Sample(void);
	void CopyImageToMaskEdge(void);
	void CopyMaskEdgeToImage(void);
	void AutoContrast(void);
	int CountWhite(void);
	void TraceMask(void);
	int CheckEdgeRight(int i, int j);
	int CheckEdgeLower(int i, int j);
	int CheckEdgeDiag(int i, int j);
	void ExtractTracedObjects(void);
	void GetTracePositions(void);
	void SortTraces(void);
	void GetLineOff(int i, int j);
	int TestForEdge(int i,int j,int contrast);
	void ResampleDecimalPoint(void);
	void PositionTrace(void);
	void FindDigitNextToDecimal(void);
	void FindDecimalPoint(int &x, int &y);
	void FindHole(int &x, int &y);
	void ExtractObject(int i, int j, int obj);
	void Tick(void);
	void DrawCircle(int i, int j);
	void DrawSquare(int p1, int p2);
	void MarkTracesForDisplay(void);
	int  GetGreatestEdge(void);
	int  GetAverageEdge(void);
	int  IsInCounterArea(int x, int y);
	int  GetCommonTopHeight(void);
	int  GetCommonLowerHeight(void);
	void DrawTopOfCounter(int i);
	
	

private:
	int bufferSize;
};