#include "Video.h"
#include "CBSAlgebra.h"


class LaserTracking
{
public:

	LaserTracking(void);
	~LaserTracking(void);
	void SetVideoPointer(pixel **im,int h,int w);
	void SearchForLaser(void);
	void SetLaserSize(int size) {laserSize=size;}
	void SetLaserSizeMin(int size) {laserSizeMin=size;}
	void SetLaserWidthHeight(int size) {laserWidth=size; laserHeight=size;}
	void SetLaserIntensity(int i) {intensity = i;}
	void SetLaserOneCoord(int x, int y) {laserOneCoord.x=x; laserOneCoord.y=y;}
	void SetLaserTwoCoord(int x, int y) {laserTwoCoord.x=x; laserTwoCoord.y=y;}
	void SetLaserCentreCoord(int x, int y) {laserCentreCoord.x=x; laserCentreCoord.y=y;}
	void SetLaserLeftSideCoord(int x, int y) {laserLeftSideCoord.x=x; laserLeftSideCoord.y=y;}
	void SetLaserRightSideCoord(int x, int y) {laserRightSideCoord.x=x; laserRightSideCoord.y=y;}
	vec2double GetLaserOneCoord(void) {return laserOneCoord;}
	vec2double GetLaserTwoCoord(void) {return laserTwoCoord;}
	vec2double GetLaserCentreCoord(void) {return laserCentreCoord;}
	vec2double GetLaserLeftSideCoord(void) {return laserLeftSideCoord;}
	vec2double GetLaserRightSideCoord(void) {return laserRightSideCoord;}
	int GetLaserTrackingSize(void) {return trackingSize;}
	void DisplaySearchBox(int red, int green, int blue);

private:
	struct Shape
	{
		int left,right,up,down; // Exrimities of the pixels in shape
		int size;				// Total number of pixels in shape
		double averageBrightness; // Average overall pixels brightness 
		int averageX, averageY; // The average x and y is also the centre of the shape
		int width,height;		// Total height and width of shape
	};
	Shape foundShapes[100];
	Shape foundLaserOne[100];
	Shape foundLaserTwo[100];
	Shape foundLaserCentre[100];
	Shape foundLaserLeftSide[100];
	Shape foundLaserRightSide[100];

	struct ImageBuffer
	{
		pixel **image;
		int height;
		int width;
	};

	int laserWidth,laserHeight;
	int trackingSize;
	int trackingSize20;
	int laserSize;
	int laserSizeMin;
	double intensity;

	pixel **imVideo; // Original video to copy the data accross
	int imHeight;
	int imWidth;
	ImageBuffer imBuffer; // Processing video image to keep seperate from original video
	ImageBuffer imLaserOne;
	ImageBuffer imLaserTwo;
	ImageBuffer imLaserCentre;
	ImageBuffer imLaserLeftSide;
	ImageBuffer imLaserRightSide;

	vec2double laserOneCoord;
	vec2double laserTwoCoord;
	vec2double laserCentreCoord;
	vec2double laserLeftSideCoord;
	vec2double laserRightSideCoord;

	int FindNextShape(ImageBuffer &imBuffer,  int x, int y,double bright, Shape &currentShape, int trackParent);
	void MarkShapeCentre(Shape currentShape);
	void MarkTarget(vec2double coord, char red, char green, char blue);
	void CopyImageToBuffer(ImageBuffer &imBuff, vec2double coord);
	void DrawLine(vec2double coord_a, vec2double coord_b, unsigned char red, unsigned char green, unsigned char blue);
	


};
