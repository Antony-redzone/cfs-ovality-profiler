#ifndef INCLUDFISHEYETRANSFORM
#define INCLUDFISHEYETRANSFORM

#include "CBSAlgebra.h"

class LensTransform{
public:
	LensTransform();
	~LensTransform();

	void Flush(void);
	void Initialize(int CurrentWidth, int CurrentHeight, int BaseWidth, int BaseHeight);
	void SetOffsets(int x,int y);
	void SetDisplayScale();
	void SetTFactor(double TF);
	void SetImageSize(void);
	void SetOriginalSize(int width,int height);
	void SetScale(double S);
	double GetScale(void);
	void SetYScale(double scale) {y_scale = scale;}	//PCN3303
	double GetYScale(void) {return y_scale;}
	void SetYPosition(double position) {y_position = position;} //PCN3303

	void TurnFishEyeOn(void);
	void TurnFishEyeOff(void);
	bool FishEyeStatus(void);

	void CalculateOldFishEyeScale(void);
	double OldFishEyeLookupTable(double Factor);

	void CreateMask(void);
	void Transform (pixel **OriginalImage);
	void Convert (double x,double y);
	void ConvertPoint(double &x, double &y);
	void ConvertPoint(vec2double &point);
	void CopyToVideo(pixel ** Destination,int xBound,int yBound);
	void CopySinglePixel(int x,int y,pixel **Destination);
	vec2int ** Mask;
	pixel ** buffer;

	int ImageWidth,ImageHeight;

	void AutoCalibrate(void);

private:

	bool Initialised;	// True if class correctly initialised
	int Active;
	double y_scale; // Used to stretch image on the y axis. Default 1, no stretch
	double y_position; // Used to reposition the video image and profile on the y axis

	vec2double grid[3][3];
	vec2double Basegrid[3][3];
	void copy_grid(void);
	
	int bestx,besty;
	double best,bestTFactor;
	double Divider;

	void ConvertPoint(double &x,double &y,int cx,int cy, double Factor);

	void Find_Centre(int x,int y,double Factor);
	double Find_TFactor(int x, int y,double Factor);
	void Load_Grid();
	double Asses_Grid(int x,int y,double f);
	double Square(vec2double ul,vec2double ur,vec2double bl, vec2double br);
	void FEGrid(int x, int y,double f);


	int xCentre,yCentre;

	int OriginalHeight,OriginalWidth;
	double TFactor;
	double Scale;
	double y_ratio,x_ratio,ratio34;

	bool MaskCreated;

};

#endif
