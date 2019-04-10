#define MASK_RES 540 // Mask Resolution. 180 points scaned is MASK_RES 180, 360 points is 360 etc
#define PROFILE_SIZE 180 // Final Profile size.
#include "CBSAlgebra.h"
#include "FishEyeTransform.h"
#include "lasertracking.h"
#include "Centre.h"


class RadialScan
{
public:
//	LaserTracking centreLaser;
//	int showLaserCentreCoord; //PCN4380

	LensTransform *FEye;
	struct profileAtom // Storage of all the data asociated with a profile point
		{
		double		radius;			// (double) The final profile point radius.
		int			atomIndex;		// The Atom that stores the profile point.
		vec2double  coordinate;		// (vec2double) float position of profile point.
		double		angle;			// (double) angle of atom in radians
		int mark;					// Mark it if it needs to be included in finding the centre
		};
	
	profileAtom	finalProfile[MASK_RES];
	
	RadialScan(void);
	~RadialScan(void);
	void Initialise(pixel **imPointer,int width, int height);
	void ShowMask(void); // Displays what the look up mask would be looking at.
	void SetOffset(int x, int y);
	void SetCounterMask(int x1, int y1, int x2, int y2);
	void SetTextMask(float x1, float y1, float x2, float y2, int setclear);
	void Process(void);
	double GetProfilePoint(int no);
	double GetProfilePointX(int no);
	double GetProfilePointY(int no);
	vec2double GetProfilePointXY(int no);
	double GetCenterX(void);
	double GetCenterY(void);
	double GetRadius(void);
	void RecalculateCoordinates(int sample);
	vec2double GetCentre(void) {return mask.offset;};
	double GetAverageRadius(int sample);
	void SetPickupLevel(int pickup) {pickupLevel = (double) pickup;} //PCNAVI 30 August 2004
	void SetThresholdLevel(int thresh) {/*threshold =  thresh;*/}
	void SetCutoffLevel(double cutoff) {cutoffLevel = cutoff;}
	void SetLaserWidth(int w) {laserWidth = w; if(laserWidth<4) laserWidth = 4;}
	void SetInternalRadius(double r) {internalRadius = r; if(internalRadius>0.96) internalRadius=0.98;}
	void SetExternalRadius(double r) {externalRadius = r; if(externalRadius<1.06) externalRadius=1.06;}


	void AdjustCentreFinalProfile(void);
	void SetFilterType(int i) {filterValue=i;};
	void SetWaterLevel(int index, int value) {egnoreList[index]=value;} //PCN2568 //PCN3219
	void SetIgnoreWaterLevel(int iWL) {ignoreWaterLevel=iWL;}
	void SetShowProfileWaterLevel(int sPWL) {showProfileWaterLevel=sPWL;}
	void SetLaserWidthOverlayOn(int onOff) {showLaserWidthOverlay=onOff;}
	void SetShowInternalProfilePoints(int onOff) {showInternalProfilePoints=onOff;}
	void SetShowProfileCandidatesOverlay(int onOff) {showProfileCandidatesOverlay=onOff;}
	void SetShowInternalCircles(int onOff) {showInternalCircle=onOff;}
	void SetShowVideoFilter(int onOff) {showVideoFilter = onOff;}
	int GetShowVideoFilter(void) {return showVideoFilter;} //PCN3215
	int IsInWaterSection(int i); //PCN2568 Is it in the water area of the profile
	int IsWaterLevelOn(void) {return ignoreWaterLevel;}
	int profileType;
	void PrintProfile(char *file); //PCN2993 was created to see the profile points as data
	void ShowPutPixels(void);
	void ShowDrawLines(void);
	void SetDebug1(int val) {debug1=val;}
	void SetDebug2(int val) {debug2=val;}
	int egnoreMaskHeight;
	unsigned char **egnoreMask; //The new mask will be pixel orrientated and not just a square regon, we will apaint where
	void SetColourAdjust(double red, double green, double blue){redAdjust = red; greenAdjust = green;blueAdjust=blue;}
	void TurnCentreOff(void) {CentreOff = true; }
	void TurnCentreOn(void) {CentreOff = false;}

	//PCN4539
	void LockDonut(void) {DonutLocked = true; }
	void UnlockDonut(void) {DonutLocked = false; }
	void SetExpectedDiameter(double d){expectedDiameter =(float) d;}
	void SetContrastBrightness(unsigned char *lookup) {memcpy(&contrastBrightness[0],lookup,256);}




private:
	void setim(int x, int y, int r, int g, int b);
	bool CentreOff; 
	bool DonutLocked; //PCN4539
	double GausianMask[7][7];
	Centre theCentre;
	int egnoreList[PROFILE_SIZE];
	float expectedDiameter; ///PCN4539
	unsigned char contrastBrightness[256];
	
	
	vec2double centreHistory[5];
	//Debug arrays added - 12 November 2004 - - - - - - - - - - - - -
	vec2double centreDebugHistory[10000]; int numberCentreDebug;	//
	char centreDebugHistoryNotes[10000][80];						//
	vec2double averageCentreHistory;								//
	vec2double previousCentre;
	int averageCentreHead;											//
	//- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	
	int debug1;
	int debug2;

	pixel	**imVideo;	// Pointer to original video image.
	pixel   **copyBuffer;
	bool	validPointer; // If the imVideo is a valid pointer mark as true.
						  // Used to stop accessing video when not initialised.
	double	cutoffLevel; // How low does laser line have to be before its not a possible profile point 
	double	pickupLevel;
	int		centreIn;
	int		centreOut;
	int		filterValue;
	int		hasPrinted;
	int		laserWidth;
	double internalRadius;
	double externalRadius;
	double centreHeight;

	vec2int counterMaskTopLeft;
	vec2int counterMaskBottomRight;
	vec2int textMaskTopLeft;
	vec2int textMaskBottomRight;
	vec2int laserWidthOverlay[MASK_RES*100]; int noLaserWidthOverlay;
	vec2int profileCandidatesOverlay[MASK_RES*100]; int noProfileCandidatesOverlay;
	struct putpixel{
		vec2double coord;
		unsigned char red;
		unsigned char blue;
		unsigned char green;
		};


	putpixel putPixels[20000];
	int noPutPixels;
	int showPutPixels;

	putpixel drawLines[20000];
	int noDrawLines;
	int showDrawLines;

	int wlLeft, wlRight; //PCN2568 
	int ignoreWaterLevel; //PCN2568 When when finding the centre egnore the water level profile points
	int showProfileWaterLevel; //PCN2568 When profiling with water level on, profile the water level profile points
	int showLaserWidthOverlay;
	int showInternalProfilePoints;
	int showProfileCandidatesOverlay;
	int showInternalCircle;
	int showVideoFilter;
	vec2double foundCentres[6000];
	int noFoundCentres;
	double redAdjust;
	double greenAdjust;
	double blueAdjust;

	// Ray Unit is what is stored in unit of a ray.
	struct rayAtom 
		{
		double		radius;			// (double) The calculated radius value for that point up the Ray
		vec2int		coordScreen;	// (vec2int) Pixel Coordinate from the video image
		int			processedImage; // (int) Processed pixel value from the raw r,g,b values
		int			originalImage;  // (int) orriginal Processd pixel before anything has been altered
		pixel		imVideo;		// (pixel) Raw video r,g,b value.
		int			edge;			// Storing the change in video image;
		double		angle;			// (double) angle of 2:3 ratio coordinate.
		vec2double	coord;			// Absolute Coordiante of 2:3 ratio.
		};

	struct ray
		{						
		rayAtom		*rayAtoms;		// (*rayAtom) Array of rayAtoms, Size of mask.rayLength
		int			lastEntry;		// (int) the last entry where there is a valid pixel to look up
		//double		profileRadius;	// (double) The final profile point radius.
		//int			atomIndex;		// The Atom that stores the profile point.
		//vec2double  coordinate;		// (vec2double) float position of profile point.
		
		profileAtom	profilePoint;	// (pointData) final profile point (radius, atomIndex, coordiante, angle)
		};

	struct profileMask
		{
		ray		rays[MASK_RES];	// Rays Array (Size of MASK_RES)
		vec2double offset;			// Center offset of the mask. 2:3 ratio
		int		rayLength;		// The lengths of the rays.
		double	averageRadius;	// The average radius of current profile points.
	} mask;

	int		imWidth;	// Width of video
	int		imHeight;	// Height of video.
	double	imRatio;	// Ratio between the Width and Height of the video, Needs to be 3:4
	vec2double		maskXYOffset;	// X and Y mask offset;


	void BuildMask(void); // Builds the lookup mask for all the rays
	void BuildRadiansTable(double *preCalculatedRadians); //PCN3156
	void GrabData(int sample); // Grabs the image data the fills the rays, sample, 1 all, 2 every 2nd, 3 every 3rd etc


	double ProcessRay(ray *singleRay, int pass, int loop, int display); // Our Black box, give it a ray, outputs a profile point
	void SmoothProfile(void);
	double Hypot(double x, double y); // returns the Hypotinuse to x and y relative to 0;
	double GetAngle(double x,double y);
	void AdjustCentre(int pass);

	void FindBestNeighbour(vec2double &curPoint, double size,  int pass);
	double GetCentreVariance(vec2double p,  int pass);
	inline double DistOfTwoPoints(vec2double pt1, vec2double pt2);
	void ShowAtom(int rays, int atom, int red, int green, int blue, int rel);
	void ShowProfilePoints(int sample, int red, int green, int blue, int centered);
	bool IsInMask(vec2int point);
	int Candidate(vec2double *edges,int in, int out, double greatestNeg, int &pos, int &neg,int sizeOfEdgeArray);
	void DownSample(void);
	int	FilterVideoPixel(int x, int y);
	double GetAverageProfile(int sample);
	vec2int FindBestCandidates(vec2int *candidates,int countCandidates,ray *singleRay);
	void RemoveRoughPoints(void); //PCN2993 Remove rough points from profile
	void RemoveFinalRoughPoints(void); //PCN2993, 3 September 2004 (Antony van Iersel)
	void SmoothLikeABabysBottom(void);

	void BlankScreen(void);
	void AddLaserWidthOverlay(vec2int coord);
	void AddProfileCandidatesOverlay(vec2int coord);
	void AddPutPixels(vec2double coord, unsigned char red, unsigned char green, unsigned char blue);
	
	// More debug drawing information 12 November 2004 ///
	void AddPutPixels(vec2int coord,  unsigned char red, unsigned char green, unsigned char blue)
	{													//
		AddPutPixels(vec2double((double) coord.x,(double) coord.y), red, green, blue);
	}													//
	//////////////////////////////////////////////////////
	
	void AddDrawLines(vec2double coord, unsigned char red, unsigned char green, unsigned char blue);
	void ShowLaserWidthOverlay(void);
	void ShowProfileCandidatesOverlay(void);
	void ShowInternalCircle(vec2double cen, double size, int red, int green, int blue, int centered);


	void DrawLine(vec2double coord_a, vec2double coord_b, unsigned char red, unsigned char green, unsigned char blue);
	void ShowVideoFilter(void);

	inline vec2double GetCoordinate(vec2double vector);
	vec2double GetVector(vec2double coord);
	void RecalculateAverageCentre(void); // PCN3122 - 12 November 2004
	void CreateGausianMask(void);
	void FindMedianXCentre(vec2double &centre);
	void AdjustCentreWithSmartDataFill(void);
};


