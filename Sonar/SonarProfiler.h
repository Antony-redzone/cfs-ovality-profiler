#include "1512USBInterface.h"
#include "Configuration.h"
#include "Definitions.h"
#include "OutlineProcessor.h"
#include "SensorData.h"
#include <time.h>


Sonar::C1512USBComm *usbComm;
COutlineProcessor	*outline;

CConfiguration m_Config;
//CConfiguration *m_Config;
EnumStepSize ss;
EnumCentreAngle ca;
EnumArcSize as;

const int CALIBRATE_SWEEPS = 10;
_SYSTEMTIME calibrateBuffer[CALIBRATE_SWEEPS + 1];
_SYSTEMTIME currentSystemTime;
int syncCount;


int imgW, imgH;
int frameCount;
unsigned char *Image;
int previewImage;
int bytesPerPixel;
char outputFileName[300];
double scanAngle;
int pos;
double homeAngle;

int lowPass;
int highPass;
float slope;
int halfWay;

int logoWidth;
int logoHeight;
int logoOffsetX;
int logoOffsetY;
int copyWidth;
bool drawLogo;

double previousMaxRad;
double maxRad;
int maxSample;

int startX, startY;
double scale;
int xCentre, yCentre;



unsigned char Pallette[256];
unsigned char LogoData[200000];
unsigned char ProfileData[10000000];

clock_t start;

int SAMPLES;
int STEPS;
int MAXDATASIZE;

int cableOffset;

int __stdcall StartSonarSweep(int width, int height, int preview, int arcSize, int centreAngle, int sampleRate, int overSamples, int samples, int sSize, int pSize);
//int __stdcall InitialiseDLL(int width, int height, int preview);
void OnSonarUpdate(void);
void SaveProfileToFile(void);
void __stdcall StopScanning();
void __stdcall StartScanning(char *directory);

void drawray(double ang, int start, int finish,unsigned char *Img);
void __stdcall drawframe(unsigned char *Img);
void __stdcall drawlineBMP(int x1, int y1,int x2, int y2, BYTE col, unsigned char *Img);
void __stdcall ReadProfileFromFile(unsigned char *Img, int bytes, char *filename, int W, int H, int *Hours, int *Mins, int *mSec, int *Dist);
void __stdcall CheckForScanning(int *status);
void __stdcall VBRay(unsigned char *Img);
void DrawSquare(unsigned char *Img,int pos, BYTE col);
void ClearImage(unsigned char * Img);
void setPixel(int x, int y, unsigned char *Img,BYTE col);
void drawArc(double radius, double ang1, double ang2,unsigned char *Img, BYTE col, int XOffset, int YOffset);
void setRadarColour(unsigned char *Img, BYTE col,int x, int y, double radius, bool large);
void __stdcall LoadPallette(unsigned char *Img);
void __stdcall LoadLogo(unsigned char *Img,int w, int h);
void __stdcall SynchronisedStart(int *secs, int *millisecs, _SYSTEMTIME *startTime, int countDown);
void __stdcall SetScanAngle(double Angle);
void __stdcall DrawCircle(unsigned char *Img, int width, int height, double radius);
void fillCircle(int xCentre, int yCentre, int width, unsigned char *Img, BYTE col);

void drawrayOLD(double ang, int start, int finish,unsigned char *Img);
void drawArcOLD(int radius, double ang1, double ang2,unsigned char *Img, BYTE col);
void __stdcall InitialiseDLL();
void __stdcall SetCablePayoutStart(int cable);
void __stdcall GetCablePayout(int *cable);
void drawAverageRay(double angle,int i,unsigned char Img);
void drawIMGframe(float centreX, float centreY, unsigned char *Img);
void getRayAddress(int &rayNo, int &rayPos, int x,int y);
void filterData(void);
