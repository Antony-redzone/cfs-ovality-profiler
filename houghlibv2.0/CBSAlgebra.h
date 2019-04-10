#ifndef PI
#define PI 3.1415926535897932384626433832795
#endif


#ifndef INCLUDEAGLEBRA
#define INCLUDEAGLEBRA

#define ON  1
#define OFF 0


#include <math.h>
#include <windows.h>

#ifndef ULONG_PTR
#define ULONG_PTR ULONG
//ULONG_PTR m_gdiplusToken;
#endif

struct ShapeArc_V10
{
    float	OriginX;
    float	OriginY;
    float	Radius;
    float	startAngle;
    float	endAngle;
    int		Colour;
};

// Type define for line shape, start x,y coord and end x,y coord
struct ShapeLine_V10
{
    float	StartX;
    float	StartY;
    float	EndX;
    float	EndY;
    int		Colour;
};

// Type define for a shape, reason for Use defined is some pipes have a different shape
// for external and intenal pipe shape
struct ReferenceShape_V10
{
    char name[256]; //This is the name of the shape type eg Egg, Circle, HorseShoe etc
    char Use[256]; //Type of use, eg "Internal", "External", "All" etc
    ShapeArc_V10 Arcs[129];
	int NoArcs;
    ShapeLine_V10 Lines[129];
	int NoLines;
    float CentreOffsetX;
    float CentreOffsetY;
    int Colour;
};

class vec2int
{
public:
	int x;	// Interger Value for 2D vector
	int y;	// Interger Value for 2D vector
	vec2int(void) {x=y=0;}
	vec2int(int a) {x=y=a;}
	vec2int(int a, int b) {x=a;y=b;}
	vec2int	operator+(vec2int p) { p.x+=x; p.y+=y; return p; }
	vec2int operator-(vec2int p) { p.x-=x; p.y-=y; return p; }
	vec2int	operator*(vec2int p) { p.x*=x; p.y*=y; return p; }
	vec2int operator/(vec2int p) { p.x/=x; p.y/=y; return p; }
	vec2int operator/(int i)     { vec2int temp; temp.x=x/i; temp.y=y/i; return temp; }
	bool	operator==(vec2int p) {return((p.x==x) && (p.y==y)) ? true:false;}
	bool	operator!=(vec2int p) {return((p.x!=x) || (p.y!=y)) ? true:false;}
	bool	operator!=(int i) {return((x!=i) && (y!=i)) ? true:false;}
	bool	operator==(int i) {return((x==i) && (y==i)) ? true:false;}
	bool	operator>(int i) { return ((x > i) || (y > i)) ? true : false;}
	bool	operator<(int i) { return ((x < i) || (y < i)) ? true : false;}
};

class vec2double
{ 
	private:
	int	   t;   // int Value of type (0 = vector / 1 = coordinate)
public:
	double x;	// Double Value for 2D vector
	double y;	// Double Value for 2D vector


	vec2double(void) {x=y=0; t=0;}
	vec2double(double a) {x=y=a; t=0;}
	vec2double(double a, double b) {x=a;y=b; t=0;}
	vec2double operator+(vec2double p) { p.x+=x; p.y+=y; return p; }
	vec2double operator-(vec2double p) { p.x=x-p.x; p.y=y-p.y; return p; }
	vec2double operator*(vec2double p) { p.x*=x; p.y*=y; return p; }
	vec2double operator/(vec2double p) { p.x/=x; p.y/=y; return p; }
	vec2double operator*(double f) { 	vec2double temp; temp.x=x*f; temp.y=y*f; return temp; }
	vec2double operator/(double f) { 	vec2double temp; temp.x=x/f; temp.y=y/f; return temp; }
	bool	operator==(vec2double p) {return((p.x==x) && (p.y==y)) ? true:false;}
	bool	operator!=(vec2double p) {return((p.x!=x) || (p.y!=y)) ? true:false;}
	bool	operator!=(double i) {return((x==i) && (y==i)) ? false:true;} 
	bool	operator==(double i) {return((x==i) && (y==i)) ? true:false;} 
	double length(void) {return sqrt((x*x)+(y*y));}
	double length(vec2double p) {return sqrt( ((x-p.x) * (x-p.x)) + ((y-p.y) * (y-p.y)) );}
	vec2double toCoordinate(void) {t=0; return vec2double(sin(x)*y,cos(x)*y);}
	vec2double toVector(void)
	{
		double adj;
		double opp;
		
		double dist = sqrt((x*x) + (y*y));
		
		adj = fabs(x);
		opp = fabs(y);
		t = 1; // Object is now vector coordinates
		
		if((x==0) && (y==0)) return vec2double(0,0);
		if((x==0) && (y>0)) return vec2double(0					  ,dist);
		if((x==0) && (y<0)) return vec2double(PI				  ,dist);
		if((x>0) && (y==0)) return vec2double((PI/2)			  ,dist);
		if((x<0) && (y==0)) return vec2double(PI+(PI/2)		      ,dist);
		if((x>0) && (y>0))  return vec2double(       atan(adj/opp),dist); // + , +
		if((x>0) && (y<0))  return vec2double(PI-    atan(adj/opp),dist); // + , -
		if((x<0) && (y<0))  return vec2double(PI+    atan(adj/opp),dist); // - , -
		if((x<0) && (y>0))  return vec2double((2*PI)-atan(adj/opp),dist); // - , +
		return 0;
	}

	vec2double rotateCoordinate(vec2double origin, double radians)
	{
		vec2double shifted;
		vec2double rotatedCoord;

		// Shift the coordinates relative to the centre to be shifted
		shifted = vec2double(x,y)-origin;
		
		// X '         = cos(theta)*x - sin(theta)*y
		// Y '         = sin(theta)*x + cos(theta)*y
		rotatedCoord.x = (shifted.x * cos(radians)) - (shifted.y * sin(radians));
		rotatedCoord.y = (shifted.x * sin(radians)) + (shifted.y * cos(radians));

		rotatedCoord=rotatedCoord+origin;
		return rotatedCoord;
	}
	vec2double rotateCoordinate(double radians) {return rotateCoordinate(vec2double(0,0),radians);}
	vec2double toEndOfVector(vec2double origin)
	{
		vec2double endPoint;
		endPoint.x = sin(x) * y;
		endPoint.y = cos(x) * y;
		return origin+endPoint;
	}
	vec2double toEndOfVector(void) {toEndOfVector(vec2double(0,0));}

};

class Area
{ 
public:
	double Triangle(vec2double a, vec2double b, vec2double c)
	{
		double area;
		area=((a.x*b.y)+(a.y*c.y)+(b.x*c.y)-(c.x*b.y)-(c.y*a.x)-(a.y*b.x))/2;
		return fabs(area);
	}
};

class Intersection
{
public:
	void TwoLines(vec2double point1, vec2double point2, 
	 				     vec2double point3, vec2double point4,
						 vec2double &answer,
						 bool &line1intersect, bool &line2intersect)
	{
		double UAnumerator,UBnumerator;
		double UA,UB;
		double denominator;
		vec2double ans;
		


		UAnumerator = (point4.x - point3.x) * (point1.y - point3.y) - (point4.y - point3.y) * (point1.x - point3.x);
		UBnumerator = (point2.x - point1.x) * (point1.y - point3.y) - (point2.y - point1.y) * (point1.x - point3.x);
		
		denominator = (point4.y - point3.y) * (point2.x - point1.x) - (point4.x - point3.x) * (point2.y - point1.y);

		if(denominator == 0) {line1intersect = false; line2intersect = false; return;}


		UA = UAnumerator / denominator;
		UB = UBnumerator / denominator;


		answer.x = point1.x + UA * (point2.x - point1.x);
		answer.y = point1.y + UA * (point2.y - point1.y);

		if(UA >= 0 && UA <= 1) line1intersect = true; else line1intersect = false;
		if(UB >= 0 && UB <= 1) line2intersect = true; else line2intersect = false;
	}
};
#endif
