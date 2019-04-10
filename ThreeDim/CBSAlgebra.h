#ifndef PI
#define PI 3.1415926535897932384626433832795
#endif
#ifndef INCLUDEAGLEBRA
#define INCLUDEAGLEBRA

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

public:
	vec2double(void) {x=y=0;}
	vec2double(double a) {x=y=a;}
	vec2double(double a, double b) {x=a;y=b;}
	double x;	// Double Value for 2D vector
	double y;	// Double Value for 2D vector
	vec2double operator+(vec2double p) { p.x+=x; p.y+=y; return p; }
	vec2double operator-(vec2double p) { p.x=x-p.x; p.y=y-p.y; return p; }
	vec2double operator*(vec2double p) { p.x*=x; p.y*=y; return p; }
	vec2double operator/(vec2double p) { p.x/=x; p.y/=y; return p; }
	vec2double operator*(double f) { 	vec2double temp; temp.x=x*f; temp.y=y*f; return temp; }
	vec2double operator/(double f) { 	vec2double temp; temp.x=x/f; temp.y=y/f; return temp; }
	bool	operator==(vec2double p) {return((p.x==x) && (p.y==y)) ? true:false;}
	bool	operator!=(vec2double p) {return((p.x!=x) || (p.y!=y)) ? true:false;}
	bool	operator!=(double i) {return((x!=i) && (y!=i)) ? true:false;} 
	bool	operator==(double i) {return((x==i) && (y==i)) ? true:false;} 
};

inline double DistOfTwoPoints(vec2double pt1, vec2double pt2)
{
	return sqrt(pow(pt1.x-pt2.x,2)+pow(pt1.y-pt2.y,2));
}
#endif