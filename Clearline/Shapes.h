#include "..\houghlibv2.0\CBSAlgebra.h"

#ifndef INCLUDSHAPES
#define INCLUDSHAPES

class Shapes
{
public:
	
	ReferenceShape_V10 *ReferenceShape;
	Shapes(ReferenceShape_V10 *_Shape,
 		   double _ShapeRadius,
		   double _ShapeCentreX,
		   double _ShapeCentreY,
		   double _ShapeRotation);
	~Shapes();
void ProfileRefShapeDistCalc(float X, 
							 float Y, 
						     double *OrthoX, 
						     double *OrthoY, 
						     double *OrthoDistance);
double ProfileRefShapeDistCalc(vec2double point);
private:

	
	vec2double ShapeCentre;
	double ShapeRotationAngle;
	double ShapeRadius;
bool ProfileRefShapeDistCalcArc(double X,
							   double Y,
							   double Radius,
							   double ArcStart,
							   double ArcEnd,
							   double &OrthoX,
							   double &OrthoY,
							   double &OrthoDistance);
bool ProfilerRefShapeDistCalcLine(vec2double Cursor,
								  double AX,
								  double AY,
								  double BX,
								  double BY,
								  double &OrthoX,
								  double &OrhtoY,
								  double &OrthoDistance);

};

#endif