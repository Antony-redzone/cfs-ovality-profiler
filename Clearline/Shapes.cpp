#include "Shapes.h"

Shapes::Shapes(ReferenceShape_V10 *_Shape,
 			   double _ShapeRadius,
			   double _ShapeCentreX,
			   double _ShapeCentreY,
			   double _ShapeRotation)
{
	ReferenceShape =	 _Shape;
	ShapeRadius =		 _ShapeRadius;
	ShapeCentre.x =		 _ShapeCentreX;
	ShapeCentre.y =		 _ShapeCentreY;
	ShapeRotationAngle = _ShapeRotation;
}

Shapes::~Shapes()
{
}

double Shapes::ProfileRefShapeDistCalc(vec2double point)
{
	double orthoDist;
	double orthoX;
	double orthoY;
	ProfileRefShapeDistCalc((float) point.x, (float) point.y,&orthoX,&orthoY,&orthoDist);
	return orthoDist;
}

void Shapes::ProfileRefShapeDistCalc(float X, 
									  float Y, 
									  double *OrthoX, 
									  double *OrthoY, 
									  double *OrthoDistance)
//'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
//'Name    : ProfileRefShapeDistCalcSemiElliptical
//'Created : 10 December 2004, PCN3055
//'Updated :
//'Prg By  : Antony van Iersel
//'Param   : X - Current cursor X position
//'          Y - Current cursor Y position
//'          OrthoX - The orthoganal X from the reference shape to (X,Y)
//'          OrthoY - The orthoganal Y from the reference shape to (X,Y)
//'          OrthoDistance - The orthoganal distance from the reference shape to (X,Y)
//'Desc    : Finds the closest point on the current reference shape to X,Y and
//'          calculates distance and angles.
//'Usage   :
//'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
{

	double DegToRad;
	vec2double Coord;
	vec2double CoordOffset;
	vec2double CoordA;
	vec2double CoordB;
	double Radius;
	double ArcStart;
	double ArcEnd;


//	int		ShortestArcIndex;
	int		NumberOfPartsFound;
	vec2double Cursor;
	vec2double ClosestOrtho;
	vec2double ClosestCoord;
	vec2double CurentOrtho;
	double ClosestDistance;	
	double CurentDistance;

	double Unit;
	int i;
	int NumberOfArcs;
	int NumberOfLines;

	Cursor.x = X;
	Cursor.y = Y;

	DegToRad = PI / 180;
	Unit = ShapeRadius;
	NumberOfPartsFound = 0;
	


	CoordOffset.x = ReferenceShape->CentreOffsetX * Unit;
	CoordOffset.y = ReferenceShape->CentreOffsetY * Unit * -1;

	// Move everything relative to shape 0,0 coordinate to calculate arcs normals
	Cursor = Cursor.rotateCoordinate(ShapeCentre,-ShapeRotationAngle);
	Cursor = Cursor - CoordOffset - ShapeCentre;



	NumberOfArcs = ReferenceShape->NoArcs;
	NumberOfLines = ReferenceShape->NoLines;

	if(NumberOfArcs!=0)
	{
		Coord.x = ReferenceShape->Arcs[0].OriginX * Unit;
		Coord.y = ReferenceShape->Arcs[0].OriginY * Unit;
		Radius = ReferenceShape->Arcs[0].Radius * Unit;
		ArcStart = (ReferenceShape->Arcs[0].startAngle) * DegToRad;
		ArcEnd = (ReferenceShape->Arcs[0].endAngle) * DegToRad;
		Coord.x = Coord.x;
		Coord.y *= -1;
		if(ProfileRefShapeDistCalcArc(Cursor.x - Coord.x, 
									  Cursor.y - Coord.y, 
									  Radius, 
									  ArcStart, 
									  ArcEnd, 
									  CurentOrtho.x, 
									  CurentOrtho.y, 
									  CurentDistance))
		{
			NumberOfPartsFound++;
			ClosestDistance = CurentDistance;
			ClosestOrtho = CurentOrtho;
			ClosestCoord = Coord;
//			ShortestArcIndex = 0;
		}
	}

	

	for(i = 1;i<NumberOfArcs;i++)
	{
		Coord.x = ReferenceShape->Arcs[i].OriginX * Unit;
		Coord.y = ReferenceShape->Arcs[i].OriginY * Unit;
		Radius = ReferenceShape->Arcs[i].Radius * Unit;
		ArcStart = (ReferenceShape->Arcs[i].startAngle) * DegToRad;
		ArcEnd = (ReferenceShape->Arcs[i].endAngle) * DegToRad;
		Coord.x = Coord.x;
		Coord.y *= -1;
		if(ProfileRefShapeDistCalcArc(Cursor.x - Coord.x, 
									  Cursor.y - Coord.y, 
									  Radius, 
									  ArcStart, 
									  ArcEnd, 
									  CurentOrtho.x, 
									  CurentOrtho.y, 
									  CurentDistance))
		{
			NumberOfPartsFound++;
			if((fabs(CurentDistance) < fabs(ClosestDistance)) || (NumberOfPartsFound == 1))
			{
				ClosestDistance = CurentDistance;
				ClosestOrtho = CurentOrtho;
				ClosestCoord = Coord;
							
//				ShortestArcIndex = i;
			}
		}
	}
	
	if((NumberOfPartsFound == 0) && NumberOfLines != 0)
	{
		CoordA.x = ReferenceShape->Lines[0].StartX * Unit;
		CoordA.y = ReferenceShape->Lines[0].StartY * Unit;
		CoordB.x = ReferenceShape->Lines[0].EndX * Unit;
		CoordB.y = ReferenceShape->Lines[0].EndY * Unit;
		CoordA.y *= -1;
		CoordB.y *= -1;

		if(ProfilerRefShapeDistCalcLine(Cursor,
										CoordA.x,
										CoordA.y,
										CoordB.x,
										CoordB.y,
										CurentOrtho.x,
										CurentOrtho.y,
										CurentDistance))
		{
			NumberOfPartsFound++;
			ClosestDistance = CurentDistance;
			ClosestOrtho = CurentOrtho;
			ClosestCoord = 0;
		}

	}


	
	for(i = 0;i<NumberOfLines;i++)
	{
		CoordA.x = ReferenceShape->Lines[i].StartX * Unit;
		CoordA.y = ReferenceShape->Lines[i].StartY * Unit;
		CoordB.x = ReferenceShape->Lines[i].EndX * Unit;
		CoordB.y = ReferenceShape->Lines[i].EndY * Unit;
		CoordA.y *= -1;
		CoordB.y *= -1;
		if(ProfilerRefShapeDistCalcLine(Cursor,
										CoordA.x,
										CoordA.y,
										CoordB.x,
										CoordB.y,
										CurentOrtho.x,
										CurentOrtho.y,
										CurentDistance))
		{
			NumberOfPartsFound++;
			if((fabs(CurentDistance) < fabs(ClosestDistance)) || (NumberOfPartsFound == 1))
			{
			ClosestDistance = CurentDistance;
			ClosestOrtho = CurentOrtho;
			ClosestCoord = 0;
			}
		}
		
	}

	// If no valid arcs are found then exit
	if(NumberOfPartsFound == 0)
	{
		*OrthoX = X;
		*OrthoY = Y;
		*OrthoDistance = 0;
		return;
	}
	vec2double OrthoRotate;
	OrthoRotate = ClosestOrtho + ShapeCentre + CoordOffset + ClosestCoord;
	OrthoRotate = OrthoRotate.rotateCoordinate(ShapeCentre,ShapeRotationAngle);

	*OrthoX = OrthoRotate.x;
	*OrthoY = OrthoRotate.y;
	*OrthoDistance = ClosestDistance;
}

bool Shapes::ProfileRefShapeDistCalcArc(
										double X,
									    double Y,
									    double Radius,
									    double ArcStart,
									    double ArcEnd,
									    double &OrthoX,
									    double &OrthoY,
									    double &OrthoDistance)
{
//'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
//'Name    : ProfileRefShapeDistCalcArc
//'Created : 13 December 2004 2004, PCN3055
//'Updated :
//'Prg By  : Antony van Iersel
//'Param   : X - Current cursor X position
//'          Y - Current cursor Y position
//'          Radius - radius of arc
//'          ArcStart - starting angle of arc (Anti Clockwise, 0 radians East)
//'          ArcEnd - ending angle of arc (Anti Clockwise, 0 radians East)
//'          OrthoX - The orthoganal X from the reference shape to (X,Y)
//'          OrthoY - The orthoganal Y from the reference shape to (X,Y)
//'          OrthoDistance - The orthoganal distance from the reference shape to (X,Y)
//'Desc    : Finds the distance from the current point to the arc tangent,
//'          Returns True or False if inside the arc,
//'          Sets Distance from Current X,Y to Ortho X,Y and -ve or +ve if inside or outside expected radius
//'          Sets OrthoX, Ortho Y, the point sitting on the arc,
//'Usage   : Used to find the normalised value of a arbatury point from an arbatury arc
//'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


	double NormalisedAngle;	//'Angle to find ortho x,y also to check to see if inside arc
	double DistanceToOrtho;	//'Distance to ortho x,y
	double DistanceToXY;	//'Distance to the passed x,y coordinates
	double DistanceToEnd;
	double DistanceToStart;
	double EndX, EndY, StartX, StartY;
	double AdjNormAngle;

	NormalisedAngle = ::vec2double(X, Y).toVector().x;
	AdjNormAngle = NormalisedAngle - (PI/2);
	
	if(ArcEnd < ArcStart) {ArcEnd += (2 * PI); AdjNormAngle += (2 * PI);}
	if(AdjNormAngle < 0) AdjNormAngle+=(2*PI);

	//'360deg added to angle to make sure the arc doesn't pass through 0deg then
	//'check if between arc start and end. If so then return false and

	if((AdjNormAngle < ArcStart) || (AdjNormAngle > ArcEnd))
	{
		StartX = cos(ArcStart) * Radius;
		StartY = sin(ArcStart) * Radius * -1;
		EndX = cos(ArcEnd) * Radius;
		EndY = sin(ArcEnd) * Radius * -1;
		DistanceToStart = sqrt(((StartX - X) * (StartX - X)) + ((StartY - Y) *(StartY - Y)));
		DistanceToEnd = sqrt(((EndX - X) * (EndX - X)) + ((EndY - Y) * (EndY - Y)));
		if(DistanceToStart < DistanceToEnd)
		{
			OrthoX = StartX;
			OrthoY = StartY;
			OrthoDistance = DistanceToStart;
		}
		else
		{
			OrthoX = EndX;
			OrthoY = EndY;
			OrthoDistance = DistanceToEnd;
		}
	}    
	else

	{
		OrthoX = sin(NormalisedAngle) * Radius;
		OrthoY = cos(NormalisedAngle) * Radius;// * -1;
		DistanceToOrtho = sqrt((OrthoX * OrthoX) + (OrthoY * OrthoY));
		DistanceToXY = sqrt((X * X) + (Y * Y));
		OrthoDistance = DistanceToXY - Radius;
	}
	//'''''''''''''''''''''''''''''''''''''''''''''
    

	return true;
}

bool Shapes::ProfilerRefShapeDistCalcLine(vec2double Cursor,
										  double AX,
								  double AY,
								  double BX,
								  double BY,
								  double &OrthoX,
								  double &OrthoY,
								  double &OrthoDistance)


{


	

	double DistanceToPointA;
	double DistanceToPointB;
	double DistanceOrthoToZero;
	double DistanceCursorToZero;

	vec2double PointA(AX,AY);
	vec2double PointB(BX,BY);
	vec2double PointC = Cursor;
	vec2double PointD(AY-BY,AX-BX) ;
	vec2double Ortho;
	bool PointABIntersected;
	bool PointCDIntersected;

	PointD = PointD + Cursor;

	::Intersection().TwoLines(PointA,PointB,PointC,PointD,Ortho,PointABIntersected,PointCDIntersected);
	

	if(PointABIntersected) 
	{ 
		OrthoX=Ortho.x; 
		OrthoY=Ortho.y; 
		OrthoDistance = sqrt(((OrthoX-Cursor.x)*(OrthoX-Cursor.x))+((OrthoY-Cursor.y)*(OrthoY-Cursor.y)));
	}
	else
	{
		DistanceToPointA = sqrt(((PointA.x-Cursor.x)*(PointA.x-Cursor.x))+((PointA.y-Cursor.y)*(PointA.y-Cursor.y)));
		DistanceToPointB = sqrt(((PointB.x-Cursor.x)*(PointB.x-Cursor.x))+((PointB.y-Cursor.y)*(PointB.y-Cursor.y)));
		if(DistanceToPointA < DistanceToPointB)
		{
			OrthoX = PointA.x;
			OrthoY = PointA.y;
			OrthoDistance = DistanceToPointA;
		}
		else
		{
			OrthoX = PointB.x;
			OrthoY = PointB.y;
			OrthoDistance = DistanceToPointB;
		}

	}
	DistanceOrthoToZero = sqrt((OrthoX * OrthoX)+(OrthoY*OrthoY));
	DistanceCursorToZero = sqrt((Cursor.x*Cursor.x)+(Cursor.y*Cursor.y));
	if(DistanceOrthoToZero>DistanceCursorToZero) OrthoDistance*=-1;



	return true;
}