#pragma once


// added for SonarProfiler
//const int MAXDATASIZE = 60211200;

//0.9 Degree =	400 * 360 / ArcSize(in Degrees) steps
//1.8 Degree = 	200 * 360 / ArcSize(in Degrees) steps
//2.7 Degree = 	133 * 360 / ArcSize(in Degrees) steps	
//3.6 Degree = 	100 * 360 / ArcSize(in Degrees) steps

//const double PI = 3.14159265358979; 'ANT its now in algebra

enum EnumArcSize
{
	eum30Degrees = '0',		eum60Degrees = '1',
	eum90Degrees = '2',		eum120Degrees = '3',
	eum150Degrees = '4',	eum180Degrees = '5',
	eum210Degrees = '6',	eum240Degrees = '7',
	eum270Degrees = '8',	eum360Degrees = '9'
};

enum EnumCentreAngle
{
	eumCentre30 = '"',		eumCentre60 = '#',
	eumCentre90 = '$',		eumCentre120 = '%',
	eumCentre150 = '&',		eumCentre180 = '\'',
	eumCentre210 = '(',		eumCentre240 = ')',
	eumCentre270 = '*',		eumCentre300 = ',',
	eumCentre330 = '.'
};

enum EnumStepSize
{
	eum09Degree = ':',		eum18Degree = ';',
	eum27Degree = '<',		eum36Degree = '>'
};

enum eEncoder
{
	SetPearnormal = 0x00,
	SetQuadEncoder = 0x10,
	ReverseEncoder = 0x08,
	ResetEncoder = 0x04

};