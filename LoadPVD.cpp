#include "LoadPVD.h"
#include "..\houghlibv2.0\CBSAlgebra.h"
#include <stdio.h>

LoadPVD::LoadPVD(char *_pvFileName,
				int _pvDataStartAddress,
				int _pvDataBlockSize,
				int _xy,
				float *_pvDataX,
                float *_pvDataY,
                double _pvDataMultiplier,
                int _fromFrame,
                int _toFrame)
{
	pvFileName			= _pvFileName;
	pvDataStartAddress	= _pvDataStartAddress;
	pvDataBlockSize		= _pvDataBlockSize,
	pvDataX				= _pvDataX;
    pvDataY				= _pvDataY;
    pvDataMultiplier	= _pvDataMultiplier;
    fromFrame			= _fromFrame;
    toFrame				= _toFrame;
	xy_data = _xy;
}

LoadPVD::~LoadPVD(void)
{
}

void LoadPVD::LoadPVDData(void)
{
	FILE *f;

	short int coordX=0,coordY=0;
	float fCoordX=0, fCoordY=0;
	long i; //Frame
	long p; //profile Point
	long c;
	int dataSize;

	fpos_t a;

	vec2double vector;

	i=0;
	c=0;

	f = fopen(pvFileName,"rb");
	
	//fseek(f,0,SEEK_CUR);
	//fseek( f, 0L, SEEK_SET );

	if(xy_data>1) dataSize=4;
	else dataSize=2;

	--pvDataStartAddress;
	for(i=0;i<toFrame;i++)
	{
		if(i==4303)
		{
			__asm nop;
		}
		//fseek(f,(long) (pvDataStartAddress + (i*pvDataBlockSize)),SEEK_CUR);
		a = (long) (pvDataStartAddress + (i*pvDataBlockSize));
		fsetpos( f,  &a);
		if(xy_data==1)
		{
			for(p=0;p<180;p++)
			{
				fread(&coordX,dataSize,1,f);
				fread(&coordY,dataSize,1,f);
				pvDataX[c]=(float) ((double) coordX*pvDataMultiplier);
				pvDataY[c]=(float) ((double) coordY*pvDataMultiplier);
				if((pvDataX[c] > 10000) || (pvDataX[c] < -10000) || (pvDataY[c] > 10000) || (pvDataY[c] < -10000))
					{
						pvDataX[c]=0; pvDataY[c]=0;
					}
				c++;
			}
		}
		else if(xy_data==2)
		{
			for(p=0;p<180;p++)
			{
				fread( &fCoordX,dataSize,1,f);
				fread( &fCoordY,dataSize,1,f);
				if((fCoordX > 10000) || (fCoordX < -10000) || (fCoordY > 10000) || (fCoordY < -10000))
					{
						fCoordX=(float) 0; fCoordY=(float) 0;
					}
				pvDataX[c]=(float) ((double) fCoordX*pvDataMultiplier);
				pvDataY[c]=(float) ((double) fCoordY*pvDataMultiplier);
//				if((pvDataX[c] > 10000) || (pvDataX[c] < -10000) || (pvDataY[c] > 10000) || (pvDataY[c] < -10000))
//					{
//						pvDataX[c]=(float) 0; pvDataY[c]=(float) 0;
//					}
				c++;
			}
		}

		else
		{
			for(p=0;p<180;p++)
			{
				fread(&coordY,2,1,f);
				if(coordY == 0)
				{
					pvDataX[c] = 0;
					pvDataY[c] = 0;
				}
				else
				{
					vector.x = ((p+91)*2)*PI / 180;
					vector.y = coordY;
					vector = vector.toCoordinate();
					pvDataX[c] =(float) (vector.x * (double) pvDataMultiplier);
					pvDataY[c] =(float) (vector.y * (double) pvDataMultiplier);
				}
				if((pvDataX[c] > 10000) || (pvDataX[c] < -10000) || (pvDataY[c] > 10000) || (pvDataY[c] < -10000))
					{
						pvDataX[c]=(float) 0; pvDataY[c]=(float) 0;
					}
				c++;
			}


		}
		
		if((pvDataX[c] > 10000) || (pvDataX[c] < -10000) || (pvDataY[c] > 10000) || (pvDataY[c] < -10000))
		{
			pvDataX[c]=(float) 0; pvDataY[c]=(float) 0;
		}
	}

	
	fclose(f);
}