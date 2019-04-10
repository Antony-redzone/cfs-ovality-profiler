#include "lasertracking.h"

LaserTracking::LaserTracking(void)
{
	int i;
	
	imVideo = NULL;
	imBuffer.image = NULL;
	imBuffer.height = 0;
	imBuffer.width = 0;
	trackingSize = 40; // Area to search
	trackingSize20 = 12;
	laserSize = 450; // Size of laser to find
	laserSizeMin = 20;
	intensity = 155; // Cutoff Intensity for laser dot
	
	// Initialise found shapes to nadda, zero, zilch etc.
	for(i=0;i<100;i++) 
	{
		foundShapes[i].averageBrightness=0;
		foundShapes[i].averageX=0;
		foundShapes[i].averageY=0;
		foundShapes[i].down=0;
		foundShapes[i].left=0;
		foundShapes[i].right=0;
		foundShapes[i].up=0;
		foundShapes[i].size=0;
	}
	laserOneCoord.x=0; laserOneCoord.y=0;
	laserTwoCoord.x=0; laserTwoCoord.y=0;
	laserCentreCoord.x=0; laserCentreCoord.y=0;
	laserLeftSideCoord=0;
	laserRightSideCoord=0;
	
}

LaserTracking::~LaserTracking(void)
{
	int i;
	if(imBuffer.image!=NULL) 
	{
		for(i=0;i<imHeight;i++) delete[] imBuffer.image[i];
		delete[] imBuffer.image;
		imBuffer.image=NULL;
	}

	if(imLaserOne.image!=0)
	{
		for(i=0;i<imLaserOne.height;i++) delete[] imLaserOne.image[i];
		delete[] imLaserOne.image;
		imLaserOne.image=0;
	}

	if(imLaserTwo.image!=0)
	{
		for(i=0;i<imLaserTwo.height;i++) delete[] imLaserTwo.image[i];
		delete[] imLaserTwo.image;
		imLaserTwo.image=0;
	}

	if(imLaserCentre.image!=0)
	{
		for(i=0;i<imLaserCentre.height;i++) delete[] imLaserCentre.image[i];
		delete[] imLaserCentre.image;
		imLaserCentre.image=0;
	}
	
	if(imLaserLeftSide.image!=0)
	{
		for(i=0;i<imLaserLeftSide.height;i++) delete[] imLaserLeftSide.image[i];
		delete[] imLaserLeftSide.image;
		imLaserLeftSide.image=0;
	}
	
	if(imLaserRightSide.image!=0)
	{
		for(i=0;i<imLaserRightSide.height;i++) delete[] imLaserRightSide.image[i];
		delete[] imLaserRightSide.image;
		imLaserRightSide.image=0;
	}

}

void LaserTracking::SetVideoPointer(pixel **im,int h,int w)
{
	int i;
//	int x,y;

	imVideo = im; // Store video pointer
	imHeight = h; // Store video height
	imWidth = w;  // Store video width
	
	// Create the imProcessing buffer ////////
	imBuffer.image = new pixel*[imHeight];	//
	for(i=0;i<imHeight;i++) imBuffer.image[i] = new pixel[imWidth];

	// Create target buffers for One and Two target lasers
	imLaserOne.image = new pixel*[trackingSize];
	imLaserTwo.image = new pixel*[trackingSize];
	imLaserCentre.image =  new pixel*[trackingSize];
	imLaserLeftSide.image = new pixel*[trackingSize];
	imLaserRightSide.image = new pixel*[trackingSize];

	for(i=0;i<trackingSize;i++) 
	{
		imLaserOne.image[i] = new pixel[trackingSize];
		imLaserTwo.image[i] = new pixel[trackingSize];
		imLaserCentre.image[i] = new pixel[trackingSize];
		imLaserLeftSide.image[i] = new pixel[trackingSize];
		imLaserRightSide.image[i] = new pixel[trackingSize];
	}
	
	//////////////////////////////////////////
	// And copy over the video image
//	for(y=0;y<imHeight;y++)
//		for(x=0;x<imWidth;x++)
//		{
//			imBuffer.image[y][x] = imVideo[y][x];
//		}
	
	imBuffer.width = imWidth;
	imBuffer.height = imHeight;
	imLaserOne.width = imLaserOne.height = trackingSize;
	imLaserTwo.width = imLaserTwo.height = trackingSize;
	imLaserCentre.width = imLaserCentre.height = trackingSize20;
	imLaserLeftSide.width = imLaserLeftSide.height = trackingSize;
	imLaserRightSide.width = imLaserRightSide.height = trackingSize;

	// Set the scan mask for both laser targets too centre of screen
	laserOneCoord.x=imWidth/2; laserOneCoord.y=imHeight/2;
	laserTwoCoord.x=imWidth/2; laserTwoCoord.y=imHeight/2;
	laserCentreCoord.x=imWidth/2; laserCentreCoord.y=imHeight/2;
	laserLeftSideCoord.x=imWidth/2; laserLeftSideCoord.y=imHeight/2;
	laserRightSideCoord.x=imWidth/2; laserRightSideCoord.y=imHeight/2;
}

void LaserTracking::SearchForLaser(void)
{
	int shapeCount=0;
	int laserOneCount=0;
	int laserTwoCount=0;
	int laserCentreCount=0;
	int laserLeftSideCount=0;
	int laserRightSideCount=0;
	int i;
	int x,y;
	if(imBuffer.image == NULL) return;
	if(imLaserOne.image == NULL) return;
	if(imLaserTwo.image == NULL) return;
	if(imLaserCentre.image == NULL) return;
	if(imLaserLeftSide.image == NULL) return;
	if(imLaserRightSide.image == NULL) return;

//	// Refresh the Video Buffer
//	for(y=0;y<imHeight;y++)
//		for(x=0;x<imWidth;x++)
//		{
//			imBuffer.image[y][x] = imVideo[y][x];
//		}

	CopyImageToBuffer(imLaserOne, laserOneCoord);
	CopyImageToBuffer(imLaserTwo, laserTwoCoord);
	CopyImageToBuffer(imLaserCentre, laserCentreCoord);
	CopyImageToBuffer(imLaserLeftSide, laserLeftSideCoord);
	CopyImageToBuffer(imLaserRightSide, laserRightSideCoord);

	// Reset the shapes data
	for(i=0;i<100;i++) 
	{
		foundShapes[i].averageBrightness=0;	foundLaserOne[i].averageBrightness=0;	foundLaserTwo[i].averageBrightness=0;	foundLaserCentre[i].averageBrightness=0;
		foundShapes[i].averageX=0;			foundLaserOne[i].averageX=0;			foundLaserTwo[i].averageX=0;			foundLaserCentre[i].averageX=0;
		foundShapes[i].averageY=0;			foundLaserOne[i].averageY=0;			foundLaserTwo[i].averageY=0;			foundLaserCentre[i].averageY=0;
		foundShapes[i].down=0;				foundLaserOne[i].down=0;				foundLaserTwo[i].down=0;				foundLaserCentre[i].down=0;
		foundShapes[i].left=0;				foundLaserOne[i].left=0;				foundLaserTwo[i].left=0;				foundLaserCentre[i].left=0;
		foundShapes[i].right=0;				foundLaserOne[i].right=0;				foundLaserTwo[i].right=0;				foundLaserCentre[i].right=0;
		foundShapes[i].up=0;				foundLaserOne[i].up=0;					foundLaserTwo[i].up=0;					foundLaserCentre[i].up=0;
		foundShapes[i].size=0;				foundLaserOne[i].size=0;				foundLaserTwo[i].size=0;				foundLaserCentre[i].size=0;


		foundLaserLeftSide[i].averageBrightness=0;	foundLaserRightSide[i].averageBrightness=0;
		foundLaserLeftSide[i].averageX=0;			foundLaserRightSide[i].averageX=0;
		foundLaserLeftSide[i].averageY=0;			foundLaserRightSide[i].averageY=0;
		foundLaserLeftSide[i].down=0;				foundLaserRightSide[i].down=0;
		foundLaserLeftSide[i].left=0;				foundLaserRightSide[i].left=0;
		foundLaserLeftSide[i].right=0;				foundLaserRightSide[i].right=0;
		foundLaserLeftSide[i].up=0;					foundLaserRightSide[i].up=0;
		foundLaserLeftSide[i].size=0;				foundLaserRightSide[i].size=0;
	}

	// Scan thru the video buffer
/*
	for(x=0;x<imWidth;x++)
		for(y=0;y<imHeight;y++)
		{
			if(FindNextShape(imBuffer, x, y, intensity, foundShapes[shapeCount], 0)) // Find the next shape
			{
				if((shapeCount<99) && // No more that 100 shapes can be stored
				   (foundShapes[shapeCount].size<laserSize) && // Is it small enough
				   (foundShapes[shapeCount].size>laserSizeMin) &&
				   (foundShapes[shapeCount].averageBrightness > intensity) && // Is it bright enough
				   (abs(foundShapes[shapeCount].width-foundShapes[shapeCount].height)<4)) 
				{
					shapeCount++; // Move onto the next shape
				}
				else
				{
					foundShapes[shapeCount].averageBrightness=0;
					foundShapes[shapeCount].averageX=0;
					foundShapes[shapeCount].averageY=0;
					foundShapes[shapeCount].down=0;
					foundShapes[shapeCount].left=0;
					foundShapes[shapeCount].right=0;
					foundShapes[shapeCount].up=0;
					foundShapes[shapeCount].size=0;
				}

			}
		}
//	for(i=0;i<shapeCount;i++) MarkShapeCentre(foundShapes[i]);
*/
	// Scan thru the video buffer
	for(x=imLaserOne.width-1;x>=0;x--)
		for(y=0;y<imLaserOne.height;y++)
		{
			if(FindNextShape(imLaserOne, x, y, intensity, foundLaserOne[laserOneCount], 0)) // Find the next shape
			{
				if(laserOneCount<99) laserOneCount++; // Move onto the next shape
			}
		}

	for(x=0;x<imLaserTwo.width;x++)
		for(y=0;y<imLaserTwo.height;y++)
		{
			if(FindNextShape(imLaserTwo, x, y, intensity, foundLaserTwo[laserTwoCount], 0)) // Find the next shape
			{
				if(laserTwoCount<99) laserTwoCount++; // Move onto the next shape
			}
		}
	
	for(x=0;x<imLaserCentre.width;x++)
		for(y=0;y<imLaserCentre.height;y++)
		{
			if(FindNextShape(imLaserCentre, x, y, intensity, foundLaserCentre[laserCentreCount], 0)) // Find the next shape
			{
				if(laserCentreCount<99) laserCentreCount++; // Move onto the next shape
			}
		}
	for(x=0;x<imLaserLeftSide.width;x++)
		for(y=0;y<imLaserLeftSide.height;y++)
		{
			if(FindNextShape(imLaserLeftSide, x, y, intensity/1.3, foundLaserLeftSide[laserLeftSideCount], 0)) // Find the next shape
			{
				if(laserLeftSideCount<99) laserLeftSideCount++; // Move onto the next shape
			}
		}

	for(x=imLaserRightSide.width-1;x>=0;x--)
		for(y=0;y<imLaserRightSide.height;y++)
		{
			if(FindNextShape(imLaserRightSide, x, y, intensity/2, foundLaserRightSide[laserRightSideCount], 0)) // Find the next shape
			{
				if(laserRightSideCount<99) laserRightSideCount++; // Move onto the next shape
			}
		}



	if(laserOneCount>0) 
	{
		laserOneCoord.x=foundLaserOne[0].averageX+laserOneCoord.x;
		laserOneCoord.y=foundLaserOne[0].averageY+laserOneCoord.y;
		laserOneCoord=laserOneCoord-(double) (trackingSize/2);
	}

	if(laserTwoCount>0)
	{
		laserTwoCoord.x=foundLaserTwo[0].averageX+laserTwoCoord.x;
		laserTwoCoord.y=foundLaserTwo[0].averageY+laserTwoCoord.y;
		laserTwoCoord=laserTwoCoord-(double) (trackingSize/2);
	}

	int lowestDown=0;
	int foundLaserCentreLowest=0;
	int highestUp=0;
	int foundLaserCentreHighest = 0;
	


	if(laserCentreCount>0)
	{

//		laserCentreCoord.x=foundLaserCentre[0].averageX+laserCentreCoord.x;
//		laserCentreCoord.y=foundLaserCentre[0].averageY+laserCentreCoord.y;
//		laserCentreCoord=laserCentreCoord-(double) (trackingSize/2);

		//lowestDown = foundLaserCentre[0].down;
		//for(i=1;i<laserCentreCount;i++)
		//{
		//	if(foundLaserCentre[i].down<lowestDown)
		//	{
		//		lowestDown=foundLaserCentre[i].down;
		//		foundLaserCentreLowest=i;
		//	}
		//}

		highestUp = foundLaserCentre[0].up;
		for(i=1;i<laserCentreCount;i++)
		{
			if(foundLaserCentre[i].down>highestUp)
			{
				highestUp=foundLaserCentre[i].up;
				foundLaserCentreHighest=i;
			}
		}

	
		laserCentreCoord.x=foundLaserCentre[foundLaserCentreHighest].averageX+laserCentreCoord.x;
		laserCentreCoord.y=foundLaserCentre[foundLaserCentreHighest].up+laserCentreCoord.y;
		laserCentreCoord=laserCentreCoord-(double) (trackingSize20/2);
	}

	if(laserLeftSideCount>0) 
	{
		laserLeftSideCoord.x=foundLaserLeftSide[0].averageX+laserLeftSideCoord.x;
//		laserLeftSideCoord.y=foundLaserLeftSide[0].averageY+laserLeftSideCoord.y;
		laserLeftSideCoord.x=laserLeftSideCoord.x-(double) (trackingSize/2);
	}

	if(laserRightSideCount>0)
	{
		laserRightSideCoord.x=foundLaserRightSide[0].averageX+laserRightSideCoord.x;
//		laserRightSideCoord.y=foundLaserRightSide[0].averageY+laserRightSideCoord.y;
		laserRightSideCoord.x=laserRightSideCoord.x-(double) (trackingSize/2);
	}


//	MarkTarget(laserOneCoord, (char) 255,(char) 0,(char) 255);
//	MarkTarget(laserTwoCoord, (char) 255,(char) 255,(char) 0);
	
}

int LaserTracking::FindNextShape(ImageBuffer &imBuffer, int x, int y, double bright, Shape &currentShape,int trackParent)
{
	double intensity;
	pixel *poi; // Pixel Of Interest
	
	if((x<0) || (y<0) || (x>=imBuffer.width) || (y>=imBuffer.height)) return false; //Out of bounds
	if(trackParent>20) return false; // Stop overflow if lots pixels highlited

	
	poi = &imBuffer.image[y][x]; // Addres the Pixel of Interest 
	intensity = (poi->blue+poi->green+poi->red)/3; // Get the average r,g,b pixel value

	if(intensity>bright) 
	{
		if(trackParent==0)
		{
			__asm nop
		};
		currentShape.averageBrightness+=intensity;
		currentShape.averageX+=x;
		currentShape.averageY+=y;
		if(trackParent==0)
		{
			currentShape.left=currentShape.right=x;
			currentShape.up=currentShape.down=y;
		}
		else{
			if(x<currentShape.left) currentShape.left=x;
			if(x>currentShape.right) currentShape.right=x;
			if(y>currentShape.up) currentShape.up=y;
			if(y<currentShape.down) currentShape.down=y;
		}
		currentShape.size++;
		poi->blue=poi->green=poi->red=0;
//		imVideo[y][x].green=255;
//		imVideo[y][x].blue=0;
//		imVideo[y][x].red=0;
	}
	else return false;

	FindNextShape(imBuffer ,x+1 ,y+1 ,bright,currentShape,trackParent+1);
	FindNextShape(imBuffer ,x+1 ,y   ,bright,currentShape,trackParent+1);
	FindNextShape(imBuffer ,x+1 ,y-1 ,bright,currentShape,trackParent+1);
	FindNextShape(imBuffer ,x-1 ,y+1 ,bright,currentShape,trackParent+1);
	FindNextShape(imBuffer ,x-1 ,y   ,bright,currentShape,trackParent+1);
	FindNextShape(imBuffer ,x-1 ,y-1 ,bright,currentShape,trackParent+1);
	FindNextShape(imBuffer ,x   ,y+1 ,bright,currentShape,trackParent+1);
	FindNextShape(imBuffer ,x   ,y-1 ,bright,currentShape,trackParent+1);

	if(trackParent==0)
	{
		currentShape.averageBrightness/=currentShape.size;
		currentShape.averageX/=currentShape.size;
		currentShape.averageY/=currentShape.size;
		currentShape.width=currentShape.right-currentShape.left;
		currentShape.height=currentShape.up-currentShape.down;
	}
	return true;

	
}

void LaserTracking::MarkShapeCentre(Shape currentShape)
{
	int x,y;
	int averageX;
	int averageY;

	averageX=(int) (currentShape.averageX+0.5);
	averageY=(int) (currentShape.averageY+0.5);
	y=averageY;
	for(x=averageX-5;x<averageX+5;x++)
	{
		if((y<0) || (y>=imHeight) || (x<0) || (x>=imWidth)) continue;
		imVideo[y][x].blue=255;
		imVideo[y][x].green=0;
		imVideo[y][x].red=0;
	}
	x=averageX;
	for(y=averageY-5;y<averageY+5;y++)
	{
		if((y<0) || (y>=imHeight) || (x<0) || (x>=imWidth)) continue;
		imVideo[y][x].blue=255;
		imVideo[y][x].green=0;
		imVideo[y][x].red=0;
	}
}

void LaserTracking::MarkTarget(vec2double coord, char red, char green, char blue)
{
	int x,y;
	int centX;
	int centY;

	centX=(int) (coord.x+0.5);
	centY=(int) (coord.y+0.5);
	
	for(x=centX-5,y=centY-5;x<centX+5;x++,y++)
	{
		if((y<0) || (y>=imHeight) || (x<0) || (x>=imWidth)) continue;
		imVideo[y][x].blue=blue;
		imVideo[y][x].green=green;
		imVideo[y][x].red=red;
	}

	for(y=centY-5,x=centX+5;y<centY+5;y++,x--)
	{
		if((y<0) || (y>=imHeight) || (x<0) || (x>=imWidth)) continue;
		imVideo[y][x].blue=blue;
		imVideo[y][x].green=green;
		imVideo[y][x].red=red;
	}
}

void LaserTracking::CopyImageToBuffer(ImageBuffer &imBuff, vec2double centreCoord)
{
	int xImVideo,yImVideo, xImBuffer,yImBuffer;
	int centreX, centreY;
	
	centreX = (int) centreCoord.x;
	centreY = (int) centreCoord.y;
	
	for(yImBuffer=0;yImBuffer<imBuff.height;yImBuffer++)
	{
		yImVideo = yImBuffer+centreY - (imBuff.height/2); // y coord of image coppied from is offset by half the size of the area to be copied
		for(xImBuffer=0;xImBuffer<imBuff.width;xImBuffer++)  
		{	
			xImVideo = xImBuffer+centreX - (imBuff.width/2); // x coord of image coppied from is offse by half the size of the area to be copied	
			if((xImVideo<0) || (xImVideo>=imWidth) || (yImVideo<1) || (yImVideo>=imHeight)) 
			{
				imBuff.image[yImBuffer][xImBuffer].blue=0;	// If the video pixel to copy is
				imBuff.image[yImBuffer][xImBuffer].green=0; // out of bounds then make the
				imBuff.image[yImBuffer][xImBuffer].red=0;	// imBuff colours all set to zero
			}
			else {// video pixel to be copied is not out of bounds then copy over the pixel colours
//				imBuff.image[yImBuffer][xImBuffer].blue = 255-imVideo[yImVideo][xImVideo].blue;
//				imBuff.image[yImBuffer][xImBuffer].green= 255-imVideo[yImVideo][xImVideo].green;
//				imBuff.image[yImBuffer][xImBuffer].red  = 255-imVideo[yImVideo][xImVideo].red;
//				imVideo[yImVideo][xImVideo].red=160;
				imBuff.image[yImBuffer][xImBuffer].blue = imVideo[yImVideo][xImVideo].blue;
				imBuff.image[yImBuffer][xImBuffer].green= imVideo[yImVideo][xImVideo].green;
				imBuff.image[yImBuffer][xImBuffer].red  = imVideo[yImVideo][xImVideo].red;

				
			}
		}
	}
}

void LaserTracking::DisplaySearchBox(int red, int green, int blue)
{
	int x, y;

	vec2double displayCoord;

	for(x=(int) laserCentreCoord.x-(trackingSize/2);x<laserCentreCoord.x+(trackingSize/2);x++)
		for(y=(int) (laserCentreCoord.y-(trackingSize/2));y<laserCentreCoord.y+(trackingSize/2);y++)
		{
			if((x<0) || (x>=imWidth) || (y<0) || (y>=imHeight)) continue;
			if(red!=-1)   imVideo[y][x].red = (unsigned char) red;
			if(green!=-1) imVideo[y][x].green = (unsigned char) green;
			if(blue!=-1)  imVideo[y][x].blue = (unsigned char) blue;
		}
	DrawLine(laserCentreCoord-(trackingSize/2),laserCentreCoord+(trackingSize/2),255,0,0);
	DrawLine(vec2double(laserCentreCoord.x-(trackingSize/2),laserCentreCoord.y+(trackingSize/2)),
		     vec2double(laserCentreCoord.x+(trackingSize/2),laserCentreCoord.y-(trackingSize/2)),255,0,0);

}

void LaserTracking::DrawLine(vec2double coord_a, vec2double coord_b, unsigned char red, unsigned char green, unsigned char blue)
{
	int x,y;
	double a,b;
	double x_size, y_size;
	double largest;
	double step_x;
	double step_y;
	double i;

	x_size=coord_b.x-coord_a.x;
	y_size=coord_b.y-coord_a.y;
	if(fabs(x_size)>fabs(y_size)) largest=fabs(x_size);
	else                          largest=fabs(y_size);

	step_x=x_size/largest;
	step_y=y_size/largest;

	a=coord_a.x;
	b=coord_a.y;

	for(i=0;i<largest;i++)
	{
		x = (int) (a+0.5);
		y = (int) (b+0.5); 
		if((x>=imWidth) || (x<0) || (y>=imHeight) || (y<0)) //ANT VOB
		{
			a=step_x;
			b=step_y;
			continue;
		}
		imVideo[y][x].red = red;
		imVideo[y][x].green = green;
		imVideo[y][x].blue = blue;
		a+=step_x;
		b+=step_y;
	}
}