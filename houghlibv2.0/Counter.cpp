#include <math.h>

#include <stdio.h>
#include <atlbase.h>

#include "Video.h"
#include "Counter.h"

#define BUFFER_SIZE 20
#define TRACE_SIZE 100
#define POINTS_SIZE 599

void MsgCounter(TCHAR *szFormat, ...)
	{
    TCHAR szBuffer[512];

    va_list pArgs;
    va_start(pArgs, szFormat);
    _vstprintf(szBuffer, szFormat, pArgs);
    va_end(pArgs);

    MessageBox(NULL, szBuffer, TEXT("LaserProfiler Message"), MB_OK | MB_ICONERROR);
	}

Counter::Counter(void)
{
	// Set im video pointer and im width & height //////////
	imVideo=NULL;// originalIm;   // Original image data from video
	imWidth=0; imHeight=0; // width & height of video

	imMask = NULL;	// Copy of image mask, PCN2639 (5 April 2004, Antony van Iersl);
	chBuff=NULL;	// buffer previous image
	threshold=40;	// 60 is good How much of change before it notices as a counter change.
					// 40 for water
	bufferValue=7;  // normally 7
	totalAverageChange=0;
	count=0;
	neg=1;
	head=0;
	tail=0;
	glitch=0;
	direction=1;
	isSet=false;
	maskWidth=0; maskHeight=0;
	contrast=77; // Normally 47
	neighContr=30;
	isDecimalFound=false;
	countDecimals=0;
	textColour=0;
	sxHistory=0;
	units = METRIC; // PCN2874 0 for meters, 1 for feet.
}

///////////////////////////////////////////////////////////////////////////////////
// PCN2639
// Name: SetCounterPointer
// Input: 2 dimensional pointer type pixel, width of video, height of video
// Output: None
//
// Description: Sets the local 2 dimensional pixel pointer to the video image,
//				sets the width and height the same as the video image width and hight.
///////////////////////////////////////////////////////////////////////////////////
void Counter::SetCounterPointer(pixel **originalIm, int w, int h)
{
	// Set im video pointer and im width & height //////////
	imVideo= originalIm;   // Original image data from video
	imWidth=w; imHeight=h; // width & height of video

	//PCN2639 Adjust Mask if not inside the new video parameters (10 June 2004, Antony van Iersel)
	if(xRight>=imWidth) xRight=imWidth -1;		// //ANT VOB
	if(xLeft>=imWidth) xLeft=imWidth-1;		////ANT VOB
	if(yTop>=imHeight) yTop=imHeight-1;		////ANT VOB
	if(yLower>=imHeight) yLower=imHeight-1;	////ANT VOB
	//////////////////////////////////////////
	
	isSet=true;
}

Counter::~Counter()
{
	int i;
	isSet=false;
//	delete [] chBuff;
//	delete [] imMask;

	//PCN3085 trying to make the profiler more stable by removing some possible memory leaks
	// the folling code replaces the commented out lines above.
	for(i=0;i<BUFFER_SIZE;i++) delete[] chBuff[i];
	delete[] chBuff;

	for(i=0;i<maskHeight;i++) delete[] imMask[i];
	delete[] imMask;
	
}

///////////////////////////////////////////////////////////////////////////////////
// PCN2639
// Name: Tick
// Input: none
// Output: none
//
// Description: If there is enough changes between tail and head, glitch increased by 1.
//				If glitch reaches a set number of changes, then the counter ticks over.
//				Up or down, depending on the direction. The number of glitches it has
//				to count up to is the lag in the counter, default is 7. Head is current
//				counter sample, tail is sample it checks for when the counter changes.
///////////////////////////////////////////////////////////////////////////////////

void Counter::Tick(void)
{

	int d;	// number of changes in the counter mask returned from sample

	if(!isSet) return; // if mask is not set then don't process

///////////////// The Following is needed only if testing mask for every frame
//	sxLeft=xLeft;	sxRight=xRight;
//	syTop=yTop;		syLower=yLower;
//	ResetMaskOverSingleDigit();
//////////////////////////////////////////////////////////////////////////////

	
	if(countDecimals<11) ResampleDecimalPoint();
	d=Sample();
//	if(countDecimals<11) ResampleDecimalPoint();

	// if number of changes 'd' are above a certain value then take notice, it might be a counter tick
	if(d>300) //300 to 400 is good,   150 for beter for water (Background of couter changes)    
		{
		glitch++; // Keep track of constant change. If its only a glitch in video, This will reset back to 0
		tail--;	// Make sure the sample is compared before the timer tick over.
		if(glitch==bufferValue) // was 7 , If there is a constant change in counter frame
			{		  //         then its a real change and not a glitch
			if(direction==0) count--; // Count down if direction 0
			else count++;			  // Count up if direction 1	
			glitch=0;		// Reset glitch tracking to 0
			tail=head-1;	// Start comparing changes one down from head
			}
		}
	else // If there is not enough changes then don't tick.
		{
		if(glitch!=0) tail=head-1;  // If there was a previous glitch tick then reset tail just behind head
		glitch=0; // Reset glitch tracking to 0
		} 
	if(tail==-1) tail=(BUFFER_SIZE-1); // If tail needs to loop to the back of buffer.

// Only for displaying purposes //////////////////
//	DrawSquare(sxLeft,sxRight);					//
//	CopyMaskEdgeToImage();						//
//	DrawCircle(decimal_x+xLeft,decimal_y+yTop);	//
//////////////////////////////////////////////////
}

///////////////////////////////////////////////////////////////////////////////////
// PCN2639
// Name: SetCounterMask
// Input: boundries of the Counter Mask, xL LEFT, xR RIGHT, yT Top, yL Bottom
// Output: none:
//
// Description: Used to set the mask where the counter class looks, to
//				a: The decimal point
//				b: Find what decimal to track for a change to incriment the counter
//				   which will be the first number on the right of the decimal point
//				Allocate the Frame Buffer to the size of the Counter Mask
///////////////////////////////////////////////////////////////////////////////////

void Counter::SetCounterMask(int xL, int xR, int yT, int yL)
{
	int i;

	if(xL<xR) {xLeft=xL;	xRight=xR;} //If Left Of Mask is < Right of Mask xLeft=xL etc
	else      {xLeft=xR;	xRight=xL;} //else reverse

	if(yT<yL) {yTop=yT;		yLower=yL;} //Same as x, this is to make sure the left is left
	else	  {yTop=yL;		yLower=yT;} //of right, and top is above bottom;

	bufferSize=(xRight-xLeft)*(yLower-yTop); //Mask size in number of pixels
	
	chBuff = new int *[BUFFER_SIZE]; //BUFFER_SIZE number of counter frames it can store to analyse
	for(i=0;i<BUFFER_SIZE;i++) chBuff[i] = new int[bufferSize+100]; //+100 just incase;

	//PCN2639 (5 April 2004, Antony van Iersl) ///////////////
	maskHeight = yLower-yTop; 
	maskWidth = xRight-xLeft;
	imMask = new edge *[maskHeight];
	for(i=0;i<maskHeight;i++) imMask[i] = new edge[maskWidth];

	sxLeft=xLeft;	sxRight=xRight;
	syTop=yTop;		syLower=yLower;
};

void Counter::ResampleDecimalPoint(void)
{
	int i;
	int swaped;
	history temp;
	
	if(!isSet) return; //PCN2639 (17 May 2004, Antony) Avoid accesing Mask when not allocated.

	// Reset sample mask size to counter mask size to rescan mask.
	sxLeft=xLeft;	sxRight=xRight;
	syTop=yTop;		syLower=yLower;
	ResetMaskOverSingleDigit();

	foundDecimals[countDecimals].x=decimal_x; 
	foundDecimals[countDecimals].y=decimal_y;
	foundDecimals[countDecimals].nextDecimal=sxLeft;

	//sxLeft=traceMask[t].i_min+xLeft; sxRight=traceMask[t].i_min+xLeft+5;

	countDecimals++;
	
	swaped=true;
	while(swaped)
		{
		swaped=false;
		for(i=0;i<countDecimals-1;i++)
			if(foundDecimals[i].nextDecimal>foundDecimals[i+1].nextDecimal)
				{
				temp       = foundDecimals[i];
				foundDecimals[i]   = foundDecimals[i+1];
				foundDecimals[i+1]   = temp;
				swaped=true;
				}
		}
	
	decimal_x=foundDecimals[countDecimals/2].x;
	decimal_y=foundDecimals[countDecimals/2].y;
	sxLeft=foundDecimals[countDecimals/2].nextDecimal;
	sxRight=sxLeft+5;
	sxHistory=sxLeft;
}

///////////////////////////////////////////////////////////////////////////////////
// PCN2639 28 April 2004 (Antony van Iersel)
// Name: Sample
// Input: none
// Output: none
//
// Description: Also finds the Decimal point and with this,
//				repositions the Mask given by the user to fit only on the Digit
//				that is to incriment the counter.
///////////////////////////////////////////////////////////////////////////////////

void Counter::ResetMaskOverSingleDigit(void)
{
	if(!isSet) return;
	CopyImageToMaskEdge();	// Get a clean mask image
	TraceMask();
	contrast=(int) ((double) GetAverageEdge()); // PCN2972 was 0.4
	ExtractTracedObjects(); 
	FindHole(decimal_x,decimal_y);
//PCN2972	FindDecimalPoint(decimal_x,decimal_y);
	FindDigitNextToDecimal();
}

///////////////////////////////////////////////////////////////////////////////////
// PCN2639
// Name: Sample
// Input: none
// Output: Returns the amount of change between head and tail
//
// Description: Takes the number of pixels that have changed between the current mask image
//				and the tail image, but only if its greater than a set threshold.
//				The number of changes are added and returned. 
///////////////////////////////////////////////////////////////////////////////////

int Counter::Sample(void)
{

	int change, totalChange = 0;
	
	if(imVideo==NULL) return 0;
	if(chBuff==NULL) return 0;
	
	int x,y,avg;
	int buffIndex=0;
	for(x=sxLeft;x<sxRight;x++)
		for(y=syTop;y<syLower;y++)
			{
			if((x>=imWidth) || (x<0)) continue; //Out of bounds checking; //ANT VOB
			if((y>=imHeight)|| (y<0)) continue; //Out of bounds checking;//ANT VOB
	
			avg=(imVideo[y][x].red+imVideo[y][x].green+imVideo[y][x].blue)/3; // Image prosses in gray
			if(tail!=head) // Buffer empty dont check for changes , normal first run.
				{
				change=abs(chBuff[tail][buffIndex]-avg);  // Get the pixel difference between head and tail
				if(change>threshold) totalChange+=change; // if the change > threshold incriment total change
				}
			chBuff[head][buffIndex++]=avg; // Add difference to head of buffer
			}

	if(tail!=head) tail=(tail+1)%BUFFER_SIZE; // Loop tail if need
	head = (head+1)%BUFFER_SIZE; // Loop head if needed

	return totalChange; // Return the number of changes between current image and tail buffer.
}


///////////////////////////////////////////////////////////////////////////////////
// PCN2639
// Name: CopyImageToMaskEdge
// Input: none
// Output: none
//
// Description: Takes a copy from the original video frame bounded my the counter mask
//				and copies it into the the imMask, working copy of the counter.
//				It only holds a black and white image of the mask, so the pixels have
//				to be averaged out.
//				The edges in the mask ask also initialised to 0.
///////////////////////////////////////////////////////////////////////////////////

void Counter::CopyImageToMaskEdge(void)
{
	if(!isSet) return; //PCN2639 (17 May 2004, Antony) kept copying mask even thou it was not allocated.
	int x,y; // Video image index, offset by the counters mask (Left, Right, Top, Lower)
	int i=0; // Mask index, i, j. (same variable name used thu out class, (I Hope),
	int j=0; // x & y normaly represent video image. i & j represent mask image.
	

	for(x=xLeft;x<xRight;x++)
		{
		j=0;
		for(y=yTop;y<yLower;y++)
			{
			imMask[j][i].avgPixel=(imVideo[y][x].red+imVideo[y][x].green+imVideo[y][x].blue)/3;
			imMask[j][i].edgeDiag=0;
			imMask[j][i].edgeLower=0;
			imMask[j][i].edgeRight=0;
			imMask[j][i].marked=0;
			imMask[j][i].display=0;
			j++;
			}
		i++;
		}
}

///////////////////////////////////////////////////////////////////////////////////
// PCN2639
// Name: CopyMaskEdgeToImage
// Input: none
// Output: none
//
// Description: Copy the mask image data back to the video image,
//				this is only used for debugging to see what the counter class sees
//				Does not effect the counter or profiler.				
///////////////////////////////////////////////////////////////////////////////////

void Counter::CopyMaskEdgeToImage(void)
{
int x,y;
	int i=0,j=0;
	for(x=xLeft;x<xRight;x++)
		{
		j=0;
		for(y=yTop;y<yLower;y++)
			{
			if(imMask[j][i].display==-1)
				{
				imVideo[y][x].blue=255;
				imVideo[y][x].red=255;
				imVideo[y][x].green=0;
				}
			if(imMask[j][i].display==-2)
				{
				imVideo[y][x].blue=255;
				imVideo[y][x].red=0;
				imVideo[y][x].green=255;
				}
			if(imMask[j][i].display==-3)
				{
				imVideo[y][x].blue=0;
				imVideo[y][x].red=255;
				imVideo[y][x].green=255;
				}

			j++;
			}
		i++;
		}
}

///////////////////////////////////////////////////////////////////////////////////
// PCN2639
// Name: AutoContrast
// Input: none
// Output: none
//
// Description: Calculates the best threshold value to use when converting an image to
//				black and white. Works well for calibration dots, not so well for
//				counter mask. To eratic for counter mask, needs work.
///////////////////////////////////////////////////////////////////////////////////

void Counter::AutoContrast(void)
{
	int currentWhite;
	double countLeft, countRight;
	int adjust=64;
	int centre;

	centre=128;
	while(adjust>1)
		{
		contrast=centre;		CopyImageToMaskEdge(); currentWhite=CountWhite();
		contrast=centre-adjust; CopyImageToMaskEdge(); countLeft=(double) CountWhite();
		contrast=centre+adjust; CopyImageToMaskEdge(); countRight=(double) CountWhite();

		countLeft=(double)  abs(currentWhite-(int) countLeft);
		countRight=(double) abs(currentWhite-(int) countRight);
		
		if(countLeft>countRight) centre=centre+adjust;
		else centre=centre-adjust;
		adjust/=2;
		}

	if(contrast<0) contrast=0;
	if(contrast>255) contrast=255;
}

///////////////////////////////////////////////////////////////////////////////////
// PCN2639
// Name: CountWhite
// Input: none
// Output: number of pixels that are white
//
// Description: Returns the number of pixels that are white,
//				this is used in auto-contrast to measure the black and white ratio.
///////////////////////////////////////////////////////////////////////////////////

int Counter::CountWhite(void)
{
	int totalWhite=0;
	int i,j;
	for(i=0;i<maskWidth;i++)
		for(j=0;j<maskHeight;j++)
			if(imMask[j][i].avgPixel==255) totalWhite++;
return totalWhite;
}

///////////////////////////////////////////////////////////////////////////////////
// PCN2639
// Name: TraceMask
// Input: none
// Output: none
//
// Description: takes the neigbouring values of each pixelReturns the number of pixels that are white,
//				this is used in auto-contrast to measure the black and white ratio.
///////////////////////////////////////////////////////////////////////////////////

void Counter::TraceMask(void)
{
	int i,j;
	CopyImageToMaskEdge();
	for(i=0;i<maskWidth-1;i++)
		for(j=0;j<maskHeight-1;j++)
			{
			imMask[j][i].edgeRight=CheckEdgeRight(i,j);
			imMask[j][i].edgeLower=CheckEdgeLower(i,j);
			imMask[j][i].edgeDiag=CheckEdgeDiag(i,j);
			}

}

int Counter::CheckEdgeRight(int i, int j)
{
	int diff;
	diff=imMask[j][i].avgPixel-imMask[j][i+1].avgPixel;
	return diff;

}

int Counter::CheckEdgeLower(int i, int j)
{
	int diff;
	diff=imMask[j][i].avgPixel-imMask[j+1][i].avgPixel;
	return diff;
}	

int Counter::CheckEdgeDiag(int i, int j)
{
	int diff;
	diff=imMask[j][i].avgPixel-imMask[j+1][i+1].avgPixel;
	return diff;
}	

///////////////////////////////////////////////////////////////////////////////////
// PCN2639
// Name: ScanForGreatestEdge
// Input: none
// Output: Int, The strongest Edge
//
// Description: First the strongest edge needs to be found to find out
//				what contrast threshold is to be used when detecting the
//				text.
///////////////////////////////////////////////////////////////////////////////////

int Counter::GetGreatestEdge(void)
{
	int greatestEdge=0;
	int ed,er,el; //diaganal, right and lower edge;
	int i,j;
	for(i=0;i<maskWidth-1;i++)
		for(j=0;j<maskHeight-1;j++)
			{
			ed=imMask[j][i].edgeDiag;
			er=imMask[j][i].edgeRight;
			el=imMask[j][i].edgeLower;
			if(ed>greatestEdge) greatestEdge=ed;
			if(er>greatestEdge) greatestEdge=er;
			if(el>greatestEdge) greatestEdge=el;
			}
	return greatestEdge;
}

///////////////////////////////////////////////////////////////////////////////////
// PCN2639
// Name: GetAverageEdge
// Input: none
// Output: Int, The average edge value for the mask area
//
// Description: Find the average edge value for the mask area
//				used when deciding if a edge is valid or not as a traced object.
///////////////////////////////////////////////////////////////////////////////////

int Counter::GetAverageEdge(void)
{
	int averageEdge=0;
	int count=0;
	int ed,er,el; //diaganal, right and lower edge;
	int i,j;
	for(i=0;i<maskWidth-1;i++)
		for(j=0;j<maskHeight-1;j++)
			{
			ed=imMask[j][i].edgeDiag;	count++;
			er=imMask[j][i].edgeRight;	count++;
			el=imMask[j][i].edgeLower;	count++;
			averageEdge+=(abs(ed)+abs(er)+abs(el));
			
			}
	if(count==0) return 0;
	return averageEdge/count;
}


///////////////////////////////////////////////////////////////////////////////////
// PCN2639
// Name: GetTextColour
// Input: none
// Output: median of text colour
//
// Description: Return the median Text Colour boarderd by the surrounded by the mask.
//				1) Trace Mask, (no threshold test)
//				2) Count and store the vertical edges (GetLineOff) set by lengthOffLine
//				3) Find the left pixel of the traced vertical edges and store, then sort
//				   from left to right.
//				4) Sort pixels in most left vertical trace from top to bottom
//				5) Return the value of the middle pixel. So this pixel is defined as follows
//					- Middle pixel from top to bottom of the left most vertical trace.
///////////////////////////////////////////////////////////////////////////////////

void Counter::ExtractTracedObjects(void)
{
	if(!isSet) return;
	int i,j;


	countTraces=0; // Keep track of the number of vertical traces that as => lengthOffLine 
	TraceMask(); // Get Trace information and store in edges.

	// Scan thru the mask image and look for contrasting edges
	for(j=0;j<maskHeight-1;j++)
		for(i=0;i<maskWidth-1;i++)
			if(TestForEdge(i,j,contrast) && !(imMask[j][i].marked))
				{
				if(countTraces==TRACE_SIZE) break;
				traceMask[countTraces].numberPoints=0;
				GetLineOff(i,j);// a recursive trace,
				imMask[j][i].marked=true;
				if(traceMask[countTraces].numberPoints<3) continue;
				countTraces++;
				}

	GetTracePositions(); // Find the position of the traces and sort from left to right.
	SortTraces();
	MarkTracesForDisplay();
}

void Counter::MarkTracesForDisplay(void)
{
	//Now mark all the diagonal edges of all the traces, for displaying purpose
	int tog=-1;
	int t,np;
	int i,j;

	for(t=0;t<countTraces;t++)
		{
		for(np=0;np<traceMask[t].numberPoints;np++)
			{
			i=traceMask[t].p[np].i;
			j=traceMask[t].p[np].j;
			if(traceMask[t].flagged) imMask[j][i].display=tog;
			else imMask[j][i].display=0;
			
			}	
		if(tog==-1) tog=-2;
		else if(tog==-2) tog=-3;
		else if(tog==-3) tog=-1;
		}
}

void Counter::GetLineOff(int i, int j)//, int countLine)
{
	if(countTraces>=TRACE_SIZE) return;
	if(imMask[j][i].marked) return;

	int numPoints;

	if(((i+1)<maskWidth) && ((i-1)>0) && ((j+1)<maskHeight) && ((j-1) >0) ) 
		{
		numPoints=traceMask[countTraces].numberPoints;
		if(numPoints>=POINTS_SIZE) return;
		traceMask[countTraces].p[numPoints].e.avgPixel=imMask[j][i+1].avgPixel; 
		traceMask[countTraces].p[numPoints].e.edgeRight=imMask[j][i].edgeRight;
		traceMask[countTraces].p[numPoints].e.edgeDiag=imMask[j][i].edgeDiag;
		traceMask[countTraces].p[numPoints].e.edgeLower=imMask[j][i].edgeLower;
		traceMask[countTraces].p[numPoints].i=i;
		traceMask[countTraces].p[numPoints].j=j;
		imMask[j][i].marked=1;
		traceMask[countTraces].numberPoints++;
		}
	else return;

//	if(TestForEdge(i-1,j+1,neighContr)) GetLineOff(i-1,j+1);
	if(TestForEdge(i-1,j  ,neighContr)) GetLineOff(i-1,j  );
//	if(TestForEdge(i-1,j-1,neighContr)) GetLineOff(i-1,j-1);
	if(TestForEdge(i  ,j+1,neighContr)) GetLineOff(i  ,j+1);
	if(TestForEdge(i  ,j-1,neighContr)) GetLineOff(i  ,j-1);
//	if(TestForEdge(i+1,j+1,neighContr)) GetLineOff(i+1,j+1);
	if(TestForEdge(i+1,j  ,neighContr)) GetLineOff(i+1,j  );
//	if(TestForEdge(i+1,j-1,neighContr)) GetLineOff(i+1,j-1);
	
	return;
}

int Counter::TestForEdge(int i,int j,int contrast)
{
	if(imMask[j][i].edgeRight>contrast) return true;
//	if(imMask[j][i].edgeDiag>contrast) return true;
	if(imMask[j][i].edgeLower>contrast) return true;
	return false;
}

void Counter::GetTracePositions(void)
{
	int i,j;
	int t,np;

	for(t=0;t<countTraces;t++)
		{
		traceMask[t].i_max=traceMask[t].p[0].i;
		traceMask[t].i_min=traceMask[t].p[0].i;
		traceMask[t].j_max=traceMask[t].p[0].j;
		traceMask[t].j_min=traceMask[t].p[0].j;
		traceMask[t].i_width=0;
		traceMask[t].j_width=0;
		traceMask[t].flagged=true;
		}

	for(t=0;t<countTraces;t++)
		{
		for(np=1;np<traceMask[t].numberPoints;np++)
			{
			i=traceMask[t].p[np].i;
			j=traceMask[t].p[np].j;
			if(i<traceMask[t].i_min) traceMask[t].i_min=i;
			if(i>traceMask[t].i_max) traceMask[t].i_max=i;
			if(j<traceMask[t].j_min) traceMask[t].j_min=j;
			if(j>traceMask[t].j_max) traceMask[t].j_max=j;
			traceMask[t].i_width=traceMask[t].i_max-traceMask[t].i_min;
			traceMask[t].j_width=traceMask[t].j_max-traceMask[t].j_min;
			}	
		}
}

void Counter::SortTraces(void)
{
	if(!isSet) return;
	int maxHeight;
	int minHeight;
	int t;
	int swap=true;
	trace tempTrace;

	while(swap)
		{
		swap=false;
		for(t=0;t<countTraces-1;t++)
			{
			if(traceMask[t].i_min>traceMask[t+1].i_min)
				{
				tempTrace=traceMask[t];
				traceMask[t]=traceMask[t+1];
				traceMask[t+1]=tempTrace;
				swap=true;
				}
			}
		}

	maxHeight=GetCommonTopHeight();
	minHeight=GetCommonLowerHeight();
	for(t=0;t<countTraces;t++)
		{
		if(traceMask[t].j_max<maxHeight-2) traceMask[t].flagged=false;
		}
	syTop=yTop+minHeight;
	syLower=yTop+maxHeight;
}

///////////////////////////////////////////////////////////////////////////////////
// PCN2972
// Name: GetCommonTopHeight
// Input: none
// Output: (int) The most common highest height
///////////////////////////////////////////////////////////////////////////////////
int Counter::GetCommonTopHeight(void)
{
	if(!isSet) return 0;
	if(maskHeight==0) return 0; //If the mask has not been initialised return 0
	
	int i,t;
	int *sampledHeights;
	sampledHeights = new int[maskHeight];
	for(i=0;i<maskHeight;i++) sampledHeights[i]=0; //Clear sampledHeights

	for(t=0;t<countTraces;t++) 
		{
		sampledHeights[traceMask[t].j_max]++;
		}
	for(i=maskHeight-1;i>0;i--)
		if((sampledHeights[i]+sampledHeights[i-1])>1) break;
//	DrawTopOfCounter(i);
	delete[] sampledHeights;
	return i;
}

///////////////////////////////////////////////////////////////////////////////////
// PCN2972
// Name: GetCommonLowerHeight
// Input: none
// Output: (int) The most common lowest height
///////////////////////////////////////////////////////////////////////////////////
int Counter::GetCommonLowerHeight(void)
{
	if(maskHeight==0) return 0; //If the mask has not been initialised return 0
	
	int i,t;
	int *sampledHeights;
	sampledHeights = new int[maskHeight];
	for(i=0;i<maskHeight;i++) sampledHeights[i]=0; //Clear sampledHeights

	for(t=0;t<countTraces;t++) 
		{
		sampledHeights[traceMask[t].j_min]++;
		}
	for(i=0;i<maskHeight-1;i++)
		if((sampledHeights[i]+sampledHeights[i-1])>1) break;
//	DrawTopOfCounter(i);
	delete[] sampledHeights;
	return i;
}

void Counter::FindDigitNextToDecimal(void)
{
	int t;
	if(countTraces==0) return; //PCN????
	if(units==METRIC)
		{
		for(t=0;t<countTraces;t++)
			{
			if(!traceMask[t].flagged) continue;
			if(traceMask[t].i_min>decimal_x) break;
			}
		sxLeft=traceMask[t].i_min+xLeft; sxRight=sxLeft+5;
		}
	else
		{
		for(t=countTraces-1;t>=0;t--)
			{
			if(!traceMask[t].flagged) continue;
			if(traceMask[t].i_max<decimal_x) break;
			}
		if(t<0) return;
		sxRight=traceMask[t].i_max+xLeft; sxLeft=sxRight-5;
		}
	
	// PCN2874 (Antony van Iersel, 8 June 2004)
	// if units are 1 (feet) then look to left of decimal point otherwise leave as is.
//	if((t>1) && (units==1)) t=t-2;
	
//	sxLeft=traceMask[t].i_min+xLeft; sxRight=traceMask[t].i_min+xLeft+5;
	
}

void Counter::FindHole(int &x, int &y)
{
	int biggestHole=0;
	int biggestHoleLeft=0;
	int biggestHoleRight=0;
	int holeLeft=0 , holeRight=0;
	int hole;
	int t,u;

	for(t=0;t<countTraces-1;t++)
		{
		if(!traceMask[t].flagged) continue;
		holeLeft=traceMask[t].i_max;
		for(u=t+1;u<countTraces;u++)
			{
			if(!traceMask[u].flagged) continue;
			holeRight=traceMask[u].i_min;
			break;
			}
		hole=holeRight-holeLeft;
		if(hole>biggestHole) 
			{
			biggestHole=hole;
			biggestHoleLeft=holeLeft;
			biggestHoleRight=holeRight;
			}
		}
	y=0;
	x=(biggestHoleLeft+biggestHoleRight)/2;
}

void Counter::FindDecimalPoint(int &x, int &y)
{
	if(countTraces<=0) {x=0;y=0;return;}

	int t;
	int i,j;
	int numberMabeyDecimalPoints=0;
	history mabeyDecimalPoints[50];
	int lowestDecimal=0;

	if(imWidth>400) pointSize=8;
	else pointSize=5;
	for(t=0;t<countTraces;t++)	
		{
		if((traceMask[t].i_width<pointSize) && (traceMask[t].j_width<pointSize) &&
		   (traceMask[t].i_width>1) && (traceMask[t].j_width>1) &&	
		   (traceMask[t].i_max>(maskWidth/4)) && 
		   (traceMask[t].i_min > traceMask[0].i_max))
			{
		   if(abs(traceMask[t].i_width-traceMask[t].j_width)<5)
				{
				if(numberMabeyDecimalPoints==49) break;	
				if(t!=countTraces)
					{
					i=traceMask[t].i_min+(traceMask[t].i_width/2);
					j=traceMask[t].j_min+(traceMask[t].j_width/2);
					mabeyDecimalPoints[numberMabeyDecimalPoints].x=i;
					mabeyDecimalPoints[numberMabeyDecimalPoints].y=j;
					numberMabeyDecimalPoints++;
					}
				}
			}
		}

	if(numberMabeyDecimalPoints==0) 
		{
		i=xLeft;j=yTop;
		return;
		}
	
	for(t=0;t<numberMabeyDecimalPoints;t++)
	{
//		DrawCircle(mabeyDecimalPoints[t].x+xLeft,
//				   mabeyDecimalPoints[t].y+yTop);
	}
	
	for(t=1;t<numberMabeyDecimalPoints;t++)
	{
		if(mabeyDecimalPoints[t].x>mabeyDecimalPoints[lowestDecimal].x) lowestDecimal=t;
	}
	i=mabeyDecimalPoints[lowestDecimal].x;
	j=mabeyDecimalPoints[lowestDecimal].y;
	x=i; y=j;
		
	i=i+xLeft;
	j=j+yTop;
//	DrawCircle(i,j);
}	

#define PI 3.14159265358979323846

void Counter::DrawCircle(int i, int j)
{
	if(!isSet) return;
	double rad;
	double x,y;
	int cx,cy;
//	int x,y;
	if((i<5) || (j<5) || (i>(imWidth-5)) || (j>(imHeight-5))) {/*MsgCounter("Draw Circle out of bounds");*/ return;}
	
	for(rad=0;rad<(2*PI);rad+=(PI/16))
		{
		x=sin(rad)*5;
		y=cos(rad)*5;
		cx=i+(int) x;
		cy=j+(int) y;

		imVideo[cy][cx].green=255;
		imVideo[cy][cx].blue=0;
		imVideo[cy][cx].red=0;
		}
}

void Counter::DrawSquare(int p1, int p2)
{
	if(!isSet) return;
	if((xLeft<0) || (xLeft>=imWidth) || (xRight<0) || (xRight>=imWidth)) return; //ANT VOB
	if((yTop<0) || (yTop>=imHeight) || (yLower<0) || (yLower>=imHeight)) return; //ANT VOB
	
	
	if((p1<xLeft) || (p1>xRight)) return;
	if((p2<xLeft) || (p2>xRight)) return;

	
	int y,x;

	for(x=xLeft;x<=xRight;x++)
		{
		imVideo[yTop][x].red=255;
		imVideo[yLower][x].red=255;
		}
	for(y=yTop;y<yLower;y++)
		{
		imVideo[y][xRight].red=255;
		imVideo[y][xLeft].red=255;
		}
		

	for(y=yTop;y<yLower;y++) 
		{
		for(x=sxLeft;x<=sxRight;x++)
			{
			imVideo[y][x].blue=255;
			imVideo[y][x].green=(unsigned char) (imVideo[y][x].green/3);


			}
		}

}

void Counter::DrawTopOfCounter(int t)
{
	if(!isSet) return;
	int i;
	for(i=xLeft;i<xRight;i++)
		{
		imVideo[yTop+t][i].red=255;
		imVideo[yTop+t][i].green=255;
		imVideo[yTop+t][i].blue=255;
		}

}

int Counter::IsInCounterArea(int x, int y)
{
	if((x<xLeft) || (x>xRight)) return false;
	if((y<yTop) || (y>yLower)) return false;
	else return true;
}




	


