#include <windows.h>
#include <stdlib.h>
#include <atlbase.h>
#include <atlbase.h>
#include <d3dx9.h>
#include <stdio.h>
#include "Camera.h"
#include "Scape.h"

#define D3DFVF_CUSTOMVERTEX_SCAPE (D3DFVF_XYZ|D3DFVF_NORMAL|D3DFVF_DIFFUSE|D3DFVF_TEX1)



void MsgScape(TCHAR *szFormat, ...)
{
    TCHAR szBuffer[512];

    va_list pArgs;
    va_start(pArgs, szFormat);
    _vstprintf(szBuffer, szFormat, pArgs);
    va_end(pArgs);

    MessageBox(NULL, szBuffer, TEXT("LaserProfiler Message"), MB_OK | MB_ICONERROR);
}

Scape::Scape(LPDIRECT3DDEVICE9 pd3dDevice, 
			 char *path,
			 languageText *language)  //PCN2473 Language pointer (Antony, 11 May 2004)
{
	languageScape = language; //PCN2693 Needed for colour calculations (Antony van Iersel, 15 March 2004)
	pd3dDeviceScape=pd3dDevice; 
	pTextureScape =NULL;
	pVBScape	  =NULL;
	node		  =NULL; //PCN2465 (8 December 2003, Antony)
	pIBScape	  =NULL;

	strcpy(textureDirectory,path);

	size=65;	// Make a land from 0 - 128 by 0 - 128
	scale=3000;
	offset=D3DXVECTOR3(15800,2000,28360);
	fade=D3DXVECTOR4(1,1,1,1);
	rotateLand=(float) 152.628;
	numVertices = size * size;
	numberStrips=size-1;
	node = new nodeStruct[numVertices]; 
	lod = (long) (320*(scale/10));
	InitScape();
}

Scape::~Scape()
{

	int i;
	if(pVBScape!=NULL) pVBScape->Release();
	if(pTextureScape!=NULL) pTextureScape->Release();
	for(i=0;i<numberStrips;i++) if(pIBScape[i]!=NULL) pIBScape[i]->Release();
	if(pIBScape!=NULL) {delete[] pIBScape; pIBScape = NULL;}
	if(node!=NULL) { delete[] node; node = NULL; } //PCN2465 (8 December 2003,Antony) PCN3085 [] and = NULL added

}

void Scape::InitGeometry(void)
{
	long i;
	char texFile[800];
	nodeStruct *p_vertex;
	p_vertex=NULL;
	
//////////////////
// Load Texture //
//////////////////

	strcpy(texFile,textureDirectory);
	strcat(texFile, "\\Maps\\brown.jpg");
	if(FAILED( D3DXCreateTextureFromFile(pd3dDeviceScape, texFile, &pTextureScape) ))
		{
		//PCN4240 MsgScape("%s %s\n%s",languageScape[15].text, 
		//					 languageScape[16].text, // PCN2473 Language pointer (Antony, 11 May 2004)
		//					 texFile); //PCN2467 texFile Added (9 Dec 2003, Antony)
		pTextureScape=NULL; //PCN2485 added to make sure pTextureScape is not release when not there	
		}
	
///////////////////////////////////////////////////////////
// Create The Land Vertex Buffer and enter it with Data.
///////////////////////////////////////////////////////////
	pd3dDeviceScape->CreateVertexBuffer(numVertices*sizeof(nodeStruct),
												  0,D3DFVF_CUSTOMVERTEX_SCAPE,
												  D3DPOOL_DEFAULT, &pVBScape, NULL);
////
	long numberPointsStrip;
	numberPointsStrip=size*2;
	numberStrips=size-1;
	pIBScape = new LPDIRECT3DINDEXBUFFER9[numberStrips+5];
	for(i=0;i<numberStrips;i++) pIBScape[i]=NULL;
	for(i=0;i<numberStrips;i++)
		{
		pd3dDeviceScape->CreateIndexBuffer(numberPointsStrip*sizeof(WORD), 0,
						D3DFMT_INDEX16,	D3DPOOL_DEFAULT, &(pIBScape[i]), NULL);
		}



////




	pVBScape->Lock(0,0,(void **) &p_vertex,0); //D3DLOCK_DISCARD);
	for(i=0;i<numVertices;i++)  // Copy Scape Data into Vertex Buffer.
		{
		p_vertex[i].vertice=node[i].vertice;
		p_vertex[i].normal=node[i].normal;
		p_vertex[i].colour=node[i].colour;
		p_vertex[i].tu=node[i].tu;
		p_vertex[i].tv=node[i].tv;
		}
	pVBScape->Unlock();


/////////////////////////////////////////////////////////////////
// Create The Pipe Index Buffers and endter it with Data...
// The Pipe is split into patches (Like a quilt) 8 x 18 squares,
// Each patch has its own Index Buffer.
/////////////////////////////////////////////////////////////////
/*
	
	long numberPointsStrip;
	numberPointsStrip=size*2;
	numberStrips=size-1;
	pIBScape = new LPDIRECT3DINDEXBUFFER9[numberStrips+5];
	for(i=0;i<numberStrips;i++) pIBScape[i]=NULL;
	for(i=0;i<numberStrips;i++)
		{
		pd3dDeviceScape->CreateIndexBuffer(numberPointsStrip*sizeof(WORD), 0,
						D3DFMT_INDEX16,	D3DPOOL_DEFAULT, &(pIBScape[i]), NULL);
		}

 */
	WORD *p_index;
	p_index=NULL;
	
	long count=0;
	long strip;
	long offset1, offset2;	


//////////////////////////////
// Filling the Index Buffer //
//////////////////////////////
	//	*--*	offsetOne,   offsetTwo
	//	|\ |
	//	| \|
	//	*--*	offsetOne,   offsetTwo

	for(strip=0;strip<numberStrips;strip++)
		{
		pIBScape[strip]->Lock(0,0,(void **) &p_index, 0);
		count=0;
		for(i=0;i<size;i++)
			{
			offset1=i+((strip  )*size);
			offset2=i+((strip+1)*size);

			// First Traingle in Square
			p_index[count++]=(WORD) offset1; // * -* 1  2
			p_index[count++]=(WORD) offset2; // 
			}
		pIBScape[strip]->Unlock();
		}

}	

void Scape::DrawScape(void)
	{
	long i;
	long numberTriangles;
	D3DXMATRIX translate;
	D3DXMATRIX rotate;
	D3DXMATRIX temp;
	numberTriangles=(size-1)*2;
	

	D3DXMatrixRotationY(&rotate,rotateLand);
	D3DXMatrixTranslation( &translate, offset.x, offset.y, offset.z);
	D3DXMatrixMultiply(&translate,&translate,&rotate);
    pd3dDeviceScape->SetTransform( D3DTS_WORLD, &translate );
	
	pd3dDeviceScape->SetTexture(0,pTextureScape);
	pd3dDeviceScape->SetTextureStageState( 0, D3DTSS_COLOROP, D3DTOP_MODULATE);
	pd3dDeviceScape->SetTextureStageState( 0, D3DTSS_COLORARG1, D3DTA_TEXTURE);
	pd3dDeviceScape->SetTextureStageState( 0, D3DTSS_COLORARG2, D3DTA_DIFFUSE);
	pd3dDeviceScape->SetTextureStageState( 0, D3DTSS_ALPHAOP, D3DTOP_DISABLE);
	pd3dDeviceScape->SetStreamSource( 0, pVBScape, 0, sizeof(nodeStruct));
	pd3dDeviceScape->SetFVF( D3DFVF_CUSTOMVERTEX_SCAPE );
	for(i=0;i<numberStrips;i++)
		{
		pd3dDeviceScape->SetIndices(pIBScape[i]);
		pd3dDeviceScape->DrawIndexedPrimitive(D3DPT_TRIANGLESTRIP, 0, 0, numVertices,
											0,numberTriangles=(size-1)*2);
		}

	D3DXMatrixTranslation( &translate, 0,0,0);
	pd3dDeviceScape->SetTransform( D3DTS_WORLD, &translate );
	}	

void Scape::InitScape(void)
	{
	float x,z;
	long p=0;
	
	for(z=0;z<size;z++)
		for(x=0;x<size;x++)
			{
			node[p].vertice=D3DXVECTOR3((x-(size/2))*scale,0,(z-(size/2))*(-scale*(float)0.60));
			node[p].normal=D3DXVECTOR3(0,-1,0);
			node[p].colour=D3DCOLOR_RGBA(255,255,255,170); // Red, Green, Blue, Alpha.
			node[p].tu=(float) x / (float) size;
			node[p].tv=(float) z / (float) size;
			p++;
			}
	}

void Scape::Shadeland(long startX, long startY, long endX, long endY)
{
//	long x, y;
//	if(startX<1) startX=1;	if(startY<1) startY=1;
//	if(endX>=size-1) endX=size-2; if(endY>=size-1) endY=size-2;
//	for(y=startY;y<endY;y++)
//		for(x=startX;x<endX;x++)
//			{
//			if( ((x%2)     && (y%2))     ) AvgAllEightSec(x,y);
//			if( (((x+1)%2) && ((y+1)%2)) ) AvgAllEightSec(x,y);
//			if( (((x+1)%2) && (y%2))     ) AvgFourSec(x,y);
//		    if( ((x%2)     && ((y+1)%2)) ) AvgFourSec(x,y);
//			}
}

#define POne   (x-1)+((y-1)*size)
#define PTwo     (x)+((y-1)*size)
#define PThree (x+1)+((y-1)*size)
#define PFour  (x-1)+(y*size)
#define PFive    (x)+(y*size)
#define PSix   (x+1)+(y*size)
#define PSeven (x-1)+((y+1)*size)
#define PEight   (x)+((y+1)*size)
#define PNine  (x+1)+((y+1)*size)
	
void Scape::AvgAllEightSec(long x, long y)
{
//	int i;
//	sgVec3 tOne,tTwo,tThree,tFour,tFive,tSix,tSeven,tEight;
//	sgVec3 *vOne, *vTwo, *vThree, *vFour, *vFive, *vSix, *vSeven, *vEight, *vNine;
//	
//	vOne=&node[POne].vertice;	   vTwo=&node[PTwo].vertice;	   vThree=&node[PThree].vertice;
//	vFour=&node[PFour].vertice;   vFive=&node[PFive].vertice;   vSix=&node[PSix].vertice;
//	vSeven=&node[PSeven].vertice; vEight=&node[PEight].vertice; vNine=&node[PNine].vertice;
//
//	sgMakeNormal(tOne,   *vTwo,   *vOne,   *vFive);
//	sgMakeNormal(tTwo,   *vThree,   *vTwo, *vFive);
//	sgMakeNormal(tThree, *vSix, *vThree,   *vFive);
//	sgMakeNormal(tFour,  *vSix,  *vFive,   *vNine);
//	sgMakeNormal(tFive,  *vNine,  *vFive,  *vEight);
//	sgMakeNormal(tSix,   *vEight,  *vFive, *vSeven);
//	sgMakeNormal(tSeven, *vSeven,  *vFive,  *vFour);
//	sgMakeNormal(tEight, *vFive,   *vOne,  *vFour);

//	for(i=0;i<3;i++) node[x+(y*size)].normal[i]=(tOne[i]+tTwo[i]+tThree[i]+tFour[i]+tFive[i]+tSix[i]+tSeven[i]+tEight[i])/8;
}

void Scape::AvgFourSec(long x, long y)
	{
//	int i;
//	sgVec3 tOne,tTwo,tThree,tFour;
//	sgVec3 *vTwo, *vFour, *vFive, *vSix, *vEight;
//	
//	vTwo=&node[PTwo].vertice;
//	vFour=&node[PFour].vertice;   vFive=&node[PFive].vertice;   vSix=&node[PSix].vertice;
//	vEight=&node[PEight].vertice;
//	
//	sgMakeNormal(tOne, *vFive, *vTwo, *vFour);
//	sgMakeNormal(tTwo, *vSix, *vTwo, *vFive);
//	sgMakeNormal(tThree, *vSix, *vFive, *vEight);
//	sgMakeNormal(tFour, *vEight, *vFive, *vFour);
//	for(i=0;i<3;i++) node[x+(y*size)].normal[i]=(tOne[i]+tTwo[i]+tThree[i]+tFour[i])/4;
	}




