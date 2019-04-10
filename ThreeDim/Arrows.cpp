#include "Pipe.h"
#include "Camera.h"
#include "Arrows.h"
#include <atlbase.h>
#include <stdio.h>

#define D3DFVF_CUSTOMVERTEX_ARROW (D3DFVF_XYZRHW|D3DFVF_DIFFUSE|D3DFVF_TEX1)

void MsgArrow(TCHAR *szFormat, ...)
{
    TCHAR szBuffer[512];

    va_list pArgs;
    va_start(pArgs, szFormat);
    _vstprintf(szBuffer, szFormat, pArgs);
    va_end(pArgs);

    MessageBox(NULL, szBuffer, TEXT("LaserProfiler Message"), MB_OK | MB_ICONERROR);
}

Arrow::Arrow(LPDIRECT3DDEVICE9 pd3dDevice,
			 char *path,
			 languageText *language)  //PCN2473 Language pointer (Antony, 11 May 2004)
{
	languageArrow = language;  //PCN2473 Language pointer (Antony, 11 May 2004)
	pd3dDeviceArrow=pd3dDevice;
	strcpy(textureDirectory,path);
	int i;
	arrowNodes[0].x = 143;  arrowNodes[0].y=55;  arrowNodes[0].tu = 0; arrowNodes[0].tv = 0;
	arrowNodes[1].x = 143;  arrowNodes[1].y=93;  arrowNodes[1].tu = 0; arrowNodes[1].tv = 0;
	arrowNodes[2].x = 143;  arrowNodes[2].y=55; arrowNodes[2].tu = 0; arrowNodes[2].tv = 0;	
	arrowNodes[3].x = 143;  arrowNodes[3].y=93; arrowNodes[3].tu = 0; arrowNodes[3].tv = 0;
	for(i=0;i<4;i++) 
		{
		arrowNodes[i].colour=0x90ffffff;	
		arrowNodes[i].rhw=1.0;
		arrowNodes[i].z=0.5;
		}
}

Arrow::~Arrow()
{
	if(pVBArrow       !=NULL)		 pVBArrow->Release();
	if(pTextureArrow  !=NULL)  pTextureArrow->Release();
}


void Arrow::InitGeometry(void)
{
	int i;
	char texFile[800];
	arrowStruct *p_vertex;
	p_vertex=NULL;
	
//////////////////
// Load Texture //
//////////////////
	strcpy(texFile, textureDirectory);
	strcat(texFile, "\\Arrow.jpg" );

if(FAILED( D3DXCreateTextureFromFile(pd3dDeviceArrow, 
						  texFile,
						  &pTextureArrow) ))
	{
	//PCN4240 MsgArrow("%s %s\n%s",languageArrow[15].text, // PCN2473 Language pointer (Antony, 11 May 2004)
	//					 languageArrow[16].text, // PCN2473 Language pointer (Antony, 11 May 2004)
	//					texFile); //PCN2467 texFile Added (9 Dec 2003, Antony)
	pTextureArrow=NULL; //PCN2485 if no texture then make sure its null to prevent release when
	}					//closing pipe. (Note Doesn't work - 15 Dec 2003, Antony)
///////////////////////////////////////////////////////////
// Create The Pipe Vertex Buffer and enter it with Data.
///////////////////////////////////////////////////////////
	pd3dDeviceArrow->CreateVertexBuffer(4*sizeof(arrowStruct),
										0,D3DFVF_CUSTOMVERTEX_ARROW,
										D3DPOOL_DEFAULT, &pVBArrow, NULL);
	pVBArrow->Lock(0,0,(void **) &p_vertex,0); //D3DLOCK_DISCARD);
	for(i=0;i<4;i++)  // Copy Scape Data into Vertex Buffer.
		{
		p_vertex[i].x = arrowNodes[i].x;
		p_vertex[i].y = arrowNodes[i].y;
		p_vertex[i].z = arrowNodes[i].z;
		p_vertex[i].rhw = arrowNodes[i].rhw;
		p_vertex[i].colour =arrowNodes[i].colour;
		p_vertex[i].tu     =arrowNodes[i].tu;
		p_vertex[i].tv	   =arrowNodes[i].tv;
		}
	pVBArrow->Unlock();

}

void Arrow::Draw()
{
	if(p==NULL) return; //PCN2461 (8 December 2003, Antony)
	int i;
	D3DXVECTOR3 distance;
	distance = ( cam->position - (p->topRing));
	length=D3DXVec3Length(&distance);
	if(length<15*p->avgRad) return;
	arrowStruct *p_vertex;

	pVBArrow->Lock(0,0,(void **) &p_vertex,0);

	for(i=0;i<2;i++)  // Copy Scape Data into Vertex Buffer.
		{
		p_vertex[i].x=p->laserRingProj[p->highestPoint].x; 
		p_vertex[i].y=p->laserRingProj[p->highestPoint].y;
		}
	pVBArrow->Unlock();

	pd3dDeviceArrow->SetTexture(0,pTextureArrow);
	pd3dDeviceArrow->SetTextureStageState( 0, D3DTSS_COLOROP, D3DTOP_MODULATE);
	pd3dDeviceArrow->SetTextureStageState( 0, D3DTSS_COLORARG1, D3DTA_TEXTURE);
	pd3dDeviceArrow->SetTextureStageState( 0, D3DTSS_COLORARG2, D3DTA_DIFFUSE);
	pd3dDeviceArrow->SetTextureStageState( 0, D3DTSS_ALPHAOP, D3DTOP_DISABLE);
	pd3dDeviceArrow->SetStreamSource( 0, pVBArrow, 0, sizeof(arrowStruct));
	pd3dDeviceArrow->SetFVF( D3DFVF_CUSTOMVERTEX_ARROW );	
	pd3dDeviceArrow->DrawPrimitive(D3DPT_TRIANGLEFAN, 0, 2);
		
}
