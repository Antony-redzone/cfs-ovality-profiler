
#include "D3DFont.h"
#include <atlbase.h>
#include <stdio.h>

#define D3DFVF_CUSTOMVERTEX_PANEL (D3DFVF_XYZRHW|D3DFVF_DIFFUSE|D3DFVF_TEX1)

void MsgFont(TCHAR *szFormat, ...)
{
    TCHAR szBuffer[512];

    va_list pArgs;
    va_start(pArgs, szFormat);
    _vstprintf(szBuffer, szFormat, pArgs);
    va_end(pArgs);

    MessageBox(NULL, szBuffer, TEXT("LaserProfiler Message"), MB_OK | MB_ICONERROR);
}

D3DFont::D3DFont(LPDIRECT3DDEVICE9 pd3dDevice, 
				 char *path, 
				 char *panel, 
				 char *boarder,
				 languageText *language)  //PCN2473 Language pointer (Antony, 11 May 2004)
{
	languageD3DFont = language;  //PCN2473 Language pointer (Antony, 11 May 2004)
	pVBPanel      = NULL;
	pTexturePanel = NULL;
	pFont		  = NULL;
	pd3dDeviceFont= NULL; //PCN2465 (8 December 2003, Antony)

	pd3dDeviceFont = pd3dDevice;
	strcpy(textureDirectory,  path);
	strcpy(texturePanel,     panel);
	strcpy(textureBoarder, boarder);
	

	int i;

	defaultFormat = DT_LEFT;
	defaultColour = 0xffff2000;

	logFont.lfHeight=18;
	logFont.lfWidth=0;
	logFont.lfEscapement=0;
	logFont.lfOrientation=0;
	logFont.lfWeight=FW_NORMAL;
	logFont.lfItalic=false;
	logFont.lfUnderline=false;
	logFont.lfStrikeOut=false;
	logFont.lfCharSet=DEFAULT_CHARSET;
	logFont.lfOutPrecision=OUT_TT_PRECIS;
	logFont.lfClipPrecision=CLIP_DEFAULT_PRECIS;
	logFont.lfQuality=PROOF_QUALITY;
	logFont.lfPitchAndFamily=DEFAULT_PITCH;
	strcpy(logFont.lfFaceName,"Arial");

	D3DXCreateFontIndirect(pd3dDeviceFont,&logFont,&pFont);

	fontPosition.top=5;
	fontPosition.left=5;
	fontPosition.right= 125;
	fontPosition.bottom=150;

	panelDim.x			   = 5;	  panelDim.y			= 5;
	panelDim.width		   = 140; panelDim.height		= 85;
	panelDim.boarderHeight = 3;	  panelDim.boarderWidth = 3;
	panelDim.colour        = D3DCOLOR_RGBA(255,255,255,90);
	panelDim.boarderColour = D3DCOLOR_RGBA(255,255,255,255);

	noText=0;

//	SetPanelDimension();
	for(i=0;i<20;i++) textArray[i]=NULL;
}

D3DFont::~D3DFont()
{
	for(int i=0;i<noText;i++) delete textArray[i]; //PCN2465 remove memory leaks PCN3085 removed brackets around testArray[i]
	if(pVBPanel       !=NULL)        pVBPanel->Release();  
    if(pVBBoarder     !=NULL)      pVBBoarder->Release();
	if(pTexturePanel  !=NULL)   pTexturePanel->Release();
	if(pTextureBoarder!=NULL) pTextureBoarder->Release();
	if(pFont          !=NULL)           pFont->Release();
//	if(pd3dDeviceFont!=NULL) pd3dDeviceFont->Release(); //PCN2465 (8 December 2003, Antony) PCN3085, not suppose to release.
}

void D3DFont::SetPanel(void)
{
	SetPanelDimension();
}

void D3DFont::SetPanelDimension()
{
	int i, count=0;

//////////////////////////////
// Extracting Panel VB Data //
//////////////////////////////

	for(i=0;i<10;i++)
		{
		panel[i].rhw=1.0;
		panel[i].colour= panelDim.colour;
		panel[i].z = 0.5;
		}
	
	panel[0].x=panelDim.x;
	panel[0].y=panelDim.y;
	panel[0].tu=0; panel[0].tv=0;

	
	panel[1].x=panelDim.x+panelDim.width;
	panel[1].y=panelDim.y;
	panel[0].tu=0; panel[0].tv=1;

	panel[2].x=panelDim.x+panelDim.width;
	panel[2].y=panelDim.y+panelDim.height;
	panel[0].tu=1; panel[0].tv=1;

	panel[3].x=panelDim.x;
	panel[3].y=panelDim.y+panelDim.height;
	panel[0].tu=1; panel[0].tv=0;


//////////////////////////////////////
// Extracting Panel Boarder VB Data //
//////////////////////////////////////
	
	for(i=0;i<5;i++)
		{
		boarder[count++].tu=0;
		boarder[count++].tu=1;
		boarder[i*2].tv=(float) i/4; boarder[(i+1)*2].tv=(float) i/4;
		}
	for(i=0;i<10;i++)
		{
		boarder[i].rhw=1.0;
		boarder[i].colour=panelDim.boarderColour;
		boarder[i].z = 0.5;
		}

	boarder[0].x  = panelDim.x-panelDim.boarderWidth;
	boarder[0].y  = panelDim.y-panelDim.boarderHeight;
	
	boarder[1].x = panelDim.x;
	boarder[1].y = panelDim.y;
	
	boarder[2].x = panelDim.x+panelDim.width+panelDim.boarderWidth;
	boarder[2].y = panelDim.y-panelDim.boarderHeight;
	
	boarder[3].x = panelDim.x+panelDim.width;
	boarder[3].y = panelDim.y;
	
	boarder[4].x = panelDim.x+panelDim.width+panelDim.boarderWidth;
	boarder[4].y = panelDim.y+panelDim.height+panelDim.boarderHeight;
	
	boarder[5].x = panelDim.x+panelDim.width;
	boarder[5].y = panelDim.y+panelDim.height;
	
	boarder[6].x = panelDim.x-panelDim.boarderWidth;
	boarder[6].y = panelDim.y+panelDim.height+panelDim.boarderHeight;
	
	boarder[7].x = panelDim.x;
	boarder[7].y = panelDim.y+panelDim.height;
	
	boarder[8].x = panelDim.x-panelDim.boarderWidth;
	boarder[8].y = panelDim.y-panelDim.boarderHeight;
	
	boarder[9].x = panelDim.x;
	boarder[9].y = panelDim.y;
}



void D3DFont::InitGeometry(void)
{
	int i;
	char texFile[800];
	panelStruct *p_vertex;
	p_vertex=NULL;
	
//////////////////
// Load Texture //
//////////////////
	strcpy(texFile, textureDirectory);
	strcat(texFile, texturePanel );
	if(FAILED( D3DXCreateTextureFromFile(pd3dDeviceFont, texFile, &pTexturePanel) ))
		{
		//PCN4240 MsgFont("%s %s\n%s", languageD3DFont[15].text, // PCN2473 Language pointer (Antony, 11 May 2004)
		//					 languageD3DFont[16].text, // PCN2473 Language pointer (Antony, 11 May 2004)
		//					 texFile); //PCN2467 texFile Added (9 Dec 2003, Antony)
		pTexturePanel=NULL; //PCN2485 added to make sure pTexturePanel is not released when not there
		}
	strcpy(texFile, textureDirectory);
	strcat(texFile, textureBoarder );
	if(FAILED( D3DXCreateTextureFromFile(pd3dDeviceFont, texFile, &pTextureBoarder) ))
		{
		//PCN4240 MsgFont("%s %s\n%s", languageD3DFont[15].text, // PCN2473 Language pointer (Antony, 11 May 2004)
		//					 languageD3DFont[16].text, // PCN2473 Language pointer (Antony, 11 May 2004)
		//					 texFile); //PCN2467 texFile Added (9 Dec 2003, Antony) 
		pTextureBoarder=NULL; //PCN2485 added to make sure pTextureBoarder is not released when not there
		}

///////////////////////////////////////////////////////////
// Create The Panel Vertex Buffer and enter it with Data.
///////////////////////////////////////////////////////////
	pd3dDeviceFont->CreateVertexBuffer(4*sizeof(panelStruct),
										0,D3DFVF_CUSTOMVERTEX_PANEL,
										D3DPOOL_DEFAULT, &pVBPanel, NULL);
	pVBPanel->Lock(0,0,(void **) &p_vertex,0); //D3DLOCK_DISCARD);
	for(i=0;i<4;i++)  // Copy Scape Data into Vertex Buffer.
		{
		p_vertex[i].x     =panel[i].x;
		p_vertex[i].y     =panel[i].y;
		p_vertex[i].z     =panel[i].z;
		p_vertex[i].rhw   =panel[i].rhw;
		p_vertex[i].colour=panel[i].colour;
		p_vertex[i].tu    =panel[i].tu;
		p_vertex[i].tv	  =panel[i].tv;
		}
	pVBPanel->Unlock();

///////////////////////////////////////////////////////////
// Create The Boarder Vertex Buffer and enter it with Data.
///////////////////////////////////////////////////////////
	pd3dDeviceFont->CreateVertexBuffer(10*sizeof(panelStruct),
										0,D3DFVF_CUSTOMVERTEX_PANEL,
										D3DPOOL_DEFAULT, &pVBBoarder, NULL);
	pVBBoarder->Lock(0,0,(void **) &p_vertex,0); //D3DLOCK_DISCARD);
	for(i=0;i<10;i++)  // Copy Scape Data into Vertex Buffer.
		{
		p_vertex[i].x     =boarder[i].x;
		p_vertex[i].y     =boarder[i].y;
		p_vertex[i].z     =boarder[i].z;
		p_vertex[i].rhw   =boarder[i].rhw;
		p_vertex[i].colour=boarder[i].colour;
		p_vertex[i].tu    =boarder[i].tu;
		p_vertex[i].tv	  =boarder[i].tv;
		}
	pVBBoarder->Unlock();
}

void D3DFont::NewText(char *t, D3DXVECTOR2 p)
{
	NewText(t, p, defaultFormat, defaultColour);
}
	
void D3DFont::NewText(int i, D3DXVECTOR2 p)
{
	char t[800];
	itoa(i, t, 10);
	NewText(t, p, DT_RIGHT, defaultColour);
}

void D3DFont::NewText(int i, D3DXVECTOR2 p, DWORD f)
{
	char t[800];
	itoa(i, t, 10);
	NewText(t, p, f, defaultColour);
}

void D3DFont::NewText(char *t, D3DXVECTOR2 p, DWORD f, DWORD c)
{
	textArray[noText] = new textStruct();
	strcpy(textArray[noText]->text,t);
	textArray[noText]->position=p;
	textArray[noText]->format=f;
	textArray[noText]->colour=c;
	textArray[noText]->length=strlen(t);
	noText++;
}

void D3DFont::UpdateText(int sel,char *t)
{
	strcpy(textArray[sel]->text, t);		
	textArray[sel]->length=strlen(textArray[sel]->text);
}

void D3DFont::UpdateText(int sel,int i)
{
	itoa(i,textArray[sel]->text,10);
	textArray[sel]->length=strlen(textArray[sel]->text);
}

void D3DFont::UpdateText(int sel, float f)
{
	sprintf(textArray[sel]->text,"%3.1f",f);
	textArray[sel]->length=strlen(textArray[sel]->text);
}

void D3DFont::Draw()
{
	int i;

	////////////////
	// Draw Panel //
	////////////////

	pd3dDeviceFont->SetTexture(0,pTexturePanel);
	pd3dDeviceFont->SetTextureStageState( 0, D3DTSS_COLOROP, D3DTOP_MODULATE);
	pd3dDeviceFont->SetTextureStageState( 0, D3DTSS_COLORARG1, D3DTA_TEXTURE);
	pd3dDeviceFont->SetTextureStageState( 0, D3DTSS_COLORARG2, D3DTA_DIFFUSE);
	pd3dDeviceFont->SetTextureStageState( 0, D3DTSS_ALPHAOP, D3DTOP_DISABLE);
	pd3dDeviceFont->SetStreamSource( 0, pVBPanel, 0, sizeof(panelStruct));
	pd3dDeviceFont->SetFVF( D3DFVF_CUSTOMVERTEX_PANEL );	
	pd3dDeviceFont->DrawPrimitive(D3DPT_TRIANGLEFAN, 0, 2);

	//////////////////
	// Draw Boarder //
	//////////////////
	
	pd3dDeviceFont->SetTexture(0,pTextureBoarder);
	pd3dDeviceFont->SetTextureStageState( 0, D3DTSS_COLOROP, D3DTOP_MODULATE);
	pd3dDeviceFont->SetTextureStageState( 0, D3DTSS_COLORARG1, D3DTA_TEXTURE);
	pd3dDeviceFont->SetTextureStageState( 0, D3DTSS_COLORARG2, D3DTA_DIFFUSE);
	pd3dDeviceFont->SetTextureStageState( 0, D3DTSS_ALPHAOP, D3DTOP_DISABLE);
	pd3dDeviceFont->SetStreamSource( 0, pVBBoarder, 0, sizeof(panelStruct));
	pd3dDeviceFont->SetFVF( D3DFVF_CUSTOMVERTEX_PANEL );	
	pd3dDeviceFont->DrawPrimitive(D3DPT_TRIANGLESTRIP, 0, 8);


	pFont->Begin();
		for(i=0;i<noText;i++)
			{
			fontPosition.left=(long) textArray[i]->position.x;
			fontPosition.top =(long) textArray[i]->position.y;
		    pFont->DrawText(textArray[i]->text,
							textArray[i]->length,
							&fontPosition,
							textArray[i]->format,
							textArray[i]->colour);
			}
	pFont->End();
}
