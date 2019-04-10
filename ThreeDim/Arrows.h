#include <d3dx9.h>

class Arrow
{
	struct arrowStruct {
		float x,y,z,rhw;
		DWORD colour;
		FLOAT tu,tv;
		};
		

public:
	struct languageText{char text[101];}; //PCN2473 Antony van Iesel (12 May 2004)
	LPDIRECT3DDEVICE9		pd3dDeviceArrow;
	LPDIRECT3DVERTEXBUFFER9 pVBArrow;
	LPDIRECT3DTEXTURE9		pTextureArrow;
	char textureDirectory[800];
	languageText *languageArrow;

	CViewpipe *p;
	Camera *cam;
	float length;

	arrowStruct arrowNodes[4];
	
//	D3DXVECTOR3 originalArrow[4];
	
	Arrow(LPDIRECT3DDEVICE9 pd3dDevice, 
		  char *path,
		  languageText *language); //PCN2473 Language pointer (Antony, 11 May 2004)
	~Arrow();
	void Draw(void);
	void SetPipe(CViewpipe *pipe) {p=pipe;};
	void SetCamera(Camera *c) {cam=c;}
	void InitGeometry(void);
};