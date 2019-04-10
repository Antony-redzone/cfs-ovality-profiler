#include <d3dx9.h>

class D3DFont
{
	struct panelDimension {
		float x,y;
		float height;
		float width;
		float boarderHeight;
		float boarderWidth;
		DWORD boarderColour;
		DWORD colour;
		};
	struct textStruct {
		LOGFONT logFont;
		char text[100];
		char length;
		D3DXVECTOR2 position;
		DWORD format;
		DWORD colour;
		};
	struct panelStruct {
		float x,y,z,rhw;
		DWORD colour;
		FLOAT tu,tv;
		};
		

public:
	struct languageText{char text[101];}; //PCN2473 Antony van Iesel (12 May 2004)
	LPDIRECT3DDEVICE9		pd3dDeviceFont;
	LPDIRECT3DVERTEXBUFFER9 pVBPanel;
	LPDIRECT3DVERTEXBUFFER9 pVBBoarder;
	LPDIRECT3DTEXTURE9		pTexturePanel;
	LPDIRECT3DTEXTURE9		pTextureBoarder;
	char textureDirectory[800];
	char texturePanel[800];
	char textureBoarder[800];

	textStruct *textArray[20];
	languageText *languageD3DFont;  //PCN2473 Language pointer (Antony, 11 May 2004)
	

	panelDimension panelDim;
	panelStruct panel[4];
	panelStruct boarder[10];
	long noText;
	DWORD defaultFormat;
	DWORD defaultColour;

	LPD3DXFONT pFont;
	LOGFONT logFont;
	RECT fontPosition;	

	D3DFont(LPDIRECT3DDEVICE9 pd3dDevice, 
			char *path,
			char *panel, 
			char *boarder,
			languageText *language); //PCN2473 Language pointer (Antony, 11 May 2004)
	~D3DFont();
	void SetPanel(void);
	void Draw(void);
	void InitGeometry(void);
	void SetPanelDimension(void);
	void NewText(char *t, D3DXVECTOR2 p);
	void NewText(int i, D3DXVECTOR2 p);
	void NewText(int i, D3DXVECTOR2 p, DWORD f);
	void NewText(float f, D3DXVECTOR2 p);
	void NewText(char *t, D3DXVECTOR2 p, DWORD format, DWORD c);
	void UpdateText(int sel, char *t);
	void UpdateText(int sel, int i);
	void UpdateText(int sel, float f);
};