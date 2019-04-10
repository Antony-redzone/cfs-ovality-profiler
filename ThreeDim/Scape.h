class Scape 
{
public:
	struct languageText{char text[101];}; //PCN2473 Antony van Iesel (12 May 2004)
	long size;
	float scale;
	D3DXVECTOR3 offset;
	D3DXVECTOR4 fade;
	float rotateLand;
	long numVertices;
	long lod;
	long numberStrips;

	struct nodeStruct {
		D3DXVECTOR3 vertice;
		D3DXVECTOR3 normal;
		DWORD colour;
		FLOAT tu,tv;
		};
	nodeStruct *node;
	Camera *scapeCam;

	LPDIRECT3DDEVICE9		pd3dDeviceScape;
	LPDIRECT3DVERTEXBUFFER9 pVBScape;
	LPDIRECT3DINDEXBUFFER9 *pIBScape;
	LPDIRECT3DTEXTURE9		pTextureScape;
	char textureDirectory[800];
	languageText *languageScape;

	Scape(LPDIRECT3DDEVICE9 pd3dDevice,
		  char *path,
		  languageText *language);
	~Scape();
	void DrawScape(void);
	void GetFan(long x, long y, long size);
	void DisplayFan(long x1, long y1, long x2, long y2, long x_cen, long y_cen);
	void SetCam(Camera *c) {scapeCam=c;};
	void LodMore(void) {lod*=2;}
	void LodLess(void) {lod/=2;}
	void InitScape(void);
	void Shadeland(long startX, long startY, long endX, long endY);
	void AvgAllEightSec(long x, long y);
	void AvgFourSec(long x, long y);
	void InitGeometry(void);

};

