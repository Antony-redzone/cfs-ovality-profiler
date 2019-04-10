#include <d3dx9.h>

class D3DObject
{
public:
	struct languageText{char text[101];}; //PCN2473 Antony van Iesel
	LPDIRECT3DDEVICE9	pd3dDeviceObject;
	LPD3DXMESH			pMesh;
	D3DMATERIAL9		*pMeshMaterials;
	LPDIRECT3DTEXTURE9	*pMeshTextures;
	DWORD				dwNumMaterials;
	Camera *objectCam;
	int	alignRotation; //Alignment Rotation to World or Camera;
	int	alignPosition; //Alignment Position to World or Camera;

	char textureDirectory[800];
	char objectFile[800];
	languageText *languageD3DObject;

	D3DXVECTOR3 position;
	D3DXVECTOR3 rotation;
	D3DXVECTOR3 scale;

	D3DObject(LPDIRECT3DDEVICE9 pd3dDevice, // Device to be rendered to.
			  char *path,					// Directory path of textures, models etc
			  char *object,					// Model Name
			  D3DXVECTOR3 rot,				// Rotation of Model
			  D3DXVECTOR3 pos,				// Position of Model
			  D3DXVECTOR3 sca,				// Scale of Model
			  int alignRot,					// Align the position with World (0) or Camera (1), 
			  int alignPos,  				// Align the rotation with World (0) or Carema (1)
			  languageText *language); //PCN2473 Language pointer (Antony, 11 May 2004)
	~D3DObject(void);

	void Draw(void);
	void SetCam(Camera *c) {objectCam=c;};
//private:
	void InitGeometry(void);
	

};