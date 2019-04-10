#include <atlbase.h>
#include <stdio.h>
#include "Camera.h"
#include "D3DObject.h"


void MsgD3DObject(TCHAR *szFormat, ...)
{
    TCHAR szBuffer[512];

    va_list pArgs;
    va_start(pArgs, szFormat);
    _vstprintf(szBuffer, szFormat, pArgs);
    va_end(pArgs);

    MessageBox(NULL, szBuffer, TEXT("LaserProfiler Message"), MB_OK | MB_ICONERROR);
}

D3DObject::D3DObject(LPDIRECT3DDEVICE9 pd3dDevice, 
					 char *path, 
					 char *object, 
					 D3DXVECTOR3 rot,
					 D3DXVECTOR3 pos,
					 D3DXVECTOR3 sca,
					 int alignRot,
					 int alignPos,
					 languageText *language)  //PCN2473 Language pointer (Antony, 11 May 2004)
{
	languageD3DObject = language;  //PCN2473 Language pointer (Antony, 11 May 2004)
	pMesh			= NULL;
	pMeshMaterials = NULL;
	pMeshTextures	= NULL;
	pd3dDeviceObject = pd3dDevice;
	dwNumMaterials	= 0;
	objectCam = NULL;

	strcpy(textureDirectory,path);
	strcpy(objectFile,object);

	position=pos;
	rotation=rot;
	scale=sca;

	alignRotation = alignRot; // 0 for World, 1 for Camera
	alignPosition = alignPos; // 0 for World, 1 for Camera

	InitGeometry();
}

D3DObject::~D3DObject(void)
{
	if(pMeshMaterials != NULL) delete[] pMeshMaterials;
	if(pMeshTextures)
		{
		for(DWORD i=0;i<dwNumMaterials;i++)
			if(pMeshTextures[i]) pMeshTextures[i]->Release();
		delete[] pMeshTextures;
		}
	if(pMesh!=NULL) pMesh->Release();
	
}

void D3DObject::InitGeometry(void)
{
	char objectFileName[800];
	char objectTextureName[800];
	char objectDirectory[800];

	strcpy(objectFileName, textureDirectory);
	strcat(objectFileName, "\\Models\\");
	strcat(objectFileName, objectFile);

	strcpy(objectDirectory, textureDirectory);
	strcat(objectDirectory, "\\Models\\");

	LPD3DXBUFFER pD3DXMtrlBuffer=NULL; // Only used temporary to extract textures and materials
	D3DXMATERIAL *d3dxMaterials;

	if(FAILED(D3DXLoadMeshFromX(objectFileName,
								D3DXMESH_SYSTEMMEM,
								pd3dDeviceObject,
								NULL,
								&pD3DXMtrlBuffer,
								NULL,
								&dwNumMaterials,
								&pMesh)));

	//	PCN4240 MsgD3DObject("%s %s\n%s",languageD3DObject[15].text, // PCN2473 Language pointer (Antony, 11 May 2004)
	//							 languageD3DObject[22].text, // PCN2473 Language pointer (Antony, 11 May 2004)
	//							 objectFileName);
//	MsgD3DObject("%s %s",languagePipe[15],languagePipe[22]);
	if ((D3DXMATERIAL *) pD3DXMtrlBuffer!=0)
		d3dxMaterials = (D3DXMATERIAL *) pD3DXMtrlBuffer->GetBufferPointer();
	pMeshMaterials = new D3DMATERIAL9[dwNumMaterials];
	pMeshTextures  = new LPDIRECT3DTEXTURE9[dwNumMaterials];
	for(DWORD i=0;i<dwNumMaterials; i++)
		{
		pMeshMaterials[i]=d3dxMaterials[i].MatD3D; // Copy the material
		pMeshMaterials[i].Ambient = pMeshMaterials[i].Diffuse;
		pMeshTextures[i]=NULL;
		if(d3dxMaterials[i].pTextureFilename != NULL && lstrlen(d3dxMaterials[i].pTextureFilename)>=0)
			{
			strcpy(objectTextureName,objectDirectory);
			strcat(objectTextureName,d3dxMaterials[i].pTextureFilename);
			if(FAILED(D3DXCreateTextureFromFile(pd3dDeviceObject, 
												objectTextureName,
												&pMeshTextures[i])));
			//	PCN4240 MsgD3DObject("%s %s\n%s",languageD3DObject[15].text, // PCN2473 Language pointer (Antony, 11 May 2004)
			//							 languageD3DObject[16].text, // PCN2473 Language pointer (Antony, 11 May 2004)
			//							 objectTextureName);
			}
		}
//	if(pD3DXMtrlBuffer!=NULL) pD3DXMtrlBuffer->Release();
}

void D3DObject::Draw(void)
{
	D3DXVECTOR3 posOffset;	// How much object has to be moved ralative to camera
	D3DXMATRIX  matTranslate;
	D3DXMATRIX  matScale;
	D3DXMATRIX  matRotationX;
	D3DXMATRIX	matRotationY;
	D3DXMATRIX	matRotationZ;
	D3DXMATRIX  matTransform;
	D3DXMATRIX  mattemp;
//	float camRotX;
//	float camRotY;

	if(objectCam!=NULL && (alignPosition==1)) posOffset=objectCam->position
											 +objectCam->direction*position.z
											 +objectCam->right*position.x
											 +objectCam->up*position.y;

	else posOffset=position;	// If not alligned to camera, take as real world coordinates
	if(objectCam!=NULL && (alignRotation==1))
		{
		rotation.x=(float) acos((double) objectCam->direction.x);
		rotation.y=(float) asin((double) objectCam->direction.y);
		rotation.z=(float) asin((double) objectCam->right.z);
		}

	
	D3DXMatrixRotationZ(&matRotationZ,D3DXToRadian(rotation.z));
	D3DXMatrixRotationX(&matRotationX,D3DXToRadian(rotation.x));
	D3DXMatrixRotationY(&matRotationY,D3DXToRadian(rotation.y));
	
	D3DXMatrixTranslation( &matTranslate, posOffset.x, posOffset.y, posOffset.z);
	D3DXMatrixScaling(&matScale,scale.x, scale.y, scale.z);

	matTransform=matRotationZ;


	D3DXMatrixMultiply(&matTransform,&matTransform,&matRotationX);
	D3DXMatrixMultiply(&matTransform,&matTransform,&matRotationY);
	
	D3DXMatrixMultiply(&matTransform,&matTransform,&matScale);
	D3DXMatrixMultiply(&matTransform,&matTransform,&matTranslate);
	
//	D3DXMatrixMultiply(&matTransform,&matScale, &matTranslate);
//	D3DXMatrixMultiply(&matTransform,&matTransform,&matRotation);

    pd3dDeviceObject->SetTransform( D3DTS_WORLD, &matTransform );

	for(DWORD i=0;i<dwNumMaterials;i++)
		{
		pd3dDeviceObject->SetMaterial(&pMeshMaterials[i]);
		pd3dDeviceObject->SetTexture(0, pMeshTextures[i]);
		pMesh->DrawSubset(i);
		}

	D3DXMatrixScaling(&matScale,1,1,1);
	D3DXMatrixTranslation( &matTranslate, 0, 0, 0);
	D3DXMatrixMultiply(&matTransform, &matScale,&matTranslate);
    pd3dDeviceObject->SetTransform( D3DTS_WORLD, &matTransform );
}

