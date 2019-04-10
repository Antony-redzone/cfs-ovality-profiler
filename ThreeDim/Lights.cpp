#include "Lights.h"

#include <atlbase.h>
#include <stdio.h>

void MsgLights(TCHAR *szFormat, ...)
{
    TCHAR szBuffer[512];

    va_list pArgs;
    va_start(pArgs, szFormat);
    _vstprintf(szBuffer, szFormat, pArgs);
    va_end(pArgs);

    MessageBox(NULL, szBuffer, TEXT("LaserProfiler Message"), MB_OK | MB_ICONERROR);
}

Lights::Lights(void)
{
	ZeroMemory( &light, sizeof(D3DLIGHT9) );
    light.Type       = D3DLIGHT_DIRECTIONAL;
    light.Diffuse.r  = 1.0f;
    light.Diffuse.g  = 1.0f;
    light.Diffuse.b  = 1.0f;
	light.Direction  = D3DXVECTOR3(0,0,1);
	light.Range      = 100000.0f;
}


void Lights::Tilt(float deg)
{
	D3DXMATRIX matRot;
	D3DXVECTOR3 axis;

	axis=D3DXVECTOR3(1,0,0);
	D3DXMatrixRotationAxis( &matRot, &axis, deg);

	D3DXVec3TransformNormal( (D3DXVECTOR3*)&light.Direction, (D3DXVECTOR3*)&light.Direction, &matRot);
}

