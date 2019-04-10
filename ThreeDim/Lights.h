#include <d3dx9.h>

class Lights
{
public:
	D3DLIGHT9 light;
	
	Lights(void);
	void Tilt(float deg);
};