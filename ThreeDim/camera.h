#include <D3DX9.h>

class Camera
{
public:
	~Camera(); // PCN2465 (Antony van Iersel, 5 December 2003) remove
	D3DXVECTOR3 test;
	D3DXVECTOR3 position;
	D3DXVECTOR3 target;
	D3DXVECTOR3 up;
	D3DXVECTOR3 direction;
	D3DXVECTOR3 motion;
	D3DXVECTOR3 right;
	D3DXVECTOR3 worldUp;
	D3DXVECTOR3 levelUp;
	D3DXVECTOR3 levelDirection;
	D3DXVECTOR3 *rail;
	float distance;
	long noRails;
	long railTarget;
	float testDist;
	float speed;
	bool TTL; // Tilt Target Limit, if True - Limits Viewing angle of Camera



	Camera(void);
	void Yaw(float deg);
	void Pan(float deg);
	void Tilt(float deg);
	void PanTarget(float deg);
	void TiltTarget(float deg);
	void YawTarget(float deg);
	void YawLevelTarget(float deg);
	void RollTargetZ(float deg);
	void RollTargetY(float deg);
	void MoveFoward(float dis);
	void MoveMotion(float dis);
	void MoveDirection(float dis);
	void MoveHeight(float dis);
	void MoveStrafe(float dis);
	void MoveRail(float dis);
	void MoveRailTarget(float dis);
	void MoveTo(float x, float y, float z);
	void MoveTo(D3DXVECTOR3 d) {MoveTo(d.x, d.y, d.z);}
	void MoveToRail(long r);
	void MoveToRailTarget(long r);
	void ZoomTarget(float dis);
	void Zoom(float dis);
	bool TiltTargetLimit(float deg);
	void Reset(void);

		
	void InitRail(long np) {rail = new D3DXVECTOR3[np+100]; noRails=np; }
	void SetMotion(void);
	void SetNormal(D3DXVECTOR3 &des, D3DXVECTOR3 pos, D3DXVECTOR3 trg);
	
};
