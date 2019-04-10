#include "Camera.h"

Camera::Camera(void)
{
	position	   = D3DXVECTOR3(0	,0  ,100);
	target		   = D3DXVECTOR3(0	,0  ,0);
	up			   = D3DXVECTOR3(0  ,1  ,0);
	levelUp		   = D3DXVECTOR3(0  ,1  ,0);
	direction	   = D3DXVECTOR3(0  ,0  ,-1);
	right		   = D3DXVECTOR3(-1	,0 	,0);
	worldUp		   = D3DXVECTOR3(0	,1	,0);
	levelDirection = D3DXVECTOR3(0	,0	,-1);
	motion		   = D3DXVECTOR3(0	,0	,-1);
	distance=100;
	speed=0;
	TTL=false;
	rail=NULL;
}

Camera::~Camera()					//PCN2465 (Antony) remove memory leaks 5 December 2003 
{									//
	if(rail!=NULL) { delete rail; rail = NULL; } 	// PCN3085
}									//

void Camera::Reset(void)
{
	position	   = D3DXVECTOR3(0	,0  ,100);
	target		   = D3DXVECTOR3(0	,0  ,0);
	up			   = D3DXVECTOR3(0  ,1  ,0);
	levelUp		   = D3DXVECTOR3(0  ,1  ,0);
	direction	   = D3DXVECTOR3(0  ,0  ,-1);
	right		   = D3DXVECTOR3(-1	,0 	,0);
	worldUp		   = D3DXVECTOR3(0	,1	,0);
	levelDirection = D3DXVECTOR3(0	,0	,-1);
	motion		   = D3DXVECTOR3(0	,0	,-1);
	distance=100;
	speed=0;
}

void Camera::Yaw(float deg)
{
	D3DXMATRIX matRot;
	D3DXMatrixRotationAxis(&matRot, &up, D3DXToRadian(deg));
	D3DXVec3TransformNormal(&direction,		 &direction,	  &matRot);
	D3DXVec3TransformNormal(&right,			 &right,		  &matRot);
	D3DXVec3TransformNormal(&levelDirection, &levelDirection, &matRot);
	target = position + (direction * distance);
}

void Camera::Pan(float deg)
{
	D3DXMATRIX matRot;
	D3DXMatrixRotationAxis(&matRot, &worldUp, D3DXToRadian(deg));
	D3DXVec3TransformNormal(&direction,		 &direction,	  &matRot);
	D3DXVec3TransformNormal(&right,			 &right,		  &matRot);
	D3DXVec3TransformNormal(&up,			 &up,			  &matRot);
	D3DXVec3TransformNormal(&levelDirection, &levelDirection, &matRot);
	target = position + (direction * distance);
}

void Camera::PanTarget(float deg)
{
	D3DXMATRIX matRot;
	D3DXMatrixRotationAxis(&matRot, &worldUp, D3DXToRadian(deg));
	D3DXVec3TransformNormal(&direction,		 &direction,	  &matRot);
	D3DXVec3TransformNormal(&right,			 &right,		  &matRot);
	D3DXVec3TransformNormal(&up,			 &up,			  &matRot);
	D3DXVec3TransformNormal(&levelDirection, &levelDirection, &matRot);
	position = target - (direction * distance);
}

void Camera::YawTarget(float deg)
{
	D3DXMATRIX matRot;
	D3DXMatrixRotationAxis(&matRot, &up, D3DXToRadian(deg));
	D3DXVec3TransformNormal(&direction,		 &direction,	  &matRot);
	D3DXVec3TransformNormal(&right,			 &right,		  &matRot);
	D3DXVec3TransformNormal(&levelUp,		 &levelUp,		  &matRot);
	D3DXVec3TransformNormal(&levelDirection, &levelDirection, &matRot);
	position = target - (direction * distance);
}

void Camera::YawLevelTarget(float deg)
{
	D3DXMATRIX matRot;
	D3DXMatrixRotationAxis(&matRot, &levelUp, D3DXToRadian(deg));
	D3DXVec3TransformNormal(&direction,		 &direction,	  &matRot);
	D3DXVec3TransformNormal(&right,			 &right,		  &matRot);
	D3DXVec3TransformNormal(&levelDirection, &levelDirection, &matRot);
	position = target - (direction * distance);
}

void Camera::Tilt(float deg)
{
	D3DXMATRIX matRot;
	D3DXMatrixRotationAxis(&matRot, &right, D3DXToRadian(deg));
	D3DXVec3TransformNormal(&direction, &direction, &matRot);
	D3DXVec3TransformNormal(&up, &up, &matRot);
	target = position + (direction * distance);
}

void Camera::TiltTarget(float deg)
{
	if(TTL && TiltTargetLimit(deg)) return;
	D3DXMATRIX matRot;
	D3DXMatrixRotationAxis(&matRot, &right, D3DXToRadian(deg));
	D3DXVec3TransformNormal(&direction, &direction, &matRot);
	D3DXVec3TransformNormal(&up, &up, &matRot);

	position = target - (direction * distance);
}
bool Camera::TiltTargetLimit(float deg)
{


/*
	double y1, y2;
	double angle;
	
	y1=levelUp.y;

	y2=direction.y;

	if(x1<0) x1*=-1;
	if(y1<0) y1*=-1;
	if(x1==0) angle=D3DX_PI/2;
	else angle=atan(y1/x1);
	
	if(x2<0) x2*=-1;
	if(y2<0) y2*=-1;
	
	x2=x2*sin(-angle);
	y2=y2*cos(-angle);

	if((deg<0) && (x2<0.133)) return true;
	if((deg>0) && (x2>0.866)) return true;
*/
  return false;

}

void Camera::RollTargetZ(float deg)
{
	D3DXMATRIX matRot;
	D3DXMATRIX matInvRot;
	D3DXMatrixRotationZ(&matRot,    D3DXToRadian(deg));
	D3DXMatrixRotationZ(&matInvRot, D3DXToRadian(deg));
	D3DXVec3TransformNormal(&direction, &direction, &matRot);
	D3DXVec3TransformNormal(&up, &up, &matRot);
	D3DXVec3TransformNormal(&levelUp, &levelUp, &matInvRot);
	D3DXVec3TransformNormal(&right, &right, &matRot);
	position = target - (direction * distance);
}

void Camera::RollTargetY(float deg)
{
	D3DXMATRIX matRot;

	D3DXMatrixRotationAxis(&matRot, &levelUp,  D3DXToRadian(deg));
	D3DXVec3TransformNormal(&direction, &direction, &matRot);
	D3DXVec3TransformNormal(&up, &up, &matRot);
//	D3DXVec3TransformNormal(&levelUp, &levelUp, &matRot);
	D3DXVec3TransformNormal(&right, &right, &matRot);
	position = target - (direction * distance);
}


void Camera::MoveFoward(float dis)
{
	position = position + (levelDirection * dis);
	target   = target   + (levelDirection * dis);
}

void Camera::MoveDirection(float dir)
{
	position = position + (direction * dir);
	target   = target   + (direction * dir);
}

void Camera::MoveMotion(float dir)
{
	position = position + (motion * dir);
	target   = target   + (motion * dir);
}

void Camera::MoveHeight(float dis)
{
	position.y+=dis;
	target.y+=dis;
}

void Camera::MoveStrafe(float dis)
{
	position = position + (right * dis);
	target   = target   + (right * dis);

}

void Camera::MoveRail(float dis)
{
	if(rail==NULL) return; //PCN2461 (8 December 2003, Antony)
	float disTo;
	float disFrom;
	
	if(dis>0) 
		{
		if(railTarget>=noRails) {speed=0;return;}
		MoveMotion(dis);
		disTo=D3DXVec3Length(&(position - rail[railTarget]));
		if(disTo<speed) railTarget++;
		testDist=disTo;
		SetNormal(motion, position, rail[railTarget]);
		}
	if(dis<0) 
		{
		if(railTarget<=0) {speed=0;return;}
		MoveMotion(dis);
		disFrom=D3DXVec3Length(&(position-rail[railTarget-1]));
		if(disFrom<-(speed)) railTarget--; 
		testDist=disFrom;
		SetNormal(motion,  rail[railTarget-1],position );
		}
}

void Camera::MoveRailTarget(float dis)
{
	if(rail==NULL) return; //PCN2461 (8 December 2003, Antony)
	float disTo;
	float disFrom;
	
	if(dis>0) 
		{
		if(railTarget>=noRails) {speed=0;return;}
		MoveMotion(dis);
		disTo=D3DXVec3Length(&(target - rail[railTarget]));
		if(disTo<speed) railTarget++;
		testDist=disTo;
		SetNormal(motion, target, rail[railTarget]);
		}
	if(dis<0) 
		{
		if(railTarget<=0) {speed=0;return;}
		MoveMotion(dis);
		disFrom=D3DXVec3Length(&(target - rail[railTarget-1]));
		if(disFrom<-(speed)) railTarget--; 
		testDist=disFrom;
		SetNormal(motion,  rail[railTarget-1],target);
		}
}

void Camera::MoveTo(float x, float y, float z)
{
	position = D3DXVECTOR3(x ,y ,z);
	target = position + (direction * distance);
}

void Camera::MoveToRail(long r)
{
	if(rail==NULL) return; //PCN2461 (8 December 2003, Antony)
	if(r>=noRails) r=noRails-1;
	if(r<0) r=0;
	MoveTo(rail[r][0], rail[r][1], rail[r][2]);
	railTarget=r+1;
	SetNormal(motion, position, rail[railTarget]);
}

void Camera::MoveToRailTarget(long r)
{
	if(rail==NULL) return; //PCN2461 (8 December 2003, Antony)
	if(r>noRails) r=noRails-1;
	if(r<0) r=0;
	MoveTo(rail[r][0], rail[r][1], rail[r][2]);
	railTarget=r+1;
	SetNormal(motion,position,rail[railTarget]);
	MoveDirection(-distance);
}
void Camera::ZoomTarget(float dis)
{
	if((distance+dis)<5) return;	
	if((distance-+dis) > 1000) return;
	distance=distance+dis;
	target = position + (direction * distance);

	
}

void Camera::Zoom(float dis)
{
	if((distance+dis)<5) return;
	distance=distance+dis;
	position = target - (direction * distance);
}	

void Camera::SetNormal(D3DXVECTOR3 &des, D3DXVECTOR3 pos, D3DXVECTOR3 trg)
{
	des = trg - pos;
	D3DXVec3Normalize(&des, &des);	
} 
	