#include <d3dx9.h>
#pragma warning( disable : 4996 )

extern "C" __declspec(dllexport) float Vec2Length ( D3DXVECTOR2* pV )
  {return D3DXVec2Length( pV );}

extern "C" __declspec(dllexport) float Vec2LengthSq ( D3DXVECTOR2* pV )
  {return D3DXVec2LengthSq( pV );}

extern "C" __declspec(dllexport) float Vec2Dot ( D3DXVECTOR2* pV1, D3DXVECTOR2* pV2 )
  {return D3DXVec2Dot( pV1, pV2 );}

extern "C" __declspec(dllexport) float Vec2CCW ( D3DXVECTOR2* pV1, D3DXVECTOR2* pV2 )
  {return D3DXVec2CCW( pV1, pV2 );}

extern "C" __declspec(dllexport) void Vec2Add ( D3DXVECTOR2* pOut, D3DXVECTOR2* pV1, D3DXVECTOR2* pV2 )
  {D3DXVec2Add( pOut, pV1, pV2 );}

extern "C" __declspec(dllexport) void Vec2Subtract ( D3DXVECTOR2* pOut, D3DXVECTOR2* pV1, D3DXVECTOR2* pV2 )
  {D3DXVec2Subtract( pOut, pV1, pV2 );}

extern "C" __declspec(dllexport) void Vec2Minimize ( D3DXVECTOR2* pOut, D3DXVECTOR2* pV1, D3DXVECTOR2* pV2 )
  {D3DXVec2Minimize( pOut, pV1, pV2 );}

extern "C" __declspec(dllexport) void Vec2Maximize ( D3DXVECTOR2* pOut, D3DXVECTOR2* pV1, D3DXVECTOR2* pV2 )
  {D3DXVec2Maximize( pOut, pV1, pV2 );}

extern "C" __declspec(dllexport) void Vec2Scale ( D3DXVECTOR2* pOut, D3DXVECTOR2* pV, FLOAT s )
  {D3DXVec2Scale( pOut, pV, s );}

extern "C" __declspec(dllexport) void Vec2Lerp ( D3DXVECTOR2* pOut, D3DXVECTOR2* pV1, D3DXVECTOR2* pV2, FLOAT s )
  {D3DXVec2Lerp( pOut, pV1, pV2, s );}

extern "C" __declspec(dllexport) void Vec2Normalize ( D3DXVECTOR2* pOut, D3DXVECTOR2* pV )
  {D3DXVec2Normalize( pOut, pV );}

extern "C" __declspec(dllexport) void Vec2Hermite ( D3DXVECTOR2* pOut, D3DXVECTOR2* pV1, D3DXVECTOR2* pT1, D3DXVECTOR2* pV2, D3DXVECTOR2* pT2, FLOAT s )
  {D3DXVec2Hermite( pOut, pV1, pT1, pV2, pT2, s );}

extern "C" __declspec(dllexport) void Vec2CatmullRom ( D3DXVECTOR2* pOut, D3DXVECTOR2* pV0, D3DXVECTOR2* pV1, D3DXVECTOR2* pV2, D3DXVECTOR2* pV3, FLOAT s )
  {D3DXVec2CatmullRom( pOut, pV0, pV1, pV2, pV3, s );}

extern "C" __declspec(dllexport) void Vec2BaryCentric ( D3DXVECTOR2* pOut, D3DXVECTOR2* pV1, D3DXVECTOR2* pV2, D3DXVECTOR2* pV3, FLOAT f, FLOAT g)
  {D3DXVec2BaryCentric( pOut, pV1, pV2, pV3, f, g );}

extern "C" __declspec(dllexport) void Vec2Transform ( D3DXVECTOR4* pOut, D3DXVECTOR2* pV, D3DXMATRIX* pM )
  {D3DXVec2Transform( pOut, pV, pM );}

extern "C" __declspec(dllexport) void Vec2TransformCoord ( D3DXVECTOR2* pOut, D3DXVECTOR2* pV, D3DXMATRIX* pM )
  {D3DXVec2TransformCoord( pOut, pV, pM );}

extern "C" __declspec(dllexport) void Vec2TransformNormal ( D3DXVECTOR2* pOut, D3DXVECTOR2* pV, D3DXMATRIX* pM )
  {D3DXVec2TransformNormal( pOut, pV, pM );}

extern "C" __declspec(dllexport) float Vec3Length ( D3DXVECTOR3* pV )
  {return D3DXVec3Length( pV );}

extern "C" __declspec(dllexport) float Vec3LengthSq ( D3DXVECTOR3* pV )
  {return D3DXVec3LengthSq( pV );}

extern "C" __declspec(dllexport) float Vec3Dot ( D3DXVECTOR3* pV1, D3DXVECTOR3* pV2 )
  {return D3DXVec3Dot( pV1, pV2 );}

extern "C" __declspec(dllexport) void Vec3Cross ( D3DXVECTOR3* pOut, D3DXVECTOR3* pV1, D3DXVECTOR3* pV2 )
  {D3DXVec3Cross( pOut, pV1, pV2 );}

extern "C" __declspec(dllexport) void Vec3Add ( D3DXVECTOR3* pOut, D3DXVECTOR3* pV1, D3DXVECTOR3* pV2 )
  {D3DXVec3Add( pOut, pV1, pV2 );}

extern "C" __declspec(dllexport) void Vec3Subtract ( D3DXVECTOR3* pOut, D3DXVECTOR3* pV1, D3DXVECTOR3* pV2 )
  {D3DXVec3Subtract( pOut, pV1, pV2 );}

extern "C" __declspec(dllexport) void Vec3Minimize ( D3DXVECTOR3* pOut, D3DXVECTOR3* pV1, D3DXVECTOR3* pV2 )
  {D3DXVec3Minimize( pOut, pV1, pV2 );}

extern "C" __declspec(dllexport) void Vec3Maximize ( D3DXVECTOR3* pOut, D3DXVECTOR3* pV1, D3DXVECTOR3* pV2 )
  {D3DXVec3Maximize( pOut, pV1, pV2 );}

extern "C" __declspec(dllexport) void Vec3Scale ( D3DXVECTOR3* pOut, D3DXVECTOR3* pV, FLOAT s)
  {D3DXVec3Scale( pOut, pV, s );}

extern "C" __declspec(dllexport) void Vec3Lerp ( D3DXVECTOR3* pOut, D3DXVECTOR3* pV1, D3DXVECTOR3* pV2, FLOAT s )
  {D3DXVec3Lerp( pOut, pV1, pV2, s );}

extern "C" __declspec(dllexport) void Vec3Normalize ( D3DXVECTOR3* pOut, D3DXVECTOR3* pV )
  {D3DXVec3Normalize( pOut, pV );}

extern "C" __declspec(dllexport) void Vec3Hermite ( D3DXVECTOR3* pOut, D3DXVECTOR3* pV1, D3DXVECTOR3* pT1, D3DXVECTOR3* pV2, D3DXVECTOR3* pT2, FLOAT s )
  {D3DXVec3Hermite( pOut, pV1, pT1, pV2, pT2, s );}

extern "C" __declspec(dllexport) void Vec3CatmullRom ( D3DXVECTOR3* pOut, D3DXVECTOR3* pV0, D3DXVECTOR3* pV1, D3DXVECTOR3* pV2, D3DXVECTOR3* pV3, FLOAT s )
  {D3DXVec3CatmullRom( pOut, pV0, pV1, pV2, pV3, s );}

extern "C" __declspec(dllexport) void Vec3BaryCentric ( D3DXVECTOR3* pOut, D3DXVECTOR3* pV1, D3DXVECTOR3* pV2, D3DXVECTOR3* pV3, FLOAT f, FLOAT g)
  {D3DXVec3BaryCentric( pOut, pV1, pV2, pV3, f, g );}

extern "C" __declspec(dllexport) void Vec3Transform ( D3DXVECTOR4* pOut, D3DXVECTOR3* pV, D3DXMATRIX* pM )
  {D3DXVec3Transform( pOut, pV, pM );}

extern "C" __declspec(dllexport) void Vec3TransformCoord ( D3DXVECTOR3* pOut, D3DXVECTOR3* pV, D3DXMATRIX* pM )
  {D3DXVec3TransformCoord( pOut, pV, pM );}

extern "C" __declspec(dllexport) void Vec3TransformNormal ( D3DXVECTOR3* pOut, D3DXVECTOR3* pV, D3DXMATRIX* pM )
  {D3DXVec3TransformNormal( pOut, pV, pM );}

extern "C" __declspec(dllexport) void Vec3Project ( D3DXVECTOR3* pOut, D3DXVECTOR3* pV, D3DVIEWPORT9* pViewport, D3DXMATRIX* pProjection, D3DXMATRIX* pView, D3DXMATRIX* pWorld)
  {D3DXVec3Project( pOut, pV, pViewport, pProjection, pView, pWorld );}

extern "C" __declspec(dllexport) void Vec3Unproject ( D3DXVECTOR3* pOut, D3DXVECTOR3* pV, D3DVIEWPORT9* pViewport, D3DXMATRIX* pProjection, D3DXMATRIX* pView, D3DXMATRIX* pWorld)
  {D3DXVec3Unproject( pOut, pV, pViewport, pProjection, pView, pWorld );}

extern "C" __declspec(dllexport) float Vec4Length ( D3DXVECTOR4* pV )
  {return D3DXVec4Length( pV );}

extern "C" __declspec(dllexport) float Vec4LengthSq ( D3DXVECTOR4* pV )
  {return D3DXVec4LengthSq( pV );}

extern "C" __declspec(dllexport) float Vec4Dot ( D3DXVECTOR4* pV1, D3DXVECTOR4* pV2 )
  {return D3DXVec4Dot( pV1, pV2 );}

extern "C" __declspec(dllexport) void Vec4Add ( D3DXVECTOR4* pOut, D3DXVECTOR4* pV1, D3DXVECTOR4* pV2)
  {D3DXVec4Add( pOut, pV1, pV2 );}

extern "C" __declspec(dllexport) void Vec4Subtract ( D3DXVECTOR4* pOut, D3DXVECTOR4* pV1, D3DXVECTOR4* pV2)
  {D3DXVec4Subtract( pOut, pV1, pV2 );}

extern "C" __declspec(dllexport) void Vec4Minimize ( D3DXVECTOR4* pOut, D3DXVECTOR4* pV1, D3DXVECTOR4* pV2)
  {D3DXVec4Minimize( pOut, pV1, pV2 );}

extern "C" __declspec(dllexport) void Vec4Maximize ( D3DXVECTOR4* pOut, D3DXVECTOR4* pV1, D3DXVECTOR4* pV2)
  {D3DXVec4Maximize( pOut, pV1, pV2 );}

extern "C" __declspec(dllexport) void Vec4Scale ( D3DXVECTOR4* pOut, D3DXVECTOR4* pV, FLOAT s)
  {D3DXVec4Scale( pOut, pV, s );}

extern "C" __declspec(dllexport) void Vec4Lerp ( D3DXVECTOR4* pOut, D3DXVECTOR4* pV1, D3DXVECTOR4* pV2, FLOAT s )
  {D3DXVec4Lerp( pOut, pV1, pV2, s );}

extern "C" __declspec(dllexport) void Vec4Cross ( D3DXVECTOR4* pOut, D3DXVECTOR4* pV1, D3DXVECTOR4* pV2, D3DXVECTOR4* pV3)
  {D3DXVec4Cross( pOut, pV1, pV2, pV3 );}

extern "C" __declspec(dllexport) void Vec4Normalize ( D3DXVECTOR4* pOut, D3DXVECTOR4* pV )
  {D3DXVec4Normalize( pOut, pV );}

extern "C" __declspec(dllexport) void Vec4Hermite ( D3DXVECTOR4* pOut, D3DXVECTOR4* pV1, D3DXVECTOR4* pT1, D3DXVECTOR4* pV2, D3DXVECTOR4* pT2, FLOAT s )
  {D3DXVec4Hermite( pOut, pV1, pT1, pV2, pT2, s );}

extern "C" __declspec(dllexport) void Vec4CatmullRom ( D3DXVECTOR4* pOut, D3DXVECTOR4* pV0, D3DXVECTOR4* pV1, D3DXVECTOR4* pV2, D3DXVECTOR4* pV3, FLOAT s )
  {D3DXVec4CatmullRom( pOut, pV0, pV1, pV2, pV3, s );}

extern "C" __declspec(dllexport) void Vec4BaryCentric ( D3DXVECTOR4* pOut, D3DXVECTOR4* pV1, D3DXVECTOR4* pV2, D3DXVECTOR4* pV3, FLOAT f, FLOAT g)
  {D3DXVec4BaryCentric( pOut, pV1, pV2, pV3, f, g );}

extern "C" __declspec(dllexport) void Vec4Transform ( D3DXVECTOR4* pOut, D3DXVECTOR4* pV, D3DXMATRIX* pM )
  {D3DXVec4Transform( pOut, pV, pM );}

extern "C" __declspec(dllexport) void MatrixIdentity ( D3DXMATRIX* pOut )
  {D3DXMatrixIdentity( pOut );}

extern "C" __declspec(dllexport) INT16 MatrixIsIdentity ( D3DXMATRIX* pM )
  {return -(INT16)D3DXMatrixIsIdentity( pM );}

extern "C" __declspec(dllexport) float MatrixDeterminant ( D3DXMATRIX* pM )
  {return D3DXMatrixDeterminant( pM );}

extern "C" __declspec(dllexport) void MatrixTranspose ( D3DXMATRIX* pOut, D3DXMATRIX* pM )
  {D3DXMatrixTranspose( pOut, pM );}

extern "C" __declspec(dllexport) void MatrixMultiply ( D3DXMATRIX* pOut, D3DXMATRIX* pM1, D3DXMATRIX* pM2 )
  {D3DXMatrixMultiply( pOut, pM1, pM2 );}

extern "C" __declspec(dllexport) void MatrixMultiplyTranspose ( D3DXMATRIX* pOut, D3DXMATRIX* pM1, D3DXMATRIX* pM2 )
  {D3DXMatrixMultiplyTranspose( pOut, pM1, pM2 );}

extern "C" __declspec(dllexport) void MatrixInverse ( D3DXMATRIX* pOut, FLOAT* pDeterminant, D3DXMATRIX* pM )
  {D3DXMatrixInverse( pOut, pDeterminant, pM );}

extern "C" __declspec(dllexport) void MatrixScaling ( D3DXMATRIX* pOut, FLOAT sx, FLOAT sy, FLOAT sz )
  {D3DXMatrixScaling( pOut, sx, sy, sz );}

extern "C" __declspec(dllexport) void MatrixTranslation ( D3DXMATRIX* pOut, FLOAT x, FLOAT y, FLOAT z )
  {D3DXMatrixTranslation( pOut, x, y, z );}

extern "C" __declspec(dllexport) void MatrixRotationX ( D3DXMATRIX* pOut, FLOAT Angle )
  {D3DXMatrixRotationX( pOut, Angle );}

extern "C" __declspec(dllexport) void MatrixRotationY ( D3DXMATRIX* pOut, FLOAT Angle )
  {D3DXMatrixRotationY( pOut, Angle );}

extern "C" __declspec(dllexport) void MatrixRotationZ ( D3DXMATRIX* pOut, FLOAT Angle )
  {D3DXMatrixRotationZ( pOut, Angle );}

extern "C" __declspec(dllexport) void MatrixRotationAxis ( D3DXMATRIX* pOut, D3DXVECTOR3* pV, FLOAT Angle )
  {D3DXMatrixRotationAxis( pOut, pV, Angle );}

extern "C" __declspec(dllexport) void MatrixRotationQuaternion ( D3DXMATRIX* pOut, D3DXQUATERNION* pQ)
  {D3DXMatrixRotationQuaternion( pOut, pQ );}

extern "C" __declspec(dllexport) void MatrixRotationYawPitchRoll ( D3DXMATRIX* pOut, FLOAT Yaw, FLOAT Pitch, FLOAT Roll )
  {D3DXMatrixRotationYawPitchRoll( pOut, Yaw, Pitch, Roll );}

extern "C" __declspec(dllexport) void MatrixTransformation ( D3DXMATRIX* pOut, D3DXVECTOR3* pScalingCenter, D3DXQUATERNION* pScalingRotation, D3DXVECTOR3* pScaling, D3DXVECTOR3* pRotationCenter, D3DXQUATERNION* pRotation, D3DXVECTOR3* pTranslation)
  {D3DXMatrixTransformation( pOut, pScalingCenter, pScalingRotation, pScaling, pRotationCenter, pRotation, pTranslation );}

extern "C" __declspec(dllexport) void MatrixAffineTransformation ( D3DXMATRIX* pOut, FLOAT Scaling, D3DXVECTOR3* pRotationCenter, D3DXQUATERNION* pRotation, D3DXVECTOR3* pTranslation)
  {D3DXMatrixAffineTransformation( pOut, Scaling, pRotationCenter, pRotation, pTranslation );}

extern "C" __declspec(dllexport) void MatrixLookAtRH ( D3DXMATRIX* pOut, D3DXVECTOR3* pEye, D3DXVECTOR3* pAt, D3DXVECTOR3* pUp )
  {D3DXMatrixLookAtRH( pOut, pEye, pAt, pUp );}

extern "C" __declspec(dllexport) void MatrixLookAtLH ( D3DXMATRIX* pOut, D3DXVECTOR3* pEye, D3DXVECTOR3* pAt, D3DXVECTOR3* pUp )
  {D3DXMatrixLookAtLH( pOut, pEye, pAt, pUp );}

extern "C" __declspec(dllexport) void MatrixPerspectiveRH ( D3DXMATRIX* pOut, FLOAT w, FLOAT h, FLOAT zn, FLOAT zf )
  {D3DXMatrixPerspectiveRH( pOut, w, h, zn, zf );}

extern "C" __declspec(dllexport) void MatrixPerspectiveLH ( D3DXMATRIX* pOut, FLOAT w, FLOAT h, FLOAT zn, FLOAT zf )
  {D3DXMatrixPerspectiveLH( pOut, w, h, zn, zf );}

extern "C" __declspec(dllexport) void MatrixPerspectiveFovRH ( D3DXMATRIX* pOut, FLOAT fovy, FLOAT Aspect, FLOAT zn, FLOAT zf )
  {D3DXMatrixPerspectiveFovRH( pOut, fovy, Aspect, zn, zf );}

extern "C" __declspec(dllexport) void MatrixPerspectiveFovLH ( D3DXMATRIX* pOut, FLOAT fovy, FLOAT Aspect, FLOAT zn, FLOAT zf )
  {D3DXMatrixPerspectiveFovLH( pOut, fovy, Aspect, zn, zf );}

extern "C" __declspec(dllexport) void MatrixPerspectiveOffCenterRH ( D3DXMATRIX* pOut, FLOAT l, FLOAT r, FLOAT b, FLOAT t, FLOAT zn, FLOAT zf )
  {D3DXMatrixPerspectiveOffCenterRH( pOut, l, r, b, t, zn, zf );}

extern "C" __declspec(dllexport) void MatrixPerspectiveOffCenterLH ( D3DXMATRIX* pOut, FLOAT l, FLOAT r, FLOAT b, FLOAT t, FLOAT zn, FLOAT zf )
  {D3DXMatrixPerspectiveOffCenterLH( pOut, l, r, b, t, zn, zf );}

extern "C" __declspec(dllexport) void MatrixOrthoRH ( D3DXMATRIX* pOut, FLOAT w, FLOAT h, FLOAT zn, FLOAT zf )
  {D3DXMatrixOrthoRH( pOut, w, h, zn, zf );}

extern "C" __declspec(dllexport) void MatrixOrthoLH ( D3DXMATRIX* pOut, FLOAT w, FLOAT h, FLOAT zn, FLOAT zf )
  {D3DXMatrixOrthoLH( pOut, w, h, zn, zf );}

extern "C" __declspec(dllexport) void MatrixOrthoOffCenterRH ( D3DXMATRIX* pOut, FLOAT l, FLOAT r, FLOAT b, FLOAT t, FLOAT zn, FLOAT zf )
  {D3DXMatrixOrthoOffCenterRH( pOut, l, r, b, t, zn, zf );}

extern "C" __declspec(dllexport) void MatrixOrthoOffCenterLH ( D3DXMATRIX* pOut, FLOAT l, FLOAT r, FLOAT b, FLOAT t, FLOAT zn, FLOAT zf )
  {D3DXMatrixOrthoOffCenterLH( pOut, l, r, b, t, zn, zf );}

extern "C" __declspec(dllexport) void MatrixShadow ( D3DXMATRIX* pOut, D3DXVECTOR4* pLight, D3DXPLANE* pPlane )
  {D3DXMatrixShadow( pOut, pLight, pPlane );}

extern "C" __declspec(dllexport) void MatrixReflect ( D3DXMATRIX* pOut, D3DXPLANE* pPlane )
  {D3DXMatrixReflect( pOut, pPlane );}

extern "C" __declspec(dllexport) float QuaternionLength ( D3DXQUATERNION* pQ )
  {return D3DXQuaternionLength( pQ );}

extern "C" __declspec(dllexport) float QuaternionLengthSq ( D3DXQUATERNION* pQ )
  {return D3DXQuaternionLengthSq( pQ );}

extern "C" __declspec(dllexport) float QuaternionDot ( D3DXQUATERNION* pQ1, D3DXQUATERNION* pQ2 )
  {return D3DXQuaternionDot( pQ1, pQ2 );}

extern "C" __declspec(dllexport) void QuaternionIdentity ( D3DXQUATERNION* pOut )
  {D3DXQuaternionIdentity( pOut );}

extern "C" __declspec(dllexport) INT16 QuaternionIsIdentity ( D3DXQUATERNION* pQ )
  {return -(INT16)D3DXQuaternionIsIdentity( pQ );}

extern "C" __declspec(dllexport) void QuaternionConjugate ( D3DXQUATERNION* pOut, D3DXQUATERNION* pQ )
  {D3DXQuaternionConjugate( pOut, pQ );}

extern "C" __declspec(dllexport) void QuaternionToAxisAngle ( D3DXQUATERNION* pQ, D3DXVECTOR3* pAxis, FLOAT* pAngle )
  {D3DXQuaternionToAxisAngle( pQ, pAxis, pAngle );}

extern "C" __declspec(dllexport) void QuaternionRotationMatrix ( D3DXQUATERNION* pOut, D3DXMATRIX* pM)
  {D3DXQuaternionRotationMatrix( pOut, pM );}

extern "C" __declspec(dllexport) void QuaternionRotationAxis ( D3DXQUATERNION* pOut, D3DXVECTOR3* pV, FLOAT Angle )
  {D3DXQuaternionRotationAxis( pOut, pV, Angle );}

extern "C" __declspec(dllexport) void QuaternionRotationYawPitchRoll ( D3DXQUATERNION* pOut, FLOAT Yaw, FLOAT Pitch, FLOAT Roll )
  {D3DXQuaternionRotationYawPitchRoll( pOut, Yaw, Pitch, Roll );}

extern "C" __declspec(dllexport) void QuaternionMultiply ( D3DXQUATERNION* pOut, D3DXQUATERNION* pQ1, D3DXQUATERNION* pQ2 )
  {D3DXQuaternionMultiply( pOut, pQ1, pQ2 );}

extern "C" __declspec(dllexport) void QuaternionNormalize ( D3DXQUATERNION* pOut, D3DXQUATERNION* pQ )
  {D3DXQuaternionNormalize( pOut, pQ );}

extern "C" __declspec(dllexport) void QuaternionInverse ( D3DXQUATERNION* pOut, D3DXQUATERNION* pQ )
  {D3DXQuaternionInverse( pOut, pQ );}

extern "C" __declspec(dllexport) void QuaternionLn ( D3DXQUATERNION* pOut, D3DXQUATERNION* pQ )
  {D3DXQuaternionLn( pOut, pQ );}

extern "C" __declspec(dllexport) void QuaternionExp ( D3DXQUATERNION* pOut, D3DXQUATERNION* pQ )
  {D3DXQuaternionExp( pOut, pQ );}

extern "C" __declspec(dllexport) void QuaternionSlerp ( D3DXQUATERNION* pOut, D3DXQUATERNION* pQ1, D3DXQUATERNION* pQ2, FLOAT t )
  {D3DXQuaternionSlerp( pOut, pQ1, pQ2, t );}

extern "C" __declspec(dllexport) void QuaternionSquad ( D3DXQUATERNION* pOut, D3DXQUATERNION* pQ1, D3DXQUATERNION* pA, D3DXQUATERNION* pB, D3DXQUATERNION* pC, FLOAT t )
  {D3DXQuaternionSquad( pOut, pQ1, pA, pB, pC, t );}

extern "C" __declspec(dllexport) void QuaternionSquadSetup ( D3DXQUATERNION* pAOut, D3DXQUATERNION* pBOut, D3DXQUATERNION* pCOut, D3DXQUATERNION* pQ0, D3DXQUATERNION* pQ1, D3DXQUATERNION* pQ2, D3DXQUATERNION* pQ3 )
  {D3DXQuaternionSquadSetup( pAOut, pBOut, pCOut, pQ0, pQ1, pQ2, pQ3 );}

extern "C" __declspec(dllexport) void QuaternionBaryCentric ( D3DXQUATERNION* pOut, D3DXQUATERNION* pQ1, D3DXQUATERNION* pQ2, D3DXQUATERNION* pQ3, FLOAT f, FLOAT g )
  {D3DXQuaternionBaryCentric( pOut, pQ1, pQ2, pQ3, f, g );}

extern "C" __declspec(dllexport) float PlaneDot ( D3DXPLANE* pP, D3DXVECTOR4* pV)
  {return D3DXPlaneDot( pP, pV );}

extern "C" __declspec(dllexport) float PlaneDotCoord ( D3DXPLANE* pP, D3DXVECTOR3* pV)
  {return D3DXPlaneDotCoord( pP, pV );}

extern "C" __declspec(dllexport) float PlaneDotNormal ( D3DXPLANE* pP, D3DXVECTOR3* pV)
  {return D3DXPlaneDotNormal( pP, pV );}

extern "C" __declspec(dllexport) void PlaneNormalize ( D3DXPLANE* pOut, D3DXPLANE* pP)
  {D3DXPlaneNormalize( pOut, pP );}

extern "C" __declspec(dllexport) void PlaneIntersectLine ( D3DXVECTOR3* pOut, D3DXPLANE* pP, D3DXVECTOR3* pV1, D3DXVECTOR3* pV2)
  {D3DXPlaneIntersectLine( pOut, pP, pV1, pV2 );}

extern "C" __declspec(dllexport) void PlaneFromPointNormal ( D3DXPLANE* pOut, D3DXVECTOR3* pPoint, D3DXVECTOR3* pNormal)
  {D3DXPlaneFromPointNormal( pOut, pPoint, pNormal );}

extern "C" __declspec(dllexport) void PlaneFromPoints ( D3DXPLANE* pOut, D3DXVECTOR3* pV1, D3DXVECTOR3* pV2, D3DXVECTOR3* pV3)
  {D3DXPlaneFromPoints( pOut, pV1, pV2, pV3 );}

extern "C" __declspec(dllexport) void PlaneTransform ( D3DXPLANE* pOut, D3DXPLANE* pP, D3DXMATRIX* pM )
  {D3DXPlaneTransform( pOut, pP, pM );}

extern "C" __declspec(dllexport) void ColorNegative (D3DXCOLOR* pOut, D3DXCOLOR* pC)
  {D3DXColorNegative( pOut, pC );}

extern "C" __declspec(dllexport) void ColorAdd (D3DXCOLOR* pOut, D3DXCOLOR* pC1, D3DXCOLOR* pC2)
  {D3DXColorAdd( pOut, pC1, pC2 );}

extern "C" __declspec(dllexport) void ColorSubtract (D3DXCOLOR* pOut, D3DXCOLOR* pC1, D3DXCOLOR* pC2)
  {D3DXColorSubtract( pOut, pC1, pC2 );}

extern "C" __declspec(dllexport) void ColorScale (D3DXCOLOR* pOut, D3DXCOLOR* pC, FLOAT s)
  {D3DXColorScale( pOut, pC, s );}

extern "C" __declspec(dllexport) void ColorModulate (D3DXCOLOR* pOut, D3DXCOLOR* pC1, D3DXCOLOR* pC2)
  {D3DXColorModulate( pOut, pC1, pC2 );}

extern "C" __declspec(dllexport) void ColorLerp (D3DXCOLOR* pOut, D3DXCOLOR* pC1, D3DXCOLOR* pC2, FLOAT s)
  {D3DXColorLerp( pOut, pC1, pC2, s );}

extern "C" __declspec(dllexport) void ColorAdjustSaturation (D3DXCOLOR* pOut, D3DXCOLOR* pC, FLOAT s)
  {D3DXColorAdjustSaturation( pOut, pC, s );}

extern "C" __declspec(dllexport) void ColorAdjustContrast (D3DXCOLOR* pOut, D3DXCOLOR* pC, FLOAT c)
  {D3DXColorAdjustContrast( pOut, pC, c );}


extern "C" __declspec(dllexport) int ARGB_ (int a, int r, int g, int b)
  {return ((a & 0xff)<<24)|((r & 0xff)<<16)|((g & 0xff)<<8)|(b & 0xff);}

extern "C" __declspec(dllexport) D3DCOLORVALUE D3DColorValue_ (UINT c)
  {
    D3DCOLORVALUE v;
    v.b = (float)(c & 0xff);
    v.g = (float)((c >> 8) & 0xff);
    v.r = (float)((c >> 16) & 0xff);
    v.a = (float)(c >> 24);
    return v;
  }
