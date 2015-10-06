Attribute VB_Name = "dCurveFittingModule"

Private Function BezierInterpolation(iRatio As Single, v1 As TV_3DVECTOR, v2 As TV_3DVECTOR, v3 As TV_3DVECTOR) As TV_3DVECTOR
'二次贝塞尔插值
If iRatio < 0 Or iRatio > 1 Then GoTo ex:

Dim a As Single, b As Single, c As Single '系数
a = (1 - iRatio) ^ 2
b = 2 * (1 - iRatio) * iRatio
c = iRatio ^ 2

Dim vec1 As TV_3DVECTOR, vec2 As TV_3DVECTOR, vec3 As TV_3DVECTOR
vec1 = Math.VScale(v1, a)
vec2 = Math.VScale(v2, b)
vec3 = Math.VScale(v3, c)

BezierInterpolation = Math.vAdd(Math.vAdd(vec1, vec2), vec3)

ex:
End Function



Public Sub BezierCurveFitting()

Dim x As Single, y As Single, z As Single, nx As Single, ny As Single, nz As Single, tu1 As Single, tv1 As Single, tu2 As Single, tv2 As Single, c As Long

'Dim MiddlePoint1 As TV_3DVECTOR, MiddlePoint2 As TV_3DVECTOR
Dim OriginPoint1 As TV_3DVECTOR, OriginPoint2 As TV_3DVECTOR, OriginPoint3 As TV_3DVECTOR
Dim ResultPoint As TV_3DVECTOR

For i = 1 To NumOfVerticalLines + 1

       For j = 2 To NumOfPointVerticalLine(i) - 1 '只平滑第二到倒数第二个点
       
       mesh.GetVertex GetPointID(Val(i), Val(j - 1)), x, y, z, nx, ny, nz, tu1, tv1, tu2, tv2, c
       OriginPoint1 = Vector3(x, y, z)
       mesh.GetVertex GetPointID(Val(i), Val(j)), x, y, z, nx, ny, nz, tu1, tv1, tu2, tv2, c
       OriginPoint2 = Vector3(x, y, z)
       mesh.GetVertex GetPointID(Val(i), Val(j + 1)), x, y, z, nx, ny, nz, tu1, tv1, tu2, tv2, c
       OriginPoint3 = Vector3(x, y, z)
       
       ResultPoint = BezierInterpolation(0.5, OriginPoint1, OriginPoint2, OriginPoint3)
       
       mesh.SetVertex GetPointID(Val(i), Val(j)), ResultPoint.x, ResultPoint.y, ResultPoint.z, nx, ny, nz, tu1, tv1, tu2, tv2, c
       Next j
       
Next i

End Sub
