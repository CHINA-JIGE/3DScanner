VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Scan3DProcessingPipeline2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'——————加载图片————————
Private TheFirstPicturePath As String '第一张图的路径

Private Type Type_ScanCamPerPicture
        ScanCamPos As TV_3DVECTOR
        ScanCamVector As TV_3DVECTOR
        ScanCamEulerX As Single
        ScanCamEulerY As Single
        ScanCamEulerZ As Single
End Type
Private mScanCamPerPicture() As Type_ScanCamPerPicture

'——————1采样————————
Public Enum CONST_LightSide '激光在CAM左还是右
Side_Left = 0
Side_Right = 1
End Enum

Public Enum CONST_Sample
Sample_MainColorFilter = 0
Sample_SingleColorChannel = 1
End Enum


Public Enum CONST_Reconstruction
RC_MappedBall = 0
End Enum

Public Enum CONST_MatrixType
Matrix_OrthoBasis = 0
Matrix_Euler = 1
End Enum


Public Type Type_ScanParameters
StandardDepth As Single '标准深度 焦点到背景板
VisibleAngleHorizontal As Single '可视角θ横
VisibleAngleVertical As Single '可视角θ竖
CamToLight As Single '镜头到红线的距离
ColorFilter As Single '颜色筛选下界
IdentificationLineLength As Single
End Type


'标识点
Private pointO As TV_2DVECTOR, pointA As TV_2DVECTOR, pointB As TV_2DVECTOR, pointC As TV_2DVECTOR





'————————————————————————————————
Sub SetMainBuffer(NewMeshBuffer As Scan3DMeshBuffer2)
Set MB2 = NewMeshBuffer
End Sub

Function GetMainBuffer() As Scan3DMeshBuffer2
Set GetMainBuffer = MB2
End Function



Sub LoadScanPicture(iPath As String, PictureID As Long) '加载图片
       If PictureID = 1 Then TheFirstPicturePath = iPath '记录下第一张图片的路径
       TexF.LoadTexture iPath, CStr(PictureID), PictureWidth, PictureHeight
       TexF.LockTexture GetTex(CStr(PictureID)), True
End Sub

Sub ClearPictureBuffer()
TexF.DeleteAllTextures
End Sub
Sub ClearMainMeshBuffer()
MB2.ClearTriangleBuffer
End Sub
Sub ClearMainPointCloudBuffer()
MB2.ClearPoint3DBuffer
End Sub
Sub ClearMainBuffer()
MB2.ClearBuffer
End Sub
Sub AddPointCloudToRenderMesh()
MB2.AddPointCloudToRenderMesh
End Sub
Sub AddTriangleToRenderMesh()
MB2.AddTriangleToRenderMesh
End Sub
Sub WeldVertices_RenderMesh()
MB2.WeldVertices
End Sub

Sub ResetRenderMesh()
GetMesh("POINT1").ResetMesh
GetMesh("TRIANGLE1").ResetMesh
End Sub










'——————————采样——————————
'流程：根据图片ID（1 to N）加载GETTEX（i），然后
'添加采样点SAMPLEPOINT,再ADDLINE
'——————————————————————
Sub SampleFromPicture(iLightSide As CONST_LightSide, PictureID As Long)

        Dim PixelTraversalStartX As Long, PixelTraversalEndX As Long '遍历像素的横坐标区间
        Dim NumberOfPoint  As Single, TotalPx As Single, RED As Single
        Dim PixelCOLOR As Long, MainColor As Single, Color1 As Single, Color2 As Single
       Dim tmpTransitionX(1 To 10) As Long '颜色转折点的X
       Dim tmpTransitionCount As Integer '一行有多少个转折点
       Dim PrevQualified As Boolean, CurrQualified As Boolean
       Dim SamplingCountPerPic As Long
       
        TotalPx = 0
        NumberOfPoint = 0
       'MaincolorLOWERBOUND = MainColorFilterLowerBound '记录下下界 LOAD_PICTUREEND要用
       
        '判断灯光在CAM的左还是右
        If iLightSide = Side_Left Then
                PixelTraversalStartX = 1
                PixelTraversalEndX = PictureWidth / 2
        Else
                PixelTraversalStartX = PictureWidth / 2
                PixelTraversalEndX = PictureWidth
        End If
       
       Dim p As Long
       If PictureID = -1 Then '-1代表最后一张图片
       p = GetTex("1")
       Else
       p = GetTex(PictureID) '1 到 N
       End If
       
       '对一张图片进行遍历
       Dim x As Long, y As Long, pcolor As TV_COLOR, i As Long 'i是遍历用的下标
       Dim tmpX As Long
              For y = 1 To PictureHeight
                        For x = PixelTraversalStartX To PixelTraversalEndX
                                   
                                PixelCOLOR = TexF.GetPixel(p, x, y)
                                pcolor = DecodeRGBA(PixelCOLOR)
                                MainColor = pcolor.r  '主颜色
                                Color1 = pcolor.g '
                                Color2 = pcolor.b

                               PrevQualified = CurrQualified '相邻两点是否入选的bool更新
                               If MainColor > MainColorFilterLowerBound And Color1 < 0.1 And Color2 < 0.1 Then '激光颜色太坑爹了
                                   CurrQualified = True '当前点入选了！
                                   'TotalPx = TotalPx + x * MainColor
                                   'NumberOfPoint = NumberOfPoint + MainColor
                               Else
                                   CurrQualified = False
                               End If
                               
                               '如果这是转折点
                               If PrevQualified = False And CurrQualified = True Then
                                   tmpTransitionCount = tmpTransitionCount + 1
                                   '就记下他的x坐标
                                   tmpTransitionX(tmpTransitionCount) = x - 5
                               End If
                               If PrevQualified = True And CurrQualified = False Then
                                   tmpTransitionCount = tmpTransitionCount + 1
                                   '就记下他的x坐标
                                   tmpTransitionX(tmpTransitionCount) = x + 5
                               End If
                        Next x
                     
              '遍历完一行后，如果有转折点
              If tmpTransitionCount > 0 And tmpTransitionCount Mod 2 = 0 Then
                     For i = 1 To tmpTransitionCount Step 2
                            For tmpX = tmpTransitionX(i) To tmpTransitionX(i + 1) '每个采样区间
                                   PixelCOLOR = TexF.GetPixel(p, tmpX, y)
                                   pcolor = DecodeRGBA(PixelCOLOR)
                                   MainColor = pcolor.r
                                   TotalPx = TotalPx + tmpX * MainColor
                                   NumberOfPoint = NumberOfPoint + MainColor
                            Next tmpX
                     If NumberOfPoint > 0 Then MB2.AddSamplingResult Vector2(TotalPx / NumberOfPoint, y)
                     Next i
                     SamplingCountPerPic = SamplingCountPerPic + (tmpTransitionCount / 2) '端点数/2 为区间数
              End If
                
              '为下一行初始化
              tmpX = 0
              tmpTransitionCount = 0
              TotalPx = 0
              NumberOfPoint = 0
                        
              Next y
              
              '每张图片有多少个点
              MB2.AddSamplingCountPerPicture SamplingCountPerPic
        
        '清理下
        'If PictureID <> GetTex("1") Then TexF.DeleteTexture GetTex(PictureID)
End Sub












'————————————————————识别标识点————————————————
Sub ComputeCamPosAndAngleY(iPictureID As Long)
'识别出2D标识点AOBC
AnalyzeIdenPoint (GetTex(iPictureID))
Dim CamAdjust As TV_3DVECTOR, CamWorld As TV_3DVECTOR
Dim Alpha As Single, Beta As Single, Cita As Single, Omega As Single, halfPhi As Single
Dim ViewAngleHori As Single, ViewAngleVert As Single, tanHalfVA_Vert As Single, tanHalfVA_Hori As Single
Dim EulerX As Single, EulerY As Single, EulerZ As Single
'可视角
ViewAngleHori = MainScanParam.VisibleAngleHorizontal * 3.1415926 / 180
ViewAngleVert = MainScanParam.VisibleAngleVertical * 3.1415926 / 180
tanHalfVA_Vert = Tan(ViewAngleVert / 2)
tanHalfVA_Hori = Tan(ViewAngleHori / 2)
'EulerY未知
EulerX = mScanCamPerPicture(iPictureID).ScanCamEulerX
EulerZ = mScanCamPerPicture(iPictureID).ScanCamEulerZ

'这里的2D point是-1到1的屏幕空间坐标
'现在要得到同样cam位置，而没有了翻滚角和仰角时的成像
'——————————————————————处理EulerZ
Dim tmpX As Single, tmpY As Single, k1 As Single, k2 As Single
Dim PW As Long, PH As Long
PW = PictureWidth
PH = PictureHeight
k1 = PH / PW
k2 = PW / PH


'TV.Clear
'scr.Draw_FilledCircle (PW / 2) * pointO.x + PW / 2, -pointO.y * (PH / 2) + PH / 2, 3, 12, RGBA(0, 0, 1, 1)
tmpX = pointO.x '一定要先储存....不然pointO会直接赋值影响pointO.y的计算
tmpY = pointO.y
pointO.x = tmpX * Cos(EulerZ) - tmpY * k1 * Sin(EulerZ)
pointO.y = tmpX * k2 * Sin(EulerZ) + tmpY * Cos(EulerZ)
'scr.Draw_FilledCircle (PW / 2) * pointO.x + PW / 2, -pointO.y * (PH / 2) + PH / 2, 3, 12, RGBA(1, 0, 0, 1)


'scr.Draw_FilledCircle (PW / 2) * pointA.x + PW / 2, -pointA.y * (PH / 2) + PH / 2, 3, 12, RGBA(0, 0, 1, 1)
tmpX = pointA.x
tmpY = pointA.y
pointA.x = tmpX * Cos(EulerZ) - tmpY * k1 * Sin(EulerZ)
pointA.y = tmpX * k2 * Sin(EulerZ) + tmpY * Cos(EulerZ)
'scr.Draw_FilledCircle (PW / 2) * pointA.x + PW / 2, -pointA.y * (PH / 2) + PH / 2, 3, 12, RGBA(1, 0, 0, 1)

'scr.Draw_FilledCircle (PW / 2) * pointB.x + PW / 2, -pointB.y * (PH / 2) + PH / 2, 3, 12, RGBA(0, 0, 1, 1)
tmpX = pointB.x
tmpY = pointB.y
pointB.x = tmpX * Cos(EulerZ) - tmpY * k1 * Sin(EulerZ)
pointB.y = tmpX * k2 * Sin(EulerZ) + tmpY * Cos(EulerZ)
'scr.Draw_FilledCircle (PW / 2) * pointB.x + PW / 2, -pointB.y * (PH / 2) + PH / 2, 3, 12, RGBA(1, 0, 0, 1)

'scr.Draw_FilledCircle (PW / 2) * pointC.x + PW / 2, -pointC.y * (PH / 2) + PH / 2, 3, 12, RGBA(0, 0, 1, 1)
tmpX = pointC.x
tmpY = pointC.y
pointC.x = tmpX * Cos(EulerZ) - tmpY * k1 * Sin(EulerZ)
pointC.y = tmpX * k2 * Sin(EulerZ) + tmpY * Cos(EulerZ)
'scr.Draw_FilledCircle (PW / 2) * pointC.x + PW / 2, -pointC.y * (PH / 2) + PH / 2, 3, 12, RGBA(1, 0, 0, 1)

'TV.RenderToScreen

'————————————————————————处理EulerX
Dim pointO3d As TV_3DVECTOR, pointA3d As TV_3DVECTOR, pointB3d As TV_3DVECTOR, pointC3d As TV_3DVECTOR
Dim L As Single, R0 As Single, R1 As Single, R2 As Single, R3 As Single '论文的东西
Dim e As Single, f As Single, k As Single, q As Single, p As Single
Dim PA_scr As TV_3DVECTOR, PO_scr As TV_3DVECTOR
PA_scr = Vector3(pointA.x * tanHalfVA_Hori, pointA.y * tanHalfVA_Vert, 1) '屏幕射线PA
PO_scr = Vector3(pointO.x * tanHalfVA_Hori, pointO.y * tanHalfVA_Vert, 1) '屏幕射线PO模
L = MainScanParam.IdentificationLineLength '标识边长
e = Math.VModulus(PA_scr)
f = Math.VModulus(PO_scr)
k = Math.VDotProduct(Math.VNormalize(PA_scr), Math.VNormalize(PO_scr))
'这个在开方时正负号或许要讨论下。。?
q = e * L * Sin(EulerX)
p = e / f
'解一元二次方程
Dim a As Single, b As Single, c As Single, Solution1 As Single, Solution2 As Single
'If (R1 / e - R0 / f) * (L * Sin(EulerX)) > 0 Then
       a = 1 + p ^ 2 - 2 * k * p
       b = 2 * p * q - 2 * k * q
       c = q ^ 2 - L ^ 2
'Else
 '      a = 1 + p ^ 2 - 2 * k * p
 '      b = -2 * p * q + 2 * k * q
 '      c = q ^ 2 - L ^ 2
'End If
'根的正负？
Solution1 = (-b + Sqr(b ^ 2 - 4 * a * c)) / (2 * a)
Solution2 = (-b - Sqr(b ^ 2 - 4 * a * c)) / (2 * a)
R0 = Solution1
'R1 = p * R0 + q
pointO3d = Math.VScale(Math.VNormalize(PO_scr), R0)
'pointA3d = Math.VScale(PA_scr, R1)
'这时还有仰角和水平角的影响
Dim MatrixTMP As TV_3DMATRIX, MatrixAdjustX_RotateY As TV_3DMATRIX
Dim MatrixAdjustX_RotateYInv As TV_3DMATRIX, MatrixAdjustX_RotateX As TV_3DMATRIX
With MatrixTMP
       .m11 = pointO3d.x
       .m21 = pointO3d.y
       .m31 = pointO3d.z
End With

'正交阵的逆等于他的转置
TV.SetAngleSystem TV_ANGLE_RADIAN
'还要把O旋转一个仰角让它和水平面平行！！！
Math.TVMatrixRotationX MatrixAdjustX_RotateX, EulerX

Math.TVMatrixMultiply MatrixTMP, MatrixAdjustX_RotateX, MatrixTMP

'这里直接把O点当局部原点，则camera为-O的坐标
CamAdjust = Vector3(-MatrixTMP.m11, -MatrixTMP.m21, -MatrixTMP.m31)

'TV.AddToLog CamAdjust.x
'TV.AddToLog CamAdjust.y
'TV.AddToLog CamAdjust.z






'////////////////3D点B的校正空间的坐标 把上面的O换B A换C即可
Dim PB_scr As TV_3DVECTOR, PC_scr As TV_3DVECTOR
PB_scr = Vector3(pointB.x * tanHalfVA_Hori, pointB.y * tanHalfVA_Vert, 1) '屏幕射线PC
PC_scr = Vector3(pointC.x * tanHalfVA_Hori, pointC.y * tanHalfVA_Vert, 1) '屏幕射线PB模
L = MainScanParam.IdentificationLineLength '标识边长
e = Math.VModulus(PC_scr)
f = Math.VModulus(PB_scr)
k = Math.VDotProduct(Math.VNormalize(PC_scr), Math.VNormalize(PB_scr))
'这个在开方时正负号或许要讨论下。。?
q = e * L * Sin(EulerX)
p = e / f
'解一元二次方程
a = 1 + p ^ 2 - 2 * k * p
b = 2 * p * q - 2 * k * q
c = q ^ 2 - L ^ 2
'根的正负？
Solution1 = (-b + Sqr(b ^ 2 - 4 * a * c)) / (2 * a)
Solution2 = (-b - Sqr(b ^ 2 - 4 * a * c)) / (2 * a)
R2 = Solution1
'R3 = p * R2 + q
pointB3d = Math.VScale(Math.VNormalize(PB_scr), R2)
'pointC3d = Math.VScale(PC_scr, R3)



'这时还有仰角和水平角的影响
With MatrixTMP
       .m11 = pointB3d.x
       .m21 = pointB3d.y
       .m31 = pointB3d.z
End With


'正交阵的逆等于他的转置
TV.SetAngleSystem TV_ANGLE_RADIAN

Math.TVMatrixMultiply MatrixTMP, MatrixAdjustX_RotateX, MatrixTMP
'TV.AddToLog "B 3d  rotated :x" & MatrixTMP.m11
'TV.AddToLog "B 3d  rotated :y" & MatrixTMP.m21
'TV.AddToLog "B 3d  rotated :z" & MatrixTMP.m31

pointB3d = Vector3(MatrixTMP.m11 + CamAdjust.x, MatrixTMP.m21 + CamAdjust.y, MatrixTMP.m31 + CamAdjust.z)
'pointB3d = Math.VSubtract(pointB3d, pointO3d) '这两个都是没旋转eulerX的





'剩下水平角影响
'————————————————处理EulerY
'把标识线段OB对齐到世界坐标系X轴
Dim MatrixLocal As TV_3DMATRIX, MatrixAlign As TV_3DMATRIX, MatrixWorld As TV_3DMATRIX
Dim d As Single
d = Sqr((pointB3d.x) ^ 2 + (pointB3d.z) ^ 2)
With MatrixLocal
       .m11 = CamAdjust.x
       .m21 = CamAdjust.y
       .m31 = CamAdjust.z
End With
With MatrixAlign
       .m11 = pointB3d.x / d
       .m21 = 0
       .m31 = -pointB3d.z / d
       .m12 = 0
       .m22 = 1
       .m32 = 0
       .m13 = pointB3d.z / d
       .m23 = 0
       .m33 = pointB3d.x / d
End With
Math.TVMatrixMultiply MatrixWorld, MatrixAlign, MatrixLocal '左乘
CamWorld.x = MatrixWorld.m11
CamWorld.y = MatrixWorld.m21
CamWorld.z = MatrixWorld.m31


'获取eulerY
'先对齐到世界空间
Dim worldViewVector As TV_3DVECTOR, v As TV_3DVECTOR
With MatrixLocal
       .m11 = 0
       .m21 = 0
       .m31 = 1
End With
Math.TVMatrixMultiply MatrixWorld, MatrixAlign, MatrixLocal '左乘
worldViewVector.x = MatrixWorld.m11 'pointB3d.z / d
worldViewVector.y = 0
worldViewVector.z = MatrixWorld.m31 'pointB3d.x / d
Dim cosEulerY As Single, sinEulerY As Single
cosEulerY = worldViewVector.z
v = Math.VCrossProduct(Vector3(0, 0, 1), worldViewVector)
sinEulerY = v.y 'Math.VCrossProduct(worldViewVector, Vector3(0, 0, 1)).y

If sinEulerY < 0 Then '逆时针
       EulerY = Math.ACos(cosEulerY)
Else
       EulerY = 2 * 3.1415926 - Math.ACos(cosEulerY)
End If
'终于解出了水平角和世界坐标
SetScanCameraPerPicture iPictureID, CamWorld, , , , EulerX, EulerY, EulerZ
End Sub






Private Sub AnalyzeIdenPoint(iTextureID As Long)
'由于分析一个图片里面的标识点可能是一坨像素，但是我们只需要一个矢量点
'所以每个标识点就当成一个组
Dim pixel() As Type_PointWithID, SamplePoint() As TV_2DVECTOR
Dim PW As Long, PH As Long
Dim GROUP_POINTER(1 To 4) As Integer, CanCreateGroup As Boolean
GROUP_POINTER(1) = 0
GROUP_POINTER(2) = 0
GROUP_POINTER(3) = 0
GROUP_POINTER(4) = 0
PW = PictureWidth
PH = PictureHeight

ReDim pixel(1 To PW, 1 To PH) As Type_PointWithID
'——————加载图片————————

'——————选点————————
Dim tmpColor As Long, tmpC As TV_COLOR, i As Long, j As Long
For i = 1 To PW
       For j = 1 To PH
       tmpColor = TexF.GetPixel(iTextureID, i, j)
       tmpC = DecodeRGBA(tmpColor)
       If tmpC.r < 0.1 And tmpC.g < 0.1 And tmpC.b > 0.9 Then pixel(i, j).IsQualified = 1 '蓝
       If tmpC.r < 0.1 And tmpC.g > 0.9 And tmpC.b < 0.1 Then pixel(i, j).IsQualified = 2 '绿
       If tmpC.r > 0.9 And tmpC.g < 0.1 And tmpC.b > 0.9 Then pixel(i, j).IsQualified = 3 '紫
       If tmpC.r < 0.1 And tmpC.g > 0.9 And tmpC.b > 0.9 Then pixel(i, j).IsQualified = 4 '天蓝
       Next j
Next i
'——————

Dim idenColor As Integer
For i = 2 To PW - 1
       For j = 2 To PH - 1
              For idenColor = 1 To 4
              If pixel(i, j).gID = 0 And pixel(i, j).IsQualified = idenColor Then '还没有分到组的拿去处理
                     CanCreateGroup = True
                     '左上方有“朋友”就不用新建一个组了 干脆就加入那个组吧
                     If pixel(i - 1, j - 1).IsQualified = idenColor And pixel(i - 1, j - 1).gID <> 0 Then
                     pixel(i, j).gID = pixel(i - 1, j - 1).gID
                     CanCreateGroup = False
                     End If
                     If pixel(i - 1, j).IsQualified = idenColor And pixel(i - 1, j).gID <> 0 Then
                     pixel(i, j).gID = pixel(i - 1, j).gID
                     CanCreateGroup = False
                     End If
                     If pixel(i, j - 1).IsQualified = idenColor And pixel(i, j - 1).gID <> 0 Then
                     pixel(i, j).gID = pixel(i, j - 1).gID
                     CanCreateGroup = False
                     End If
                     If pixel(i - 1, j + 1).IsQualified = idenColor And pixel(i - 1, j + 1).gID <> 0 Then
                     pixel(i, j).gID = pixel(i - 1, j + 1).gID
                     CanCreateGroup = False
                     End If
                     If CanCreateGroup = True Then                                   '需要新建组
                     GROUP_POINTER(idenColor) = GROUP_POINTER(idenColor) + 1
                     pixel(i, j).gID = GROUP_POINTER(idenColor)                   'idencolor色的新建点组的index
                     End If
              End If
              Next idenColor
       Next j
Next i

'如果没生成组那就返回吧
For idenColor = 1 To 4
       If GROUP_POINTER(idenColor) > 0 Then
              ReDim IdenPoint(1 To GROUP_POINTER(idenColor)) As TV_2DVECTOR '某种颜色的点组
              Dim GroupPointCount() As Long
              ReDim GroupPointCount(1 To GROUP_POINTER(idenColor)) As Long     '点组内的像素数

              For i = 1 To PW
                     For j = 1 To PH
                     If pixel(i, j).gID <> 0 And pixel(i, j).IsQualified = idenColor Then '如果像素已经分组
                            IdenPoint(pixel(i, j).gID).x = IdenPoint(pixel(i, j).gID).x + Val(i)             '像素的坐标累加 之后一除就得到平均点坐标了
                            IdenPoint(pixel(i, j).gID).y = IdenPoint(pixel(i, j).gID).y + Val(j)
                            GroupPointCount(pixel(i, j).gID) = GroupPointCount(pixel(i, j).gID) + 1 '统计组内点数
                     End If
                     Next j
              Next i

              '————————得出矢量标识点——————
              Dim tmpFinalX As Single, tmpFinalY As Single
              For i = 1 To GROUP_POINTER(idenColor)
              tmpFinalX = tmpFinalX + IdenPoint(i).x / GroupPointCount(i)
              tmpFinalY = tmpFinalY + IdenPoint(i).y / GroupPointCount(i)
              Next i
              '多个标识点再取平均 一种颜色只有一个标识点 因为在做世界坐标解算时 4个标识点都有不同颜色
              tmpFinalX = tmpFinalX / GROUP_POINTER(idenColor)
              tmpFinalY = tmpFinalY / GROUP_POINTER(idenColor)
              '————————————————————

              '屏幕空间转到投影空间
              tmpFinalX = (tmpFinalX - PictureWidth / 2) / (PictureWidth / 2)
              tmpFinalY = -(tmpFinalY - PictureHeight / 2) / (PictureHeight / 2) 'y轴往下啊= =

              Select Case idenColor
                     Case 1 '蓝 左下
                     pointO = Vector2(tmpFinalX, tmpFinalY)
                     Case 2 '绿 右下
                     pointB = Vector2(tmpFinalX, tmpFinalY)
                     Case 3 '紫 左上
                     pointA = Vector2(tmpFinalX, tmpFinalY)
                     Case 4 '天蓝 右上
                     pointC = Vector2(tmpFinalX, tmpFinalY)
              End Select
       End If
       tmpFinalX = 0
       tmpFinalY = 0
Next idenColor
'TV.AddToLog "X:" & pointO.x & "Y:" & pointO.y
End Sub














'————————————————————生成点云——————————————————
Sub Generate3DPointCloud(iLightSide As CONST_LightSide, iMatrixType As CONST_MatrixType)

Select Case iMatrixType
       Case CONST_MatrixType.Matrix_OrthoBasis
       Gen3DPoint_OrthoBasis (iLightSide)
       Case CONST_MatrixType.Matrix_Euler
       Gen3DPoint_Euler (iLightSide)
End Select

End Sub
'正交基矩阵
Private Sub Gen3DPoint_OrthoBasis(iLightSide As CONST_LightSide)
Dim MatrixTrans As TV_3DMATRIX
Dim MatrixLocal As TV_3DMATRIX
Dim MatrixWorld As TV_3DMATRIX
        Dim LocalX            As Single, LocalY As Single, LocalZ As Single
        Dim x                      As Single, y As Single, z As Single, RealPictureH As Single, DEPTH As Single
        Dim d                     As Single, Cita As Single, TurnCenterToWall As Single
        Dim CamToLight As Single
       Dim tmpPointID As Long

        d = MainScanParam.StandardDepth '标准深度
        Cita = MainScanParam.VisibleAngleVertical * 3.1415926 / 180 '上下可视角
        RealPictureH = d * Tan(0.5 * Cita) * 2    '求出可视竖直范围的实际长度  但是跟z坐标有关联的..z会影响y轴的偏移 视角问题
        CamToLight = MainScanParam.CamToLight
        
        Dim PROCESSINGPICTUREID As Long, PROCESSINGPIXEL As Long
        For PROCESSINGPICTUREID = 1 To MB2.GetPictureCount
                For PROCESSINGPIXEL = 1 To MB2.GetSamplingCount(PROCESSINGPICTUREID)   '这幅图有多少个点
                     
                     tmpPointID = tmpPointID + 1
                       Dim SamplePoint As TV_2DVECTOR '获取采样点
                       SamplePoint = MB2.GetSamplingResult(tmpPointID) '逐点遍历
                        DEPTH = GetDepthFromOffset(MainScanParam, SamplePoint.x, iLightSide)
                        'x = 深度 - 盒宽一半
                        LocalX = -CamToLight '///LOCALZ为零的限制条件是激光过圆心
                        LocalY = RealPictureH / 2 - RealPictureH * ((SamplePoint.y / PictureHeight) + DEPTH / (2 * d) - (DEPTH * SamplePoint.y) / (d * PictureHeight))
                       'LocalZ = TurnCenterToWall - DEPTH
                       LocalZ = d - DEPTH '标准深度减
                       
                        With MatrixLocal '局部坐标 就用第一列了
                                .m11 = LocalX
                                .m21 = LocalY
                                .m31 = LocalZ
                                .m41 = 1
                        End With
       
                      '处理散乱点云时的摄像机参数
                     Dim ScanCamX As Single, ScanCamY As Single, ScanCamZ As Single
                     Dim ScanCamAxisX As TV_3DVECTOR, ScanCamAxisY As TV_3DVECTOR, ScanCamAxisZ As TV_3DVECTOR
                     '虚拟扫描摄像机的POSITION
                     ScanCamX = mScanCamPerPicture(PROCESSINGPICTUREID).ScanCamPos.x
                     ScanCamY = mScanCamPerPicture(PROCESSINGPICTUREID).ScanCamPos.y
                     ScanCamZ = mScanCamPerPicture(PROCESSINGPICTUREID).ScanCamPos.z
                     '标准正交基
                     'Z是摄像机向量 跟一开始摄像机对着背景墙的方向一致
                     Dim x3 As Single, y3 As Single, z3 As Single
                     x3 = mScanCamPerPicture(PROCESSINGPICTUREID).ScanCamVector.x
                     y3 = mScanCamPerPicture(PROCESSINGPICTUREID).ScanCamVector.y
                     z3 = mScanCamPerPicture(PROCESSINGPICTUREID).ScanCamVector.z
                     ScanCamAxisZ = Vector3(x3, y3, z3)
                     ScanCamAxisX = Vector3(z3, 0, -x3)
                     ScanCamAxisY = Math.VCrossProduct(ScanCamAxisX, ScanCamAxisZ)
                     ScanCamAxisY = Math.VScale(ScanCamAxisY, -1) '标准数学左手系的Y
                     '正交基归一化
                     ScanCamAxisX = Math.VNormalize(ScanCamAxisX)
                     ScanCamAxisY = Math.VNormalize(ScanCamAxisY)
                     ScanCamAxisZ = Math.VNormalize(ScanCamAxisZ)
                     
                        With MatrixTrans
                                .m11 = ScanCamAxisX.x
                                .m21 = ScanCamAxisX.y
                                .m31 = ScanCamAxisX.z
                                .m41 = 0
                                .m12 = ScanCamAxisY.x
                                .m22 = ScanCamAxisY.y
                                .m32 = ScanCamAxisY.z
                                .m42 = 0
                                .m13 = ScanCamAxisZ.x
                                .m23 = ScanCamAxisZ.y
                                .m33 = ScanCamAxisZ.z
                                .m43 = 0
                                .m14 = ScanCamX
                                .m24 = ScanCamY
                                .m34 = ScanCamZ
                                .m44 = 1
                        End With
                            
                     '用列向量
                        Math.TVMatrixMultiply MatrixWorld, MatrixTrans, MatrixLocal
       
                        x = MatrixWorld.m11
                        y = MatrixWorld.m21
                        z = MatrixWorld.m31

                         MB2.AddPoint x, y, z  '新增顶点

                Next PROCESSINGPIXEL
        Next PROCESSINGPICTUREID
End Sub

'欧拉角矩阵
Private Sub Gen3DPoint_Euler(iLightSide As CONST_LightSide)
Dim MatrixEulerY As TV_3DMATRIX
Dim MatrixEulerX As TV_3DMATRIX
Dim MatrixDirCos As TV_3DMATRIX
Dim MatrixEulerZ As TV_3DMATRIX
Dim MatrixTranslation As TV_3DMATRIX
Dim MatrixTMP As TV_3DMATRIX
Dim MatrixWorld As TV_3DMATRIX

        Dim LocalX            As Single, LocalY As Single, LocalZ As Single
        Dim x                      As Single, y As Single, z As Single, RealPictureH As Single, DEPTH As Single
        Dim d                     As Single, Cita As Single, TurnCenterToWall As Single
        Dim CamToLight As Single
       Dim tmpPointID As Long
       Dim LocalPointPerPicture() As TV_3DVECTOR '每一列点的局部坐标 搞个bezier先
       ReDim LocalPointPerPicture(1 To PictureHeight) '其实应该预留多一点空间的


        d = MainScanParam.StandardDepth '标准深度
        Cita = MainScanParam.VisibleAngleVertical * 3.1415926 / 180 '上下可视角
        RealPictureH = d * Tan(0.5 * Cita) * 2    '求出可视竖直范围的实际长度  但是跟z坐标有关联的..z会影响y轴的偏移 视角问题
        CamToLight = MainScanParam.CamToLight
        
        Dim PROCESSINGPICTUREID As Long, PROCESSINGPIXEL As Long
        For PROCESSINGPICTUREID = 1 To MB2.GetPictureCount
                For PROCESSINGPIXEL = 1 To MB2.GetSamplingCount(PROCESSINGPICTUREID)   '这幅图有多少个点
                     
                     tmpPointID = tmpPointID + 1
                       Dim SamplePoint As TV_2DVECTOR '获取采样点
                       SamplePoint = MB2.GetSamplingResult(tmpPointID) '逐点遍历
                        DEPTH = GetDepthFromOffset(MainScanParam, SamplePoint.x, iLightSide)
                        'x = 深度 - 盒宽一半
                        LocalX = -CamToLight '///LOCALZ为零的限制条件是激光过圆心
                        LocalY = RealPictureH / 2 - RealPictureH * ((SamplePoint.y / PictureHeight) + DEPTH / (2 * d) - (DEPTH * SamplePoint.y) / (d * PictureHeight))
                       LocalZ = d - DEPTH '标准深度减
                       
                       '////////////////////////bezier////////////////////
                      ' LocalPointPerPicture(PROCESSINGPIXEL) = Vector3(LocalX, LocalY, LocalZ)


                        With MatrixTMP '局部坐标 就用第一列了
                                .m11 = LocalX 'LocalPointPerPicture(PROCESSINGPIXEL).x
                                .m21 = LocalY 'LocalPointPerPicture(PROCESSINGPIXEL).y
                                .m31 = LocalZ 'LocalPointPerPicture(PROCESSINGPIXEL).z
                                .m41 = 1
                        End With
       
       
                      '处理散乱点云时的摄像机参数
                     Dim EulerX As Single, EulerY As Single, EulerZ As Single
                     Dim ScanCamX As Single, ScanCamY As Single, ScanCamZ As Single
                     EulerX = mScanCamPerPicture(PROCESSINGPICTUREID).ScanCamEulerX
                     EulerY = mScanCamPerPicture(PROCESSINGPICTUREID).ScanCamEulerY '+ 3.1415926 / 2
                     EulerZ = mScanCamPerPicture(PROCESSINGPICTUREID).ScanCamEulerZ
                     ScanCamX = mScanCamPerPicture(PROCESSINGPICTUREID).ScanCamPos.x
                     ScanCamY = mScanCamPerPicture(PROCESSINGPICTUREID).ScanCamPos.y
                     ScanCamZ = mScanCamPerPicture(PROCESSINGPICTUREID).ScanCamPos.z
                     
                     With MatrixDirCos
                            .m11 = Cos(EulerY)
                            .m21 = 0
                            .m31 = Sin(EulerY)
                            .m41 = 0
                            .m12 = Sin(EulerX) * Sin(EulerY)
                            .m22 = Cos(EulerX)
                            .m32 = -Cos(EulerY) * Sin(EulerX)
                            .m42 = 0
                            .m13 = -Sin(EulerY) * Cos(EulerX)
                            .m23 = Sin(EulerX)
                            .m33 = Cos(EulerX) * Cos(EulerY)
                            .m43 = 0
                            .m14 = 0
                            .m24 = 0
                            .m34 = 0
                            .m44 = 1
                     End With
                     With MatrixEulerZ
                            .m11 = Cos(EulerZ)
                            .m21 = Sin(EulerZ)
                            .m12 = -Sin(EulerZ)
                            .m22 = Cos(EulerZ)
                            .m33 = 1
                            .m44 = 1
                     End With
                     With MatrixTranslation
                                .m11 = 1
                                .m22 = 1
                                .m33 = 1
                                .m14 = ScanCamX
                                .m24 = ScanCamY
                                .m34 = ScanCamZ
                                .m44 = 1
                     End With
                     '用列向量
                     
                    'Math.TVMatrixRotationAxis MatrixEulerZ, _
                                                               Vector3(MatrixDirCos.m13, _
                                                                MatrixDirCos.m23, _
                                                                 MatrixDirCos.m33), _
                                                                EulerZ
                     'Math.TVMatrixRotationYawPitchRoll MatrixEulerZ, EulerY, -EulerX, -EulerZ
                     
                    'Math.TVMatrixRotationX MatrixEulerX, -EulerX
                     'Math.TVMatrixRotationY MatrixEulerY, -EulerY
                     'TV.AddToLog MatrixEulerX.m32
                     'TV.AddToLog MatrixEulerY.m31
                     'TV.AddToLog "——————" & PROCESSINGPICTUREID & "————————"
                     'TV.AddToLog "local---"
                     'TV.AddToLog MatrixTMP.m11
                     'TV.AddToLog MatrixTMP.m21
                     'TV.AddToLog MatrixTMP.m31
                      'Math.TVMatrixMultiply MatrixTMP, MatrixEulerZ, MatrixTMP
                     'TV.AddToLog "z---"
                     ''TV.AddToLog MatrixTMP.m11
                    ' TV.AddToLog MatrixTMP.m21
                     'TV.AddToLog MatrixTMP.m31
                     
                     'Math.TVMatrixTranspose MatrixDirCos, MatrixDirCos
                     Math.TVMatrixMultiply MatrixTMP, MatrixEulerZ, MatrixTMP
                     Math.TVMatrixMultiply MatrixTMP, MatrixDirCos, MatrixTMP
                     'TV.AddToLog "x---"
                     'TV.AddToLog MatrixTMP.m11
                     'TV.AddToLog MatrixTMP.m21
                     'TV.AddToLog MatrixTMP.m31
                     
                     'Math.TVMatrixMultiply MatrixTMP, MatrixEulerZ, MatrixTMP
                     'TV.AddToLog "y---"
                    ' TV.AddToLog MatrixTMP.m11
                    ' TV.AddToLog MatrixTMP.m21
                     'TV.AddToLog MatrixTMP.m31
                     'Math.TVMatrixMultiply MatrixTMP, MatrixEulerZ, MatrixTMP
                     
                    
                        Math.TVMatrixMultiply MatrixWorld, MatrixTranslation, MatrixTMP
                    ' TV.AddToLog "world---"
                    ' TV.AddToLog MatrixWorld.m11
                   '  TV.AddToLog MatrixWorld.m21
                     'TV.AddToLog MatrixWorld.m31
                        x = MatrixWorld.m11
                        y = MatrixWorld.m21
                        z = MatrixWorld.m31

                         MB2.AddPoint x, y, z  '新增顶点
                            'MB2.AddPoint ScanCamX, ScanCamY, ScanCamZ
                Next PROCESSINGPIXEL
        Next PROCESSINGPICTUREID
End Sub

Sub SetPictureCount(iCount As Long)
ReDim Preserve mScanCamPerPicture(1 To iCount + 1)
End Sub
'每张图片都要有对应的camera信息
Sub SetScanCameraPerPicture(picID As Long, CamPos As TV_3DVECTOR, _
                                                 Optional CamVectorX As Single = 0, _
                                                 Optional CamVectorY As Single = 0, _
                                                 Optional CamVectorZ As Single = 1, _
                                                 Optional CamEulerX As Single = 0, _
                                                 Optional CamEulerY As Single = 0, _
                                                 Optional CamEulerZ As Single = 0)
With mScanCamPerPicture(picID)
.ScanCamPos = CamPos
.ScanCamVector = Vector3(CamVectorX, CamVectorY, CamVectorZ)
.ScanCamEulerX = CamEulerX
.ScanCamEulerY = CamEulerY
.ScanCamEulerZ = CamEulerZ
End With
End Sub




'——————————————形成闭合体————————————
Sub MeshReconstruction(iMethod As CONST_Reconstruction)
Select Case iMethod
       Case CONST_Reconstruction.RC_MappedBall
       
       
End Select
End Sub






'————————————————保存STL——————————————
Sub SaveSTL(ObjectName As String, Path As String, CoverOriginFile As Boolean)
Dim i As Long
Dim FileStart As String * 80
Dim TotalTriangles As Long '
Dim TriangleEnd As String * 2
Dim Normal As TV_3DVECTOR, v1 As TV_3DVECTOR, v2 As TV_3DVECTOR

FileStart = "Solid " & ObjectName
TotalTriangles = MB2.GetTriangleAmount


If Dir(Path) <> "" And CoverOriginFile = True Then Kill Path '删除源文件
Open Path For Binary As #1
Put #1, , FileStart '文件头 80字节
Put #1, , TotalTriangles '三角形数 4字节
For i = 1 To TotalTriangles
       v1 = Math.VSubtract(MB2.GetTriangleVertex(i, 2), MB2.GetTriangleVertex(i, 1))
       v2 = Math.VSubtract(MB2.GetTriangleVertex(i, 3), MB2.GetTriangleVertex(i, 1))
       Normal = Math.VCrossProduct(v1, v2) '法向量
       Put #1, , Normal.x 'UP软件是XZY
       Put #1, , Normal.z
       Put #1, , Normal.y
       Put #1, , MB2.GetTriangleVertex(i, 2).x
       Put #1, , MB2.GetTriangleVertex(i, 2).z
       Put #1, , MB2.GetTriangleVertex(i, 2).y
       Put #1, , MB2.GetTriangleVertex(i, 1).x
       Put #1, , MB2.GetTriangleVertex(i, 1).z
       Put #1, , MB2.GetTriangleVertex(i, 1).y
       Put #1, , MB2.GetTriangleVertex(i, 3).x
       Put #1, , MB2.GetTriangleVertex(i, 3).z
       Put #1, , MB2.GetTriangleVertex(i, 3).y
       Put #1, , TriangleEnd '每个三角形的末尾 2字节
Next i
Close #1

End Sub



'————————————————保存STL——————————————
Sub SaveAsciiPointCloud(Path As String, CoverOriginFile As Boolean)
Dim i As Long
Dim tmpPoint As TV_3DVECTOR
Dim TotalTriangles As Long '
TotalTriangles = MB2.GetTriangleAmount


If Dir(Path) <> "" And CoverOriginFile = True Then Kill Path '删除源文件

Open Path For Append As #1
Print #1, "Triangle Amount:" & TotalTriangles '文件头 80字节
For i = 1 To MB2.GetPointAmount
       tmpPoint = MB2.GetPointByID(i)
       Print #1, "(" & tmpPoint.x & "," & tmpPoint.y & "," & tmpPoint.z & ")"
Next i
Close #1

End Sub


Sub LoadtvmToPointCloud(iPath As String)
MB2.LoadtvmToPointCloud iPath
End Sub


Sub AddToLog(aText As String)
TV.AddToLog aText
End Sub
Sub SetScanParameters(SP As Type_ScanParameters)
MainScanParam = SP
MainColorFilterLowerBound = SP.ColorFilter '颜色筛选下界
End Sub
Sub AppSleep(mSec As Long)
Sleep (mSec)
End Sub

