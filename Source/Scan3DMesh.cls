VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Scan3DProcessingPipeLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True




'——————加载图片————————
Private TheFirstPicturePath As String '第一张图的路径

Private Type Type_ScanCamPerPicture
        ScanCamPos As TV_3DVECTOR
       ScanCamVector As TV_3DVECTOR
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

Public Type Type_ScanParameters
TurningCenterToWall As Single '旋转中心到背景板
StandardDepth As Single '标准深度 焦点到背景板
VisibleAngleHorizontal As Single '可视角θ横
VisibleAngleVertical As Single '可视角θ竖
CamToLight As Single '镜头到红线的距离
ColorFilter As Single '颜色筛选下界
End Type

'Private MaincolorLOWERBOUND As Single
'——————2生成顶点——————





'————————流水线开始——————
'
'处理MAINMeshBuffer
'
'——————————————————
Sub SetMainBuffer(NewMeshBuffer As Scan3DMeshBuffer)
Set MB = NewMeshBuffer
End Sub

Function GetMainBuffer() As Scan3DMeshBuffer
Set GetMainBuffer = MB
End Function




Sub LoadScanPicture_START()
        TexF.SetTextureMode TV_TEXTUREMODE_16BITS
End Sub

Sub LoadScanPicture(iPath As String, PictureID As Long) '加载图片
       If PictureID = 1 Then TheFirstPicturePath = iPath '记录下第一张图片的路径
       TexF.LoadTexture iPath, CStr(PictureID), PictureWidth, PictureHeight
       TexF.LockTexture GetTex(CStr(PictureID)), True
End Sub

Sub LoadScanPicture_END()
       SampleFromPicture Side_Left, 1  '其实是对第一张图片采样 要重叠
End Sub

Sub ClearPictureBuffer()
TexF.DeleteAllTextures
End Sub
Sub ClearMainMeshBuffer()
MB.ClearTriangleBuffer
End Sub
Sub ClearMainPointCloudBuffer()
MB.ClearPoint3DBuffer
End Sub
Sub ClearMainBuffer()
MB.ClearBuffer
End Sub
Sub AddPointCloudToRenderMesh()
MB.AddPointCloudToRenderMesh
End Sub
Sub AddTriangleToRenderMesh()
MB.AddTriangleToRenderMesh
End Sub
Sub WeldVertices_RenderMesh()
MB.WeldVertices
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
        Dim NumberOfPoint        As Single, TotalPx As Single, RED As Single, VlinepointCounter As Long
        Dim PixelCOLOR As Long, MainColor As Single, Color1 As Single, Color2 As Single

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
       
       Dim P As Long
       If PictureID = -1 Then '-1代表最后一张图片
       P = GetTex("1")
       Else
       P = GetTex(PictureID) '1 到 N
       End If

       Dim x As Long, y As Long, pcolor As TV_COLOR
                For y = 1 To PictureHeight
                        For x = PixelTraversalStartX To PixelTraversalEndX
                                   
                                PixelCOLOR = TexF.GetPixel(P, x, y)
                                pcolor = DecodeRGBA(PixelCOLOR)
                                MainColor = pcolor.r  '取色
                                Color1 = pcolor.g
                                Color2 = pcolor.b

                                '筛选颜色
                                'If MainColor * (1 - Color1) ^ 2 * (1 - Color2) ^ 2 > MaincolorLOWERBOUND Then
                                 '       TotalPx = TotalPx + x * MainColor
                                '        NumberOfPoint = NumberOfPoint + MainColor
                                'End If
                               'If MainColor > MainColorFilterLowerBound And Color1 < 0.2 And Color2 < 0.2 Then
                               If MainColor > MainColorFilterLowerBound And Color1 < 0.2 And 0.4 < Color2 < 0.55 Then '激光颜色太坑爹了
                                     TotalPx = TotalPx + x * MainColor
                                     NumberOfPoint = NumberOfPoint + MainColor
                               End If
                        Next x
                     
                        '生成采样点
                        If NumberOfPoint <> 0 Then
                            '////////////////////返回函数值
                                MB.AddSamplingResult Vector2(TotalPx / NumberOfPoint, y)
                                VlinepointCounter = VlinepointCounter + 1
                        End If
                
                        TotalPx = 0
                        NumberOfPoint = 0
                        
                Next y
                
        MB.AddLine VlinepointCounter  '扫完了一幅图片 新增一列点
        VlinepointCounter = 0
        
        '清理下
        'If PictureID <> GetTex("1") Then TexF.DeleteTexture GetTex(PictureID)
End Sub








'————————————生成3D顶点———————————————

Sub Generate3DPointCloud(iLightSide As CONST_LightSide, IsStructed As Boolean)
Select Case IsStructed
       Case True '是结构点云模式
       Call_Gen3DPointCloud_Structed (iLightSide)
       Case False
       Call_Gen3DPointCloud_Irregular (iLightSide)
End Select
End Sub
Sub Call_Gen3DPointCloud_Structed(iLightSide As CONST_LightSide)

Dim MatrixTrans As TV_3DMATRIX
Dim MatrixLocal As TV_3DMATRIX
Dim MatrixWorld As TV_3DMATRIX


        Dim LocalX          As Single, LocalY As Single, LocalZ As Single

        Dim x               As Single, y As Single, z As Single, RealPictureH As Single, DEPTH As Single

        Dim d               As Single, cita As Single, TurnCenterToWall As Single

        Dim NumOfPictures   As Long '照片数

        Dim PictureStartID  As Long, PictureEndID As Long

        Dim ProcessingAngle As Single '正在处理照片的角度


        PictureStartID = 1 'GetTex("1")
        PictureEndID = MB.GetLineAmount - 1 '不记得为啥 但是以前写的不要改了= =
        NumOfPictures = MB.GetLineAmount - 1
        d = MainScanParam.StandardDepth '标准深度
        cita = MainScanParam.VisibleAngleVertical * 3.1415926 / 180 '上下可视角
        TurnCenterToWall = MainScanParam.TurningCenterToWall '必加
        RealPictureH = d * Tan(0.5 * cita) * 2    '求出可视竖直范围的实际长度  但是跟z坐标有关联的..z会影响y轴的偏移 视角问题
        
        
        
        Dim PROCESSINGPICTUREID As Long, PROCESSINGPIXEL As Long
        
        For PROCESSINGPICTUREID = PictureStartID To PictureEndID + 1
                For PROCESSINGPIXEL = 1 To MB.GetLine(PROCESSINGPICTUREID) '这列有多少个点
 
                       Dim SamplePoint As TV_2DVECTOR '获取采样点
                       SamplePoint = MB.GetSamplingResult(MB.GetPointID(PROCESSINGPICTUREID, PROCESSINGPIXEL))
                        DEPTH = GetDepthFromOffset(MainScanParam, SamplePoint.x, iLightSide)
                        
                        'x = 深度 - 盒宽一半
                        LocalX = 0 'CamToLight ///LOCALZ为零的限制条件是激光过圆心

                        LocalY = RealPictureH * ((SamplePoint.y / PictureHeight) + DEPTH / (2 * d) - (DEPTH * SamplePoint.y) / (d * PictureHeight))
        
                       LocalZ = TurnCenterToWall - DEPTH
                       
                       
                        With MatrixLocal '局部坐标 就用第一列了
                                .m11 = LocalX
                                .m21 = LocalY
                                .m31 = LocalZ
                        End With
        
                        '正在处理的图片的摄像机角度(俯视图)
                        ProcessingAngle = 2 * 3.1415926 * (PROCESSINGPICTUREID - PictureStartID) / NumOfPictures
       
                        With MatrixTrans '变换矩阵 其实是格式是4X4的不过用3X3够了
                                .m11 = Cos(ProcessingAngle)
                                .m12 = 0
                                .m13 = -Sin(ProcessingAngle)
                                .m21 = 0
                                .m22 = -1
                                .m23 = 0
                                .m31 = Sin(ProcessingAngle)
                                .m32 = 0
                                .m33 = Cos(ProcessingAngle)
                        End With

                        Math.TVMatrixMultiply MatrixWorld, MatrixTrans, MatrixLocal  'Local坐标变换
       
                        x = MatrixWorld.m11
                        y = MatrixWorld.m21
                        z = MatrixWorld.m31

                         MB.AddPoint x, y + RealPictureH / 2, z  '新增顶点

                Next PROCESSINGPIXEL
        Next PROCESSINGPICTUREID

End Sub

Sub Call_Gen3DPointCloud_Irregular(iLightSide As CONST_LightSide)
Dim MatrixTrans As TV_3DMATRIX
Dim MatrixLocal As TV_3DMATRIX
Dim MatrixWorld As TV_3DMATRIX
        Dim LocalX            As Single, LocalY As Single, LocalZ As Single
        Dim x                      As Single, y As Single, z As Single, RealPictureH As Single, DEPTH As Single
        Dim d                     As Single, cita As Single, TurnCenterToWall As Single
        Dim NumOfPictures   As Long '照片数
        Dim PictureStartID  As Long, PictureEndID As Long


        PictureStartID = 1 'GetTex("1")
        PictureEndID = MB.GetLineAmount '- 1 'GetTex("END") - GetTex("1")
        NumOfPictures = MB.GetLineAmount - 1
        d = MainScanParam.StandardDepth '标准深度
        cita = MainScanParam.VisibleAngleVertical * 3.1415926 / 180 '上下可视角
        TurnCenterToWall = MainScanParam.TurningCenterToWall '必加
        RealPictureH = d * Tan(0.5 * cita) * 2    '求出可视竖直范围的实际长度  但是跟z坐标有关联的..z会影响y轴的偏移 视角问题
        
        
        Dim PROCESSINGPICTUREID As Long, PROCESSINGPIXEL As Long
        For PROCESSINGPICTUREID = PictureStartID To PictureEndID '+ 1
                For PROCESSINGPIXEL = 1 To MB.GetLine(PROCESSINGPICTUREID) '这列有多少个点
 
                       Dim SamplePoint As TV_2DVECTOR '获取采样点
                       SamplePoint = MB.GetSamplingResult(MB.GetPointID(PROCESSINGPICTUREID, PROCESSINGPIXEL))
                        DEPTH = GetDepthFromOffset(MainScanParam, SamplePoint.x, iLightSide)
                        'x = 深度 - 盒宽一半
                        LocalX = 0 'CamToLight ///LOCALZ为零的限制条件是激光过圆心
                        LocalY = RealPictureH * ((SamplePoint.y / PictureHeight) + DEPTH / (2 * d) - (DEPTH * SamplePoint.y) / (d * PictureHeight))
                       LocalZ = TurnCenterToWall - DEPTH
                       
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

                         MB.AddPoint x, y + RealPictureH / 2, z  '新增顶点

                Next PROCESSINGPIXEL
        Next PROCESSINGPICTUREID

End Sub

Sub SetPictureCount(iCount As Long)
ReDim Preserve mScanCamPerPicture(1 To iCount + 1)
End Sub

Sub SetScanCameraPerPicture(picID As Long, CamPos As TV_3DVECTOR, CamVector As TV_3DVECTOR)
With mScanCamPerPicture(picID)
.ScanCamPos = CamPos
.ScanCamVector = CamVector
End With
End Sub










'————————————————BEZIER曲线校正————————————————————
Sub BezierCurveFittingForPoint(RepeatTimes As Long)
Dim OriginPoint1 As TV_3DVECTOR, OriginPoint2 As TV_3DVECTOR, OriginPoint3 As TV_3DVECTOR
Dim ResultPoint As TV_3DVECTOR
Dim r As Long, i As Long, j As Long
For r = 1 To RepeatTimes
       For i = 1 To MB.GetLineAmount
              For j = 2 To MB.GetLine(i) - 1  '只平滑第二到倒数第二个点
              OriginPoint1 = MB.GetPoint(i, j - 1)
              OriginPoint2 = MB.GetPoint(i, j)
              OriginPoint3 = MB.GetPoint(i, j + 1)

              ResultPoint = BezierInterpolation(0.5, OriginPoint1, OriginPoint2, OriginPoint3) 'bezier函数在GLOBAL模块
              MB.SetPoint MB.GetPointID(i, j), ResultPoint.x, ResultPoint.y, ResultPoint.z
              Next j
       Next i
Next r

ResetRenderMesh
MB.AddPointCloudToRenderMesh
End Sub










'——————————————形成闭合体————————————

Sub GenerateClosedObject()
       
        DoEvents
        '————————————————连起邻列三角形——————————————
        
        Dim MinOfTotalPoint As Long

        Dim p1              As TV_3DVECTOR, p2 As TV_3DVECTOR, p3 As TV_3DVECTOR

       Dim i As Long, j As Long
        For i = 1 To MB.GetLineAmount - 1
        
                MinOfTotalPoint = MIN(MB.GetLine(i), MB.GetLine(i + 1))   '两列点谁少点
                For j = 1 To MAX(MB.GetLine(i), MB.GetLine(i + 1))

                        Select Case Val(j)
                                Case Is < MinOfTotalPoint  '先判断两列谁的点数少
                                        p1 = MB.GetPoint(i, j)   'Line1FirstPointID
                                        p2 = MB.GetPoint(i + 1, j)  'Line2FirstPointID + j
                                        p3 = MB.GetPoint(i, j + 1) 'Line1FirstPointID + j + 1
                                        MB.AddTriangleByVector p1, p2, p3
                                        
                                        p1 = MB.GetPoint(i + 1, j)
                                        p2 = MB.GetPoint(i + 1, j + 1)
                                        p3 = MB.GetPoint(i, j + 1)
                                        MB.AddTriangleByVector p1, p2, p3
                                        
                                        
                                Case Is > MinOfTotalPoint '对应点连完到多出来的点了
                                        If MB.GetLine(i) > MB.GetLine(i + 1) Then
                                                '第一列的点多第二列点少的时候
                                                p1 = MB.GetPoint(i, j - 1)
                                                p2 = MB.GetPoint(i + 1, MinOfTotalPoint)  '第二列最低点做定点
                                                p3 = MB.GetPoint(i, j)
                                        Else
                                                p1 = MB.GetPoint(i, MinOfTotalPoint)  '-1
                                                p2 = MB.GetPoint(i + 1, j - 1)
                                                p3 = MB.GetPoint(i + 1, j)
                                        End If
                                   
                            MB.AddTriangleByVector p1, p2, p3
                            End Select

                Next j
        Next i
       
       
        '——————————封顶与封底——————————————
        Dim TotalVLine As Long
        TotalVLine = MB.GetLineAmount - 1 '总列数
        
        For i = 1 To Int((MB.GetLineAmount - 1) / 2) - 1 'int是向下取整  注意：第一列和最后一列是重叠的，所以减1
                '——顶
                p1 = MB.GetPoint(i, 1)
                p2 = MB.GetPoint(TotalVLine - i + 1, 1)
                p3 = MB.GetPoint(TotalVLine - i, 1)
                MB.AddTriangleByVector p1, p2, p3
        
                p1 = MB.GetPoint(i, 1)
                p2 = MB.GetPoint(i + 1, 1)
                p3 = MB.GetPoint(TotalVLine - i, 1)
                MB.AddTriangleByVector p2, p1, p3
       
                '——底
                p1 = MB.GetPoint(i, MB.GetLine(i))
                p2 = MB.GetPoint(TotalVLine - i + 1, MB.GetLine(TotalVLine - i + 1))
                p3 = MB.GetPoint(TotalVLine - i, MB.GetLine(TotalVLine - i))
                MB.AddTriangleByVector p2, p1, p3
        
                p1 = MB.GetPoint(i, MB.GetLine(i))
                p2 = MB.GetPoint(i + 1, MB.GetLine(i + 1))
                p3 = MB.GetPoint(TotalVLine - i, MB.GetLine(TotalVLine - i))
                MB.AddTriangleByVector p1, p2, p3 '顶点储存顺序是由UP软件提示修改的= =
       Next i
        
        
        
        Dim BeforeHalf As Long
        BeforeHalf = Int((TotalVLine - 1) / 2)
        If Int(NumOfVerticalLines - 1 / 2) Mod 2 = 1 Then '除去最后重复列后   列数是奇数
                p1 = MB.GetPoint(BeforeHalf, 1)  '半圈前的最后一个点
                p2 = MB.GetPoint(BeforeHalf + 1, 1)
                p3 = MB.GetPoint(BeforeHalf + 2, 1)
                MB.AddTriangleByVector p1, p2, p3
        
                p1 = MB.GetPoint(BeforeHalf, MB.GetLine(BeforeHalf))  '半圈前的最后一个点
                p2 = MB.GetPoint(BeforeHalf + 1, MB.GetLine(BeforeHalf + 1))
                p3 = MB.GetPoint(BeforeHalf + 2, MB.GetLine(BeforeHalf + 2))
                MB.AddTriangleByVector p1, p2, p3
       End If

                '————————————————————————————————————————————————
       MB.WeldVertices
End Sub






'————————————————保存STL——————————————
Sub SaveSTL(ObjectName As String, Path As String, CoverOriginFile As Boolean)
Dim i As Long
Dim FileStart As String * 80
Dim TotalTriangles As Long '
Dim TriangleEnd As String * 2
Dim Normal As TV_3DVECTOR, v1 As TV_3DVECTOR, v2 As TV_3DVECTOR

FileStart = "Solid " & ObjectName
TotalTriangles = MB.GetTriangleAmount


If Dir(Path) <> "" And CoverOriginFile = True Then Kill Path '删除源文件
Open Path For Binary As #1
Put #1, , FileStart '文件头 80字节
Put #1, , TotalTriangles '三角形数 4字节
For i = 1 To TotalTriangles
       v1 = Math.VSubtract(MB.GetTriangleVertex(i, 2), MB.GetTriangleVertex(i, 1))
       v2 = Math.VSubtract(MB.GetTriangleVertex(i, 3), MB.GetTriangleVertex(i, 1))
       Normal = Math.VCrossProduct(v1, v2) '法向量
       Put #1, , Normal.x 'UP软件是XZY
       Put #1, , Normal.z
       Put #1, , Normal.y
       Put #1, , MB.GetTriangleVertex(i, 2).x
       Put #1, , MB.GetTriangleVertex(i, 2).z
       Put #1, , MB.GetTriangleVertex(i, 2).y
       Put #1, , MB.GetTriangleVertex(i, 1).x
       Put #1, , MB.GetTriangleVertex(i, 1).z
       Put #1, , MB.GetTriangleVertex(i, 1).y
       Put #1, , MB.GetTriangleVertex(i, 3).x
       Put #1, , MB.GetTriangleVertex(i, 3).z
       Put #1, , MB.GetTriangleVertex(i, 3).y
       Put #1, , TriangleEnd '每个三角形的末尾 2字节
Next i
Close #1

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
