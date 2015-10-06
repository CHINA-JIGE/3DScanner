Attribute VB_Name = "cGenVertexModule"
'Public TopPoint As TV_3DVECTOR, BottomPoint As TV_3DVECTOR '封顶 封底点



Private MatrixTrans As TV_3DMATRIX
Private MatrixLocal As TV_3DMATRIX
Private MatrixWorld As TV_3DMATRIX

Private VectorWorldOffset As TV_3DVECTOR '
Private VectorStart As TV_3DVECTOR '背景板的世界坐标点

Public Sub GenerateVertex(ProcessingPictureID As Long, ProcessingPixelY As Long)

        Dim LocalX As Single, LocalY As Single, LocalZ As Single

        Dim x             As Single, y As Single, z As Single, RealPictureH As Single, DEPTH As Single

        Dim d             As Single, cita As Single, LightDistance As Single, TurnCenterToWall As Single

        Dim NumOfPictures As Long '照片数
        
        Dim ProcessingAngle As Single '正在处理照片的角度


        PictureStartID = Val(Form2.Text_START)
        d = CamParam1.StandardLength '标准深度
        cita = CamParam1.VisibleAngleVertical '上下可视角
        NumOfPictures = Val(Form2.Text_END - Form2.Text_START + 1)
        TurnCenterToWall = Val(Form2.Text_CenterToWall) '必加
        RealPictureH = d * Tan(0.5 * cita) * 2    '求出可视竖直范围的实际长度  但是跟z坐标有关联的..z会影响y轴的偏移 视角问题
        DEPTH = GetDepthFromOffset(CamParam1, SamplingPx(ProcessingPixelY), Side_Left)
        
        
       'x = 深度 - 盒宽一半
        LocalX = TurnCenterToWall - DEPTH
        
        LocalZ = 0 'CamToLight
        ' y = WorldY 不需要变换
        LocalY = -RealPictureH * ((ProcessingPixelY / PictureHeight) + DEPTH / (2 * d) - (DEPTH * ProcessingPixelY) / (d * PictureHeight)) '上半段公式
        'Else上下半段公式一样
        'y = -RealPictureH * ((ProcessingPixelY / PictureHeight) + z / (2 * d) - (z * ProcessingPixelY) / (d * PictureHeight))
        'End If
        

        With MatrixLocal '局部坐标 就用第一列了
                .m11 = LocalX
                .m21 = LocalY
                .m31 = LocalZ
        End With
        
       '正在处理的图片的摄像机角度(俯视图)
       ProcessingAngle = 2 * 3.1415926 * (ProcessingPictureID - Val(Form2.Text_START)) / NumOfPictures
       
       With MatrixTrans '变换矩阵 其实是格式是4X4的不过用3X3够了
              .m11 = Cos(ProcessingAngle)
              .m12 = 0
              .m13 = -Sin(ProcessingAngle)
              .m21 = 0
              .m22 = 1
              .m23 = 0
              .m31 = Sin(ProcessingAngle)
              .m32 = 0
              .m33 = Cos(ProcessingAngle)
       End With

       Math.TVMatrixMultiply MatrixWorld, MatrixTrans, MatrixLocal  'Local坐标变换
       
       x = MatrixWorld.m11
       y = MatrixWorld.m21
       z = MatrixWorld.m31


       mesh.AddVertex x, y + RealPictureH / 2, z, -Cos(ProcessingAngle), 0, -Sin(ProcessingAngle), 0, 0 '新增顶点

End Sub




'Public Sub GenerateTopAndBottomCenterPoint() '用于封顶和封底
'        Dim x As Single, y As Single, z As Single, nx As Single, ny As Single, nz As Single, tu1 As Single, tv1 As Single, tu2 As Single, tv2 As Single, c As Long
 '      Dim LineFirstPointID As Long, LineLastPointID As Long
'
'       LineFirstPointID = 0
'       LineLastPointID = NumOfPointVerticalLine(1) - 1
''       For i = 1 To NumOfVerticalLines
'       mesh.GetVertex LineFirstPointID, x, y, z, nx, ny, nz, tu1, tv1, tu2, tv2, c
 '      TopPoint.y = TopPoint.y + y
 '      mesh.GetVertex LineLastPointID, x, y, z, nx, ny, nz, tu1, tv1, tu2, tv2, c
 '      BottomPoint.y = BottomPoint.y + y
'       '更新处理ID
 '      LineFirstPointID = LineLastPointID + 1 '
 '      LineLastPointID = LineLastPointID + NumOfPointVerticalLine(i + 1) '本列最后顶点ID+下列顶点数
 '      Next i
'
'TopPoint = Vector3(0, TopPoint.y / NumOfVerticalLines, 0) '封顶点
'BottomPoint = Vector3(0, BottomPoint.y / NumOfVerticalLines, 0) '封底点

'End Sub






