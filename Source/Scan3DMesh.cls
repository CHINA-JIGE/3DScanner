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




'����������������ͼƬ����������������
Private TheFirstPicturePath As String '��һ��ͼ��·��

Private Type Type_ScanCamPerPicture
        ScanCamPos As TV_3DVECTOR
       ScanCamVector As TV_3DVECTOR
End Type
Private mScanCamPerPicture() As Type_ScanCamPerPicture

'������������1��������������������
Public Enum CONST_LightSide '������CAM������
Side_Left = 0
Side_Right = 1
End Enum

Public Enum CONST_Sample
Sample_MainColorFilter = 0
Sample_SingleColorChannel = 1
End Enum

Public Type Type_ScanParameters
TurningCenterToWall As Single '��ת���ĵ�������
StandardDepth As Single '��׼��� ���㵽������
VisibleAngleHorizontal As Single '���ӽǦȺ�
VisibleAngleVertical As Single '���ӽǦ���
CamToLight As Single '��ͷ�����ߵľ���
ColorFilter As Single '��ɫɸѡ�½�
End Type

'Private MaincolorLOWERBOUND As Single
'������������2���ɶ��㡪����������





'������������������ˮ�߿�ʼ������������
'
'����MAINMeshBuffer
'
'������������������������������������
Sub SetMainBuffer(NewMeshBuffer As Scan3DMeshBuffer)
Set MB = NewMeshBuffer
End Sub

Function GetMainBuffer() As Scan3DMeshBuffer
Set GetMainBuffer = MB
End Function




Sub LoadScanPicture_START()
        TexF.SetTextureMode TV_TEXTUREMODE_16BITS
End Sub

Sub LoadScanPicture(iPath As String, PictureID As Long) '����ͼƬ
       If PictureID = 1 Then TheFirstPicturePath = iPath '��¼�µ�һ��ͼƬ��·��
       TexF.LoadTexture iPath, CStr(PictureID), PictureWidth, PictureHeight
       TexF.LockTexture GetTex(CStr(PictureID)), True
End Sub

Sub LoadScanPicture_END()
       SampleFromPicture Side_Left, 1  '��ʵ�ǶԵ�һ��ͼƬ���� Ҫ�ص�
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










'��������������������������������������������
'���̣�����ͼƬID��1 to N������GETTEX��i����Ȼ��
'���Ӳ�����SAMPLEPOINT,��ADDLINE
'��������������������������������������������
Sub SampleFromPicture(iLightSide As CONST_LightSide, PictureID As Long)

        Dim PixelTraversalStartX As Long, PixelTraversalEndX As Long '�������صĺ���������
        Dim NumberOfPoint        As Single, TotalPx As Single, RED As Single, VlinepointCounter As Long
        Dim PixelCOLOR As Long, MainColor As Single, Color1 As Single, Color2 As Single

        TotalPx = 0
        NumberOfPoint = 0
       'MaincolorLOWERBOUND = MainColorFilterLowerBound '��¼���½� LOAD_PICTUREENDҪ��
       
        '�жϵƹ���CAM��������
        If iLightSide = Side_Left Then
                PixelTraversalStartX = 1
                PixelTraversalEndX = PictureWidth / 2
        Else
                PixelTraversalStartX = PictureWidth / 2
                PixelTraversalEndX = PictureWidth
        End If
       
       Dim P As Long
       If PictureID = -1 Then '-1�������һ��ͼƬ
       P = GetTex("1")
       Else
       P = GetTex(PictureID) '1 �� N
       End If

       Dim x As Long, y As Long, pcolor As TV_COLOR
                For y = 1 To PictureHeight
                        For x = PixelTraversalStartX To PixelTraversalEndX
                                   
                                PixelCOLOR = TexF.GetPixel(P, x, y)
                                pcolor = DecodeRGBA(PixelCOLOR)
                                MainColor = pcolor.r  'ȡɫ
                                Color1 = pcolor.g
                                Color2 = pcolor.b

                                'ɸѡ��ɫ
                                'If MainColor * (1 - Color1) ^ 2 * (1 - Color2) ^ 2 > MaincolorLOWERBOUND Then
                                 '       TotalPx = TotalPx + x * MainColor
                                '        NumberOfPoint = NumberOfPoint + MainColor
                                'End If
                               'If MainColor > MainColorFilterLowerBound And Color1 < 0.2 And Color2 < 0.2 Then
                               If MainColor > MainColorFilterLowerBound And Color1 < 0.2 And 0.4 < Color2 < 0.55 Then '������ɫ̫�ӵ���
                                     TotalPx = TotalPx + x * MainColor
                                     NumberOfPoint = NumberOfPoint + MainColor
                               End If
                        Next x
                     
                        '���ɲ�����
                        If NumberOfPoint <> 0 Then
                            '////////////////////���غ���ֵ
                                MB.AddSamplingResult Vector2(TotalPx / NumberOfPoint, y)
                                VlinepointCounter = VlinepointCounter + 1
                        End If
                
                        TotalPx = 0
                        NumberOfPoint = 0
                        
                Next y
                
        MB.AddLine VlinepointCounter  'ɨ����һ��ͼƬ ����һ�е�
        VlinepointCounter = 0
        
        '������
        'If PictureID <> GetTex("1") Then TexF.DeleteTexture GetTex(PictureID)
End Sub








'����������������������������3D���㡪����������������������������

Sub Generate3DPointCloud(iLightSide As CONST_LightSide, IsStructed As Boolean)
Select Case IsStructed
       Case True '�ǽṹ����ģʽ
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

        Dim NumOfPictures   As Long '��Ƭ��

        Dim PictureStartID  As Long, PictureEndID As Long

        Dim ProcessingAngle As Single '���ڴ�����Ƭ�ĽǶ�


        PictureStartID = 1 'GetTex("1")
        PictureEndID = MB.GetLineAmount - 1 '���ǵ�Ϊɶ ������ǰд�Ĳ�Ҫ����= =
        NumOfPictures = MB.GetLineAmount - 1
        d = MainScanParam.StandardDepth '��׼���
        cita = MainScanParam.VisibleAngleVertical * 3.1415926 / 180 '���¿��ӽ�
        TurnCenterToWall = MainScanParam.TurningCenterToWall '�ؼ�
        RealPictureH = d * Tan(0.5 * cita) * 2    '���������ֱ��Χ��ʵ�ʳ���  ���Ǹ�z�����й�����..z��Ӱ��y���ƫ�� �ӽ�����
        
        
        
        Dim PROCESSINGPICTUREID As Long, PROCESSINGPIXEL As Long
        
        For PROCESSINGPICTUREID = PictureStartID To PictureEndID + 1
                For PROCESSINGPIXEL = 1 To MB.GetLine(PROCESSINGPICTUREID) '�����ж��ٸ���
 
                       Dim SamplePoint As TV_2DVECTOR '��ȡ������
                       SamplePoint = MB.GetSamplingResult(MB.GetPointID(PROCESSINGPICTUREID, PROCESSINGPIXEL))
                        DEPTH = GetDepthFromOffset(MainScanParam, SamplePoint.x, iLightSide)
                        
                        'x = ��� - �п�һ��
                        LocalX = 0 'CamToLight ///LOCALZΪ������������Ǽ����Բ��

                        LocalY = RealPictureH * ((SamplePoint.y / PictureHeight) + DEPTH / (2 * d) - (DEPTH * SamplePoint.y) / (d * PictureHeight))
        
                       LocalZ = TurnCenterToWall - DEPTH
                       
                       
                        With MatrixLocal '�ֲ����� ���õ�һ����
                                .m11 = LocalX
                                .m21 = LocalY
                                .m31 = LocalZ
                        End With
        
                        '���ڴ�����ͼƬ��������Ƕ�(����ͼ)
                        ProcessingAngle = 2 * 3.1415926 * (PROCESSINGPICTUREID - PictureStartID) / NumOfPictures
       
                        With MatrixTrans '�任���� ��ʵ�Ǹ�ʽ��4X4�Ĳ�����3X3����
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

                        Math.TVMatrixMultiply MatrixWorld, MatrixTrans, MatrixLocal  'Local����任
       
                        x = MatrixWorld.m11
                        y = MatrixWorld.m21
                        z = MatrixWorld.m31

                         MB.AddPoint x, y + RealPictureH / 2, z  '��������

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
        Dim NumOfPictures   As Long '��Ƭ��
        Dim PictureStartID  As Long, PictureEndID As Long


        PictureStartID = 1 'GetTex("1")
        PictureEndID = MB.GetLineAmount '- 1 'GetTex("END") - GetTex("1")
        NumOfPictures = MB.GetLineAmount - 1
        d = MainScanParam.StandardDepth '��׼���
        cita = MainScanParam.VisibleAngleVertical * 3.1415926 / 180 '���¿��ӽ�
        TurnCenterToWall = MainScanParam.TurningCenterToWall '�ؼ�
        RealPictureH = d * Tan(0.5 * cita) * 2    '���������ֱ��Χ��ʵ�ʳ���  ���Ǹ�z�����й�����..z��Ӱ��y���ƫ�� �ӽ�����
        
        
        Dim PROCESSINGPICTUREID As Long, PROCESSINGPIXEL As Long
        For PROCESSINGPICTUREID = PictureStartID To PictureEndID '+ 1
                For PROCESSINGPIXEL = 1 To MB.GetLine(PROCESSINGPICTUREID) '�����ж��ٸ���
 
                       Dim SamplePoint As TV_2DVECTOR '��ȡ������
                       SamplePoint = MB.GetSamplingResult(MB.GetPointID(PROCESSINGPICTUREID, PROCESSINGPIXEL))
                        DEPTH = GetDepthFromOffset(MainScanParam, SamplePoint.x, iLightSide)
                        'x = ��� - �п�һ��
                        LocalX = 0 'CamToLight ///LOCALZΪ������������Ǽ����Բ��
                        LocalY = RealPictureH * ((SamplePoint.y / PictureHeight) + DEPTH / (2 * d) - (DEPTH * SamplePoint.y) / (d * PictureHeight))
                       LocalZ = TurnCenterToWall - DEPTH
                       
                        With MatrixLocal '�ֲ����� ���õ�һ����
                                .m11 = LocalX
                                .m21 = LocalY
                                .m31 = LocalZ
                                .m41 = 1
                        End With
       
                      '����ɢ�ҵ���ʱ�����������
                     Dim ScanCamX As Single, ScanCamY As Single, ScanCamZ As Single
                     Dim ScanCamAxisX As TV_3DVECTOR, ScanCamAxisY As TV_3DVECTOR, ScanCamAxisZ As TV_3DVECTOR
                     '����ɨ���������POSITION
                     ScanCamX = mScanCamPerPicture(PROCESSINGPICTUREID).ScanCamPos.x
                     ScanCamY = mScanCamPerPicture(PROCESSINGPICTUREID).ScanCamPos.y
                     ScanCamZ = mScanCamPerPicture(PROCESSINGPICTUREID).ScanCamPos.z
                     '��׼������
                     'Z����������� ��һ��ʼ��������ű���ǽ�ķ���һ��
                     Dim x3 As Single, y3 As Single, z3 As Single
                     x3 = mScanCamPerPicture(PROCESSINGPICTUREID).ScanCamVector.x
                     y3 = mScanCamPerPicture(PROCESSINGPICTUREID).ScanCamVector.y
                     z3 = mScanCamPerPicture(PROCESSINGPICTUREID).ScanCamVector.z
                     ScanCamAxisZ = Vector3(x3, y3, z3)
                     ScanCamAxisX = Vector3(z3, 0, -x3)
                     ScanCamAxisY = Math.VCrossProduct(ScanCamAxisX, ScanCamAxisZ)
                     '��������һ��
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
                            
                     '��������
                        Math.TVMatrixMultiply MatrixWorld, MatrixTrans, MatrixLocal
       
                        x = MatrixWorld.m11
                        y = MatrixWorld.m21
                        z = MatrixWorld.m31

                         MB.AddPoint x, y + RealPictureH / 2, z  '��������

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










'��������������������������������BEZIER����У������������������������������������������
Sub BezierCurveFittingForPoint(RepeatTimes As Long)
Dim OriginPoint1 As TV_3DVECTOR, OriginPoint2 As TV_3DVECTOR, OriginPoint3 As TV_3DVECTOR
Dim ResultPoint As TV_3DVECTOR
Dim r As Long, i As Long, j As Long
For r = 1 To RepeatTimes
       For i = 1 To MB.GetLineAmount
              For j = 2 To MB.GetLine(i) - 1  'ֻƽ���ڶ��������ڶ�����
              OriginPoint1 = MB.GetPoint(i, j - 1)
              OriginPoint2 = MB.GetPoint(i, j)
              OriginPoint3 = MB.GetPoint(i, j + 1)

              ResultPoint = BezierInterpolation(0.5, OriginPoint1, OriginPoint2, OriginPoint3) 'bezier������GLOBALģ��
              MB.SetPoint MB.GetPointID(i, j), ResultPoint.x, ResultPoint.y, ResultPoint.z
              Next j
       Next i
Next r

ResetRenderMesh
MB.AddPointCloudToRenderMesh
End Sub










'�����������������������������γɱպ��塪����������������������

Sub GenerateClosedObject()
       
        DoEvents
        '���������������������������������������������Ρ���������������������������
        
        Dim MinOfTotalPoint As Long

        Dim p1              As TV_3DVECTOR, p2 As TV_3DVECTOR, p3 As TV_3DVECTOR

       Dim i As Long, j As Long
        For i = 1 To MB.GetLineAmount - 1
        
                MinOfTotalPoint = MIN(MB.GetLine(i), MB.GetLine(i + 1))   '���е�˭�ٵ�
                For j = 1 To MAX(MB.GetLine(i), MB.GetLine(i + 1))

                        Select Case Val(j)
                                Case Is < MinOfTotalPoint  '���ж�����˭�ĵ�����
                                        p1 = MB.GetPoint(i, j)   'Line1FirstPointID
                                        p2 = MB.GetPoint(i + 1, j)  'Line2FirstPointID + j
                                        p3 = MB.GetPoint(i, j + 1) 'Line1FirstPointID + j + 1
                                        MB.AddTriangleByVector p1, p2, p3
                                        
                                        p1 = MB.GetPoint(i + 1, j)
                                        p2 = MB.GetPoint(i + 1, j + 1)
                                        p3 = MB.GetPoint(i, j + 1)
                                        MB.AddTriangleByVector p1, p2, p3
                                        
                                        
                                Case Is > MinOfTotalPoint '��Ӧ�����굽������ĵ���
                                        If MB.GetLine(i) > MB.GetLine(i + 1) Then
                                                '��һ�еĵ��ڶ��е��ٵ�ʱ��
                                                p1 = MB.GetPoint(i, j - 1)
                                                p2 = MB.GetPoint(i + 1, MinOfTotalPoint)  '�ڶ�����͵�������
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
       
       
        '���������������������ⶥ���ס���������������������������
        Dim TotalVLine As Long
        TotalVLine = MB.GetLineAmount - 1 '������
        
        For i = 1 To Int((MB.GetLineAmount - 1) / 2) - 1 'int������ȡ��  ע�⣺��һ�к����һ�����ص��ģ����Լ�1
                '������
                p1 = MB.GetPoint(i, 1)
                p2 = MB.GetPoint(TotalVLine - i + 1, 1)
                p3 = MB.GetPoint(TotalVLine - i, 1)
                MB.AddTriangleByVector p1, p2, p3
        
                p1 = MB.GetPoint(i, 1)
                p2 = MB.GetPoint(i + 1, 1)
                p3 = MB.GetPoint(TotalVLine - i, 1)
                MB.AddTriangleByVector p2, p1, p3
       
                '������
                p1 = MB.GetPoint(i, MB.GetLine(i))
                p2 = MB.GetPoint(TotalVLine - i + 1, MB.GetLine(TotalVLine - i + 1))
                p3 = MB.GetPoint(TotalVLine - i, MB.GetLine(TotalVLine - i))
                MB.AddTriangleByVector p2, p1, p3
        
                p1 = MB.GetPoint(i, MB.GetLine(i))
                p2 = MB.GetPoint(i + 1, MB.GetLine(i + 1))
                p3 = MB.GetPoint(TotalVLine - i, MB.GetLine(TotalVLine - i))
                MB.AddTriangleByVector p1, p2, p3 '���㴢��˳������UP������ʾ�޸ĵ�= =
       Next i
        
        
        
        Dim BeforeHalf As Long
        BeforeHalf = Int((TotalVLine - 1) / 2)
        If Int(NumOfVerticalLines - 1 / 2) Mod 2 = 1 Then '��ȥ����ظ��к�   ����������
                p1 = MB.GetPoint(BeforeHalf, 1)  '��Ȧǰ�����һ����
                p2 = MB.GetPoint(BeforeHalf + 1, 1)
                p3 = MB.GetPoint(BeforeHalf + 2, 1)
                MB.AddTriangleByVector p1, p2, p3
        
                p1 = MB.GetPoint(BeforeHalf, MB.GetLine(BeforeHalf))  '��Ȧǰ�����һ����
                p2 = MB.GetPoint(BeforeHalf + 1, MB.GetLine(BeforeHalf + 1))
                p3 = MB.GetPoint(BeforeHalf + 2, MB.GetLine(BeforeHalf + 2))
                MB.AddTriangleByVector p1, p2, p3
       End If

                '������������������������������������������������������������������������������������������������
       MB.WeldVertices
End Sub






'������������������������������������STL����������������������������
Sub SaveSTL(ObjectName As String, Path As String, CoverOriginFile As Boolean)
Dim i As Long
Dim FileStart As String * 80
Dim TotalTriangles As Long '
Dim TriangleEnd As String * 2
Dim Normal As TV_3DVECTOR, v1 As TV_3DVECTOR, v2 As TV_3DVECTOR

FileStart = "Solid " & ObjectName
TotalTriangles = MB.GetTriangleAmount


If Dir(Path) <> "" And CoverOriginFile = True Then Kill Path 'ɾ��Դ�ļ�
Open Path For Binary As #1
Put #1, , FileStart '�ļ�ͷ 80�ֽ�
Put #1, , TotalTriangles '�������� 4�ֽ�
For i = 1 To TotalTriangles
       v1 = Math.VSubtract(MB.GetTriangleVertex(i, 2), MB.GetTriangleVertex(i, 1))
       v2 = Math.VSubtract(MB.GetTriangleVertex(i, 3), MB.GetTriangleVertex(i, 1))
       Normal = Math.VCrossProduct(v1, v2) '������
       Put #1, , Normal.x 'UP������XZY
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
       Put #1, , TriangleEnd 'ÿ�������ε�ĩβ 2�ֽ�
Next i
Close #1

End Sub



Sub AddToLog(aText As String)
TV.AddToLog aText
End Sub
Sub SetScanParameters(SP As Type_ScanParameters)
MainScanParam = SP
MainColorFilterLowerBound = SP.ColorFilter '��ɫɸѡ�½�
End Sub
Sub AppSleep(mSec As Long)
Sleep (mSec)
End Sub