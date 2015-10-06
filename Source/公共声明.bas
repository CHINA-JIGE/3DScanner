Attribute VB_Name = "公共过程与声明"
'————————————API——————————————
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'————————————Constant——————————
Public Const PictureWidth = 640
Public Const PictureHeight = 480

'————————————Var————————
Public RenderLevel As Integer

Public NumOfVerticalLines As Long '竖列总数

Public NumOfPointVerticalLine() As Long '每列多少个点 用于连接点为三角形  在主程序定义

Public Point3D() As TV_3DVECTOR


'————————————Type————————————
Public Type CameraParameters

        VisibleAngleHorizontal As Single '可视角θ横
        VisibleAngleVertical As Single '可视角θ竖
        StandardLength As Single '标准深度 焦点到背景板
        CamToLight As Single '镜头到红线的距离

End Type

'————————————Enum————————————

Public Enum LightSideType
       Side_Left = 0
        Side_Right = 1
End Enum

Public Enum SaveFileType
FILE_STL = 0
FILE_ASC_POINT_CLOUD = 1
End Enum



Public Sub ——————Init——————()
       'Set OutList = New OutputTriangleList
        CamParam1.StandardLength = Val(Form2.Text1) 'mm
        CamParam1.VisibleAngleHorizontal = (Form2.Text2 / 360) * 2 * 3.1415926
        CamParam1.VisibleAngleVertical = (Form2.Text6 / 360) * 2 * 3.1415926
        CamParam1.CamToLight = Form2.Text3


        'PictureWidth = 480 '加载图片的尺寸（像素）
        'PictureHeight = 320
        
        InitTV
       
       Form2.Show
       Form1.Show '渲染窗 里面有主循环
End Sub





Public Sub SavePointCloud()
        Dim x As Single, y As Single, z As Single, nx As Single, ny As Single, nz As Single, tu1 As Single, tv1 As Single, tu2 As Single, tv2 As Single, c As Long
        Open App.Path & "\Output.txt" For Append As #2
        For a = 0 To mesh.GetVertexCount - 1
                mesh.GetVertex a, x, y, z, nx, ny, nz, tu1, tv1, tu2, tv2, c
                Print #2, x; "," & y & "," & z
        Next a
        Close #2
       
       Form2.List1.AddItem "保存点云完成..."
        MsgBox "ASCII点云储存完成！", vbInformation
End Sub

Public Sub SaveSTL()
Dim FileStart As String * 80
Dim TotalTriangles As Long '
Dim TriangleEnd As String * 2
Dim Normal As TV_3DVECTOR, v1 As TV_3DVECTOR, v2 As TV_3DVECTOR

FileStart = "Solid OBJECT1"
TotalTriangles = Mesh2.GetTriangleCount

If Dir(App.Path & "\Output.stl") <> "" Then Kill App.Path & "\Output.stl"
Open App.Path & "\Output.stl" For Binary As #1
Put #1, , FileStart '文件头 80字节
Put #1, , TotalTriangles '三角形数 4字节
For i = 1 To OutList.GetUpperBound
       v1 = Math.VSubtract(OutList.GetVertex(i, 2), OutList.GetVertex(i, 1))
       v2 = Math.VSubtract(OutList.GetVertex(i, 3), OutList.GetVertex(i, 1))
       Normal = Math.VCrossProduct(v1, v2) '法向量
       Put #1, , Normal.x 'UP软件是XZY
       Put #1, , Normal.z
       Put #1, , Normal.y
       Put #1, , OutList.GetVertex(i, 2).x
       Put #1, , OutList.GetVertex(i, 2).z
       Put #1, , OutList.GetVertex(i, 2).y
       Put #1, , OutList.GetVertex(i, 1).x
       Put #1, , OutList.GetVertex(i, 1).z
       Put #1, , OutList.GetVertex(i, 1).y
       Put #1, , OutList.GetVertex(i, 3).x
       Put #1, , OutList.GetVertex(i, 3).z
       Put #1, , OutList.GetVertex(i, 3).y
       Put #1, , TriangleEnd '每个三角形的末尾 2字节
Next i
Close #1

Form2.List1.AddItem "stl保存完成..."
MsgBox ".stl文件储存完成！", vbInformation
End Sub






Public Function MIN(Value1 As Single, Value2 As Single) As Single

If Value1 <= Value2 Then
MIN = Value1
Else
MIN = Value2
End If

End Function


Public Function MAX(Value1 As Variant, Value2 As Variant) As Single

If Value1 >= Value2 Then
MAX = Value1
Else
MAX = Value2
End If

End Function



Public Function GetPointID(iVerticalLine As Long, HorizontalID As Long) As Long '本质上是在遍历线性表
Dim ID As Long
'iVerticalLine 是 1 to N
If iVerticalLine > 1 Then
For i = 1 To iVerticalLine - 1 '-1是因为不用加上当前列的点数
ID = ID + NumOfPointVerticalLine(i)
Next i
End If

GetPointID = ID + HorizontalID - 1 'TV3D的MESH 顶点ID从0开始
End Function
