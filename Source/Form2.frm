VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "3D扫描 V1.12 - 参数设置"
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   8010
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   8010
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command_保存点云 
      Caption         =   "保存点云"
      Enabled         =   0   'False
      Height          =   735
      Left            =   2520
      TabIndex        =   40
      Top             =   6960
      Width           =   855
   End
   Begin VB.TextBox Text_IdenLine 
      Height          =   270
      Left            =   360
      TabIndex        =   39
      Text            =   "40"
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CommandButton Command_曲面细分 
      Caption         =   "曲面细分"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2520
      TabIndex        =   37
      Top             =   4800
      Width           =   1935
   End
   Begin VB.CommandButton Command_取消生成顶点 
      Caption         =   "点云生成取消"
      Height          =   735
      Left            =   5400
      TabIndex        =   36
      Top             =   6960
      Width           =   1815
   End
   Begin VB.CommandButton Command_清空网格 
      Caption         =   "清空网格"
      Enabled         =   0   'False
      Height          =   855
      Left            =   2520
      TabIndex        =   35
      Top             =   6000
      Width           =   855
   End
   Begin VB.CommandButton Command_测距 
      Caption         =   "测距"
      Height          =   420
      Left            =   2520
      TabIndex        =   34
      Top             =   2400
      Width           =   1935
   End
   Begin VB.TextBox Text_Bezier 
      Height          =   270
      Left            =   360
      TabIndex        =   33
      Text            =   "2"
      Top             =   7320
      Width           =   1575
   End
   Begin VB.CommandButton Command_Bezier 
      Caption         =   "Bezier点云平滑"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2520
      TabIndex        =   31
      Top             =   3600
      Width           =   1935
   End
   Begin VB.TextBox Text_ScrShotPath 
      Height          =   270
      Left            =   360
      TabIndex        =   29
      Text            =   "d:\1.jpg"
      Top             =   7920
      Width           =   1575
   End
   Begin VB.CommandButton Command_网格重构 
      Caption         =   "网格重构"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2520
      TabIndex        =   28
      Top             =   4200
      Width           =   1935
   End
   Begin VB.CommandButton Command_清空主缓存 
      Caption         =   "清空主缓存"
      Enabled         =   0   'False
      Height          =   855
      Left            =   3600
      TabIndex        =   27
      Top             =   6000
      Width           =   855
   End
   Begin VB.TextBox Text_CenterToWall 
      Height          =   270
      Left            =   360
      TabIndex        =   26
      Text            =   "0"
      Top             =   6000
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "渲染状态"
      Height          =   1935
      Left            =   2400
      TabIndex        =   21
      Top             =   360
      Width           =   2055
      Begin VB.OptionButton Option1 
         Caption         =   "采样渲染"
         Height          =   180
         Left            =   120
         TabIndex        =   24
         Top             =   480
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         Caption         =   "3D闭合体渲染"
         Height          =   300
         Left            =   120
         TabIndex        =   23
         Top             =   1080
         Width           =   1575
      End
      Begin VB.OptionButton Option3 
         Caption         =   "3D点云渲染"
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   720
         Width           =   1575
      End
   End
   Begin VB.CommandButton Command_保存STL 
      Caption         =   "保存STL"
      Enabled         =   0   'False
      Height          =   720
      Left            =   3600
      TabIndex        =   20
      Top             =   6960
      Width           =   855
   End
   Begin VB.TextBox Text_LowerBound 
      Height          =   270
      Left            =   360
      TabIndex        =   19
      Text            =   "0.2"
      Top             =   6720
      Width           =   1575
   End
   Begin VB.TextBox Text6 
      Height          =   270
      Left            =   360
      TabIndex        =   17
      Text            =   "46.8"
      Top             =   1800
      Width           =   1455
   End
   Begin VB.TextBox Text7 
      Height          =   270
      Left            =   360
      TabIndex        =   15
      Text            =   ".png"
      Top             =   5280
      Width           =   1455
   End
   Begin VB.TextBox Text_END 
      Height          =   270
      Left            =   360
      TabIndex        =   13
      Text            =   "99"
      Top             =   4680
      Width           =   1455
   End
   Begin VB.TextBox Text_START 
      Height          =   270
      Left            =   360
      TabIndex        =   12
      Text            =   "10"
      Top             =   4320
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      Height          =   270
      Left            =   360
      TabIndex        =   10
      Text            =   "00"
      Top             =   3720
      Width           =   1455
   End
   Begin VB.ListBox List1 
      Height          =   6060
      IntegralHeight  =   0   'False
      ItemData        =   "Form2.frx":0000
      Left            =   4920
      List            =   "Form2.frx":0007
      TabIndex        =   7
      Top             =   720
      Width           =   2775
   End
   Begin VB.CommandButton Command_生成点云 
      Caption         =   "生成点云"
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   3000
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Left            =   360
      TabIndex        =   5
      Text            =   "20"
      Top             =   2400
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   360
      TabIndex        =   3
      Text            =   "60"
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   360
      TabIndex        =   0
      Text            =   "65"
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label11 
      Caption         =   "标识正方形边长(mm):"
      Height          =   255
      Left            =   240
      TabIndex        =   38
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label Label14 
      Caption         =   "Bezier迭代次数："
      Height          =   255
      Left            =   360
      TabIndex        =   32
      Top             =   7080
      Width           =   1575
   End
   Begin VB.Label Label9 
      Caption         =   "截图保存路径："
      Height          =   255
      Left            =   360
      TabIndex        =   30
      Top             =   7680
      Width           =   1575
   End
   Begin VB.Label Label8 
      Caption         =   "中心到底板(mm):"
      Height          =   255
      Left            =   360
      TabIndex        =   25
      Top             =   5760
      Width           =   1695
   End
   Begin VB.Label Label12 
      Caption         =   "颜色筛选下界(0-1)："
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   6480
      Width           =   1815
   End
   Begin VB.Label Label10 
      Caption         =   "镜头可视角竖(角度)："
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   2160
      X2              =   2160
      Y1              =   120
      Y2              =   7920
   End
   Begin VB.Label Label7 
      Caption         =   "相片后缀："
      Height          =   255
      Left            =   360
      TabIndex        =   14
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "相片前缀："
      Height          =   255
      Left            =   360
      TabIndex        =   11
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "相片编号跨度："
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "状态栏："
      Height          =   255
      Left            =   4920
      TabIndex        =   8
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "CamToLight(mm):"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "镜头可视角横(角度)："
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "标准深度(mm)："
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   1455
   End
   Begin VB.Menu Menu_文件 
      Caption         =   "文件.."
      Begin VB.Menu Menu_加载文件 
         Caption         =   "加载tvm文件到点云"
      End
   End
   Begin VB.Menu Menu_硬件 
      Caption         =   "硬件.."
      Begin VB.Menu Menu_拍照 
         Caption         =   "定时拍照开始..."
      End
   End
   Begin VB.Menu Menu_预设数据 
      Caption         =   "预设数据.."
      Begin VB.Menu 预设_原型机参数 
         Caption         =   "佳能1100D&原型机参数"
      End
      Begin VB.Menu 预设_3DMAX圆柱形 
         Caption         =   "3DMAX模拟_圆柱形扫描参数"
      End
      Begin VB.Menu 预设_3DMAX散乱 
         Caption         =   "3DMAX模拟_散乱点云扫描参数"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Private PipeLine As Scan3DProcessingPipeLine
Private Pipeline2 As Scan3DProcessingPipeline2
Private Scene As Scan3DScene
Private Prepare As Scan3DPreparation
Private ScanPARAM As Type_ScanParameters
Private RENDERLEVEL As Integer
Private CancelGeneration As Boolean



Private Sub Command_测距_Click()
'MsgBox Hardware.InitSerialPort(5, True)
Command_生成点云.Enabled = True
Prepare.AnalyzeIdentificationPoint
ScanPARAM.StandardDepth = Prepare.GetStandardDepth
ScanPARAM.VisibleAngleHorizontal = Prepare.GetVisibleAngleHorizontal
ScanPARAM.VisibleAngleVertical = Prepare.GetVisibleAngleVertical
Text1 = ScanPARAM.StandardDepth
Text2 = ScanPARAM.VisibleAngleHorizontal
Text6 = ScanPARAM.VisibleAngleVertical
End Sub


Private Sub Command_清空网格_Click()
Pipeline2.ClearMainMeshBuffer
Pipeline2.ResetRenderMesh
Pipeline2.AddPointCloudToRenderMesh
End Sub

Private Sub Command_取消生成顶点_Click() '取消生成顶点
CancelGeneration = True
List1.AddItem "点云生成取消......."
End Sub


Private Sub Form_Load()
       'PictureWidth = 1920
       'PictureHeight = 1280
       Form1.Show
       
       Set Pipeline2 = New Scan3DProcessingPipeline2
       Set Scene = New Scan3DScene
       Set Prepare = New Scan3DPreparation

        '——————————初始化———————————

       ScanPARAM.CamToLight = Val(Text3)
       ScanPARAM.StandardDepth = Val(Text1)
       ScanPARAM.VisibleAngleVertical = Val(Text6)
       ScanPARAM.VisibleAngleHorizontal = Val(Text2)
       ScanPARAM.ColorFilter = Val(Text_LowerBound)
       ScanPARAM.IdentificationLineLength = Val(Text_IdenLine)
        '————————————————————————
       Me.Show
       Scene.Init 320, 240, 640, 480, Form1.hWnd
       'Form1.Width = 720 * Screen.TwipsPerPixelX '渲染窗口大小固定
       'Form1.Height = 480 * Screen.TwipsPerPixelY
       'Scene.ResizeRenderWindow
       
Do
DoEvents
Select Case RENDERLEVEL
Case 1
Scene.SimulateMovement
Scene.Render3D RENDER_POINTCLOUD
Case 2
Scene.SimulateMovement
Scene.Render3D RENDER_MESHLINE
End Select
Loop

End Sub




Private Sub Command_生成点云_Click() '生成顶点
        ScanPARAM.CamToLight = Val(Text3)
        ScanPARAM.StandardDepth = Val(Text1)
        ScanPARAM.VisibleAngleVertical = Val(Text6)
        ScanPARAM.VisibleAngleHorizontal = Val(Text2)
        ScanPARAM.ColorFilter = Val(Text_LowerBound)
        ScanPARAM.IdentificationLineLength = Val(Text_IdenLine)
       Pipeline2.SetScanParameters ScanPARAM
       
        Option1.Value = True '采样渲染
        List1.Clear

        '————————————流水线——————————————
 
        List1.AddItem "初始化完成..."
        List1.Refresh

        Dim LPath As String, PStart As Long, PEnd As Long, prefix As String, suffix As String
        PStart = Val(Text_START)
        PEnd = Val(Text_END)
        prefix = Text4
        suffix = Text7
        '有多少张图
       Pipeline2.SetPictureCount PEnd - PStart + 1
       
       Dim i As Long
        For i = PStart To PEnd '加载图片 & 采样
        DoEvents
                LPath = App.Path & "\Group\" & prefix & CStr(i) & suffix
                If Dir(LPath) = "" Then GoTo Err: '路径不存在就调到最后
                If CancelGeneration = True Then
                Pipeline2.ClearMainBuffer
                Pipeline2.ClearPictureBuffer
                CancelGeneration = False
                GoTo ex:
                End If
                
                Pipeline2.LoadScanPicture LPath, i - PStart + 1
                
                '结果加入MB里
                Pipeline2.SampleFromPicture Side_Left, i - PStart + 1
                
                '给出传感器的参数（转角）
                '仰角太大容易没有根
                Pipeline2.SetScanCameraPerPicture i - PStart + 1, Vector3(0, 0, 0), , , , -3.1415926 * 3 / 18, 0, 0
                
                '根据图片的标识正方形计算cam坐标和EulerY
                Pipeline2.ComputeCamPosAndAngleY i - PStart + 1
                
                Scene.RenderSampling i - PStart + 1, RGBA(0, 1, 0, 1)
                
                List1.AddItem "处理图片完成.." & i
                List1.ListIndex = List1.ListCount - 1
        Next i
        
        
        List1.AddItem "图片采样完成...."
        Pipeline2.Generate3DPointCloud Side_Left, Matrix_Euler '生成顶点
        Pipeline2.AddPointCloudToRenderMesh
        Pipeline2.WeldVertices_RenderMesh
        '——————————————————————————
        'PipeLine.ClearPictureBuffer
        List1.AddItem "网格顶点数: " & Pipeline2.GetMainBuffer.GetPointAmount '任务信息
        List1.ListIndex = List1.ListCount - 1
        Command_清空网格.Enabled = True
        Command_清空主缓存.Enabled = True
        Command_网格重构.Enabled = True
        Command_Bezier.Enabled = True
        Command_保存点云.Enabled = True
        Beep
        Option3.Value = True
        Scene.MoveAndLookatMesh
        Command_生成点云.Enabled = False


        GoTo ex:
        '————加载图片错误的处理————
Err:
        MsgBox "图片路径不存在！！", vbCritical
        '————————————————
ex:

End Sub




Private Sub Command_保存STL_Click() '储存文件
If MsgBox("是否保存STL??", vbYesNo, "保存文件") = vbYes Then
Pipeline2.SaveSTL "OBJECT1", App.Path & "\OUTPUT.STL", True
MsgBox "保存完成！", vbOKOnly, "完成！"
List1.AddItem "STL文件保存完成..."
List1.ListIndex = List1.ListCount - 1
End If
Beep
End Sub


Private Sub Command_保存点云_Click()
If MsgBox("是否保存ASCII点云??", vbYesNo, "保存文件") = vbYes Then
Pipeline2.SaveAsciiPointCloud App.Path & "\OUTPUT.txt", True
List1.AddItem "ASCII点云保存完成..."
List1.ListIndex = List1.ListCount - 1
MsgBox "保存完成！", vbOKOnly, "完成！"
End If
Beep
End Sub


Private Sub Command_清空主缓存_Click() '清空
Pipeline2.ClearMainBuffer
Pipeline2.ClearPictureBuffer
Command_生成点云.Enabled = True
Command_网格重构.Enabled = False
Command_保存STL.Enabled = False
Command_取消生成顶点.Enabled = False
Command_清空网格.Enabled = False
Command_Bezier.Enabled = False
Command_保存点云.Enabled = False
End Sub



Private Sub Command_网格重构_Click() '闭合体
List1.AddItem "开始生成三角形面片..."
Pipeline2.MeshReconstruction RC_MappedBall
Pipeline2.AddTriangleToRenderMesh
List1.AddItem "完成..."
List1.ListIndex = List1.ListCount - 1
Option2.Value = True
RENDERLEVEL = 2
Command_保存STL.Enabled = True '保存文件
Command_Bezier.Enabled = False 'BEZIER
Command_网格重构.Enabled = False
Command_清空网格.Enabled = False '清空网格
End Sub





Private Sub Menu_加载文件_Click()

       ScanPARAM.CamToLight = Val(Text3)
        ScanPARAM.StandardDepth = Val(Text1)
        'ScanPARAM.TurningCenterToWall = Val(Text_CenterToWall)
        ScanPARAM.VisibleAngleVertical = Val(Text6)
        ScanPARAM.VisibleAngleHorizontal = Val(Text2)
       ScanPARAM.ColorFilter = Val(Text_LowerBound)
      Pipeline2.SetScanParameters ScanPARAM
       Pipeline2.LoadtvmToPointCloud App.Path & "\Mesh\1.tvm"
       Scene.MoveAndLookatMesh
       Option3.Value = True '点云渲染
       Command_测距.Enabled = False
       Command_生成点云.Enabled = False
      Command_清空网格.Enabled = True
      Command_清空主缓存.Enabled = True
       Command_Bezier.Enabled = True
       Command_网格重构.Enabled = True
End Sub

Private Sub Menu_拍照_Click()
Prepare.InitSerialPort 0, True
Prepare.TakePhoto_START
Pipeline2.AppSleep (60000)
Prepare.TakePhoto_END
End Sub

Private Sub Option1_Click()
RENDERLEVEL = 0 '不管
End Sub

Private Sub Option2_Click()
RENDERLEVEL = 2 '闭合体
End Sub

Private Sub Option3_Click() '3D渲染
RENDERLEVEL = 1 '点云
End Sub







Private Sub Form_Unload(Cancel As Integer)
Set Pipeline2 = Nothing
Set Scene = Nothing
End
End Sub



Private Sub 预设_3DMAX散乱_Click()
Text1 = "65" 'standardlength
Text2 = "60" '横
Text6 = "46.8" '竖角
Text3 = "20" 'camtolight
Text4 = "00" '前缀
Text_START = "10"
Text_END = "99"
Text7 = ".png"
Text_LowerBound = "0.2"
Text_IdenLine = "20"
End Sub

Private Sub 预设_3DMAX圆柱形_Click()
Text1 = "65" 'standardlength
Text2 = "90" '横
Text6 = "73.4" '竖角
Text3 = "20" 'camtolight
Text4 = "00" '前缀
Text_START = "10"
Text_END = "99"
Text7 = ".png"
Text_CenterToWall = "0"
Text_LowerBound = "0.17"
End Sub

Private Sub 预设_原型机参数_Click()
Text1 = "" 'standardlength
Text2 = "" '横
Text6 = "" '竖角
Text3 = "95" 'camtolight
Text4 = "IMG_" '前缀
Text_START = "5619"
Text_END = "5779"
Text7 = ".JPG"
Text_CenterToWall = "0"
Text_LowerBound = "0.6"
End Sub
