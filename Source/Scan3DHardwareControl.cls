VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Scan3DPreparation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private MSComm1 As MSComm   'MSCOMM类
Private JPGInfoReader As ExifReader 'JPG文件的EXIF  只获取其中的焦距信息

'——————————测距————————————
Public Type Type_PointWithID
IsQualified As Integer
gID As Integer
End Type

Private IdenPointCount As Integer '标识点个数
Private IdenPoint() As TV_2DVECTOR '标识点
Private RealVisibleWidth As Single '可视宽度
Private RealVisibleHeight As Single '可视高度

Sub AnalyzeIdentificationPoint()
'由于分析一个图片里面的标识点可能是一坨像素，但是我们只需要一个矢量点
'所以每个标识点就当成一个组
Dim pixel() As Type_PointWithID, SamplePoint() As TV_2DVECTOR
Dim GROUP_POINTER As Integer, CanCreateGroup As Boolean
GROUP_POINTER = 0
Dim PW As Long, PH As Long
PW = PictureWidth
PH = PictureHeight
ReDim pixel(1 To PW, 1 To PH) As Type_PointWithID
'——————加载图片————————
TexF.LoadTexture App.Path & "\Group\Measure.jpg", "Measure", PW, PH
TexF.LockTexture GetTex("Measure")
'——————选点————————
Dim tmpColor As Long, tmpC As TV_COLOR, i As Long, j As Long
For i = 1 To PW
       For j = 1 To PH
       tmpColor = TexF.GetPixel(GetTex("Measure"), i, j)
       tmpC = DecodeRGBA(tmpColor)
       If tmpC.r > 0.7 And tmpC.g < 0.4 And tmpC.b < 0.4 Then '红
       pixel(i, j).IsQualified = 1                                                  '像素偏向？？ 就是选中的点
       End If
       '.gid = 0
       Next j
Next i

'——————

For i = 2 To PW - 1
       For j = 2 To PH - 1
       If pixel(i, j).gID = 0 And pixel(i, j).IsQualified = 1 Then '还没有分到组的拿去处理
              CanCreateGroup = True
               '左上方有“朋友”就不用新建一个组了
              If pixel(i - 1, j - 1).IsQualified = 1 And pixel(i - 1, j - 1).gID <> 0 Then
              pixel(i, j).gID = pixel(i - 1, j - 1).gID
              CanCreateGroup = False
              End If
              If pixel(i - 1, j).IsQualified = 1 And pixel(i - 1, j).gID <> 0 Then
              pixel(i, j).gID = pixel(i - 1, j).gID
              CanCreateGroup = False
              End If
              If pixel(i, j - 1).IsQualified = 1 And pixel(i, j - 1).gID <> 0 Then
              pixel(i, j).gID = pixel(i, j - 1).gID
              CanCreateGroup = False
              End If
              If pixel(i - 1, j + 1).IsQualified = 1 And pixel(i - 1, j + 1).gID <> 0 Then
              pixel(i, j).gID = pixel(i - 1, j + 1).gID
              CanCreateGroup = False
              End If
              If CanCreateGroup = True Then                                   '需要新建组
              GROUP_POINTER = GROUP_POINTER + 1
              pixel(i, j).gID = GROUP_POINTER                                   '新建组的ID
              End If
       End If
       Next j
Next i

ReDim IdenPoint(1 To GROUP_POINTER) As TV_2DVECTOR
Dim GroupPointCount() As Long
ReDim GroupPointCount(1 To GROUP_POINTER) As Long                              '组内的点数
For i = 1 To PW
       For j = 1 To PH
       If pixel(i, j).gID <> 0 Then                                                                             '如果像素已经分组
       IdenPoint(pixel(i, j).gID).x = IdenPoint(pixel(i, j).gID).x + Val(i)             '像素的坐标累加 之后一除就得到平均点坐标了
       IdenPoint(pixel(i, j).gID).y = IdenPoint(pixel(i, j).gID).y + Val(j)
       GroupPointCount(pixel(i, j).gID) = GroupPointCount(pixel(i, j).gID) + 1 '统计组内点数
       End If
       Next j
Next i



'————————画标识点——————
TV.Clear
scr.Draw_Texture GetTex("Measure"), 0, 0, WindowWidth, WindowHeight
For i = 1 To GROUP_POINTER
IdenPoint(i).x = (IdenPoint(i).x / GroupPointCount(i))
IdenPoint(i).y = IdenPoint(i).y / GroupPointCount(i)

'windowwidth/PW是从图片缓存到屏幕的坐标转换系数
scr.Draw_FilledCircle IdenPoint(i).x * (WindowWidth / PW), IdenPoint(i).y * (WindowHeight / PH), 3, 20, RGBA(1, 0, 0, 1)
Next i
IdenPointCount = GROUP_POINTER '标识点数
'————————————————————


'———————标识框——————————
Dim RectLeft As Long, RectRight As Long, RectTop As Long, RectBottom As Long
RectLeft = PictureWidth
RectRight = 0
RectTop = PictureHeight
RectBottom = 0 '初值
For i = 1 To GROUP_POINTER
'生成标识点框框 即求9个点的边界值
If IdenPoint(i).x < RectLeft Then RectLeft = IdenPoint(i).x
If IdenPoint(i).x > RectRight Then RectRight = IdenPoint(i).x
If IdenPoint(i).y < RectTop Then RectTop = IdenPoint(i).y
If IdenPoint(i).y > RectBottom Then RectBottom = IdenPoint(i).y
Next i

 '100mm乘比例  实际可视高度
 '其中100mm是九宫格的边长 所以照片里面必须拍到所有点
RealVisibleWidth = 100 * PictureWidth / (RectRight - RectLeft)
RealVisibleHeight = 100 * PictureHeight / (RectBottom - RectTop)


'标识矩形边长
Dim RectP1 As TV_2DVECTOR, RectP2 As TV_2DVECTOR
RectP1 = Vector2(RectLeft * (WindowWidth / PW), RectTop * (WindowHeight / PH))
RectP2 = Vector2(RectRight * (WindowWidth / PW), RectBottom * (WindowHeight / PH))
ScrText.NormalFont_DrawText "标识矩形: " & CStr(RectRight - RectLeft) & " x " & CStr(RectBottom - RectTop), 10, 50, RGBA(0, 1, 1, 1), 1
'画标识矩形
scr.Draw_Box RectP1.x, RectP1.y, RectP2.x, RectP2.y
'输出
TV.RenderToScreen
End Sub



Function GetIdenPoint(ID As Integer) As TV_2DVECTOR
GetIdenPoint = IdenPoint(ID)
End Function
Function GetIdenPointCount() As Integer  '获取标识点个数
GetIdenPointCount = IdenPointCount
End Function

Function GetVisibleAngleHorizontal() As Single
Dim hFocalLength As Single
hFocalLength = GetFocalLength '从EXIF中读取
Dim a As Single, b As Single, f As Single
a = 36 / 1.59 '1.59等效焦距 那个什么跟135胶片成的比例关系
b = RealVisibleWidth
f = hFocalLength
GetVisibleAngleHorizontal = 2 * Math.ATan(a * b / (2 * f * (a + b))) * 180 / 3.1415926
End Function

Function GetVisibleAngleVertical() As Single
Dim vFocalLength As Single
vFocalLength = GetFocalLength '从EXIF中读取
Dim a As Single, b As Single, f As Single
a = 24 / 1.59
b = RealVisibleHeight
f = vFocalLength
GetVisibleAngleVertical = 2 * Math.ATan(a * b / (2 * f * (a + b))) * 180 / 3.1415926 '给用户是角度制
End Function

Function GetStandardDepth() As Single
Dim FocalLength As Single
 '从EXIF中读取
FocalLength = GetFocalLength

Dim a As Single, b As Single, C1 As Single, C2 As Single
a = 36 / 1.59 'EOS感光底片的宽度 1.59是这部机子的等效焦距转换常数
b = RealVisibleWidth
C1 = (1 + (b / a)) * FocalLength
a = 24 / 1.59
b = RealVisibleHeight
C2 = (1 + (b / a)) * FocalLength
GetStandardDepth = (C1 + C2) / 2
End Function

'Function GetImageDistance(dFocalLength As Single) As Single
'Dim a As Single, b As Single, C1 As Single, C2 As Single
'a = 36 / 1.6
'b = RealVisibleWidth
'C1 = (1 + (a / b)) * dFocalLength
'GetImageDistance = C1
'End Function





'——————————————串口通信————————————


Function InitSerialPort(SpecificPortID As Integer, Optional TryOtherWhenError As Boolean = True) As Boolean
Dim s As Long
s = SpecificPortID - 1
Set MSComm1 = New MSComm ' CreateObject("MSCommLib.MSComm")

RetryMark:
On Error Resume Next
s = s + 1 '第一次是
MSComm1.CommPort = Val(s)
MSComm1.Settings = "9600,N,8,1" '参数：比特率，奇偶位校验(?)，8位数据位，1位停止位
MSComm1.PortOpen = True
'错误重试
If TryOtherWhenError = True And s < 15 And Err.Number <> 0 Then
Err.Clear
GoTo RetryMark:
End If

'返回值
If Err.Number = 0 Then
InitSerialPort = True
Else
InitSerialPort = False
End If

End Function

Sub TakePhoto_START()
MSComm1.OutBufferCount = 0 '清空输出寄存器
MSComm1.Output = "1"
End Sub

Sub TakePhoto_END()
MSComm1.OutBufferCount = 0 '清空输出寄存器
MSComm1.Output = "0"
End Sub





'————————————EXIF——————————————
Function GetFocalLength() As Single
Set JPGInfoReader = New ExifReader
JPGInfoReader.Load App.Path & "\Group\Measure.jpg"
GetFocalLength = Val(Left(JPGInfoReader.Tag(FocalLength), 2))
End Function
