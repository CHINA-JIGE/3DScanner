Attribute VB_Name = "GlobalModule"

Public Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public TV As TVEngine
Public TexF As TVTextureFactory
 Public Cam As TVCamera
Public TVSCENE1 As TVScene
 Public InputE As TVInputEngine
 Public scr As TVScreen2DImmediate
 Public ScrText As TVScreen2DText
 Public Math As TVMathLibrary


 Public RenderHWND As Long '渲染窗句柄
 Public PictureWidth As Long, PictureHeight As Long '图片缓存区大小
 Public WindowWidth As Long, WindowHeight As Long '窗口大小
 Public MainScanParam As Type_ScanParameters '扫描参数
 Public MainColorFilterLowerBound As Single
Public MB2 As Scan3DMeshBuffer2

 

 
 
Public Type Triangle
v1 As TV_3DVECTOR
v2 As TV_3DVECTOR
v3 As TV_3DVECTOR
End Type
Sub InitGlobal(PicturePixelWidth As Long, PicturePixelHeight As Long, _
                           WindowPixelWidth As Long, WindowPixelHeight As Long, _
                           hRENDERHWND As Long)
                                          
PictureWidth = PicturePixelWidth
PictureHeight = PicturePixelHeight
WindowWidth = WindowPixelWidth
WindowHeight = WindowPixelHeight
RenderHWND = hRENDERHWND '此句柄是公共的

'先设置window大小
MoveWindow hRENDERHWND, -1, -1, WindowWidth, WindowHeight, -1
        Set TV = New TVEngine
        '据DEBUG说多线程设置要在初始化之前
        TV.AllowMultithreading True
        TV.Init3DWindowed RenderHWND
        TV.DisplayFPS True
        TV.SetVSync True
        TV.SetAngleSystem TV_ANGLE_RADIAN
       TV.SetDebugFile App.Path & "\Procedure.log"
       TV.AddToLog "——————————————S T A R T ——————————————"

       
        Set TexF = New TVTextureFactory
        TexF.SetTextureMode TV_TEXTUREMODE_32BITS

        Set scr = New TVScreen2DImmediate
        Set ScrText = New TVScreen2DText
        ScrText.NormalFont_Create "宋体", "宋体", 16, True, False, False
        ScrText.NormalFont_Create "宋体", "大字", 64, True, False, False
        
        Set TVSCENE1 = New TVScene
        Set Math = New TVMathLibrary
        Set InputE = New TVInputEngine
        InputE.Initialize True, True



        Set Cam = New TVCamera
        CamX = -10
        CamY = 0
        CamZ = 0
        CamLX = 0
        CamLY = 0
        CamLZ = 0
        CamAngleX = 0
        CamAngleY = 0
        CamAngleZ = 0
        Cam.SetCamera CamX, CamY, CamZ, CamLX, CamLY, CamLZ
        
        '————————主缓存————————
       Set MB2 = New Scan3DMeshBuffer2
        MB2.Init     '要用TVScene
        '——————————————————
        

        'TV.ResizeDevice
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


Public Function Cross(v1 As TV_3DVECTOR, v2 As TV_3DVECTOR) As TV_3DVECTOR '叉乘
Cross.x = v1.y * v2.z - v1.z * v2.y
Cross.y = v1.z * v2.x - v1.x * v2.z
Cross.z = v1.x * v2.y - v1.y * v2.x
End Function



Public Function Dot(dotV1 As TV_3DVECTOR, dotV2 As TV_3DVECTOR) As Single '点乘
Dot = dotV1.x * dotV2.x + dotV1.y + dotV2.y + dotV1.z * dotV2.z
End Function




Public Function GetDepthFromOffset(SCANPARAM As Type_ScanParameters, _
                                   OffsetPx As Single, _
                                   LightSide As CONST_LightSide) As Single

        Dim Cita As Single, H As Single, d As Single
       Dim OriginPx As Single
       
        Cita = SCANPARAM.VisibleAngleHorizontal * 3.1415926 / 180
        H = SCANPARAM.StandardDepth
        d = SCANPARAM.CamToLight

        Select Case LightSide '红线在相机的左边还是右边

                Case 0 'left
                        OriginPx = 0.5 * PictureWidth * (1 - d / (H * Tan(0.5 * Cita))) '左
                        GetDepthFromOffset = (2 * (H ^ 2) * CSng(Tan(Cita * 0.5)) * (OriginPx - OffsetPx)) / (PictureWidth * d + 2 * H * CSng(Tan(Cita * 0.5)) * (OriginPx - OffsetPx))

                Case 1 'right
                        OriginPx = 0.5 * PictureWidth * (1 + d / (H * Tan(0.5 * Cita))) '右
                        GetDepthFromOffset = (2 * (H ^ 2) * CSng(Tan(Cita * 0.5)) * (OffsetPx - OriginPx)) / (PictureWidth * d + 2 * H * CSng(Tan(Cita * 0.5)) * (OffsetPx - OriginPx))
        End Select
End Function




Public Function BezierInterpolation(iRatio As Single, v1 As TV_3DVECTOR, v2 As TV_3DVECTOR, v3 As TV_3DVECTOR) As TV_3DVECTOR
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





