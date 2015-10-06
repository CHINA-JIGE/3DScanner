VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Scan3DScene"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Enum RenderLevel
RENDER_POINTCLOUD = 1
RENDER_MESHLINE = 2
End Enum



Private CamX As Single, CamY As Single, CamZ As Single
Private CamLX As Single, CamLY As Single, CamLZ As Single
Private CamAngleX As Single, CamAngleY As Single, CamAngleZ As Single




Function GetPictureWidth()
GetPictureWidth = PictureWidth
End Function

Function GetPictureHeight()
GetPictureHeight = PictureHeight
End Function

Function GetRenderHWND()
GetRenderHWND = RenderHWND
End Function


Sub Init(iPicturePixelWidth As Long, iPicturePixelHeight As Long, nWindowWIdth As Long, nWindowHeight As Long, iRENDERHWND As Long)
InitGlobal iPicturePixelWidth, iPicturePixelHeight, nWindowWIdth, nWindowHeight, iRENDERHWND
End Sub

Sub ResizeRenderWindow()
TV.ResizeDevice
End Sub




Sub SimulateMovement()
Dim tmpMouseX As Long, tmpMouseY As Long, IsMouseLeftPressed As Boolean
Dim VF As Single, VR As Single, DeltaVConst As Single

InputE.GetMouseState tmpMouseX, tmpMouseY, IsMouseLeftPressed
If IsMouseLeftPressed = True Then '按了鼠标左键才移
CamAngleX = CamAngleX - tmpMouseX * TV.TimeElapsed / 3000  '鼠标移动 镜头旋转
CamAngleY = CamAngleY - tmpMouseY * TV.TimeElapsed / 3000
End If

Select Case CamAngleY '纵向角度
Case Is > 1.57
CamAngleY = 1.57
Case Is < -1.57
CamAngleY = -1.57
End Select

DeltaVConst = (TV.TimeElapsed ^ 2) * 0.01
'If IsJumping = False Then '在空中不可以自由的移动呢 惯性嘛
   If InputE.IsKeyPressed(TV_KEY_W) = True Then
   VF = VF + DeltaVConst
   End If

   If InputE.IsKeyPressed(TV_KEY_S) = True Then
   VF = VF - DeltaVConst
   End If

   If InputE.IsKeyPressed(TV_KEY_A) = True Then
   VR = VR - DeltaVConst
   End If

   If InputE.IsKeyPressed(TV_KEY_D) = True Then
   VR = VR + DeltaVConst
   End If
'End If
'///////////////////////////////////////////////////////

Select Case VF '限制速度 &减速度
Case Is > TV.TimeElapsed * 1.5
VF = TV.TimeElapsed * 1.5
Case Is < -TV.TimeElapsed * 1.5
VF = -TV.TimeElapsed * 1.5
Case Is > 0
VF = VF - DeltaVConst * 0.5
       If VF < 0 Then VF = 0
Case Is < 0
VF = VF + DeltaVConst * 0.5
       If VF > 0 Then VF = 0
End Select

Select Case VR '限制速度 &减速度
Case Is > TV.TimeElapsed * 1.5
VR = TV.TimeElapsed * 1.5
Case Is < -TV.TimeElapsed * 1.5
VR = -TV.TimeElapsed * 1.5
Case Is > 0
VR = VR - DeltaVConst * 0.5
If VR < 0 Then VR = 0
Case Is < 0
VR = VR + DeltaVConst * 0.5
If VR > 0 Then VR = 0
End Select


CamX = CamX + VF * Cos(CamAngleX) + VR * Sin(CamAngleX)
CamY = CamY + VF * Sin(CamAngleY)
CamZ = CamZ + VF * Sin(CamAngleX) - VR * Cos(CamAngleX)

CamLX = CamX + 200 * Cos(CamAngleX)
CamLY = CamY + 200 * Sin(CamAngleY)
CamLZ = CamZ + 200 * Sin(CamAngleX)

'上升下降
If InputE.IsKeyPressed(TV_KEY_SPACE) Then CamY = CamY + DeltaVConst
If InputE.IsKeyPressed(TV_KEY_LEFTCONTROL) Then CamY = CamY - DeltaVConst

Cam.SetCamera CamX, CamY, CamZ, CamLX, CamLY, CamLZ
End Sub





'————————采样点渲染————————
'大致流程：根据PICTUREID和用Getline函数
'计算出每张PICTURE对应采样点的ID段
'然后用ID找到坐标画点
'——————————————————————
Sub RenderSampling(PictureID As Long, SamplingPointColor As Long)

              DoEvents
                     TV.Clear
                      '转换系数
                     Dim Buffer2ScreenX As Single, Buffer2ScreenY As Single
                     Buffer2ScreenX = WindowWidth / PictureWidth
                     Buffer2ScreenY = WindowHeight / PictureHeight
                     
                     '采样图片
                        scr.Draw_Texture GetTex(CStr(PictureID)), 0, 0, WindowWidth, WindowHeight
                        
                        
                        Dim sampleID1 As Long, sampleID2 As Long, SP As TV_2DVECTOR '本列采样第一个点，最后一个，采样2D点
                        Dim i As Long
                        For i = 1 To PictureID
                            If i = 1 Then
                                   sampleID1 = 1
                            Else
                                   sampleID1 = sampleID1 + MB2.GetSamplingCount(i)
                            End If
                        Next i
                        sampleID2 = sampleID1 + MB2.GetSamplingCount(PictureID) - 1 '本列最后一个点
                        
                        Dim c As Long
                        For c = sampleID1 To sampleID2 '同一幅图里面的点 的ID跨度
                        SP = MB2.GetSamplingResult(Val(c))
                         scr.Draw_Point SP.x * Buffer2ScreenX, SP.y * Buffer2ScreenY, SamplingPointColor
                        Next c

                        ScrText.NormalFont_DrawText "新增点数：" & MB2.GetSamplingCount(PictureID), 10, 50, RGBA(1, 1, 1, 1), 1
                        TV.RenderToScreen
End Sub




Sub Render3D(iRenderLevel As RenderLevel)
DoEvents
                        TV.Clear

                        scr.Draw_Line3D 0, 0, 0, 100, 0, 0, RGBA(1, 0, 0, 1) 'X轴
                        scr.Draw_Line3D 0, 0, 0, 0, 100, 0, RGBA(0, 1, 0, 1) 'Y
                        scr.Draw_Line3D 0, 0, 0, 0, 0, 100, RGBA(0, 0, 1, 1) 'Z
                        
                        Select Case iRenderLevel
                        Case 1
                        TVSCENE1.SetRenderMode TV_POINT
                        GetMesh("POINT1").Render '画点云
                        
                        Case 2
                        TVSCENE1.SetRenderMode TV_LINE
                        GetMesh("TRIANGLE1").Render '画网格
                        End Select
                        
                        TVSCENE1.SetRenderMode TV_SOLID '这样才可渲染SAMPLE
                        
'ScrText.NormalFont_DrawText "按RIGHTSHIFT截图", 20, 20, RGBA(1, 0, 0, 1), 1
ScrText.NormalFont_DrawText "CamX: " & CamX, 20, 50, RGBA(1, 1, 1, 1), 1
ScrText.NormalFont_DrawText "CamY: " & CamY, 20, 80, RGBA(1, 1, 1, 1), 1
ScrText.NormalFont_DrawText "CamZ: " & CamZ, 20, 110, RGBA(1, 1, 1, 1), 1
ScrText.NormalFont_DrawText "有效顶点数: " & MB2.GetRenderMesh_PointMesh.GetVertexCount, 20, 140, RGBA(1, 1, 1, 1), 1
ScrText.NormalFont_DrawText "实际顶点数: " & MB2.GetRenderMesh_TriangleMesh.GetVertexCount, 20, 170, RGBA(1, 1, 1, 1), 1
ScrText.NormalFont_DrawText "三角面片数: " & MB2.GetRenderMesh_TriangleMesh.GetTriangleCount, 20, 200, RGBA(1, 1, 1, 1), 1
                        
 'If InputE.IsKeyPressed(TV_KEY_RIGHTSHIFT) = True Then  '截图
'            ScrText.NormalFont_DrawText "已截图！！！！！", 20, 230, RGBA(0, 1, 0, 1), 1
'             TV.Screenshot Form2.Text_ScrShotPath, TV_IMAGE_JPG
 '            Beep
'End If
                        
                        TV.RenderToScreen
End Sub





Sub MoveAndLookatMesh()

       Dim x As Single, y As Single, z As Single, nx As Single, ny As Single, nz As Single, tu1 As Single, tv1 As Single, tu2 As Single, tv2 As Single, c As Long
       If MB2.GetRenderMesh_PointMesh.GetVertexCount > 0 Then
       MB2.GetRenderMesh_PointMesh.GetVertex 2, x, y, z, nx, ny, nz, tu1, tv1, tu2, tv2, c
        CamLX = x
        CamLY = y
        CamLZ = z
        CamAngleX = 0 'Math.ATan((CamLX - CamX) / 200)
        CamAngleY = 0 'Math.ATan((CamLY - CamY) / 200)
        CamAngleZ = 0 ' Math.ATan((CamLZ - CamZ) / 200)
        CamX = CamLX - 30 * Cos(CamAngleX)
        CamY = CamLY + 30 * Sin(CamAngleY)
        CamZ = CamLZ - 30 * Sin(CamAngleX)
        Cam.SetCamera CamX, CamY, CamZ, CamLX, CamLY, CamLZ
        
        End If
        
        'Form1.SetFocus
End Sub






