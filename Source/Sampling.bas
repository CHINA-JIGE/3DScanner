Attribute VB_Name = "aSampling_Module"


Public SamplingPx(1 To PictureHeight) As Single
Private NumberOfPoint As Single, TotalPx As Single
Private LowerBound As Single, RED As Single

Public Sub Sampling(TexID As Long, iLightSide As LightSideType)

        TotalPx = 0
        NumberOfPoint = 0
       LowerBound = Val(Form2.Text_LowerBound.Text) / 255

       
        If iLightSide = Side_Right Then
        '从2开始或减一是为了main2的模糊处理
       '————RIGHT————————————
        SamplingMain2 TexID
        Else
        '————LEFT————————
       SamplingMain1 TexID
        End If
       
End Sub

Private Sub SamplingMain1(iTexID1 As Long) '左

        For y = 1 To PictureHeight
                For x = 1 To PictureWidth / 2
                        RED = DecodeRGBA(TexF.GetPixel(iTexID1, x + 1, y)).r

                        '筛选颜色
                        If RED > LowerBound Then
                                TotalPx = TotalPx + x * RED
                                NumberOfPoint = NumberOfPoint + RED
                        End If

                Next x

                '生成采样点
                If NumberOfPoint <> 0 Then
                        SamplingPx(y) = TotalPx / NumberOfPoint
                Else
                        '负数会被剔除 不增加顶点
                        SamplingPx(y) = -1
                End If
                
       TotalPx = 0
        NumberOfPoint = 0
        
        Next y
        
End Sub



Private Sub SamplingMain2(iTexId2 As Long)
        For y = 1 To PictureHeight
                For x = PictureWidth / 2 To PictureWidth
                        RED = DecodeRGBA(TexF.GetPixel(iTexId2, x, y)).r

                        '筛选颜色
                        If RED > LowerBound Then
                                TotalPx = TotalPx + x * RED
                                NumberOfPoint = NumberOfPoint + RED
                        End If

                Next x

                '生成采样点
                If NumberOfPoint <> 0 Then
                        SamplingPx(y) = TotalPx / NumberOfPoint
                Else
                        '负数会被剔除 不增加顶点
                        SamplingPx(y) = -1
                End If
                
       TotalPx = 0
        NumberOfPoint = 0
        
        Next y
                
End Sub













Public Sub DrawSamplingPoint(pID As Long)
                        TV.clear
                        
                        TexF.UnlockTexture GetTex("Picture" & pID)
                        scr.Draw_Texture GetTex("Picture" & pID), 0, 0, PictureWidth, PictureHeight
                        
                        For c = 1 To PictureHeight '画采样点
                               If SamplingPx(c) > 0 Then scr.Draw_Point SamplingPx(c), c, RGBA(0, 1, 0, 1)
                        Next c

                        TV.RenderToScreen
End Sub


