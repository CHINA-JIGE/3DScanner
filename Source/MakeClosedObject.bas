Attribute VB_Name = "MakeClosedObject"

Public NumOfPointVerticalLine() As Long '每列多少个点 用于连接点为三角形  在主程序定义

Private Point3D() As TV_3DVECTOR

Public Sub MakeClosedObjectFromPointCloud()
        'Dim NumOfVerticalLines As Long '共多少的竖列
        'NumOfVerticalLines = Val(Text_END) - Val(Text_START)
       
        DoEvents
        Dim Line1FirstPointID As Long, Line2FirstPointID As Long '第二列

        Line1FirstPointID = 0
        Line2FirstPointID = Line1FirstPointID + NumOfPointVerticalLine(1)

        ReDim Point3D(mesh.GetVertexCount - 1) '先get好顶点

        Dim x As Single, y As Single, z As Single, nx As Single, ny As Single, nz As Single, tu1 As Single, tv1 As Single, tu2 As Single, tv2 As Single, c As Long
       
        '为了方便用顶点先get好顶点
        For i = 0 To mesh.GetVertexCount - 1
                mesh.GetVertex i, x, y, z, nx, ny, nz, tu1, tv1, tu2, tv2, c
                Point3D(i).x = x
                Point3D(i).y = y
                Point3D(i).z = z
        Next i

        '
        '
        '————————————————连起邻列三角形——————————————
        
        Dim MinOfTotalPoint As Long

        Dim p1              As Long, p2 As Long, p3 As Long

        For i = 1 To NumOfVerticalLines - 1
                MinOfTotalPoint = MIN(Val(NumOfPointVerticalLine(i)), Val(NumOfPointVerticalLine(i + 1)))  '两列点谁少点

                For j = 0 To MAX(NumOfPointVerticalLine(i), NumOfPointVerticalLine(i + 1)) - 2 '一列多少个点

                        '每列最下面的序号为N-1，j 是从0到 N-2(要往下连三角形)
                        Select Case Val(j)

                                Case Is < MinOfTotalPoint - 1 '先判断两列谁的点数少
                                        p1 = Line1FirstPointID + j
                                        p2 = Line2FirstPointID + j
                                        p3 = Line1FirstPointID + j + 1
                                        Mesh2.AddTriangle 0, Point3D(p1).x, Point3D(p1).y, Point3D(p1).z, Point3D(p2).x, Point3D(p2).y, Point3D(p2).z, Point3D(p3).x, Point3D(p3).y, Point3D(p3).z
                                        OutList.AddToList Vector3(Point3D(p1).x, Point3D(p1).y, Point3D(p1).z), Vector3(Point3D(p2).x, Point3D(p2).y, Point3D(p2).z), Vector3(Point3D(p3).x, Point3D(p3).y, Point3D(p3).z)
                                        
                                        p1 = Line2FirstPointID + j
                                        p2 = Line2FirstPointID + j + 1
                                        p3 = Line1FirstPointID + j + 1
                                        Mesh2.AddTriangle 0, Point3D(p1).x, Point3D(p1).y, Point3D(p1).z, Point3D(p2).x, Point3D(p2).y, Point3D(p2).z, Point3D(p3).x, Point3D(p3).y, Point3D(p3).z
                                        OutList.AddToList Vector3(Point3D(p1).x, Point3D(p1).y, Point3D(p1).z), Vector3(Point3D(p2).x, Point3D(p2).y, Point3D(p2).z), Vector3(Point3D(p3).x, Point3D(p3).y, Point3D(p3).z)
                                                                          
                                Case Is > MinOfTotalPoint '对应点连完到多出来的点了 注意下标判断是大于号

                                        If NumOfPointVerticalLine(i) > NumOfPointVerticalLine(i + 1) Then
                                                '第一列的点多的时候
                                                p1 = Line1FirstPointID + j - 1
                                                p2 = Line2FirstPointID + MinOfTotalPoint - 1 '第二列最低点做定点
                                                p3 = Line1FirstPointID + j
                                        Else
                                                p1 = Line1FirstPointID + MinOfTotalPoint - 1 '第一列最低点
                                                p2 = Line2FirstPointID + j - 1 '少点的那列最低点做定点
                                                p3 = Line2FirstPointID + j
                                        End If

                                        Mesh2.AddTriangle 0, Point3D(p1).x, Point3D(p1).y, Point3D(p1).z, Point3D(p2).x, Point3D(p2).y, Point3D(p2).z, Point3D(p3).x, Point3D(p3).y, Point3D(p3).z
                                        OutList.AddToList Vector3(Point3D(p1).x, Point3D(p1).y, Point3D(p1).z), Vector3(Point3D(p2).x, Point3D(p2).y, Point3D(p2).z), Vector3(Point3D(p3).x, Point3D(p3).y, Point3D(p3).z)
                        End Select

                Next j
                
                '封顶与封底
                If Line2FirstPointID < mesh.GetVertexCount Then
                        p1 = Line1FirstPointID
                        p2 = Line2FirstPointID
                        Mesh2.AddTriangle 0, TopPoint.x, TopPoint.y, TopPoint.z, Point3D(p2).x, Point3D(p2).y, Point3D(p2).z, Point3D(p1).x, Point3D(p1).y, Point3D(p1).z
                        OutList.AddToList Vector3(TopPoint.x, TopPoint.y, TopPoint.z), Vector3(Point3D(p2).x, Point3D(p2).y, Point3D(p2).z), Vector3(Point3D(p1).x, Point3D(p1).y, Point3D(p1).z)
       
                        p1 = Line1FirstPointID + NumOfPointVerticalLine(i) - 1
                        p2 = Line2FirstPointID + NumOfPointVerticalLine(i + 1) - 1
                        Mesh2.AddTriangle 0, BottomPoint.x, BottomPoint.y, BottomPoint.z, Point3D(p2).x, Point3D(p2).y, Point3D(p2).z, Point3D(p1).x, Point3D(p1).y, Point3D(p1).z
                        OutList.AddToList Vector3(BottomPoint.x, BottomPoint.y, BottomPoint.z), Vector3(Point3D(p2).x, Point3D(p2).y, Point3D(p2).z), Vector3(Point3D(p1).x, Point3D(p1).y, Point3D(p1).z)
       
                End If
       
                '更新序号
                Line1FirstPointID = Line2FirstPointID '更新正在处理列 第i列更新到i+1
                Line2FirstPointID = Line2FirstPointID + NumOfPointVerticalLine(i + 1) 'i+1更新到i+2

        Next i

        '——————————把最后一列和第一列连起来————————

        '连起最后一列和第一列三角形
        Line1FirstPointID = mesh.GetVertexCount - NumOfPointVerticalLine(NumOfVerticalLines)  '最后一列
        Line2FirstPointID = 0 '第一列

        For j = 0 To MAX(NumOfPointVerticalLine(NumOfVerticalLines), NumOfPointVerticalLine(1)) - 2
                MinOfTotalPoint = MIN(Val(NumOfPointVerticalLine(NumOfVerticalLines)), Val(NumOfPointVerticalLine(1)))

                Select Case Val(j)

                        Case Is < MinOfTotalPoint - 1 '
                                p1 = Line1FirstPointID + j '|/
                                p2 = Line2FirstPointID + j
                                p3 = Line1FirstPointID + j + 1
                                Mesh2.AddTriangle 0, Point3D(p1).x, Point3D(p1).y, Point3D(p1).z, Point3D(p2).x, Point3D(p2).y, Point3D(p2).z, Point3D(p3).x, Point3D(p3).y, Point3D(p3).z
                                OutList.AddToList Vector3(Point3D(p1).x, Point3D(p1).y, Point3D(p1).z), Vector3(Point3D(p2).x, Point3D(p2).y, Point3D(p2).z), Vector3(Point3D(p3).x, Point3D(p3).y, Point3D(p3).z)
                                                               
                                p1 = Line2FirstPointID + j
                                p2 = Line2FirstPointID + j + 1
                                p3 = Line1FirstPointID + j + 1
                                Mesh2.AddTriangle 0, Point3D(p1).x, Point3D(p1).y, Point3D(p1).z, Point3D(p2).x, Point3D(p2).y, Point3D(p2).z, Point3D(p3).x, Point3D(p3).y, Point3D(p3).z
                                OutList.AddToList Vector3(Point3D(p1).x, Point3D(p1).y, Point3D(p1).z), Vector3(Point3D(p2).x, Point3D(p2).y, Point3D(p2).z), Vector3(Point3D(p3).x, Point3D(p3).y, Point3D(p3).z)
                                                               
                        Case Is > MIN(Val(NumOfPointVerticalLine(i)), Val(NumOfPointVerticalLine(1)))

                                If NumOfPointVerticalLine(i) > NumOfPointVerticalLine(1) Then
                                        '第一列的点多的时候
                                        p1 = Line1FirstPointID + j - 1
                                        p2 = Line2FirstPointID + MinOfTotalPoint - 1 '第二列最低点做定点
                                        p3 = Line1FirstPointID + j
                                Else
                                        p1 = Line1FirstPointID + MinOfTotalPoint - 1 '第一列最低点
                                        p2 = Line2FirstPointID + j - 1 '少点的那列最低点做定点
                                        p3 = Line2FirstPointID + j
                                End If
                                
                                Mesh2.AddTriangle 0, Point3D(p1).x, Point3D(p1).y, Point3D(p1).z, Point3D(p2).x, Point3D(p2).y, Point3D(p2).z, Point3D(p3).x, Point3D(p3).y, Point3D(p3).z
                End Select
              
        Next j

        '要顺时针生成顶点 法线才能正确
        p1 = Line1FirstPointID
        p2 = Line2FirstPointID
        Mesh2.AddTriangle 0, TopPoint.x, TopPoint.y, TopPoint.z, Point3D(p2).x, Point3D(p2).y, Point3D(p2).z, Point3D(p1).x, Point3D(p1).y, Point3D(p1).z
        OutList.AddToList Vector3(TopPoint.x, TopPoint.y, TopPoint.z), Vector3(Point3D(p2).x, Point3D(p2).y, Point3D(p2).z), Vector3(Point3D(p1).x, Point3D(p1).y, Point3D(p1).z)
        
        p1 = Line1FirstPointID + NumOfPointVerticalLine(i) - 1
        p2 = Line2FirstPointID + NumOfPointVerticalLine(1) - 1
        Mesh2.AddTriangle 0, BottomPoint.x, BottomPoint.y, BottomPoint.z, Point3D(p2).x, Point3D(p2).y, Point3D(p2).z, Point3D(p1).x, Point3D(p1).y, Point3D(p1).z
        OutList.AddToList Vector3(BottomPoint.x, BottomPoint.y, BottomPoint.z), Vector3(Point3D(p2).x, Point3D(p2).y, Point3D(p2).z), Vector3(Point3D(p1).x, Point3D(p1).y, Point3D(p1).z)
        '
        '
        '————————————————————————————————————————————————
        Mesh2.WeldVertices 0
End Sub

