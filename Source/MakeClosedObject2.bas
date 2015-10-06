Attribute VB_Name = "MakeClosedObject2"
Private VLineFirstPointID() As Long
Private VLineLastPointID() As Long




Public Sub MakeClosedObjectFromPointCloud()
       
        DoEvents
        
        ReDim VLineFirstPointID(1 To NumOfVerticalLines)
        ReDim VLineLastPointID(1 To NumOfVerticalLines)
        ReDim Point3D(mesh.GetVertexCount - 1) '先get好顶点
       
        Dim x As Single, y As Single, z As Single, nx As Single, ny As Single, nz As Single, tu1 As Single, tv1 As Single, tu2 As Single, tv2 As Single, c As Long

        '为了方便用顶点先get好顶点
        For i = 0 To mesh.GetVertexCount - 1
                mesh.GetVertex i, x, y, z, nx, ny, nz, tu1, tv1, tu2, tv2, c
                Point3D(i).x = x
                Point3D(i).y = y
                Point3D(i).z = z
        Next i
        
        Dim Line1FirstPointID As Long, Line2FirstPointID As Long '第二列

        Line1FirstPointID = 0
        Line2FirstPointID = Line1FirstPointID + NumOfPointVerticalLine(1)

        '
        '
        '————————————————连起邻列三角形——————————————
        
        Dim MinOfTotalPoint As Long

        Dim p1              As Long, p2 As Long, p3 As Long

        For i = 1 To NumOfVerticalLines
                MinOfTotalPoint = MIN(Val(NumOfPointVerticalLine(i)), Val(NumOfPointVerticalLine(i + 1)))  '两列点谁少点

                For j = 0 To MAX(NumOfPointVerticalLine(i), NumOfPointVerticalLine(i + 1)) - 1

                        '每列最下面的序号为N-1，j 是从0到 N-2(要往下连三角形)
                        Select Case Val(j)

                                Case Is < MinOfTotalPoint - 1 '先判断两列谁的点数少
                                        p1 = Line1FirstPointID
                                        p2 = Line2FirstPointID + j
                                        p3 = Line1FirstPointID + j + 1
                                        Mesh2.AddTriangle 0, Point3D(p1).x, Point3D(p1).y, Point3D(p1).z, Point3D(p2).x, Point3D(p2).y, Point3D(p2).z, Point3D(p3).x, Point3D(p3).y, Point3D(p3).z
                                        OutList.AddToList Point3D(p1), Point3D(p2), Point3D(p3)
                                        
                                        p1 = Line2FirstPointID + j
                                        p2 = Line2FirstPointID + j + 1
                                        p3 = Line1FirstPointID + j + 1
                                        Mesh2.AddTriangle 0, Point3D(p1).x, Point3D(p1).y, Point3D(p1).z, Point3D(p2).x, Point3D(p2).y, Point3D(p2).z, Point3D(p3).x, Point3D(p3).y, Point3D(p3).z
                                        OutList.AddToList Point3D(p1), Point3D(p2), Point3D(p3)
                                        
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
                                        OutList.AddToList Point3D(p1), Point3D(p2), Point3D(p3)
                        End Select

                Next j
                
                '更新序号
                VLineFirstPointID(i) = Line1FirstPointID
                VLineLastPointID(i) = Line2FirstPointID - 1
                Line1FirstPointID = Line2FirstPointID '更新正在处理列 第i列更新到i+1
                Line2FirstPointID = Line2FirstPointID + NumOfPointVerticalLine(i + 1) 'i+1更新到i+2
                
        Next i
       
       
       
        '——————————封顶与封底——————————————
        For i = 1 To Int((NumOfVerticalLines - 1) / 2) - 1 'int是向下取整  注意：第一列和最后一列是重叠的，所以减1
                '——顶
                p1 = VLineFirstPointID(i)
                p2 = VLineFirstPointID(NumOfVerticalLines - i + 1)
                p3 = VLineFirstPointID(NumOfVerticalLines - i)
                Mesh2.AddTriangle 0, Point3D(p1).x, Point3D(p1).y, Point3D(p1).z, Point3D(p2).x, Point3D(p2).y, Point3D(p2).z, Point3D(p3).x, Point3D(p3).y, Point3D(p3).z
                OutList.AddToList Point3D(p1), Point3D(p2), Point3D(p3)
        
                p1 = VLineFirstPointID(i)
                p2 = VLineFirstPointID(i + 1)
                p3 = VLineFirstPointID(NumOfVerticalLines - i)
                Mesh2.AddTriangle 0, Point3D(p1).x, Point3D(p1).y, Point3D(p1).z, Point3D(p2).x, Point3D(p2).y, Point3D(p2).z, Point3D(p3).x, Point3D(p3).y, Point3D(p3).z
                OutList.AddToList Point3D(p2), Point3D(p1), Point3D(p3)
       
                '——底
                p1 = VLineLastPointID(i)
                p2 = VLineLastPointID(NumOfVerticalLines - i + 1)
                p3 = VLineLastPointID(NumOfVerticalLines - i)
                Mesh2.AddTriangle 0, Point3D(p1).x, Point3D(p1).y, Point3D(p1).z, Point3D(p2).x, Point3D(p2).y, Point3D(p2).z, Point3D(p3).x, Point3D(p3).y, Point3D(p3).z
                OutList.AddToList Point3D(p2), Point3D(p1), Point3D(p3)
        
                p1 = VLineLastPointID(i)
                p2 = VLineLastPointID(i + 1)
                p3 = VLineLastPointID(NumOfVerticalLines - i)
                Mesh2.AddTriangle 0, Point3D(p1).x, Point3D(p1).y, Point3D(p1).z, Point3D(p2).x, Point3D(p2).y, Point3D(p2).z, Point3D(p3).x, Point3D(p3).y, Point3D(p3).z
                OutList.AddToList Point3D(p1), Point3D(p2), Point3D(p3) '顶点顺序是根据UP软件的指示调整的= =
        
        Next i
        
        
        
        If Int(NumOfVerticalLines - 1 / 2) Mod 2 = 1 Then '除去最后一列列数是奇数
                p1 = VLineFirstPointID(Int((NumOfVerticalLines) / 2))  '半圈前的最后一个点
                p2 = VLineFirstPointID(Int((NumOfVerticalLines) / 2) + 1)
                p3 = VLineFirstPointID(Int((NumOfVerticalLines) / 2) + 2)
                Mesh2.AddTriangle 0, Point3D(p1).x, Point3D(p1).y, Point3D(p1).z, Point3D(p2).x, Point3D(p2).y, Point3D(p2).z, Point3D(p3).x, Point3D(p3).y, Point3D(p3).z
                OutList.AddToList Point3D(p1), Point3D(p2), Point3D(p3)
        
                p1 = VLineLastPointID(Int((NumOfVerticalLines) / 2))  '半圈前的最后一个点
                p2 = VLineLastPointID(Int((NumOfVerticalLines) / 2) + 1)
                p3 = VLineLastPointID(Int((NumOfVerticalLines) / 2) + 2)
                Mesh2.AddTriangle 0, Point3D(p1).x, Point3D(p1).y, Point3D(p1).z, Point3D(p2).x, Point3D(p2).y, Point3D(p2).z, Point3D(p3).x, Point3D(p3).y, Point3D(p3).z
                OutList.AddToList Point3D(p1), Point3D(p2), Point3D(p3)
        End If

        '————————————————————————————————————————————————
        
        Mesh2.WeldVertices 0
        
End Sub

