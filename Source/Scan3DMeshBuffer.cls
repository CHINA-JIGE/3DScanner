VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Scan3DMeshBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private SamplingPoint() As TV_2DVECTOR
Private Sampling_Pointer As Long

Private Point3D() As TV_3DVECTOR
Private Point3D_Pointer As Long

Private NumOfVLines As Long
Private NumOfPointInVLines() As Long
Private VLine_Pointer As Long

Private TriangleList() As Triangle
Private Triangle_Pointer As Long

Private Mesh_Point As TVMesh
Private Mesh_Triangle As TVMesh


Sub Init()
'Set TVSCENE1 = New TVScene
       Set Mesh_Point = TVSCENE1.CreateMeshBuilder("POINT1")
       Mesh_Point.SetCullMode TV_DOUBLESIDED
       Set Mesh_Triangle = TVSCENE1.CreateMeshBuilder("TRIANGLE1")
       Mesh_Triangle.SetCullMode TV_DOUBLESIDED
       Set Mesh_ProjectionBox = TVSCENE1.CreateMeshBuilder
       Mesh_ProjectionBox.SetCullMode TV_DOUBLESIDED
End Sub



Function GetPointID(VerticalLineID As Long, HorizontalLineID As Long) As Long '由行列得ID
 '本质上是在遍历线性表
Dim ID As Long
'iVerticalLine 是 1 to N
If VerticalLineID > 1 Then
For i = 1 To VerticalLineID - 1 '-1是因为不用加上当前列的点数
ID = ID + NumOfPointInVLines(i)
Next i
End If

GetPointID = ID + HorizontalLineID  'TV3D的MESH 顶点ID从1开始
End Function

Function GetLineAmount() As Long '列数
GetLineAmount = VLine_Pointer
End Function

Function GetPointAmount() As Long '总点数
Dim LastLineFirstID As Long
LastLineFirstID = GetPointID(GetLineAmount, 1)
GetPointAmount = LastLineFirstID + GetLine(GetLineAmount) - 1
End Function

Function GetSamplingResult(ID As Long) As TV_2DVECTOR
GetSamplingResult = SamplingPoint(ID)
End Function

Function GetTriangleAmount() As Long '总三角形数
GetTriangleAmount = Triangle_Pointer
End Function

Function GetPoint(iVerticalLineID As Long, iHorizontalLineID As Long) As TV_3DVECTOR
GetPoint = Point3D(GetPointID(iVerticalLineID, iHorizontalLineID))
End Function

Function GetPointByID(ID As Long) As TV_3DVECTOR
GetPointByID = Point3D(ID)
End Function

Function GetLine(LineID As Long) As Long '每列多少点
GetLine = NumOfPointInVLines(LineID)
End Function

Function GetTriangleVertex(TriangleID As Long, VertexID As Integer) As TV_3DVECTOR
Select Case VertexID
Case 1
GetTriangleVertex = TriangleList(TriangleID).v1
Case 2
GetTriangleVertex = TriangleList(TriangleID).v2
Case 3
GetTriangleVertex = TriangleList(TriangleID).v3
End Select
End Function









Sub AddSamplingResult(ResultPoint As TV_2DVECTOR)  '采样点
Sampling_Pointer = Sampling_Pointer + 1
ReDim Preserve SamplingPoint(1 To Sampling_Pointer) As TV_2DVECTOR
SamplingPoint(Sampling_Pointer) = ResultPoint
End Sub


Sub AddLine(NewLineTotalPoints As Long) '添加竖列
VLine_Pointer = VLine_Pointer + 1
ReDim Preserve NumOfPointInVLines(1 To VLine_Pointer)
NumOfPointInVLines(VLine_Pointer) = NewLineTotalPoints
End Sub

Sub AddPoint(x As Single, y As Single, z As Single) '添加顶点
Point3D_Pointer = Point3D_Pointer + 1
ReDim Preserve Point3D(1 To Point3D_Pointer) As TV_3DVECTOR
Point3D(Point3D_Pointer).x = x
Point3D(Point3D_Pointer).y = y
Point3D(Point3D_Pointer).z = z
End Sub
'用顶点ID新增三角形
Sub AddTriangleByPointID(ID1 As Long, ID2 As Long, ID3 As Long)
Triangle_Pointer = Triangle_Pointer + 1
ReDim Preserve TriangleList(1 To Triangle_Pointer) As Triangle
TriangleList(Triangle_Pointer).v1 = Point3D(ID1)
TriangleList(Triangle_Pointer).v2 = Point3D(ID2)
TriangleList(Triangle_Pointer).v3 = Point3D(ID3)
End Sub
'用向量新增三角形
Sub AddTriangleByVector(v1 As TV_3DVECTOR, v2 As TV_3DVECTOR, v3 As TV_3DVECTOR)
Triangle_Pointer = Triangle_Pointer + 1
ReDim Preserve TriangleList(1 To Triangle_Pointer) As Triangle
TriangleList(Triangle_Pointer).v1 = v1
TriangleList(Triangle_Pointer).v2 = v2
TriangleList(Triangle_Pointer).v3 = v3
End Sub









Sub SetSamplingResult(rID As Long, newResultPoint As TV_2DVECTOR)
If rID > 0 Or rID < Sampling_Pointer Then
SamplingPoint(rID) = newResultPoint
End If
End Sub

Sub SetPoint(ID As Long, x As Single, y As Single, z As Single) '修改顶点
If ID > 0 Or ID < Point3D_Pointer Then
Point3D(ID).x = x
Point3D(ID).y = y
Point3D(ID).z = z
End If
End Sub

Sub SetLine(LineID As Long, newLineTotalPoint As Long)
If LineID > 0 Or LineID < VLine_Pointer Then
NumOfPointInVLines(LineID) = newLineTotalPoint
End If
End Sub

Sub SetTriangle(TriangleID As Long, v1 As TV_3DVECTOR, v2 As TV_3DVECTOR, v3 As TV_3DVECTOR)
If TriangleID > 0 Or TriangleID < Triangle_Pointer Then
TriangleList(TriangleID).v1 = v1
TriangleList(TriangleID).v2 = v2
TriangleList(TriangleID).v3 = v3
End If
End Sub








Sub AddPointCloudToRenderMesh()
For i = 1 To GetPointAmount
       Dim P As TV_3DVECTOR
       P = GetPointByID(Val(i))
       Mesh_Point.AddVertex P.x, P.y, P.z, 0, 1, 0, 0, 0
Next i
End Sub

Sub AddTriangleToRenderMesh()
For i = 1 To GetTriangleAmount
       Dim p1 As TV_3DVECTOR, p2 As TV_3DVECTOR, p3 As TV_3DVECTOR
       p1 = GetTriangleVertex(Val(i), 1)
       p2 = GetTriangleVertex(Val(i), 2)
       p3 = GetTriangleVertex(Val(i), 3)
       Mesh_Triangle.AddTriangle 0, p1.x, p1.y, p1.z, p2.x, p2.y, p2.z, p3.x, p3.y, p3.z
Next i
End Sub

Function GetRenderMesh_PointMesh() As TVMesh
Set GetRenderMesh_PointMesh = GetMesh("POINT1")
End Function

 Function GetRenderMesh_TriangleMesh() As TVMesh
Set GetRenderMesh_TriangleMesh = GetMesh("TRIANGLE1")
End Function

 Sub WeldVertices()
Mesh_Triangle.WeldVertices
Mesh_Point.WeldVertices
End Sub



 Sub ClearPoint3DBuffer() '清空缓存
Point3D_Pointer = 0
ReDim Point3D(1 To 1) As TV_3DVECTOR
Mesh_Point.ResetMesh
End Sub

Sub ClearLineBuffer()
VLine_Pointer = 0
ReDim NumOfPointInVLines(1 To 1) As Long
End Sub

Sub ClearTriangleBuffer()
Triangle_Pointer = 0
ReDim TriangleList(1 To 1) As Triangle
Mesh_Triangle.ResetMesh
End Sub

Sub ClearSamplingBuffer()
Sampling_Pointer = 0
ReDim SamplingPoint(1 To 1) As TV_2DVECTOR
End Sub

Sub ClearBuffer()
ClearSamplingBuffer
ClearPoint3DBuffer
 ClearLineBuffer
ClearTriangleBuffer
End Sub
