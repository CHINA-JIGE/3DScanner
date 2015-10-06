VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Scan3DMeshBuffer2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private SamplingPoint() As TV_2DVECTOR
Private Sampling_Pointer As Long
Private Sampling_Upperbound As Long

Private SamplingCount_Upperbound As Long
Private SamplingCountPerPic() As Long

Private Point3D() As TV_3DVECTOR
Private Point3D_Pointer As Long
Private Point3D_Upperbound As Long

Private TriangleList() As Triangle
Private Triangle_Pointer As Long
Private Triangle_Upperbound As Long

Private Mesh_Point As TVMesh
Private Mesh_Triangle As TVMesh
Private Mesh_tmp As TVMesh


Sub Init()
'Set TVSCENE1 = New TVScene
       Set Mesh_Point = TVSCENE1.CreateMeshBuilder("POINT1")
       Mesh_Point.SetCullMode TV_DOUBLESIDED
       Set Mesh_Triangle = TVSCENE1.CreateMeshBuilder("TRIANGLE1")
       Mesh_Triangle.SetCullMode TV_DOUBLESIDED
       Set Mesh_tmp = TVSCENE1.CreateMeshBuilder("tmpMesh")
       Mesh_tmp.SetCullMode TV_DOUBLESIDED
End Sub

Sub LoadtvmToPointCloud(iPath As String)
Dim i As Long
Dim x As Single, y As Single, z As Single, nx As Single, ny As Single, nz As Single, tv1 As Single, tu1 As Single, tv2 As Single, tu2 As Single, c As Long

If Dir(iPath) <> "" Then
       Mesh_tmp.LoadTVM iPath
       For i = 1 To Mesh_tmp.GetVertexCount
              Mesh_tmp.GetVertex i, x, y, z, nx, ny, nz, tu1, tv1, tu2, tv2, c
              AddPoint x, y, z '本类有管理数组大小的函数
              Mesh_Point.AddVertex x, y, z, 0, 1, 0, 0, 0
       Next i
'Point3D_Pointer = Mesh_Point.GetVertexCount '上界
Mesh_tmp.ResetMesh
Else
       MsgBox "找不到模型！", vbCritical
End If
End Sub











'——————————————————GET————————————————

Function GetPointAmount() As Long '总点数
GetPointAmount = Point3D_Pointer
End Function

Function GetSamplingResult(ID As Long) As TV_2DVECTOR
GetSamplingResult = SamplingPoint(ID)
End Function

Function GetSamplingCount(iPictureID As Long) As Long
GetSamplingCount = SamplingCountPerPic(iPictureID)
End Function

Function GetPictureCount() As Long
GetPictureCount = SamplingCount_Upperbound
End Function

Function GetTriangleAmount() As Long '总三角形数
GetTriangleAmount = Triangle_Pointer
End Function

Function GetPointByID(ID As Long) As TV_3DVECTOR
GetPointByID = Point3D(ID)
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








'——————————————ADD——————————————————
Sub AddSamplingResult(ResultPoint As TV_2DVECTOR)  '采样点
Sampling_Pointer = Sampling_Pointer + 1
'每增5000个点重新分配一次内存
If Sampling_Pointer >= Sampling_Upperbound - 1 Then
       Sampling_Upperbound = Sampling_Upperbound + 200000
       ReDim Preserve SamplingPoint(1 To Sampling_Upperbound) As TV_2DVECTOR
End If
SamplingPoint(Sampling_Pointer) = ResultPoint
End Sub


Sub AddSamplingCountPerPicture(iCount As Long)
'每一张图片共采样了多少个点
       SamplingCount_Upperbound = SamplingCount_Upperbound + 1
       ReDim Preserve SamplingCountPerPic(1 To SamplingCount_Upperbound) As Long
       SamplingCountPerPic(SamplingCount_Upperbound) = iCount
End Sub


Sub AddPoint(x As Single, y As Single, z As Single) '添加顶点
Point3D_Pointer = Point3D_Pointer + 1
'每增5000个点重新分配一次内存
If Point3D_Pointer > Point3D_Upperbound Then
       Point3D_Upperbound = Point3D_Upperbound + 5000
       ReDim Preserve Point3D(1 To Point3D_Upperbound) As TV_3DVECTOR
End If
Point3D(Point3D_Pointer).x = x
Point3D(Point3D_Pointer).y = y
Point3D(Point3D_Pointer).z = z
End Sub

'用顶点ID新增三角形
Sub AddTriangleByPointID(ID1 As Long, ID2 As Long, ID3 As Long)
Triangle_Pointer = Triangle_Pointer + 1
'每5000个三角形重新分配内存
If Triangle_Pointer > Triangle_Upperbound Then
       Triangle_Upperbound = Triangle_Upperbound + 5000
       ReDim Preserve TriangleList(1 To Triangle_Upperbound) As Triangle
End If
TriangleList(Triangle_Pointer).v1 = Point3D(ID1)
TriangleList(Triangle_Pointer).v2 = Point3D(ID2)
TriangleList(Triangle_Pointer).v3 = Point3D(ID3)
End Sub

'用向量新增三角形
Sub AddTriangleByVector(v1 As TV_3DVECTOR, v2 As TV_3DVECTOR, v3 As TV_3DVECTOR)
Triangle_Pointer = Triangle_Pointer + 1
If Triangle_Pointer > Triangle_Upperbound Then
       Triangle_Upperbound = Triangle_Upperbound + 5000
       ReDim Preserve TriangleList(1 To Triangle_Upperbound) As Triangle
End If
TriangleList(Triangle_Pointer).v1 = v1
TriangleList(Triangle_Pointer).v2 = v2
TriangleList(Triangle_Pointer).v3 = v3
End Sub







'————————————SET——————————————————
Sub SetPoint(ID As Long, x As Single, y As Single, z As Single) '修改顶点
If ID > 0 Or ID < Point3D_Pointer Then
Point3D(ID).x = x
Point3D(ID).y = y
Point3D(ID).z = z
End If
End Sub

Sub SetTriangle(TriangleID As Long, v1 As TV_3DVECTOR, v2 As TV_3DVECTOR, v3 As TV_3DVECTOR)
If TriangleID > 0 Or TriangleID < Triangle_Pointer Then
TriangleList(TriangleID).v1 = v1
TriangleList(TriangleID).v2 = v2
TriangleList(TriangleID).v3 = v3
End If
End Sub







'——————————————————ADD_renderMESH——————————————
Sub AddPointCloudToRenderMesh()
Dim i As Long
For i = 1 To GetPointAmount
       Dim p As TV_3DVECTOR
       p = Point3D(i)
       Mesh_Point.AddVertex p.x, p.y, p.z, 0, 1, 0, 0, 0
Next i
End Sub

Sub AddTriangleToRenderMesh()
Dim i As Long
For i = 1 To GetTriangleAmount
       Dim p1 As TV_3DVECTOR, p2 As TV_3DVECTOR, p3 As TV_3DVECTOR
       p1 = GetTriangleVertex(Val(i), 1)
       p2 = GetTriangleVertex(Val(i), 2)
       p3 = GetTriangleVertex(Val(i), 3)
       Mesh_Triangle.AddTriangle 0, p1.x, p1.y, p1.z, p2.x, p2.y, p2.z, p3.x, p3.y, p3.z
Next i
End Sub


'————————————————GetRenderMESH——————————
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



'——————————————CLEAR——————————
 Sub ClearPoint3DBuffer() '清空缓存
Point3D_Pointer = 0
ReDim Point3D(1 To 1) As TV_3DVECTOR
Mesh_Point.ResetMesh
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
ClearTriangleBuffer
End Sub
