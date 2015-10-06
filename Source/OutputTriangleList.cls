VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OutputTriangleList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Type Triangle
v1 As TV_3DVECTOR
v2 As TV_3DVECTOR
v3 As TV_3DVECTOR
End Type
Private Pointer As Long
Private list() As Triangle


Public Sub AddToList(Vertex1 As TV_3DVECTOR, Vertex2 As TV_3DVECTOR, Vertex3 As TV_3DVECTOR)
Pointer = Pointer + 1
ReDim Preserve list(Pointer) As Triangle
list(Pointer).v1 = Vertex1
list(Pointer).v2 = Vertex2
list(Pointer).v3 = Vertex3
End Sub


Public Sub clear()
Pointer = 0
ReDim list(0) As Triangle
End Sub

Public Function GetUpperBound() As Long
GetUpperBound = Pointer
End Function



Public Function GetVertex(TriangleID As Variant, vID As Long) As TV_3DVECTOR
Select Case vID
Case 1
GetVertex = list(TriangleID).v1
Case 2
GetVertex = list(TriangleID).v2
Case 3
GetVertex = list(TriangleID).v3
End Select
End Function
