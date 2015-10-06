Attribute VB_Name = "MCsTxtTracing"
' MCsTxtTracing - PLEASE DO NOT DELETE THIS LINE
'                 AND DO NOT ALTER THE METHOD NAMES IN ANY WAY

'CSEH: Skip

Option Explicit

' Specify the position where the tracing instructions are inserted
Public Enum ProcPosition
    ProcEnter
    ProcExit
    ProcInside
End Enum

' Text indentation information
Private m_Indent As Long

' File information
Private m_iFile  As Integer

Public Const E_ERR_CS_TRACING_INIT = vbObjectError + 1265

Private Const S_ERR_CS_TRACING_INIT = "Could not initiate tracing"

' Member      : AxCsInitiateTrace
' Description : Instantiate the text file data
Public Sub AxCsInitiateTrace()

    Static bErrReported As Boolean

    On Error GoTo hErr
    
    ' Create a new file
    On Error Resume Next

    m_iFile = FreeFile
    
    Open "F:\新建文本文
    
    If Err.Number <> 0 Then
        If Not bErrReported Then
            bErrReported = True
            Err.Raise E_ERR_CS_TRACING_INIT + 1, "AxCsInitiateTrace", _
                "Could not open the text file."
        End If
        Exit Sub
    End If
    
    Exit Sub

hErr:
    Err.Raise E_ERR_CS_TRACING_INIT, "AxCsInitiateTrace", S_ERR_CS_TRACING_INIT

End Sub

' Member      : AxCsTerminateTrace
' Description : Used for cleaning purposes
Public Sub AxCsTerminateTrace()
    
    Close #m_iFile
    
    m_iFile = 0
    
End Sub

' Member      : AxCsTrace
' Description : Send information to the text file regarding the current method being processed
' Parameters  : ProjectName   - Name of the project which contains the method
'               ComponentName - Name of the component which contains the method
'               MemberName    - Name of the method being processed
'               TracePosition - Indicates the position within method body, either at start, at exit,
'                               or inside it, where the AxCsTxtTraceWatch method is called
' Notes       : For inside-member calls of AxCsTrace method, you can use the ProjectName
'               parameter to send tracing information.
Public Sub AxCsTrace(ByVal ProjectName As String, _
                     Optional ByVal ComponentName As String = "", _
                     Optional ByVal MemberName As String = "", _
                     Optional ByVal TracePosition As ProcPosition = ProcInside)

    Dim sTemp As String
    
    ' Make sure that tracing is initiated
    If m_iFile = 0 Then AxCsInitiateTrace

    ' The color, bold state and suffix can be customized below
    If TracePosition = ProcEnter Then
        sTemp = " - enter"
    ElseIf TracePosition = ProcExit Then
        sTemp = " - exit"
    Else
        sTemp = ""
    End If
    
    ' Decrease the indent for the current level of tracing information
    If TracePosition = ProcExit Then
        m_Indent = m_Indent - 4
    End If

    If Not (TracePosition = ProcInside) Then
        sTemp = ProjectName & "." & ComponentName & "." & MemberName & sTemp
    Else
        sTemp = ProjectName
    End If
    
    ' Add tracing information to the text file
    On Error Resume Next
    Print #m_iFile, CStr(Time) & "   " & String$(m_Indent, ".") & sTemp
    
    ' Increase the indent for the next level of tracing information
    If TracePosition = ProcEnter Then
        m_Indent = m_Indent + 4
    End If
    
    AxCsTerminateTrace
    
End Sub

' Member      : AxCsDumpParamValue
' Description : Returns a formatted information about each method parameter
' Parameters  : sParamName    - Parameter name
'               vParamValue   - Parameter value
Public Function AxCsDumpParamValue(sParamName As String, vParamValue As Variant) As String

    Dim sRet$
    
    If IsObject(vParamValue) Then
    
        sRet = "Object"
        
    ElseIf IsArray(vParamValue) Then
    
        sRet = "Array"
        
    Else
    
        On Error Resume Next
        sRet = CStr(vParamValue)
        
        If Err.Number <> 0 Then
            sRet = "Not determined"
        Else
            If Len(sRet) > 13 Then sRet = Left$(sRet, 10) & "..."
        End If
        
    End If
    
    AxCsDumpParamValue = sParamName & ": [" & sRet & "]"

End Function

