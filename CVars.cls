VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CVars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Variable Storage
Dim Vars() As Variant
Dim MyStack As CStack

Function CreateVariableIndex(VariableName As String) As Integer
    If InStr(VariableName, "[") = 0 Then
        'It is not a array, proceed, to search through
        ' the array of variablenames. If one is found
        ' return that variable index, otherwise, return
        ' the next variable index.
        
        CreateVariableIndex = -1
        For a = 1 To UBound(Vars)
            If VariableName = Vars(a) Then
                CreateVariableIndex = a
                Exit Function
            End If
        Next a
        If CreateVariableIndex = -1 Then
            CreateVariableIndex = UBound(Vars) + 1
            ReDim Preserve Vars(UBound(Vars) + 1)
            Vars(UBound(Vars)) = VariableName
        End If
    Else
        tmpArrayNames = Split(GenerateArrayNames(VariableName, InStr(VariableName, "[")), vbCrLf)
        ReDim Preserve Vars(UBound(Vars) + UBound(tmpArrayNames))
        For Counter = 1 To UBound(tmpArrayNames)
            Vars(UBound(Vars) - Counter) = tmpArrayNames(Counter)
        Next Counter
    End If
End Function

Function GetVariableIndex(VariableName As String) As Integer
    If VariableName <> "" Then
        GetVariableIndex = -1
        For a = 1 To UBound(Vars)
            If VariableName = Vars(a) Then
                GetVariableIndex = a
                Exit Function
            End If
        Next a
        If GetVariableIndex = -1 Then
            MsgBox "Non existant variable: " & VariableName
        End If
    Else
        GetVariableIndex = -2
    End If
End Function

'Sub SetVariable(Index As Integer, Data As Variant)
'    Vars(Index) = Data
'End Sub

'Function GetVariable(Index As Integer) As Variant
'    GetVariable = Vars(Index)
'End Function

Sub DumpVars()
    Print "Variables -------------"
    For a = 1 To UBound(Vars)
        Debug.Print a & "     "; Vars(a)
    Next a
End Sub

Public Sub AddVarBarrier()
    Call MyStack.Push(UBound(Vars) + 1)
End Sub

Public Sub RemoveVarBarrier()
    tmpPos = MyStack.Pop
    If tmpPos <= UBound(Vars) Then
        For a = tmpPos To UBound(Vars)
            Vars(a) = "#NonExistant"
        Next a
    End If
End Sub

Private Sub Class_Initialize()
    ReDim Vars(0)
    Set MyStack = New CStack
    Call MyStack.Initilize(50)
End Sub

Private Function GenerateArrayNames(ArrayInfo As String, Position As Integer) As String
    If Position < Len(ArrayInfo) Then
        tmpPos = InStr(Mid(ArrayInfo, Position), "]")
        For Counter = 1 To Mid(ArrayInfo, Position + 1, tmpPos - 2)
            tmpArrayInfo = Left(ArrayInfo, Position) & Counter & Right(ArrayInfo, Len(ArrayInfo) - Position - tmpPos + 2)
            GenerateArrayNames = GenerateArrayNames(CStr(tmpArrayInfo), tmpPos + Position) & vbCrLf & GenerateArrayNames
        Next Counter
    Else
        GenerateArrayNames = ArrayInfo
    End If
End Function

