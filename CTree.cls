VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Tree with two nodes

Public LeftNode As CTree
Public RightNode As CTree
Public Value As String
Public CodeLine As String

Private Sub Class_Initialize()
    Set LeftNode = Me
    Set RightNode = Me
    Value = ""
    CodeLine = ""
End Sub

Public Function DumpTree(Indent As Integer)
    On Error Resume Next
    frmOutput.txtOutput = frmOutput.txtOutput & Space(Indent) & Indent / 2 & ") " & Value & Space(20 - Len(Value)) & "Codeline: " & MyReplace(CodeLine & vbCrLf, "  ", " ")
    
    If ObjPtr(LeftNode) <> ObjPtr(Me) Then
        Call LeftNode.DumpTree(Indent + 2)
    End If
    
    If ObjPtr(RightNode) <> ObjPtr(Me) Then
        Call RightNode.DumpTree(Indent + 2)
    End If
End Function

Public Function TotalNodes()
    Number = 1
    If ObjPtr(LeftNode) <> ObjPtr(Me) Then
        Number = Number + LeftNode.TotalNodes
    End If
    
    If ObjPtr(RightNode) <> ObjPtr(Me) Then
        Number = Number + RightNode.TotalNodes
    End If
    TotalNodes = Number
End Function

Private Function MyReplace(Expression As String, FindString As String, ReplaceString As String)
    For Counter = 1 To Len(Expression) - Len(FindString)
        If Mid(Expression, Counter, Len(FindString)) = FindString Then
            Expression = Left(Expression, Counter - 1) & ReplaceString & Right(Expression, Len(Expression) - Counter - Len(FindString) + 1)
            Counter = Counter - Len(FindString) + 1
        End If
    Next Counter
    MyReplace = Expression
End Function
