VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCodeTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'CodeTree - Given a stream of tokens, turn it into a tree

Dim Tokener As CTokens
Public Variables As CVars
Dim CurrentSubroutine As String

Function CreateTree(Code As String) As CTree
    Dim Location As Integer
    Dim tmpToken As String
    
    Dim tmpTree As CTree
    Set tmpTree = New CTree
    Code = Trim(Code)
    
    tmpTree.CodeLine = Code
    If Tokener.ContainArguments(CStr(Code)) = False Then
        If IsNumeric(Code) Then
            tmpTree.Value = Code
        ElseIf Left(Code, 1) = Chr(34) Then ' a string
            tmpTree.Value = Code
        Else
            tmpVar = Variables.GetVariableIndex(CStr(Code))
            If tmpVar >= 0 Then
                tmpTree.Value = "VAR" & Variables.GetVariableIndex(CStr(Code))
            End If
        End If
    Else
        Location = Tokener.GetLowestPrecedence(Code, tmpToken)
        If Location = 1 Then
            Select Case LCase(tmpToken)
            Case "if"
                Part1End = FindClosing(Right(Code, Len(Code) - 2))
                tmpTree.Value = "if"
                Set tmpTree.LeftNode = CreateTree(Mid(Code, 3, Part1End))
                Set tmpTree.RightNode = CreateTree(Mid(Code, Part1End + 3, Len(Code) - Part1End - 2))
                
            Case "while"
                Part1End = FindClosing(Right(Code, Len(Code) - 5))
                tmpTree.Value = "while"
                Set tmpTree.LeftNode = CreateTree(Mid(Code, 6, Part1End))
                Set tmpTree.RightNode = CreateTree(Mid(Code, Part1End + 6, Len(Code) - Part1End - 5))
            
            Case "for"
                Part1End = FindClosing(Right(Code, Len(Code) - 3))
                tmpTree.Value = "for"
                Set tmpTree.LeftNode = CreateTree(Mid(Code, 4, Part1End))
                Set tmpTree.RightNode = CreateTree(Mid(Code, Part1End + 4, Len(Code) - Part1End - 3))
                
            Case "*"
                tmpTree.Value = "Deref"
                Set tmpTree.LeftNode = CreateTree(Right(Code, Len(Code) - 1))
                
            Case "(", "{"
                Variables.AddVarBarrier
                Set tmpTree = CreateTree(Mid(Code, 2, Len(Code) - 2))
                Variables.RemoveVarBarrier
                
            Case "-"
                tmpTree.Value = "Neg"
                Set tmpTree.LeftNode = CreateTree(Right(Code, Len(Code) - 1))
                
            Case "break"
                tmpTree.Value = "break"
                
            Case "int", "void"
                tmpTree.Value = tmpToken
                If FindOpening(Code) <> -1 Then
                    Set tmpTree.LeftNode = New CTree
                    tmpTree.LeftNode.Value = Mid(Code, Len(tmpToken) + 2, FindOpening(Code) - Len(tmpToken) - 2)
                    Set tmpTree.RightNode = New CTree
                    tmpTree.RightNode.Value = "("
                    
                    Variables.AddVarBarrier
                    GetCurrentSubroutine = Mid(Code, Len(tmpToken) + 2, FindOpening(Code) - Len(tmpToken) - 2)
                    'Generate parameters that are passed in sub
                    Set tmpTree.RightNode.LeftNode = CreatePopTree(Mid(Code, FindOpening(Code) + 1, FindClosing(Code) - FindOpening(Code) - 1))
                    
                    'Generate code in module
                    tmp1 = FindOpening(Right(Code, Len(Code) - FindClosing(Code)))
                    tmp2 = FindClosing(Right(Code, Len(Code) - FindClosing(Code)))
                    Set tmpTree.RightNode.RightNode = CreateTree(Mid(Code, tmp1 + FindClosing(Code), tmp2 + FindClosing(Code)))
                    Variables.RemoveVarBarrier
                Else
                    Call Variables.CreateVariableIndex(Trim(Mid(Code, Len(tmpToken) + 1, Len(Code) - Len(tmpToken))))
                End If
                
                Case "return"
                    Set tmpTree = CreateTree(Right(Code, Len(Code) - 6))
                    
            End Select
        Else
            If tmpToken = "(" Then
                If CurrentSubroutine = Left(Code, Location - 1) Then
                
                Else
                    tmpTree.Value = tmpToken
                    Set tmpTree.LeftNode = New CTree
                    tmpTree.LeftNode.Value = Left(Code, Location - 1)
                    Set tmpTree.RightNode = CreateTree(Right(Code, Len(Code) - Location - Len(tmpToken) + 2))
                End If
            Else
                tmpTree.Value = tmpToken
                Set tmpTree.LeftNode = CreateTree(Left(Code, Location - 1))
                Set tmpTree.RightNode = CreateTree(Right(Code, Len(Code) - Location - Len(tmpToken) + 1))
            End If
        End If
    End If
    
    Set CreateTree = tmpTree
End Function

Private Function FindClosing(Code As String) As Integer
    Dim Bracket As Integer
    Start = False
    FindClosing = -1
    
    For a = 1 To Len(Code)
        Select Case Mid(Code, a, 1)
        Case "(", "{"
            Bracket = Bracket + 1
            Start = True
        Case ")", "}"
            Bracket = Bracket - 1
        End Select
        
        If Bracket = 0 And Start = True Then
            FindClosing = a
            Exit Function
        End If
    Next a
End Function

Private Function FindOpening(Code)
    FindOpening = -1
    
    For a = 1 To Len(Code)
        Select Case Mid(Code, a, 1)
        Case "(", "{"
            FindOpening = a
            Exit Function
        End Select
    Next a
            
End Function

Private Sub Class_Initialize()
    Set Tokener = New CTokens
End Sub

Private Function CreatePopTree(Code As String) As CTree
    Dim Location As Integer
    Dim tmpToken As String

    Dim tmpTree As CTree
    Set tmpTree = New CTree
    Code = Trim(Code)
    If Code <> "" Then
        Location = Tokener.GetLowestPrecedence(Code, tmpToken)
        Select Case tmpToken
        Case ","
            Set tmpTree.RightNode = CreatePopTree(Left(Code, Location - 1))
            Set tmpTree.LeftNode = CreatePopTree(Right(Code, Len(Code) - Location))
            tmpTree.Value = ","
        Case "int", "void"
            tmpVarNo = Variables.CreateVariableIndex(Trim(Right(Code, Len(Code) - Len(tmpToken))))
            tmpTree.Value = "VAR" & tmpVarNo
        End Select
    End If
    Set CreatePopTree = tmpTree
End Function
