VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CGenIntCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'GenIntCode - Generate Intermediate code
'   GenerateCode - creates code, given a tree. uses recursion
'   OutputMachne - sets setting wether to output human-readable form, or the coded version

Dim UniqueNo As Integer
Dim MyRegisters As CRegisters
Public MyVars As CVars
Dim Machine As Boolean 'Output human form, or intermediate code

Function GenerateCode(Tree As CTree) As String
    Select Case Tree.Value
    Case "-", "+", "*", "/", "&&", "||", ">", "<"
        GenerateCode = GenerateCode(Tree.LeftNode)
        GenerateCode = GenerateCode & GenerateCode(Tree.RightNode)
        GenerateCode = GenerateCode & GetOpCode(Tree.Value) & GetOpCode(vbCrLf)
        
    Case ","
        GenerateCode = GenerateCode(Tree.LeftNode)
        GenerateCode = GenerateCode & GenerateCode(Tree.RightNode)
    
    Case "="
        GenerateCode = GenerateCode & GenerateCode(Tree.RightNode)
        GenerateCode = GenerateCode & GetOpCode("POP") & Right(Tree.LeftNode.Value, Len(Tree.LeftNode.Value) - 3) & GetOpCode(vbCrLf)
    
    Case "Deref", "Neg"
        GenerateCode = GenerateCode(Tree.LeftNode)
        GenerateCode = GenerateCode & GetOpCode(Tree.Value) & GetOpCode(vbCrLf)
    
    Case "if"
        tmpUnique1 = GetUniqueNo
        tmpUnique2 = GetUniqueNo
        
        GenerateCode = GenerateCode(Tree.LeftNode)
        GenerateCode = GenerateCode & GetOpCode("JF") & "EndBlock_" & tmpUnique1 & GetOpCode(vbCrLf)
        If Tree.RightNode.Value = "else" Then
            'IF with the ELSE
            GenerateCode = GenerateCode & GenerateCode(Tree.RightNode.LeftNode)
            GenerateCode = GenerateCode & GetOpCode("JMP") & "EndBlock_" & tmpUnique2 & GetOpCode(vbCrLf)
            GenerateCode = GenerateCode & GetOpCode("JMPD") & "EndBlock_" & tmpUnique1 & GetOpCode(Chr(22))
            GenerateCode = GenerateCode & GenerateCode(Tree.RightNode.RightNode)
            GenerateCode = GenerateCode & GetOpCode("JMPD") & "EndBlock_" & tmpUnique2 & GetOpCode(Chr(22))
        Else
            'IF without the ELSE
            GenerateCode = GenerateCode & GenerateCode(Tree.RightNode)
            GenerateCode = GenerateCode & GetOpCode("JMPD") & "EndBlock_" & tmpUnique1 & GetOpCode(Chr(22))
        End If
        
    Case "while"
        tmpUnique1 = GetUniqueNo
        tmpUnique2 = GetUniqueNo
        
        GenerateCode = GetOpCode("JMPD") & "BeginBlock_" & tmpUnique1 & GetOpCode(Chr(22))
        GenerateCode = GenerateCode & GenerateCode(Tree.LeftNode)
        GenerateCode = GenerateCode & GetOpCode("JF") & "EndBlock_" & tmpUnique2 & GetOpCode(vbCrLf)
        GenerateCode = GenerateCode & GenerateCode(Tree.RightNode)
        GenerateCode = GenerateCode & GetOpCode("JMP") & "BeginBlock_" & tmpUnique1 & GetOpCode(vbCrLf)
        GenerateCode = GenerateCode & GetOpCode("JMPD") & "EndBlock_" & tmpUnique2 & GetOpCode(Chr(22))
    
    Case "for"
        tmpUnique1 = GetUniqueNo
        tmpUnique2 = GetUniqueNo
        
        GenerateCode = GenerateCode(Tree.LeftNode.LeftNode)
        GenerateCode = GenerateCode & GetOpCode("JMPD") & "BeginBlock_" & tmpUnique1 & GetOpCode(Chr(22))
        GenerateCode = GenerateCode & GenerateCode(Tree.LeftNode.RightNode.LeftNode)
        GenerateCode = GenerateCode & GetOpCode("JF") & "EndBlock_" & tmpUnique2 & GetOpCode(vbCrLf)
        GenerateCode = GenerateCode & GenerateCode(Tree.RightNode)
        GenerateCode = GenerateCode & GenerateCode(Tree.LeftNode.RightNode.RightNode)
        GenerateCode = GenerateCode & GetOpCode("JMP") & "BeginBlock_" & tmpUnique1 & GetOpCode(vbCrLf)
        GenerateCode = GenerateCode & GetOpCode("JMPD") & "EndBlock_" & tmpUnique2 & GetOpCode(Chr(22))
    
    Case "("
        If Tree.RightNode.Value <> "" Then
            GenerateCode = GenerateCode(Tree.RightNode)
        End If
        GenerateCode = GenerateCode & GetOpCode("Call") & Tree.LeftNode.Value & GetOpCode(vbCrLf)
    
    Case ";"
        GenerateCode = GenerateCode(Tree.LeftNode)
        GenerateCode = GenerateCode & GenerateCode(Tree.RightNode)
    
    Case ""
        'Ignore
        
    'Case "break"
    '    GenerateCode = GetOpCode("JMP") & " EndBlock_" & UniqueNo - 1 & GetOpCode(vbCrLf)
    '    GenerateCode = GenerateCode & "Still FAULTY - Break" & GetOpCode(vbCrLf)
        
    Case "void", "int"
        If Tree.RightNode.Value = "(" Then
            GenerateCode = GetOpCode("Sub") & Tree.LeftNode.Value & GetOpCode(Chr(22)) '& GetOpCode(vbCrLf)
            If Tree.RightNode.LeftNode.Value <> "" Then
                GenerateCode = GenerateCode & GeneratePopCalls(Tree.RightNode.LeftNode)
            End If
            GenerateCode = GenerateCode & GenerateCode(Tree.RightNode.RightNode)
            GenerateCode = GenerateCode & GetOpCode("RET") & GetOpCode(vbCrLf)
        Else
            'GenerateCode = "CreateVar: " & Tree.LeftNode.Value & GetOpCode(vbCrLf)
        End If
        
    Case Else
        If IsNumeric(Tree.Value) Then
            GenerateCode = GetOpCode("PUSH") & Tree.Value & GetOpCode(vbCrLf)
        ElseIf Left(Tree.Value, 1) = Chr(34) Then
            GenerateCode = GetOpCode("PUSH") & Tree.Value & GetOpCode(vbCrLf)
        Else
            GenerateCode = GetOpCode("PUSHV") & Right(Tree.Value, Len(Tree.Value) - 3) & GetOpCode(vbCrLf)
        End If
    End Select
End Function

Private Function GetUniqueNo() As Integer
    GetUniqueNo = UniqueNo
    UniqueNo = UniqueNo + 1
End Function

Private Function GeneratePopCalls(Tree As CTree)
    GeneratePopCalls = GeneratePopCalls & GetOpCode("POP") & Right(Tree.LeftNode.Value, Len(Tree.LeftNode.Value) - 3) & GetOpCode(vbCrLf)
    If Tree.Value = "," Then
        GeneratePopCalls = GeneratePopCalls & GeneratePopCalls(Tree.RightNode)
    End If
End Function

Private Function GeneratePushCalls(Tree As CTree)
    If Tree.Value = "," Then
        GeneratePushCalls = GenerateCode(Tree.LeftNode)
        GeneratePushCalls = GeneratePushCalls & GeneratePushCalls(Tree.RightNode)
    Else
        GeneratePushCalls = GenerateCode(Tree)
    End If
End Function

Private Sub Class_Initialize()
    Set MyRegisters = New CRegisters
    Machine = True
End Sub

Property Let OutputMachine(Data As Boolean)
    Machine = Data
End Property

Private Function GetOpCode(Code As String) As String
    If Machine Then
        Code = UCase(Code)
        Select Case Code
        Case "-", "+", "*", "/", ">", "<"
            GetOpCode = Code
        Case "&&"
            GetOpCode = Chr(161)
        Case "||"
            GetOpCode = Chr(162)
        Case "DEREF"
            GetOpCode = Chr(163)
        Case "NEG"
            GetOpCode = Chr(164)
        Case "JF"
            GetOpCode = Chr(165)
        Case "JMP"
            GetOpCode = Chr(166)
        Case "CALL"
            GetOpCode = Chr(167)
        Case "SUB"
            GetOpCode = Chr(168)
        Case "RET"
            GetOpCode = Chr(169)
        Case "PUSH"
            GetOpCode = Chr(170)
        Case "PUSHV" 'Push Var
            GetOpCode = Chr(171)
        Case "POP"
            GetOpCode = Chr(172)
        Case "JMPD" 'Jump Destination label name
            GetOpCode = Chr(173)
        Case "CALLF"
            'Special "own functions" - eg sin and cos
            'By the way, this is not done in here
            'it is done in the CFilter class
            GetOpCode = Chr(174)
        Case Chr(22)
            GetOpCode = Chr(22)
        Case vbCrLf
            GetOpCode = vbCrLf
        End Select
    Else
        GetOpCode = Code
        If Code = Chr(22) Then GetOpCode = vbCrLf
    End If
End Function

