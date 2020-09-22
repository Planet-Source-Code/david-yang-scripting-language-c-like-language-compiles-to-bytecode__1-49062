VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim MyStack As CStack
Dim CallStack As CStack
Dim Variables() As Variant
Dim OwnFunctions() As String

Sub RunCode(Code As String)
    OwnFunctions = Split("sin|cos|tan|abs|log|oct|hex|sqr|round|rnd|msgbox|inputbox|left|right|mid|trim|date|day|hour|minute|month|time|timer", "|")
    Dim Counter As Integer
    Dim Op1 As Variant, Op2 As Variant
    Dim InMain As Integer
    Set MyStack = New CStack
    Set CallStack = New CStack
    Call CallStack.Initilize(50)
    Call MyStack.Initilize(50)
    
    For Counter = 1 To Len(Code)
        frmOutput.Output vbCrLf & "Line " & Counter
        Select Case Mid(Code, Counter, 1)
        Case "-", "+", "*", "/", ">", "<", Chr(161), Chr(162)
            Op1 = MyStack.Pop
            Op2 = MyStack.Pop
            Select Case Mid(Code, Counter, 1)
            Case "-": Call MyStack.Push(Op2 - Op1)
            Case "+": Call MyStack.Push(Op2 + Op1)
            Case "*": Call MyStack.Push(Op2 * Op1)
            Case "/": Call MyStack.Push(Op2 / Op1)
            Case ">": Call MyStack.Push(Op2 > Op1)
            Case "<": Call MyStack.Push(Op2 < Op1)
            Case Chr(161): Call MyStack.Push(Op2 And Op1)
            Case Chr(162): Call MyStack.Push(Op2 Or Op1)
            End Select
frmOutput.Output "Do operation: " & Mid(Code, Counter, 1) & " Operators: " & Op1 & " & " & Op2
        'Case "DEREF" /chr(163)
        '    GetOpCode = Chr(163)
        Case Chr(164) 'NEGITIVE SIGN
            Op1 = MyStack.Pop
            Call MyStack.Push(-Op1)
frmOutput.Output "Negative Variable: Operators: " & Op1
        Case Chr(165) 'Jump if false
frmOutput.Output "Jump if false: Location: " & GetNumber(Mid(Code, Counter + 1, 2)) + 1
            Op1 = MyStack.Pop
            If Op1 = False Then
                Counter = GetNumber(Mid(Code, Counter + 1, 2)) + 1
            Else
                Counter = Counter + 2
            End If
        Case Chr(166) 'Jump
frmOutput.Output "Jump: Location: " & GetNumber(Mid(Code, Counter + 1, 2)) + 1
            Counter = GetNumber(Mid(Code, Counter + 1, 2)) + 1
        Case Chr(167) 'Call
frmOutput.Output "Call: Location: " & GetNumber(Mid(Code, Counter + 1, 2)) + 1
            Call CallStack.Push(Counter + 3)
            Counter = GetNumber(Mid(Code, Counter + 1, 2)) + 1
            InMain = InMain + 1
        'Case "SUB"
            'Chr(168)
        Case Chr(169) 'RET
            InMain = InMain - 1
            If InMain = -1 Then
frmOutput.Output "Ret in main sub... exiting execution"
                Exit Sub
            Else
                Counter = CallStack.Pop - 1
frmOutput.Output "Ret in normal sub... returning"
            End If
        Case Chr(170) 'PUSH the number type
            Call MyStack.Push(GetNumber(Mid(Code, Counter + 1, 2)))
            Counter = Counter + 2
frmOutput.Output "Push number/text: Data: " & GetNumber(Mid(Code, Counter - 1, 2))
        Case Chr(171) 'PUSH the variabel type
            Call MyStack.Push(GetVariable(GetNumber(Mid(Code, Counter + 1, 2))))
            Counter = Counter + 2
        Case Chr(172) 'POP the variable type
            Call SetVariable(GetNumber(Mid(Code, Counter + 1, 2)), MyStack.Pop)
            Counter = Counter + 2
        Case "JMPD" 'Jump Destination label name
            'GetOpCode = Chr(173)
        Case Chr(174)
            Select Case OwnFunctions(Asc(Mid(Code, Counter + 1, 1)))
            Case "sin": Op1 = MyStack.Pop: MyStack.Push (Sin(Op1))
            Case "cos": Op1 = MyStack.Pop: MyStack.Push (Cos(Op1))
            Case "tan": Op1 = MyStack.Pop: MyStack.Push (Tan(Op1))
            Case "abs": Op1 = MyStack.Pop: MyStack.Push (Abs(Op1))
            Case "log": Op1 = MyStack.Pop: MyStack.Push (Log(Op1))
            Case "hex": Op1 = MyStack.Pop: MyStack.Push (Hex(Op1))
            Case "oct": Op1 = MyStack.Pop: MyStack.Push (Oct(Op1))
            Case "sqr": Op1 = MyStack.Pop: MyStack.Push (Sqr(Op1))
            Case "round": Op1 = MyStack.Pop: MyStack.Push (Int(Op1))
            Case "msgbox": Op1 = MyStack.Pop: MsgBox (Op1)
            End Select
            Counter = Counter + 1
        Case Else
            MsgBox "Error"
            Exit Sub
        End Select
        MyStack.DumpStack
    Next Counter
End Sub

Private Function GetNumber(HexCode As String) As Integer
    For a = 1 To Len(HexCode)
        GetNumber = GetNumber * 256
        GetNumber = GetNumber + Asc(Mid(HexCode, a, 1))
    Next a
End Function

Private Function GetVariable(Index As Integer) As Variant
    If Index > UBound(Variables) Then
        ReDim Preserve Variables(0 To Index)
    End If
    GetVariable = Variables(Index)
frmOutput.Output "Get Variable: var" & Index
End Function

Private Function SetVariable(Index As Integer, Value As Variant)
    If Index > UBound(Variables) Then
        ReDim Preserve Variables(0 To Index)
    End If
    Variables(Index) = Value
frmOutput.Output "Set Variable: var" & Index & " to " & Value
End Function

Private Sub Class_Initialize()
    ReDim Variables(0)
End Sub

Sub RunCodeNoLog(Code As String)
    OwnFunctions = Split("sin|cos|tan|abs|log|oct|hex|sqr|round|rnd|msgbox|inputbox|left|right|mid|trim|date|day|hour|minute|month|time|timer", "|")
    Dim Counter As Integer
    Dim Op1 As Variant, Op2 As Variant
    Dim InMain As Integer
    Set MyStack = New CStack
    Set CallStack = New CStack
    Call CallStack.Initilize(50)
    Call MyStack.Initilize(50)
    
    For Counter = 1 To Len(Code)
        Select Case Mid(Code, Counter, 1)
        Case "-", "+", "*", "/", ">", "<", Chr(161), Chr(162)
            Op1 = MyStack.Pop
            Op2 = MyStack.Pop
            Select Case Mid(Code, Counter, 1)
            Case "-": Call MyStack.Push(Op2 - Op1)
            Case "+": Call MyStack.Push(Op2 + Op1)
            Case "*": Call MyStack.Push(Op2 * Op1)
            Case "/": Call MyStack.Push(Op2 / Op1)
            Case ">": Call MyStack.Push(Op2 > Op1)
            Case "<": Call MyStack.Push(Op2 < Op1)
            Case Chr(161): Call MyStack.Push(Op2 And Op1)
            Case Chr(162): Call MyStack.Push(Op2 Or Op1)
            End Select
        Case Chr(163) 'Dereference
            Op1 = MyStack.Pop
            Call MyStack.Push(GetVariable(CInt(Op1)))
        Case Chr(164) 'NEGITIVE SIGN
            Op1 = MyStack.Pop
            Call MyStack.Push(-Op1)
        Case Chr(165) 'Jump if false
            Op1 = MyStack.Pop
            If Op1 = False Then
                Counter = GetNumber(Mid(Code, Counter + 1, 2)) + 1
            Else
                Counter = Counter + 2
            End If
        Case Chr(166) 'Jump
            Counter = GetNumber(Mid(Code, Counter + 1, 2)) + 1
        Case Chr(167) 'Call
            Call CallStack.Push(Counter + 3)
            Counter = GetNumber(Mid(Code, Counter + 1, 2)) + 1
            InMain = InMain + 1
        'Case "SUB"
            'Chr(168)
        Case Chr(169) 'RET
            InMain = InMain - 1
            If InMain = -1 Then
                Exit Sub
            Else
                Counter = CallStack.Pop - 1
            End If
        Case Chr(170) 'PUSH the number type
            Call MyStack.Push(GetNumber(Mid(Code, Counter + 1, 2)))
            Counter = Counter + 2
        Case Chr(171) 'PUSH the variabel type
            Call MyStack.Push(GetVariableNoLog(GetNumber(Mid(Code, Counter + 1, 2))))
            Counter = Counter + 2
        Case Chr(172) 'POP the variable type
            Call SetVariableNoLog(GetNumber(Mid(Code, Counter + 1, 2)), MyStack.Pop)
            Counter = Counter + 2
        'Case "JMPD" Jump Destination label name
            'GetOpCode = Chr(173)
        Case Chr(174)
            Select Case OwnFunctions(Asc(Mid(Code, Counter + 1, 1)))
            Case "sin": Op1 = MyStack.Pop: MyStack.Push (Sin(Op1))
            Case "cos": Op1 = MyStack.Pop: MyStack.Push (Cos(Op1))
            Case "tan": Op1 = MyStack.Pop: MyStack.Push (Tan(Op1))
            Case "abs": Op1 = MyStack.Pop: MyStack.Push (Abs(Op1))
            Case "log": Op1 = MyStack.Pop: MyStack.Push (Log(Op1))
            Case "hex": Op1 = MyStack.Pop: MyStack.Push (Hex(Op1))
            Case "oct": Op1 = MyStack.Pop: MyStack.Push (Oct(Op1))
            Case "sqr": Op1 = MyStack.Pop: MyStack.Push (Sqr(Op1))
            Case "round": Op1 = MyStack.Pop: MyStack.Push (Int(Op1))
            Case "msgbox": Op1 = MyStack.Pop: MsgBox (Op1)
            End Select
            Counter = Counter + 1
        Case Else
            MsgBox "Error"
            Exit Sub
        End Select
    Next Counter
End Sub

Private Function GetVariableNoLog(Index As Integer) As Variant
    If Index > UBound(Variables) Then
        ReDim Preserve Variables(0 To Index)
    End If
    GetVariableNoLog = Variables(Index)
End Function

Private Function SetVariableNoLog(Index As Integer, Value As Variant)
    If Index > UBound(Variables) Then
        ReDim Preserve Variables(0 To Index)
    End If
    Variables(Index) = Value
End Function
