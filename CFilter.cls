VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This dereferences all the jumps, and does other things in
' a second step after the intcode is generated
Dim MyIntCode() As String
Dim JumpLabels() As String
Dim MyNewIntCode() As String
Dim CodeLength() As Integer

Function Filter(IntCode As String) As String
    MyIntCode = Split(IntCode, vbCrLf)
    ReDim MyNewIntCode(UBound(MyIntCode))
    ReDim JumpLabels(UBound(MyIntCode))
    ReDim CodeLength(UBound(MyIntCode))
    Dim OwnFunctions() As String
    OwnFunctions = Split("sin|cos|tan|abs|log|oct|hex|sqr|round|rnd|msgbox|inputbox|left|right|mid|trim|date|day|hour|minute|month|time|timer", "|")
    
    'Search for subs, and jmp locs
    For a = 0 To UBound(MyIntCode)
        'Get jump labels
        Do
            If Left(MyIntCode(a), 1) = Chr(168) Or Left(MyIntCode(a), 1) = Chr(173) Then
                For b = 1 To Len(MyIntCode(a))
                    If Asc(Mid(MyIntCode(a), b, 1)) = 22 Then
                        tmpPos = b
                        Exit For
                    End If
                Next b
                JumpLabels(a) = Mid(MyIntCode(a), 2, b - 2) ' & "|"
                'NOT DONE - WHAT IF MULTIPLE JUMPS LAND ON SAME SPOT???
                MyIntCode(a) = Mid(MyIntCode(a), b + 1, Len(MyIntCode(a)) - b + 1)
                IsChecked = True
            Else
                IsChecked = False
            End If
        Loop Until Not IsChecked
        
        If Left(MyIntCode(a), 1) = Chr(167) Then
            'Get owner functions
            Found = False
            For b = 0 To UBound(OwnFunctions)
                If Right(MyIntCode(a), Len(MyIntCode(a)) - 1) = OwnFunctions(b) Then
                    Found = True
                    tmpPos = b
                    Exit For
                End If
            Next b
            
            If Found = True Then
            'Owner function
                MyIntCode(a) = Chr(174) & Chr(tmpPos)
            End If
        End If
    Next a
    
    'Find code lengths for the jumps
    CodeLength(0) = 0
    For a = 1 To UBound(MyIntCode)
        CodeLength(a) = GetLength(MyIntCode(a - 1)) + CodeLength(a - 1)
    Next a

    For a = 0 To UBound(MyIntCode)
        Select Case Left(MyIntCode(a), 1)
        Case "=", "-", "+", "*", "/", ">", "<", Chr(161), Chr(162), Chr(163), Chr(164), Chr(169), Chr(174)
            'the one operator no operands
            MyNewIntCode(a) = MyIntCode(a)
        Case Chr(165), Chr(166), Chr(167)
        'JF, JMP, Call
        'Search for the correct jump destination
            tmpPos = -1
            Label = Right(MyIntCode(a), Len(MyIntCode(a)) - 1)
            For b = 0 To UBound(MyIntCode)
                If JumpLabels(b) = Label Then
                    tmpPos = b
                End If
            Next b
            
            If tmpPos <> -1 Then
                MyNewIntCode(a) = Left(MyIntCode(a), 1) & GetHex(CodeLength(CInt(tmpPos)) - 1, 2)
            Else
                MsgBox "Error - Jump label not found"
            End If
        Case Chr(170), Chr(171), Chr(172) 'PUSH and PUSHV(Push Variable), and POP(pop from stack into var)
            MyNewIntCode(a) = Left(MyIntCode(a), 1)
            MyNewIntCode(a) = MyNewIntCode(a) & GetHex(Right(MyIntCode(a), Len(MyIntCode(a)) - 1), 2)
        End Select
    Next a
    Filter = Join(MyNewIntCode, "")
End Function

Private Function GetHex(Number As Integer, Characters As Integer) As String
'This returns a base 255 number - for example:
'   1 returns chr(1)
'  18 returns chr(18)
'255  returns chr(1) & chr(01)
'256  returns chr(1) & chr(02)
'258  returns chr(1) & chr(04)
'works by repetitly dividing by 255
    
    Do While Number > 255
        Result = Int(Number / 256)
        Remainder = Number - Result * 256
        GetHex = Chr(Remainder) & GetHex
        Number = Result
    Loop
    GetHex = Chr(Number) & GetHex
    
    Do While Len(GetHex) < Characters
        GetHex = Chr(0) & GetHex
    Loop
End Function

Private Function GetLength(Code As String)
    Select Case Left(Code, 1)
    Case "=", "-", "+", "*", "/", ">", "<", Chr(161), Chr(162), Chr(163), Chr(164), Chr(169)
        GetLength = 1
    Case Chr(174)
        GetLength = 2
    Case Chr(165), Chr(166), Chr(167) 'JF,JMP,CALL all use 3 bytes - 1 for op code, 2 for jump code
        GetLength = 3
    Case Chr(170), Chr(171), Chr(172) 'PUSH, PUSHV,POP all have 3 bytes - 1 for op code, 2 for the number/variable index number
        GetLength = 3
    End Select
End Function
