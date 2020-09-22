Attribute VB_Name = "modOld"
Function PrintTokens(Text As String)
    Dim der As New CTokens
    Set der = New CTokens
    Dim dagg As String
    
    'der.Initilise (Text)
    
    'Do While der.EndCode = False
    '    dagg = der.GetToken
    '    Debug.Print der.GetType(dagg), dagg
    'Loop
End Function

Function ParseLineOld(Text As String)
    'Dim Tokener As New CTokens
    'Set Tokener = New CTokens
    'Dim tmpString As String
    'Dim MyStack As New CStack
    
    MyStack.Initilize (20)
    Tokener.Initilise (Text)
    
    Do While Tokener.EndCode = False
        tmpString = Tokener.GetToken
        StringType = Tokener.GetType(tmpString)
        Select Case StringType
            Case "Statement "
                MyStack.Push tmpString
                tmpString = Tokener.GetToken
                If Tokener.GetType(tmpString) <> "(" Then
                    MsgBox "error on line"
                End If
                
            Case "Variable", "Number"
                Debug.Print "push " & tmpString
                If MyStack.StackPointer >= 0 Then
                    tmp = MyStack.Pop
                    If tmp = "=" Then
                        MyStack.Push tmp
                    Else
                        Debug.Print tmp
                    End If
                End If
                
            Case "Inequality", "Logical Op", "Maths Op", "Assign"
                MyStack.Push tmpString
                
            Case "Grouping"
                
            Case "End Line"
                If MyStack.StackPointer >= 0 Then
                    tmp = MyStack.Pop
                    If tmp = "=" Then
                        Debug.Print "Assign"
                    Else
                        MyStack.Push tmp
                    End If
                Else
                    Debug.Print "Clear"
                End If
        End Select
            
    Loop
End Function

Function GetTokena(Text As String, ByRef Token, ByRef TypeOfToken)
    lstToken = "while|if|for"
    lstToken = Split(Token, "|")
    For a = 1 To Len(Text)
        For b = 1 To UBound(lstToken)
            If Mid(Text, a, Len(lstToken(b))) = lstToken(b) Then
                Debug.Print Mid(Text, a, Len(lstToken(b)))
                a = a + Len(lstToken(b)) + 1
            End If
        Next b
        If Mid(Text, a, 1) = " " Then
            Debug.Print "variable: " & tmpVar
            tmpVar = ""
        Else
            tmpVar = tmpVar & Mid(Text, a, 1)
        End If
    Next a
End Function

Private Sub OldAttempt()
    MyStack.Initilize (20)
    
    'Dim RootNode As New CNode
    'Set RootNode = New CNode
    'RootNode.Value = "list"
    
    'Dim CurrentNode As New CNode
    'Set CurrentNode = RootNode
    
    'Set MyNode = New CNode
    'For a = 1 To Len(Text1.Text)
    '    tempTxt = tempTxt & Mid(Text1.Text, a, 1)
    '    Select Case LCase(tempTxt)
    '        Case "if"
    '            Set MyNode = New CNode
    '            MyNode.Value = "if"
    '            Set CurrentNode.MyNodes(1) = MyNode
    '            Set CurrentNode = MyNode
    '            tempTxt = ""
    '
    '        Case "print"
    '            Set MyNode = New CNode
    '            MyNode.Value = "print"
    '            MyStack.Push "print"
    '        Case "'"
    '            If TextMode Then
    '                b = MyStack.Pop
    '                Set MyNode = New CNode
    '                MyNode.Value = b
    '                Set tmpNode(1) = New CNode
    '                tmpNode(1).Value = Mytext
    '            End If
    '            TextMode = True
    '        Case Else
    '            If Mid(Text1, a, 1) = " " Then
    '                If Not TextMode Then
    '                    tempTxt = ""
    '                End If
    '            End If
    '            If TextMode Then
    '                Mytext = Mytext & Mid(Text1, a, 1) = " "
    '            End If
    '    End Select
    'Next a
    'RootNode.DumpNode
End Sub

Function GetHex(Number As Integer, Characters As Integer) As String
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

Function GetNumber(HexCode As String) As Integer
    For a = 1 To Len(HexCode)
        GetNumber = GetNumber * 255
        GetNumber = GetNumber + Asc(Mid(HexCode, a, 1))
    Next a
End Function
