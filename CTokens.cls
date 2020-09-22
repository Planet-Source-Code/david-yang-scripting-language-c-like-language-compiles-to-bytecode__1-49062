VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTokens"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Tokeniser - Given a string, it gives tokens
' The only accessable function are
'   GetLowestPrecedence - Give location and token with lowest precedence in string
'   ContainArguments - Give wether there is a token in the string

Private Location As Integer
Private Mytext As String
Private lstAll
Private lstLegal
Public MyVars As New CVars

Private Function GetToken() As String
    Dim tmpVar As String
    ReDim Vars(0)
    
    'Skip empty spaces
    Do While Mid(Mytext, Location, 1) = " "
        Location = Location + 1
    Loop
    
    'Scan for the statements
    For b = 0 To UBound(lstAll)
        If Mid(Mytext, Location, Len(lstAll(b))) = lstAll(b) Then
            GetToken = lstAll(b)
            Location = Location + Len(lstAll(b))
            Exit Function
        End If
    Next b
    
    'Scan for variables
    For a = Location To Len(Mytext)
        For b = 0 To UBound(lstAll)
            If Mid(Mytext, a, Len(lstAll(b))) = lstAll(b) Then
                GetToken = Mid(Mytext, Location, a - Location)
                Location = a
                Exit Function
            End If
        Next b
    Next a
    
    'This runs only if the token is the last token
    GetToken = Mid(Mytext, Location, Len(Mytext) - Location + 1)
    Location = Len(Mytext) + 1
End Function

Private Function EndCode() As Boolean
    If Location > Len(Mytext) Then
        EndCode = True
    Else
        EndCode = False
    End If
End Function

Function GetLowestPrecedence(Text As String, ByRef LowestToken) As Integer
'Returns location of lowest Precedence token
'GetLowestPrecedence - Location of lowest token
'LowestPrecedence - The value of the lowest token so far in the loop
    Dim LowestPrecedence As Integer
    Dim InBrackets As Integer
    Dim tmpPrecedence As Integer
    Dim tmpToken As String
    
    Mytext = Text
    Location = 1
    
    'Initilize values
    LowestToken = GetToken
    LowestPrecedence = GetPrecedence(CStr(LowestToken))
    GetLowestPrecedence = 1
    Location = 1
    InBrackets = 0
    
    'If Left(Text, 1) = "(" Or Left(Text, 1) = "{" Then
    '    Exit Function
    'End If
    
    Do While Not (EndCode)
        tmpToken = GetToken
        tmpPrecedence = GetPrecedence(CStr(tmpToken))
        
        'Increase the precedence of tokens that are in a bracket
        Select Case tmpToken
        Case "(", "{"
            InBrackets = InBrackets + tmpPrecedence
            tmpPrecedence = InBrackets
        Case ")", "}"
            tmp = InBrackets
            InBrackets = InBrackets - tmpPrecedence
            tmpPrecedence = tmp
        Case Else
            tmpPrecedence = tmpPrecedence + InBrackets
        End Select
        
        If tmpPrecedence = 0 Then
            GetLowestPrecedence = Location - Len(tmpToken)
            LowestToken = tmpToken
            Exit Function
        ElseIf LowestPrecedence > tmpPrecedence Then
            GetLowestPrecedence = Location - Len(tmpToken)
            LowestToken = tmpToken
            LowestPrecedence = tmpPrecedence
        End If
    Loop
End Function

Function GetPrecedence(Token As String) As Integer
'Precedences go up in 10s so I can insert more if required
    GetPrecedence = 100
    Select Case Token
    Case ";"
        GetPrecedence = 1
    Case ","
        GetPrecedence = 2
    Case "if", "while", "for", "void", "int", "return"
        GetPrecedence = 3
    Case "else", "break"
        GetPrecedence = 4
    Case "="
        GetPrecedence = 10
    Case "||"
        GetPrecedence = 20
    Case "&&"
        GetPrecedence = 30
    Case "==", "!="
        GetPrecedence = 40
    Case "<", "<=", ">", ">="
        GetPrecedence = 50
    Case "+", "-"
        GetPrecedence = 60
    Case "*", "/"
        GetPrecedence = 70
    Case "(", ")"
        GetPrecedence = 80
    Case "{", "}"
        GetPrecedence = 90
    End Select
End Function

Private Sub Class_Initialize()
    lstAll = "while`if`for`else`break`int`void`return`<`>`+`-`*`/`=`(`)`{`}`;`==`&&`||`,"
    lstAll = Split(lstAll, "`")
    
    Set MyVars = New CVars
End Sub

Public Function ContainArguments(Text As String) As Boolean
'Returns wether there are any things in the text
' eg "while", "+", "break", ...
    ContainArguments = False
    For a = 1 To Len(Text)
        For b = 0 To UBound(lstAll)
            If lstAll(b) = Mid(Text, a, Len(lstAll(b))) Then
                ContainArguments = True
                Exit Function
            End If
        Next b
    Next a
End Function
