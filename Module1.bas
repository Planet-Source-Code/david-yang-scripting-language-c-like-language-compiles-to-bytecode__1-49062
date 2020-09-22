Attribute VB_Name = "modmain"
'If condition
'Usage call ParseLine("if (a> 4 &&b <5) a=a+1;")
'Usage call parseline("if (a > 4 && b < 5) {a=a+1} else {a=a-1};")

'For loops
'Usage Call parseline("for (a=1;a<10;a=a+1){a=a-1}")

'While Loops
'Usage call parseline("while (a>2) {a=a+1;if (c>d&&e>f){e=f+d;}}")

'Modules
'Usage call parseline("int main(){hitme(23);}; int hitme(int a) {a=a+1;};")

'Steps of adding a command
'   1) Change CTokens
'       a) Change the lstAll definition in Class_Initilize
'       b) Give a precedence to the command
'       c) Change CCodeTree to comply with new command
'       d) Change GenIntCode in the main Select Case
'       e) Change CFilter as neccessary

Function ParseLine(Code As String, Optional Digital As Boolean = False, Optional Log As Boolean = True)
    Dim MyTree As CTree
    Dim MyCodeTree As CCodeTree
    Dim MyGenIntCode As CGenIntCode
    Dim MyRun As CRun
    Dim MyIntCode As String
    Dim MyFilter As CFilter
    Dim MyVars As CVars
    
    Set MyCodeTree = New CCodeTree
    Set MyGenIntCode = New CGenIntCode
    Set MyRun = New CRun
    Set MyFilter = New CFilter
    Set MyVars = New CVars
    
    frmOutput.txtOutput = "Code..." & vbCrLf
    frmOutput.txtOutput = frmOutput.txtOutput & frmMain.Text1.Text & vbCrLf
    frmOutput.txtOutput = frmOutput.txtOutput & "Generating Code tree..." & vbCrLf
    'Generate tree
    Set MyCodeTree.Variables = MyVars
    Set MyTree = MyCodeTree.CreateTree(Code)
    'Dump tree
    Call MyTree.DumpTree(0)
    frmOutput.txtOutput = frmOutput.txtOutput & "total number of nodes: " & "  " & MyTree.TotalNodes & vbCrLf
    
    'Generate Code
    frmOutput.txtOutput = frmOutput.txtOutput & "Generating Intermediate Code..." & vbCrLf
    Let MyGenIntCode.OutputMachine = Digital
    Set MyGenIntCode.MyVars = MyVars
    Let MyIntCode = MyGenIntCode.GenerateCode(MyTree)
    
    If Not Digital Then
        'Print Code
        frmOutput.txtOutput = frmOutput.txtOutput & MyIntCode & vbCrLf
    Else
        'Filter for jump locations etc
        Let MyIntCode = MyFilter.Filter(MyIntCode)

        'Run code
        frmOutput.txtOutput = frmOutput.txtOutput & "Running Intermediate Code..." & vbCrLf
        If Log Then
            Call MyRun.RunCode(MyIntCode)
        Else
            Call MyRun.RunCodeNoLog(MyIntCode)
        End If
    End If
End Function

