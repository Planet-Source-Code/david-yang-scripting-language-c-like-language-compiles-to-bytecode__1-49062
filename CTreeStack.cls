VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTreeStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Dynamic Stack

Dim MyStack() As CTree
Public StackPointer As Integer

Public Sub Initilize(Size As Integer)
    ReDim MyStack(Size)
    StackPointer = -1
End Sub

Public Function Push(Data As CTree) As Integer
    StackPointer = StackPointer + 1
    If StackPointer > UBound(MyStack) Then
        ReDim Preserve MyStack(UBound(MyStack) + 10)
    End If
    Set MyStack(StackPointer) = Data
    Push = UBound(MyStack) - StackPointer - 1
End Function

Public Function Pop() As CTree
    Set Pop = MyStack(StackPointer)
    StackPointer = StackPointer - 1
End Function

Public Sub DumpStack()
    Debug.Print "Stack Contents ---"
    For a = 0 To StackPointer
        Debug.Print "               " & MyStack(a).Value
    Next a
    Debug.Print "               ---"
End Sub

Private Sub Class_Terminate()
    Erase MyStack
End Sub


