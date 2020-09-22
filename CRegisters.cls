VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRegisters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim CurrentRegister As Integer

Private Sub Class_Initialize()
    Dim Counter As Integer
    CurrentRegister = 0
End Sub

Public Function GetNewRegister() As Integer
    CurrentRegister = CurrentRegister + 1
    GetRegister = CurrentRegister
End Function

Public Function GetCurrentRegister() As Integer
    GetCurrentRegister = CurrentRegister
End Function

Public Sub RemoveRegisters()
    CurrentRegister = 0
End Sub

Public Function NameRegister() As String
    CurrentRegister = CurrentRegister + 1
    NameRegister = Chr(Asc("a") + CurrentRegister - 1) & "x"
End Function
