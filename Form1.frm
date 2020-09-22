VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   5820
   ClientLeft      =   3315
   ClientTop       =   3000
   ClientWidth     =   7500
   LinkTopic       =   "Form1"
   ScaleHeight     =   5820
   ScaleWidth      =   7500
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkLog 
      Caption         =   "Log Run"
      Height          =   255
      Left            =   6480
      TabIndex        =   4
      Top             =   5160
      Width           =   975
   End
   Begin VB.CommandButton cmdOutput 
      Caption         =   "Show Output"
      Height          =   615
      Left            =   6480
      TabIndex        =   3
      Top             =   4440
      Width           =   975
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Runnable Output"
      Height          =   375
      Left            =   6480
      TabIndex        =   2
      Top             =   5400
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   5775
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "Form1.frx":0000
      Top             =   0
      Width           =   6375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "'Compile' it"
      Height          =   4335
      Left            =   6480
      TabIndex        =   0
      Top             =   0
      Width           =   975
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MyStack As New CStack

Private Sub cmdOutput_Click()
    frmOutput.Show
End Sub

Private Sub Command1_Click()
    'Do While GetToken(Text1.Text, MyToken, MyType) <> -1
    a = Split(Text1.Text, vbCrLf)
    a = Join(a, " ")
    a = Split(a, vbTab)
    a = Join(a, "")
    Call ParseLine(CStr(a), Check1.Value, chkLog.Value)
    frmOutput.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub
