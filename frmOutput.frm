VERSION 5.00
Begin VB.Form frmOutput 
   Caption         =   "Output of Run"
   ClientHeight    =   6180
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9270
   LinkTopic       =   "Form2"
   ScaleHeight     =   6180
   ScaleWidth      =   9270
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtOutput 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   5295
   End
End
Attribute VB_Name = "frmOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub Output(Text As String)
    txtOutput.Text = txtOutput.Text & vbCrLf & Text
End Sub

Private Sub Form_Resize()
    If Me.WindowState <> vbMinimized Then
        txtOutput.Width = Me.Width - 120
        txtOutput.Height = Me.Height - 405
    End If
End Sub
