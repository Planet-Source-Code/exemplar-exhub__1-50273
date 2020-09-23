VERSION 5.00
Begin VB.Form frmPM 
   Caption         =   "PM"
   ClientHeight    =   4140
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7575
   ScaleHeight     =   4140
   ScaleWidth      =   7575
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtSend 
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   3840
      Width           =   7575
   End
   Begin VB.TextBox txtRecv 
      Height          =   3855
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   7575
   End
End
Attribute VB_Name = "frmPM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    txtSend.SetFocus
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    txtRecv.Width = Me.Width - 120
    txtRecv.Height = Me.Height - txtSend.Height * 2 - 120
    txtSend.Top = txtRecv.Top + txtRecv.Height
    txtSend.Width = txtRecv.Width
End Sub

Private Sub txtRecv_Change()
    txtRecv.SelStart = Len(txtRecv.Text)
End Sub

Private Sub txtSend_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Trim(txtSend.Text) <> "" Then
        txtRecv.Text = txtRecv.Text & "<Hub> " & txtSend.Text & vbCrLf
        frmMain.SendPM Me.Tag, txtSend.Text
        txtSend.Text = ""
    End If
End Sub
