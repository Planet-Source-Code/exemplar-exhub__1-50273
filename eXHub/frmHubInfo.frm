VERSION 5.00
Begin VB.Form frmHubInfo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Dialog Caption"
   ClientHeight    =   1575
   ClientLeft      =   2760
   ClientTop       =   3705
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbSpeed 
      Height          =   315
      ItemData        =   "frmHubInfo.frx":0000
      Left            =   1080
      List            =   "frmHubInfo.frx":0016
      TabIndex        =   10
      Text            =   "Choose Speed"
      Top             =   480
      Width           =   3495
   End
   Begin VB.CheckBox chkSendWhenDone 
      Caption         =   "Send When Done"
      Height          =   375
      Left            =   4680
      TabIndex        =   9
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox txtShareSize 
      Height          =   285
      Left            =   1080
      TabIndex        =   8
      Top             =   1200
      Width           =   3495
   End
   Begin VB.TextBox txtEmail 
      Height          =   285
      Left            =   1080
      TabIndex        =   7
      Top             =   840
      Width           =   3495
   End
   Begin VB.TextBox txtInterest 
      Height          =   285
      Left            =   1080
      TabIndex        =   6
      Top             =   120
      Width           =   3495
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblEmail 
      Caption         =   "Email:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label lblSpeed 
      Caption         =   "Speed:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label lblInterest 
      Caption         =   "Interest:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblShareSize 
      Caption         =   "Share Size:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   1575
   End
End
Attribute VB_Name = "frmHubInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    txtInterest.Text = HubInfo.Interest
    cmbSpeed.Text = HubInfo.Speed
    txtEmail.Text = HubInfo.Email
    txtShareSize.Text = HubInfo.ShareSize
End Sub

Private Sub OKButton_Click()
    HubInfo.Interest = txtInterest.Text
    HubInfo.Speed = cmbSpeed.Text
    HubInfo.Email = txtEmail.Text
    HubInfo.ShareSize = txtShareSize.Text
    
    '$MyINFO $ALL <nick> <interest>$ $<speed>$<e-mail>$<sharesize>$
    If chkSendWhenDone.Value = vbChecked Then frmMain.SendToAll "$MyINFO $ALL Hub " & HubInfo.Interest & "$ $" & HubInfo.Speed & "$" & HubInfo.Email & "$" & HubInfo.ShareSize & "$"
    
    Unload Me
End Sub
