VERSION 5.00
Begin VB.Form frmNotifyConfig 
   Caption         =   "HotMon Notify Configuration"
   ClientHeight    =   2580
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4470
   LinkTopic       =   "Form1"
   ScaleHeight     =   2580
   ScaleWidth      =   4470
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdTest 
      Caption         =   "&Test"
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   2160
      Width           =   855
   End
   Begin VB.CommandButton cmdDisable 
      Caption         =   "&Disable Notification"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Connection"
      Height          =   1335
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   4455
      Begin VB.TextBox txtSMTP 
         Height          =   285
         Left            =   1200
         TabIndex        =   0
         Top             =   240
         Width           =   3135
      End
      Begin VB.TextBox txtSystemName 
         Height          =   285
         Left            =   1200
         TabIndex        =   1
         Top             =   600
         Width           =   3135
      End
      Begin VB.TextBox txtSender 
         Height          =   285
         Left            =   1200
         TabIndex        =   2
         Top             =   960
         Width           =   3135
      End
      Begin VB.Label Label1 
         Caption         =   "SMTP-server:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   300
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Systemname:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   645
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Sender e-mail:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1005
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Recepient"
      Height          =   615
      Left            =   0
      TabIndex        =   7
      Top             =   1440
      Width           =   4455
      Begin VB.TextBox txtRecepient 
         Height          =   285
         Left            =   1200
         TabIndex        =   3
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label Label4 
         Caption         =   "Recepient:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   300
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      Top             =   2160
      Width           =   855
   End
End
Attribute VB_Name = "frmNotifyConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDisable_Click()
    NotifyEnabled = False
    frmHotMon.mnuNotification.Checked = False
    frmHotMon.StatusBar1.Panels.Item(3).Text = ""
    Unload Me
End Sub

Private Sub cmdOk_Click()
  Recepient = txtRecepient.Text
  Sender = txtSender.Text
  SystemName = txtSystemName.Text
  SMTPServer = txtSMTP.Text
  If Recepient = "" Or Sender = "" Or SystemName = "" Or SMTPServer = "" Then GoTo ErrorH
  NotifyEnabled = True
  frmHotMon.SaveSMTPConfig
  Unload Me
  Exit Sub
ErrorH:
MsgBox ("Required fields missing!")
End Sub


Private Sub cmdTest_Click()
    Recepient = txtRecepient.Text
    Sender = txtSender.Text
    SystemName = txtSystemName.Text
    SMTPServer = txtSMTP.Text
    If Recepient = "" Or Sender = "" Or SystemName = "" Or SMTPServer = "" Then GoTo ErrorH
    NotifyEnabled = True
    frmHotMon.SaveSMTPConfig
    frmHotMon.Notify ("TEST-SYSTEM")
    MsgBox ("Test Message Sent!")
    Exit Sub
ErrorH:
MsgBox ("Required fields missing!")
End Sub

Private Sub Form_Load()
  txtRecepient.Text = Recepient
  txtSender.Text = Sender
  txtSystemName.Text = SystemName
  txtSMTP.Text = SMTPServer
End Sub

Private Sub txtRecepient_gotfocus()
  txtRecepient.SelStart = 0
  txtRecepient.SelLength = Len(txtRecepient.Text)
End Sub

Private Sub txtSender_gotfocus()
  txtSender.SelStart = 0
  txtSender.SelLength = Len(txtSender.Text)
End Sub

Private Sub txtSMTP_GotFocus()
  txtSMTP.SelStart = 0
  txtSMTP.SelLength = Len(txtSMTP.Text)
End Sub

Private Sub txtSystemName_gotfocus()
  txtSystemName.SelStart = 0
  txtSystemName.SelLength = Len(txtSystemName.Text)
End Sub

