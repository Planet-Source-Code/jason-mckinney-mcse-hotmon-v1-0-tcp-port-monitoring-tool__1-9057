VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmHotMon 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HotMon v1.0 by Jason McKinney"
   ClientHeight    =   4530
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   4650
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   4650
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock sckNotify 
      Left            =   240
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete Entry"
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin MSFlexGridLib.MSFlexGrid grdHotGrid 
      Height          =   2775
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   4895
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
   End
   Begin VB.ComboBox cbxPortSel 
      Height          =   315
      Left            =   2160
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   120
      Width           =   1455
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   4275
      Width           =   4650
      _ExtentX        =   8202
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3175
            MinWidth        =   3175
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3175
            MinWidth        =   3175
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Left            =   840
      Top             =   480
   End
   Begin MSWinsockLib.Winsock sckHotMon 
      Left            =   1560
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdMonitor 
      Caption         =   "&Monitor"
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Top             =   480
      Width           =   855
   End
   Begin VB.TextBox txtIpAddress 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNotification 
         Caption         =   "&Notification"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmHotMon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RowClicked As Boolean
Dim OkToTransmit As Boolean
Dim NewLine As String
Dim AbruptStop As Boolean

Const TIpAddress = 0
Const TPortNum = 1
Const TLastState = 2

'Modify Form_Load and cmdAdd_Click to add or change ports
Private Enum Ports
    WWW = 80
    FTP = 21
    TELN = 23
    SMTP = 25
    DNS = 53
    POP3 = 110
    IMAP = 443
End Enum


'Adds the host and port to the flexgrid
Private Sub cmdAdd_Click()
Dim RowAdded As Long
Dim Port As Integer

    grdHotGrid.Row = 1
    If grdHotGrid.Text = "" Then
        grdHotGrid.Row = 1
        grdHotGrid.Col = 0
        grdHotGrid.Text = txtIpAddress.Text
        grdHotGrid.Col = 1
        Select Case cbxPortSel.ListIndex
            Case 0: Port = WWW
            Case 1: Port = FTP
            Case 2: Port = TELN
            Case 3: Port = SMTP
            Case 4: Port = DNS
            Case 5: Port = POP3
            Case 6: Port = IMAP
        End Select
        grdHotGrid.Text = Port
    Else
        grdHotGrid.AddItem txtIpAddress.Text
        RowAdded = grdHotGrid.Rows - 1
        grdHotGrid.Row = RowAdded
        grdHotGrid.Col = 1
        Select Case cbxPortSel.ListIndex
            Case 0: Port = WWW
            Case 1: Port = FTP
            Case 2: Port = TELN
            Case 3: Port = SMTP
            Case 4: Port = DNS
            Case 5: Port = POP3
            Case 6: Port = IMAP
        End Select
        grdHotGrid.Text = Port
    End If
    txtIpAddress.Text = ""
    SaveConfig

End Sub

'Removes a selected line from the flexgrid
'If it's the last line, just clears teh fields (a flexgrid requirement)
Private Sub cmdDelete_Click()
    If RowClicked = True Then
        If grdHotGrid.Rows = 2 Then
            grdHotGrid.Col = TIpAddress
            grdHotGrid.Text = ""
            grdHotGrid.Col = TPortNum
            grdHotGrid.Text = ""
            grdHotGrid.Col = TLastState
            grdHotGrid.Text = ""
        Else
            grdHotGrid.RemoveItem (grdHotGrid.Row)
        End If
    Else
        MsgBox ("You must select a host to remove from the table!")
    End If
    RowClicked = False
    grdHotGrid.Row = 0
    SaveConfig
End Sub

'Starts/Stops the monitoring event
Private Sub cmdMonitor_Click()
    If cmdMonitor.Caption = "&Monitor" Then
        AbruptStop = False
        StatusBar1.Panels.Item(1).Text = "Monitoring..."
        StatusBar1.Panels.Item(2).Text = ""
        Timer1.Interval = 60000
        Timer1.Enabled = True
        cmdMonitor.Caption = "&Stop"
        LogEvent "HotMon Started"
        Monitor
    Else
        Timer1.Enabled = False
        cmdMonitor.Caption = "&Monitor"
        StatusBar1.Panels.Item(1).Text = "HotMon v1.0"
        StatusBar1.Panels.Item(2).Text = ""
        If grdHotGrid.Enabled = False Then grdHotGrid.Enabled = True
        AbruptStop = True
        LogEvent "HotMon Stopped"
    End If
End Sub

'Brings up the Notification Configuration box
Private Sub cmdNotifyConfig_Click()
    frmNotifyConfig.Show vbModal
End Sub

Private Sub Form_Load()
    LogEvent "HotMon is loading..."
    NotifyEnabled = False
    NewLine = Chr$(13) + Chr$(10)
    StatusBar1.Panels.Item(1).Text = "HotMon v1.0"
    
    cbxPortSel.AddItem "WWW   (80)", 0
    cbxPortSel.AddItem "FTP   (21)", 1
    cbxPortSel.AddItem "TELN  (23)", 2
    cbxPortSel.AddItem "SMTP  (25)", 3
    cbxPortSel.AddItem "DNS   (53)", 4
    cbxPortSel.AddItem "POP3 (110)", 5
    cbxPortSel.AddItem "IMAP (443)", 6
    cbxPortSel.Text = cbxPortSel.List(0)
    cbxPortSel.ListIndex = 0
    
    grdHotGrid.ColAlignment(TIpAddress) = flexAlignLeftCenter
    grdHotGrid.ColAlignment(TLastState) = flexAlignCenterCenter
    grdHotGrid.ColWidth(0) = 1875
    grdHotGrid.ColWidth(1) = 600
    grdHotGrid.ColWidth(2) = 1875
    grdHotGrid.Row = 0
    grdHotGrid.Col = 0
    grdHotGrid.Text = "IP Address"
    grdHotGrid.Col = 1
    grdHotGrid.Text = "Port"
    grdHotGrid.Col = 2
    grdHotGrid.Text = "Last State"
    grdHotGrid.SelectionMode = flexSelectionByRow
    
    LoadConfig
    LogEvent "HotMon load complete!"
End Sub


'This is the Winsock monitoring code
Private Sub Monitor()
Dim Port As Ports
Dim RowPos As Integer
    grdHotGrid.Enabled = False
    
    For RowPos = 1 To grdHotGrid.Rows - 1
        If AbruptStop = True Then Exit Sub
        If sckHotMon.State <> 0 Then sckHotMon.Close
        
        grdHotGrid.Col = TIpAddress
        grdHotGrid.Row = RowPos
        sckHotMon.RemoteHost = grdHotGrid.Text
        StatusBar1.Panels.Item(1).Text = grdHotGrid.Text
        grdHotGrid.Col = TPortNum
        sckHotMon.RemotePort = Val(grdHotGrid.Text)
        sckHotMon.Connect
        grdHotGrid.Col = TIpAddress
        
        While sckHotMon.State = 4
            StatusBar1.Panels.Item(2).Text = "Resolving..."
            DoEvents
            If sckHotMon.State = 5 Then StatusBar1.Panels.Item(2).Text = "Resolved"
        Wend
        
        While sckHotMon.State = 6
            StatusBar1.Panels.Item(2).Text = "Connecting..."
            DoEvents
        Wend
        
        Select Case sckHotMon.State
            Case 0:
                StatusBar1.Panels.Item(2).Text = "Closed"
            Case 1:
                grdHotGrid.Col = TLastState
                grdHotGrid.Text = "ALIVE! (1)"
                grdHotGrid.Col = TIpAddress
                StatusBar1.Panels.Item(2).Text = "ALIVE! (1)"
            Case 2:
                StatusBar1.Panels.Item(2).Text = "Listening"
            Case 3:
                StatusBar1.Panels.Item(2).Text = "Pending"
            Case 4:
                StatusBar1.Panels.Item(2).Text = "Resolving"
            Case 5:
                StatusBar1.Panels.Item(2).Text = "Resolved"
            Case 6:
                StatusBar1.Panels.Item(2).Text = "Connecting"
                While sckHotMon.State = 6
                    DoEvents
                Wend
            Case 7:
                grdHotGrid.Col = TLastState
                grdHotGrid.Text = "ALIVE! (7)"
                grdHotGrid.Col = TIpAddress
                StatusBar1.Panels.Item(2).Text = "Connected"
                StatusBar1.Panels.Item(2).Text = "ALIVE! (7)"
            Case 8:
                StatusBar1.Panels.Item(2).Text = "Closing"
            Case 9:
                grdHotGrid.Col = TLastState
                grdHotGrid.Text = "DOWN! (9)"
                grdHotGrid.Col = TIpAddress
                StatusBar1.Panels.Item(2).Text = "DOWN! (9)"
                LogEvent grdHotGrid.Text & " is DOWN! (9)"
                If NotifyEnabled = True Then Notify (grdHotGrid.Text)
        End Select
        sckHotMon.Close
    Next
    grdHotGrid.Enabled = True
End Sub

'On Terminate, make sure winsock connections are closed
Private Sub Form_Terminate()
    LogEvent "HotMon has Exited"
    If sckHotMon.State <> 0 Then sckHotMon.Close
    End
End Sub

'On Unload, make sure winsock connections are closed
Private Sub Form_Unload(Cancel As Integer)
    LogEvent "HotMon has Exited"
    If sckHotMon.State <> 0 Then sckHotMon.Close
    End
End Sub

'Keep track of whether or not a row was clicked for delete procedure
Private Sub grdHotGrid_Click()
    RowClicked = True
End Sub

'Exit the program
Private Sub mnuExit_Click()
    LogEvent "HotMon has Exited"
    Unload Me
    End
End Sub

'Show notification dialog
Private Sub mnuNotification_Click()
    frmNotifyConfig.Show
End Sub

'Timer event spawn of monitoring
Private Sub Timer1_Timer()
    Monitor
End Sub

'Makes sure responses are recieved from SMTP server to regulate transmition of SMTP notifications
Private Sub sckNotify_DataArrival(ByVal bytesTotal As Long)
Dim Data As String
    
    sckNotify.GetData Data, vbString
    OkToTransmit = True
End Sub

'Loads the monitored host file and smtp config
Private Sub LoadConfig()
Dim IpAddress As String
Dim PortNum As String
Dim RowAdded As Integer

On Error GoTo ErrorH
    Open "HotMon.ini" For Input As #1
    While Not EOF(1)
        Input #1, IpAddress, PortNum
        grdHotGrid.Row = 1
        If grdHotGrid.Text = "" Then
            grdHotGrid.Col = TIpAddress
            grdHotGrid.Text = IpAddress
            grdHotGrid.Col = 1
            grdHotGrid.Text = PortNum
        Else
            grdHotGrid.AddItem IpAddress
            RowAdded = grdHotGrid.Rows - 1
            grdHotGrid.Row = RowAdded
            grdHotGrid.Col = TPortNum
            grdHotGrid.Text = PortNum
        End If
    Wend
    Close #1
    LogEvent "HotMon Configuraton Loaded"
    LoadSMTPConfig
    Exit Sub
ErrorH:
LogEvent "HotMon Configuration Load Error"
LoadSMTPConfig
End Sub

'Saves the flexgrid of hosts to a file
Private Sub SaveConfig()
Dim TotalRows As Integer
Dim CurrentRow As Integer
Dim IpAddress As String
Dim PortNum As String

On Error GoTo ErrorH
    TotalRows = grdHotGrid.Rows - 1
    Open "HotMon.ini" For Output As #1
    For CurrentRow = 1 To TotalRows
        grdHotGrid.Row = CurrentRow
        grdHotGrid.Col = TIpAddress
        IpAddress = grdHotGrid.Text
        grdHotGrid.Col = TPortNum
        PortNum = grdHotGrid.Text
        Write #1, IpAddress; PortNum
    Next
    Close #1
    Exit Sub

ErrorH:
    LogEvent "HotMon Configuration Save Error"
    MsgBox ("Config File Error, please contact technical support.")
End Sub

'SMTP notification code
Public Sub Notify(DownSystem As String)
Dim Data As String
    sckNotify.RemoteHost = SMTPServer
    sckNotify.RemotePort = 25
    sckNotify.Protocol = sckTCPProtocol
    sckNotify.Connect
    While sckNotify.State <> 7
        DoEvents
    Wend

    OkToTransmit = False
    While OkToTransmit = False
        DoEvents
    Wend
    OkToTransmit = False
    sckNotify.SendData "Helo" & NewLine
    While OkToTransmit = False
        DoEvents
    Wend
    OkToTransmit = False
    sckNotify.SendData "Mail From: " & Sender & NewLine
    While OkToTransmit = False
        DoEvents
    Wend
    OkToTransmit = False
    sckNotify.SendData "Rcpt To: " & Recepient & NewLine
    While OkToTransmit = False
        DoEvents
    Wend
    OkToTransmit = False
    sckNotify.SendData "Data" & NewLine
    While OkToTransmit = False
        DoEvents
    Wend
    OkToTransmit = False
    sckNotify.SendData "Subject: HotMon Notification" & NewLine & DownSystem & " has not responded to HotMon as of " & NewLine & Date & " " & Time & "   and is possibly DOWN!" & NewLine & "." & NewLine
    While OkToTransmit = False
        DoEvents
    Wend
    OkToTransmit = False
    LogEvent "SMTP Notification Message Send for host " & DownSystem
    sckNotify.Close
End Sub

'Load SMTP Config
Private Sub LoadSMTPConfig()
On Error GoTo ErrorH
    
    Open "HotSMTP.ini" For Input As #1
    Input #1, SMTPServer, SystemName, Sender, Recepient
    Close #1
    StatusBar1.Panels.Item(3).Text = "Notify ON"
    NotifyEnabled = True
    mnuNotification.Checked = True
    LogEvent "SMTP Configuration Loaded"
    Exit Sub
    
ErrorH:
LogEvent "SMTP Configuration Load Error"
End Sub

'Save SMTP config
Public Sub SaveSMTPConfig()
On Error GoTo ErrorH
    
    Open "HotSMTP.ini" For Output As #1
    Write #1, SMTPServer, SystemName, Sender, Recepient
    Close #1
    StatusBar1.Panels.Item(3).Text = "Notify ON"
    mnuNotification.Checked = True
    LogEvent "SMTP Configuration Saved"
    Exit Sub
    
ErrorH:
LogEvent "SMTP Configuration Save Error"
End Sub

'Log event to log
Public Sub LogEvent(WhatHappened As String)
Dim WhatToPrint As String
On Error GoTo ErrorH
    
    Open "HotMon.Log" For Append As #500
    WhatToPrint = Date & " " & Time & "   " & WhatHappened
    Print #500, WhatToPrint
    Close #500
    Exit Sub

ErrorH:
Open "HotMon.Log" For Input As #500
Resume Next
End Sub

