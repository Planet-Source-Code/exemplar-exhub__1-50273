VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "eXHub"
   ClientHeight    =   7155
   ClientLeft      =   7050
   ClientTop       =   2850
   ClientWidth     =   7095
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "DC"
   MaxButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   7095
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkMD5Passwords 
      Caption         =   "MD5 Passwords"
      Height          =   255
      Left            =   5520
      TabIndex        =   17
      Top             =   6120
      Width           =   1455
   End
   Begin VB.CommandButton cmdRedirectClient 
      Caption         =   "Redirect User"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   6480
      Width           =   1695
   End
   Begin VB.CommandButton cmdHubInfo 
      Caption         =   "Hub Info"
      Height          =   255
      Left            =   1920
      TabIndex        =   15
      Top             =   6480
      Width           =   1695
   End
   Begin VB.CommandButton cmdKick 
      Caption         =   "Kick User"
      Height          =   255
      Left            =   3720
      TabIndex        =   14
      Top             =   6120
      Width           =   1695
   End
   Begin VB.CommandButton cmdSendPM 
      Caption         =   "Send A PM"
      Height          =   255
      Left            =   3720
      TabIndex        =   12
      Top             =   5760
      Width           =   1695
   End
   Begin VB.CommandButton cmdRedirectAll 
      Caption         =   "Redirect All Clients"
      Height          =   255
      Left            =   1920
      TabIndex        =   11
      Top             =   5760
      Width           =   1695
   End
   Begin VB.CheckBox chkPopupPM 
      Caption         =   "Popup PM's"
      Height          =   255
      Left            =   5520
      TabIndex        =   10
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CheckBox chkDisableSharing 
      Caption         =   "Disable Sharing"
      Height          =   255
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6120
      Width           =   1695
   End
   Begin VB.CheckBox chkDisableTalking 
      Caption         =   "Disable Talking"
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6120
      Width           =   1695
   End
   Begin VB.CommandButton cmdFakeAllQuit 
      Caption         =   "Send Fake All Quit"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   5760
      Width           =   1695
   End
   Begin VB.CheckBox chkClearChat 
      Caption         =   "Clear On Send"
      Height          =   255
      Left            =   5520
      TabIndex        =   6
      Top             =   5400
      Width           =   1455
   End
   Begin VB.CommandButton cmdSendCommand 
      Caption         =   "Send As Command"
      Height          =   255
      Left            =   1920
      TabIndex        =   5
      Top             =   5400
      Width           =   1695
   End
   Begin VB.CommandButton cmdMassMsg 
      Caption         =   "Send As Mass PM"
      Height          =   255
      Left            =   3720
      TabIndex        =   4
      Top             =   5400
      Width           =   1695
   End
   Begin VB.CommandButton cmdSendChat 
      Caption         =   "Send As Chat"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   5400
      Width           =   1695
   End
   Begin VB.TextBox txtChat 
      Height          =   285
      Left            =   120
      MaxLength       =   500
      TabIndex        =   2
      Top             =   5040
      Width           =   5295
   End
   Begin VB.CheckBox chkLogCommands 
      Caption         =   "Log Commands"
      Height          =   255
      Left            =   5520
      TabIndex        =   1
      Top             =   5040
      Width           =   1455
   End
   Begin VB.Timer tmrRestart 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   6600
      Top             =   600
   End
   Begin VB.Timer tmrTimeOut 
      Interval        =   1000
      Left            =   6600
      Top             =   120
   End
   Begin VB.TextBox txtLog 
      Height          =   4815
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   120
      Width           =   6855
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   6840
      Width           =   6855
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuShow 
         Caption         =   "&Show"
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRehash 
         Caption         =   "&Rehash"
      End
      Begin VB.Menu mnuRestart 
         Caption         =   "R&estart"
      End
      Begin VB.Menu mnuClearLog 
         Caption         =   "&Clear Log"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

'SocketWise stuff
Private WithEvents Server       As clsServer
Attribute Server.VB_VarHelpID = -1
Private EndPoint                As clsEndPoint

'Socket number for the listening socket
Private lSock                   As Long

'Where logs of activity are kept
Private LogFile                 As String

'Configuration file, default app.path & "\hub.conf"
' But can not make app.path a constant :(
Private ConfFile                As String

'Hold whether the hub is active or not.
Private HubRunning              As Boolean

'Hub related settings from file
Private HubName                 As String   ' Pretty self explanatory i think,
Private HubPort                 As String   ' - This too.
Private MOTDFile                As String   ' The path to the file containing MOTD
Private UseMOTD                 As Boolean  ' Whether a MOTD exists
Private MOTDString              As String   ' Hub loads MOTD into string, to use Replace() on %hubname% and %nick%
Private MinShare                As String   ' Minimum share for clients
Private HubStartTime            As Date     ' The time the hub starts, used to calculate uptime
Private UseLog                  As Boolean  ' Whether to log to file or not
Private StartMin                As Integer  ' Start hub minimised to systray

'Collection of Clients
Private Clients As Collection

'Dummy counting variable used in for each ... next loops
Dim Dummy As clsUser

'Used for MD5'ing passwords
Private MD5 As New clsMD5

'Var for holding file numbers, when logging
Dim fNum As Integer

Private Sub chkDisableSharing_Click()
    'Disable/Enable sharing
    SendToAll "<Hub> Sharing has been " & IIf(chkDisableSharing.value = vbChecked, "disabled.", "enabled.")
End Sub

Private Sub chkDisableTalking_Click()
    'Disable/Enable speaking in main chat
    SendToAll "<Hub> Talking has been " & IIf(chkDisableTalking.value = vbChecked, "disabled.", "enabled.")
End Sub

Private Sub cmdFakeAllQuit_Click()
    'This sometimes makes it look like everyone has parted from hub, even though they might not have
    For Each Dummy In Clients
        If Dummy.LogedIn = True Then SendToAll "$Quit " & Dummy.Nick
    Next
End Sub

Private Sub cmdHubInfo_Click()
    frmHubInfo.Show
End Sub

Private Sub cmdKick_Click()
    'Kick an ip/user
    Dim strKick As String
    
    strKick = InputBox("Enter Nick/IP to kick:")
    If Trim(strKick) = "" Then Exit Sub
    
    For Each Dummy In Clients
        If Dummy.Nick = strKick Or Dummy.RemoteIP = strKick Then
            'Call SendToAll("<Hub> " & strKick & " is being kicked by hub.") ' this triggers a message in the next line for some reason
            Server_OnClose Dummy.Socket
            Exit For
        End If
    Next
End Sub

Private Sub cmdMassMsg_Click()
    'Sends a mass pm message to all clients
    '$To: <othernick> From: <nick> $<<nick>> <message>
    If Trim(txtChat.Text) = "" Then Exit Sub
    For Each Dummy In Clients
        If Dummy.LogedIn = True Then Send Dummy.Socket, "$To: " & Dummy.Nick & " From: Hub $<Hub> " & txtChat.Text
    Next
    AddLog "Hub > All As Mass PM - " & txtChat.Text
    If chkClearChat.value = vbChecked Then txtChat.Text = ""
End Sub

Private Sub cmdRedirectAll_Click()
    'Redirects all clients (choice for ops or not) to server because reason
    Dim NewIP As String
    Dim ReasonMSG As String
    Dim Ops As VbMsgBoxResult
    
    NewIP = InputBox("Enter server to redirect to:")
    If NewIP = "" Then Exit Sub
    
    ReasonMSG = InputBox("Enter reason message:")
    If ReasonMSG = "" Then Exit Sub
    
    Ops = MsgBox("Redirect Ops as well?", vbYesNo)
    
    SendToAll "<Hub> Mass " & IIf(Ops = vbNo, "non-operator", "client") & " redirect to " & NewIP & ", because: " & ReasonMSG
    
    For Each Dummy In Clients ' send them the message, give timeout to disconnect
        If Ops = vbYes Then
            Send Dummy.Socket, "$ForceMove " & NewIP, False
            Dummy.TimeOut = 5 ' give them five seconds for their client to do this
        ElseIf Ops = vbNo Then
            If Dummy.Operator = False Then
                Send Dummy.Socket, "$ForceMove " & NewIP, False
                Dummy.TimeOut = 5 ' give them five seconds for their client to do this
            End If
        End If
    Next
End Sub

Private Sub cmdRedirectClient_Click()
    Dim RedirNick As String, NewServer As String, Reason As String
    Dim NickExists As Boolean
    
    RedirNick = InputBox("Enter nick you want to redirect:")
    If Trim(RedirNick) = "" Then Exit Sub
    
    For Each Dummy In Clients ' make sure nick exists
        If Dummy.Nick = RedirNick Then
            NickExists = True
            Exit For
        End If
    Next
    
    If NickExists = False Then
        MsgBox "Nick does not exist!"
        Exit Sub
    End If
    
    NewServer = InputBox("Enter server to redirect to:")
    If NewServer = "" Then Exit Sub
    Reason = InputBox("Enter reason for redirection:")
    
    '$OpForceMove $Who:<victimNick>$Where:<newIp>$Msg:<reasonMsg>
    SendToNick RedirNick, "$OpForceMove $Who:" & RedirNick & "$Where:" & NewServer & "$Msg:" & Reason
End Sub

Private Sub cmdSendChat_Click()
    'Sends a chate message from hub
    If Trim(txtChat.Text) = "" Then Exit Sub
    SendToAll "<Hub> " & txtChat.Text, True
    If chkClearChat.value = vbChecked Then txtChat.Text = ""
End Sub

Private Sub cmdSendCommand_Click()
    'Sends the message to all clients as a command
    If Trim(txtChat.Text) = "" Then Exit Sub
    SendToAll txtChat.Text
    If chkClearChat.value = vbChecked Then txtChat.Text = ""
End Sub

Private Sub cmdSendPM_Click()
    'Sends a PM to a specified nick, using frmPM
    Dim ToNick As String
    
    ToNick = InputBox("Enter to nick:")
    If Trim(ToNick) = "" Then Exit Sub
    
    For Each Dummy In Clients
        If Dummy.Nick = ToNick Then
            Dim NewPM As New frmPM
            NewPM.Tag = Dummy.Nick
            NewPM.Caption = "PM - " & Dummy.Nick
            NewPM.Show
            Exit For
        End If
    Next
End Sub

Private Sub Form_Load()

    If Not IsIDE Then On Error GoTo errHandle
    
    lblMessage.Caption = App.Title & " v" & App.Major & "." & App.Minor & "." & App.Revision & " - " & App.FileDescription

    'Configuration file path, this should remain the same.
    ConfFile = App.Path & "\hub.conf"
    
    If Dir(ConfFile) <> "" Then ' Make conf file sure it exists
        'Load settings from the config file
        Call LoadSettings
    Else ' Use default settings
        AddLog "*** ERROR: Configuration file not found: " & ConfFile, False
        AddLog "*** Using default settings..."
        
        HubName = "eXHub"
        HubPort = 411
        UseMOTD = False
        MinShare = 0
    End If
    
    'Add stuff to logfile
    AddLog "-- Hubname : " & HubName
    AddLog "-- Hubport : " & HubPort
    AddLog "-- Minimum share : " & MinShare
    If UseMOTD = True Then AddLog "-- MOTD File : " & MOTDFile Else AddLog "-- MOTD not in use"
    
    'Program text
    Me.Caption = "eXHub - " & HubName
    
    'New random seed - why not.
    Randomize
    
    'Initialise SocketWise
    Set Server = New clsServer
    Set EndPoint = New clsEndPoint
    
    'Initialise new collection of clients
    Set Clients = New Collection
        
    'Add to system tray
    Call AddSystray(Me, Me.Caption)
    
    'Start the Hub
    Call StartHub
    
    Exit Sub
errHandle: ' Usually triggered if missing socket.dll, or vb runtime files
    MsgBox "FATAL ERROR #" & Err.Number & vbCrLf & Err.Description & vbCrLf & "Source:" & Err.Source
    End
End Sub

Sub LoadSettings()

    On Error Resume Next ' Hmmmm, probably shouldn't use this
    
    'Get settings from config file
    
    'LogFile stuff
    UseLog = CBool(Val(Trim(ReadINI("Settings", "UseLog", ConfFile))))
    If UseLog = True Then
        LogFile = App.Path & "\exhub." & Replace(Date, "/", "-") & ".log"
        
        txtLog.Text = txtLog.Text & "Logging to " & LogFile & vbCrLf
        
        'begin new section in log file
        fNum = FreeFile
        Open LogFile For Append As fNum
            Print #fNum, vbCrLf & "New session started at " & Now
        Close fNum
    Else
        txtLog.Text = txtLog.Text & "Not logging to file." & vbCrLf
    End If
    
    'HubName
    HubName = ReadINI("Settings", "HubName", ConfFile)
    If Trim(HubName) = "" Then
        AddLog "*** Warning: HubName not set in config file, using default (eXHub)"
        HubName = "eXHub"
    End If
    
    'HubPort
    HubPort = ReadINI("Settings", "HubPort", ConfFile)
    If Trim(HubPort) = "" Then
        AddLog "*** Warning: HubPort not set in config file, using default. (411)"
        HubPort = 411
    ElseIf Val(HubPort) < 1 Then
        AddLog "*** Warning: HubPort specified in config file is not valid, using default. (411)"
        HubPort = 411
    End If
    
    'MOTDFile
    MOTDFile = ReadINI("Settings", "MOTDFile", ConfFile)
    If Trim(MOTDFile) = "" Then
        AddLog "*** Warning: MOTDFile not set in config file, not using a MOTD."
        UseMOTD = False
    ElseIf Dir(MOTDFile) = "" Then
        AddLog "*** Warning: MOTDFile does not exist, not using a MOTD."
        UseMOTD = False
    ElseIf FileLen(MOTDFile) = 0 Then
        AddLog "*** Warning: MOTDFile contains nothing, not using a MOTD."
        UseMOTD = False
    Else
        UseMOTD = True
    End If
    
    'Check MOTDFile exists
    If UseMOTD = True Then 'If there is a motd, read it into MOTDString
        Dim tmpMOTD As String
        fNum = FreeFile
        
        Open MOTDFile For Input As fNum ' Load motd file contents into a string
            Do
                Line Input #fNum, tmpMOTD
                MOTDString = MOTDString & tmpMOTD & vbCrLf
            Loop While Not EOF(fNum)
        Close fNum
        
        MOTDString = Mid(MOTDString, 1, Len(MOTDString) - 1) ' remove last vbcrlf
        MOTDString = Replace(MOTDString, "%hubname%", HubName) ' put hubname in place of %hubname%, %nick% is temporarily replaced when someone joins
    End If
   
    'Minimum share
    MinShare = ReadINI("Settings", "MinShare", ConfFile)
    If Trim(MinShare) = "" Or Val(MinShare) < 0 Then
        AddLog "*** Warning: MinShare not set/valid. Using none (0)"
        MinShare = 0
    End If
    
    StartMin = Val(ReadINI("Settings", "StartMin", ConfFile)) ' If program should start minimised
    If StartMin = 1 Then Me.WindowState = vbMinimized
    
    chkLogCommands.value = Val(ReadINI("Settings", "LogCommands", ConfFile))
    chkPopupPM.value = Val(ReadINI("Settings", "PopupPM", ConfFile))
    
    chkMD5Passwords.value = Val(ReadINI("Settings", "MD5Passwords", ConfFile))
    
    'Hub Info
    HubInfo.Speed = ReadINI("HubInfo", "Speed", ConfFile)
    HubInfo.Email = ReadINI("HubInfo", "Email", ConfFile)
    HubInfo.Interest = ReadINI("HubInfo", "Interest", ConfFile)
    HubInfo.ShareSize = Val(ReadINI("HubInfo", "ShareSize", ConfFile))
End Sub

Sub StartHub()
    If HubRunning = True Then Call StopHub
    'Initialise new collection of clients
    Set Clients = New Collection
    'Create the new socket
    lSock = Server.CreateSocket
    'Listen on port HubPort
    Server.Listen lSock, CLng(HubPort)
    'Add event to log
    AddLog "-- Running " & App.Title & ", Version " & App.Major & "." & App.Minor & "." & App.Revision
    AddLog "-- Hub started, Socket " & lSock
    'New hub start time
    HubStartTime = Now
    'Hub is running, well, it should be....
    HubRunning = True
End Sub

Sub StopHub()
    If HubRunning = False Then Exit Sub
    'Close the listening socket
    Server.CloseSocket lSock
    'Add event to log
    AddLog "-- Hub stopped."
    'Hub stopped.
    HubRunning = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    'Used to popupmenu when rightclick in systray (or otherwise on form) and to show when left click on systray icon
    If Button = vbLeftButton Then
        Me.Show
        Me.WindowState = vbNormal
    ElseIf Button = vbRightButton Then
        Me.PopupMenu mnuMain
    End If
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Me.Hide ' hide to systray on minimise
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Prompt on unload
    Dim ask As VbMsgBoxResult
    ask = MsgBox("Are you sure you want to quit?", vbQuestion + vbYesNo, "Quitting")
    If ask = vbNo Then
        Cancel = True
        Exit Sub
    End If
    
    'Stop the hub, not needed but meh
    Call StopHub
    
    'Remove from system tray
    Call RemoveSystray
    
    'Close anyone remaining open forms
    End
End Sub

Private Sub mnuClearLog_Click()
    DeleteFile LogFile ' Deleting the log file, beats wiping it
    AddLog "*** Log Cleared."
End Sub

Private Sub mnuExit_Click()
    Unload Me ' Unload the form, firing Form_Unload sub.
End Sub

Private Sub mnuRehash_Click() ' Reload settings from config file.
    Call LoadSettings
    AddLog "-- Rehashed config."
End Sub

Private Sub mnuRestart_Click() ' Restart the hub in 10 seconds.
    SendToAll "<Hub> This hub is restarting in 10 seconds."
    tmrRestart.Enabled = True
End Sub

Private Sub mnuShow_Click() ' Show the form
    If Me.WindowState = vbMinimized Then ' Why is this IF here? does it need to be? meh.
        Me.WindowState = vbNormal
        Me.Show
    End If
End Sub

Private Sub Server_OnClose(lngSocket As Long)
    ' Triggers when a remote client closes connection, or when we want to close it :D
    
    On Error GoTo remClient ' remove the socket on error

    If lngSocket <= 0 Then Exit Sub
    
    'If the client who quit was logedin, send to all loged in clients that it quit.
    If Clients.Item(Str(lngSocket)).LogedIn = True Then
        SendToAll "$Quit " & Clients.Item(Str(lngSocket)).Nick
        AddLog "-- " & Clients.Item(Str(lngSocket)).Nick & " disconnected."
    Else
        AddLog "-- " & Clients.Item(Str(lngSocket)).RemoteIP & " disconnected."
    End If

    'Close the socket and remove the client
    Clients.Remove Str(lngSocket)
    Server.CloseSocket lngSocket
    
    Exit Sub
    
remClient: ' Get rid of erroneous connection, BUT this does not remove a nick from the list if it has one, hrm, veeeeeeerrryy strange.... might need fatal error exit message here?
    Server.CloseSocket lngSocket
End Sub

Private Sub Server_OnConnectRequest(lngSocket As Long)
    'Quick BatMan! To the Batmobile! someone is attempting to join!
    Dim NewSocket   As Long
    Dim NewUser     As clsUser
    
    Set NewUser = New clsUser
    
    'Accept the connection
    NewSocket = Server.Accept(lSock)
    
    'Set up new client data to add to collection of clients
    NewUser.Socket = NewSocket
    NewUser.RemoteIP = EndPoint.GetRemoteIP(NewSocket)
    NewUser.TimeOut = 10
    NewUser.LockUsed = CreateLock
    
    'Add client to collection
    Clients.Add NewUser, Str(NewUser.Socket)
    
    'Add to log
    AddLog "-- " & NewUser.RemoteIP & " connected, Socket " & NewUser.Socket
    
    'New user has joined, so send Lock+Pk challenge
    '$Lock <lock> Pk=<pk>
    'This begins login process...
    Send NewUser.Socket, "<Hub> Login in process..." ' Do we really need this?
    Send NewUser.Socket, "$Lock " & NewUser.LockUsed & " Pk=" & CreatePK & "|$HubName " & HubName
End Sub

Function AddLog(txt As String, Optional IsCommand As Boolean = False)
    If IsCommand = True And chkLogCommands.value = vbUnchecked Then Exit Function
    'add text to log box & file if logging to file
    txtLog.Text = txtLog.Text & Time & ": " & txt & vbCrLf
    
    If UseLog = True Then
        'Write data to log file
        fNum = FreeFile
        Open LogFile For Append As fNum
            Print #fNum, Time & ": " & txt
        Close fNum
    End If
End Function

Private Sub Server_OnDataArrive(lngSocket As Long)
    'When data is recieved, get it, make sure it contains a command (|) then process it.
    Dim rData As String
    
    'Get the data
    Server.Recv lngSocket, rData
    
    'If it does not contain a command character '|' do nothing
    If InStr(rData, "|") = 0 Then Exit Sub
    
    'Handle data recieved
    ProcessData rData, lngSocket
End Sub

Private Sub Server_OnError(lngRetCode As Long, strDescription As String)
    'When SocketWise has an error, exit, preferably erroneously.
    SendToAll "<Hub> Server has encountered an error, and needs to close - " & strDescription
    AddLog "*** SERVER ERROR #" & lngRetCode & ", " & strDescription
    MsgBox "eXHub has encountered an error, and will now close." & vbCrLf & vbCrLf & "Server Error #" & lngRetCode & ", " & strDescription
    Call RemoveSystray
    End
End Sub

Private Sub tmrRestart_Timer() ' This is fired after initial warning of 10 seconds
    tmrRestart.Enabled = False
    SendToAll "<Hub> This hub is now restarting. Please reconnect."
    Call RestartHub
End Sub

Function RestartHub() ' This sub restarts the hub, rather than calling stophub, then starthub
    HubRunning = False
    
    'Close the listening socket
    Server.CloseSocket lSock
    
    'Disconnect all clients
    For Each Dummy In Clients
        Server.CloseSocket Dummy.Socket
    Next
    
    'Initialise new collection of clients
    Set Clients = New Collection
    
    'Create the new socket
    lSock = Server.CreateSocket
    
    'Listen on port HubPort (Defaut 411)
    Server.Listen lSock, CLng(HubPort)
    
    'Add event to log
    HubRunning = True
    
    'Add to log
    AddLog "-- Server restarted."
End Function

Private Sub tmrTimeOut_Timer()
    'This timer will disconnect a user after x seconds, specified by .TimeOut in clsUser
    'Also used for .Flood, flooding the server with chat messages
    '.Flood works by each time a user sends a message, .Flood + 1
    'This timer reduces .Flood by 1 per second.
    'If .Flood >= 5, then kick them
    For Each Dummy In Clients
        If Dummy.Flood > 0 And Dummy.Flood < 5 Then Dummy.Flood = Dummy.Flood - 1
        If Dummy.TimeOut > 0 Then Dummy.TimeOut = Dummy.TimeOut - 1
        If Dummy.TimeOut = 0 Or Dummy.Flood >= 5 Then Server_OnClose Dummy.Socket
    Next
End Sub

Sub ProcessData(Data2Proc As String, lngSocket As Long)
    'This is where it all happens. Center Of Operations, where recieved data is processed.
    'Teh Process Data Sub
    If Not IsIDE Then On Error GoTo errHandle
    If Trim(Data2Proc) = "" Then Exit Sub
    
    Dim rCmds() As String ' Holds first word recieved (usually command like $Kick)
    Dim rFunc() As String ' Holds the data split into spaces
    Dim rRecv() As String ' Holds data, split at $ sign
    Dim rAscii() As String  ' Data split at Chr(5)
    Dim i As Integer ' Counting variable
    Dim User As clsUser ' Rather than having to keep going Clients.Item(Str(lngSocket)) every time needed.
    Set User = Clients.Item(Str(lngSocket))
    
    'Temporary data storage vars
    Dim tmpBool As Boolean
    Dim strTemp As String
    Dim strTemp1 As String
    Dim strArray() As String
    
    rCmds = Split(Data2Proc, "|") ' Split commands at command pipe
    
    For i = LBound(rCmds) To UBound(rCmds)
        
        If Trim(rCmds(i)) = "" Then GoTo NextFunc ' If there is nothing recieved, do nothing
        rFunc = Split(rCmds(i), " ") ' Split data accordingly
        rRecv = Split(rCmds(i), "$")
        rAscii = Split(rCmds(i), Chr(5))
        
        'Recieved Chat Message, make sure it contains < , > and users nick
        If Mid(rFunc(0), 1, 1) = "<" And Mid(rFunc(0), 2, Len(User.Nick)) = User.Nick And Mid(rFunc(0), 2 + Len(User.Nick), 1) = ">" Then
            If User.LogedIn = False Then Exit Sub ' Must be loged in to send chats
            If chkDisableTalking.value = vbChecked Then Exit Sub
            
            'Add to log, not as command
            AddLog IIf(User.Nick = "", EndPoint.GetRemoteIP(User.Socket), User.Nick) & " > Hub - " & rCmds(i) & "|"
            
            'Some commands available to Ops/Users
            If rFunc(1) = "!uptime" Then
                SendToAll (rCmds(i)) ' Send chat to all clients
                SendToAll "<Hub> UpTime: " & Seconds2Time(DateDiff("s", HubStartTime, Now))
            ElseIf rFunc(1) = "!rehash" And User.Operator = True Then 'Rehash config files, dont really need
                Call mnuRehash_Click
                SendToAll "<Hub> Config file rehashed."
            ElseIf rFunc(1) = "!restart" And User.Operator = True Then  'Restart server
                Call mnuRestart_Click
            ElseIf rFunc(1) = "!clearlog" And User.Operator = True Then   'Clear log
                Call mnuClearLog_Click
            Else
                SendToAll rCmds(i) ' Send chat to all clients
            End If
            
            If User.Flood < 5 Then ' if user is flooding server
                User.Flood = User.Flood + 1
            Else
                Server_OnClose User.Socket
                Exit Sub
            End If
            
        Else ' it must have been a command we got, not a chat
            
            'Add to log
            AddLog IIf(User.Nick = "", EndPoint.GetRemoteIP(User.Socket), User.Nick) & " > Hub - " & rCmds(i) & "|", True
        
            Select Case rFunc(0) ' Recieved function
                
                'Recieved Key Challenge
                '$Key <key>
                Case "$Key"
                    strTemp = Mid(rCmds(i), Len("$Key ") + 1) ' Key = after '$Key ' onwards
                    If Lock2Key(User.LockUsed) = strTemp Or DC1_Lock2Key(User.LockUsed) = strTemp Then
                        User.GotKey = True
                        User.TimeOut = 10
                    Else ' Maybe a bad key, maybe their algorithm is off, maybe this algorithm is off, either way, they can always reconnect
                        Send User.Socket, "<Hub> Invalid key recieved, please try reconnecting."
                        Server_OnClose User.Socket
                        Exit Sub
                    End If
                
                'Validating Nick
                '$ValidateNick <nick>
                Case "$ValidateNick"
                    If User.GotKey = False Then GoTo NextFunc
                    If UBound(rFunc) = 1 Then ' Both fields are present
                        If User.GotNick = True Then GoTo NextFunc ' make sure not trying to muck with our minds
                        'Check to make sure nick is not in use
                        For Each Dummy In Clients ' make sure nick is not in use
                            If Dummy.Nick = rFunc(1) Then
                                Send User.Socket, "$ValidateDenide"
                                Server_OnClose User.Socket
                                Exit Sub
                            End If
                        Next
                        If Trim(rFunc(1)) <> "" And rFunc(1) <> "Hub" Then ' If not not used, and not blank, and not "Hub"
                            User.Nick = rFunc(1)
                            User.GotNick = True
                            User.TimeOut = 10 ' still to wait for myinfo or password
                            
                            'Read passwords from file
                            strTemp = ReadINI("Ops", rFunc(1), ConfFile)
                            strTemp1 = ReadINI("Users", rFunc(1), ConfFile)
                            strTemp = Trim(strTemp)
                            strTemp1 = Trim(strTemp1)
                            
                            If strTemp <> "" Then 'If user is an oper
                                User.Password = strTemp
                                User.WaitPass = True
                                User.Operator = True
                                User.Registered = False
                                Send User.Socket, "$GetPass"  ' Send request for password
                            ElseIf Trim(strTemp1) <> "" Then ' if user is registered
                                User.Password = strTemp1
                                User.WaitPass = True
                                User.Operator = False
                                User.Registered = True
                                Send User.Socket, "$GetPass"  ' Send request for password
                            Else ' If they are neither
                                Send User.Socket, "$Hello " & User.Nick 'Send hello msg to get client to respond with myinfo
                            End If
                        Else ' Nick is in use, send denide and close them.
                            Send User.Socket, "$ValidateDenide"
                            Server_OnClose User.Socket
                            Exit Sub
                        End If
                    End If
                
                'Revieved a password
                '$MyPass <password>
                Case "$MyPass"
                    If User.WaitPass = False Then GoTo NextFunc
                    If UBound(rFunc) = 1 Then ' password is there, don't forget, NO SPACES!
                        'MD5 Passwords
                        If chkMD5Passwords.value = vbChecked Then
                            strTemp = MD5.DigestStrToHexStr(rFunc(1))
                        Else
                            strTemp = rFunc(1)
                        End If
                        If strTemp = User.Password Then
                            User.TimeOut = -1
                            User.WaitPass = False
                            If User.Operator = True Then Send User.Socket, "$LogedIn " & User.Nick ' Send logedin if they are an op
                            Send User.Socket, "$Hello " & User.Nick 'Send hello msg to get client to respond with myinfo
                        Else
                            'Wrong password, does it really need a message?
                            Send User.Socket, "<Hub> The password for " & User.Nick & " is invalid.|$BadPass"
                            AddLog "-- Invalid password recieved for " & User.Nick & ", from " & User.RemoteIP
                            Server_OnClose User.Socket
                            Exit Sub
                        End If
                    Else
                        Send User.Socket, "$BadPass"
                        Server_OnClose User.Socket
                        Exit Sub
                    End If
                    
                'MyInfo recieved
                '$MyINFO $ALL <nick> <interest>$ $<speed>$<e-mail>$<sharesize>$
                Case "$MyINFO"
                    If User.GotNick = False And User.WaitPass = False Then GoTo NextFunc
                    If UBound(rRecv) = 7 Then ' If all fields are present
                        'Make sure they are sharing enough according to minshare
                        If Val(MinShare) > Val(rRecv(6)) Then
                            Send User.Socket, "<Hub> The minimum share is: " & MinShare & " bytes, you do not meet this!"
                            Server_OnClose User.Socket
                            Exit Sub
                        End If
                        
                        User.TimeOut = -1 ' Make sure they do not get disconnected because of timeout
                        User.ShareSize = rRecv(6) ' For use in some min share size later
                        User.InfoString = rCmds(i) 'Store the whole infostring, rather then seperating parts of it
                            
                        If User.LogedIn = False Then ' If first time logging in
                            SendToAll "$Hello " & User.Nick ' Tell everyone someone joined
                            User.LogedIn = True
                            Send User.Socket, "<Hub> Login successful."
                            Send User.Socket, "<Hub> Welcome to " & HubName & ", this hub is running " & App.Title & ", Version " & App.Major & "." & App.Minor & "." & App.Revision & " [UpTime: " & Seconds2Time(DateDiff("s", HubStartTime, Now)) & "]"
                            If UseMOTD = True Then 'if there is a motd, send it
                                strTemp = Replace(MOTDString, "%nick%", User.Nick)
                                'Don't add motd to log, else log could get quite big
                                Send User.Socket, "<Hub> " & strTemp, False
                                Send User.Socket, "<Hub> End of MOTD", False
                            End If
                            AddLog "-- " & User.Nick & " logged in."
                        End If
                        
                        SendToAll User.InfoString ' Should we bother? Yeah, why not.
                    End If
                    
                'Request for a users info
                '$GetINFO <othernick> <nick>
                'respond with: $MyINFO $ALL <nick> <interest>$ $<speed>$<e-mail>$<sharesize>$
                'This is stored in user.infostring
                Case "$GetINFO"
                    If User.GotNick = False Then GoTo NextFunc
                    Call SendInfoString(User.Socket, rFunc(1))
                    
                'Request for nick list
                '$GetNickList
                'Respond with:
                '   $NickList <nick1>$$<nick2>$$<nick3>$$...
                '   $OpList <op1>$$<op2>$$<op3>$$...
                Case "$GetNickList"
                    If User.GotNick = False Then GoTo NextFunc
                    Call SendNickList(User.Socket)
                    
                'A client attempts a connection to another client
                '$ConnectToMe <remoteNick> <senderIp>:<senderPort>
                'Send exact message to <remoteNick>
                'Sharing must be enabled to get past this as well
                Case "$ConnectToMe"
                    If User.LogedIn = False Then GoTo NextFunc
                    If chkDisableSharing.value <> vbChecked Then ' make sure sharing is enabled / disabled
                        If UBound(rFunc) = 2 Then 'All fields are present
                            SendToNick rFunc(1), rCmds(i)
                        End If
                    End If
                    
                'A passive client attempts a connection to another client
                '$RevConnectToMe <nick> <remoteNick>
                'Send exact message to <remoteNick>
                Case "$RevConnectToMe"
                    If User.LogedIn = False Then GoTo NextFunc
                    If UBound(rFunc) = 2 Then 'All fields are present
                        SendToNick rFunc(2), rCmds(i)
                    End If
                    
                'Active Client is attempting a search
                '$Search <clientip>:<clientport> <searchstring>
                'or
                'Passive Client is attempting to search
                '$Search Hub:<searchingNick> <searchstring>
                Case "$Search", "$Search Hub:"
                    If User.LogedIn = False Then GoTo NextFunc
                    SendToAll rCmds(i)
                
                'Search Response for Passive Clients
                '$SR <resultNick> <filepath><filesize> <freeslots>/<totalslots><hubname> (<hubhost>[:<hubport>])<searchingNick>
                ' = Chr(5)
                Case "$SR"
                    If User.LogedIn = False Then GoTo NextFunc
                    strArray = Split(rCmds(i), Chr(5))
                    If UBound(strArray) < 2 Then GoTo NextFunc 'all fields are present
                    SendToNick Trim(strArray(UBound(strArray))), Mid(rCmds(i), 1, Len(rCmds(i)) - Len(strArray(UBound(strArray))) - 1)  ' - 1  for  before <searchingNick>
                    'NB instead of the wierd mid command,
                    'Len(rCmds(i)) - Len(strArray(UBound(strArray))) - 1
                    'may be used, but is not tested.
                
                'Private message
                '$To: <othernick> From: <nick> $<<nick>> <message>
                Case "$To:"
                    If User.LogedIn = False Then GoTo NextFunc
                    If UBound(rRecv) = 2 Then ' if all fields are present
                        If rFunc(1) = "Hub" Then ' if a pm to the Hub
                            If chkPopupPM.value = vbChecked Then Call GetPM(rFunc(3), rRecv(2))
                            GoTo NextFunc
                        End If
                        'Find <othernick> and send them the message
                        SendToNick rFunc(1), rCmds(i)
                    End If
                    
                'An op is attempting to move a client to another hub,
                '$OpForceMove $Who:<victimNick>$Where:<newIp>$Msg:<reasonMsg>
                Case "$OpForceMove"
                    If User.LogedIn = False Then GoTo NextFunc
                    If User.Operator = False Then GoTo NextFunc
                    Call RedirectUser(User.Nick, rCmds(i))
                    
                'If an operator has kicked a user
                '$Kick <VictimNick>
                Case "$Kick"
                    If User.LogedIn = False Then GoTo NextFunc
                    If User.Operator = False Then GoTo NextFunc
                    If UBound(rFunc) = 1 Then ' All fields are present
                        If rFunc(1) <> "Hub" Then 'Cant kick hub!
                            For Each Dummy In Clients
                                If Dummy.Nick = rFunc(1) Then ' Close them, and exit for loop, as no duplicate nicks on hub
                                    Server_OnClose Dummy.Socket
                                    Exit For
                                End If
                            Next
                        End If
                    End If
                    
                Case "$Version"
                    'Ignore this.
                Case "$Quit"
                    Server_OnClose User.Socket ' Kill their connection, Maybe not needed?
                    Exit Sub
                    
                Case "$MultiConnectToMe"
                    'Dont bother with multi hubs, not (yet) implemented.
                Case "$MultiSearch"
                    'Dont bother with multi hubs, not (yet) implemented.
                
                Case Else
                    AddLog "*** " & IIf(User.Nick = "", EndPoint.GetRemoteIP(User.Socket), User.Nick) & " sent invalid command: " & rCmds(i)
                    Server_OnClose User.Socket
                    Exit Sub
                    'Bit harsh? Might disconnect on new/not present commands
                    'Probably should alert an op of this? even all clients .... meh.
                
            End Select
            
        End If ' end if for chat/command
        
NextFunc: ' Used to jump to next command if recieved an error or something in one.
    DoEvents ' just because
        
    Next i
    
    Exit Sub
    
'When an error occurs during this sub, and running from a exe, not ide, EVIL DIE!
errHandle:
    SendToAll "<Hub> Server has encountered an error, and needs to close - " & Err.Description
    AddLog "ERROR #" & Err.Number & ", Source:" & Err.Source & ", " & Err.Description
    MsgBox "eXHub has encountered an error, and will now close." & vbCrLf & vbCrLf & "Error #" & Err.Number & ", Source:" & Err.Source & ", " & Err.Description
    Call RemoveSystray
    End
End Sub

Function SendToAll(strMSG As String, Optional IsChat As Boolean = False)
    If HubRunning = False Then Exit Function
    If Trim(strMSG) = "" Then Exit Function
    'Send msg to all loged in clients
    For Each Dummy In Clients
        If Dummy.LogedIn = True Then Send Dummy.Socket, strMSG, False
    Next
    'Add to log etc
    AddLog "Hub > All - " & strMSG & "|", (Not IsChat)
End Function

Function Send(lngSocket As Long, strMSG As String, Optional AddToLog As Boolean = True)
    If HubRunning = False Then Exit Function
    If Trim(strMSG) = "" Then Exit Function
    'Send msg to lngSocket
    Server.Send lngSocket, strMSG & "|"
    If AddToLog = True Then AddLog "Hub > " & IIf(Clients(Str(lngSocket)).Nick = "", EndPoint.GetRemoteIP(lngSocket), Clients(Str(lngSocket)).Nick) & " - " & strMSG & "|", True
End Function

Function SendToNick(strNick As String, strMSG As String, Optional MustBeLogedIn As Boolean = False)
    'Sends message to specified nick
    If HubRunning = False Then Exit Function
    If Trim(strMSG) = "" Or Trim(strNick) = "" Then Exit Function
    For Each Dummy In Clients
        If Dummy.Nick = strNick Then
            ' if they have to be loged in to send them the msg
            ' Don't think this is needed but meh
            If MustBeLogedIn = True And Dummy.LogedIn = False Then Exit Function
            Send Dummy.Socket, strMSG
            Exit Function
        End If
    Next
End Function

Function SendInfoString(lngSocket As Long, strNick As String)
    'send strNick's InfoString requested by lngSocket
    If HubRunning = False Then Exit Function
    If Trim(strNick) = "" Then Exit Function
    
    If strNick = "Hub" Then ' requesting hub info
        Send lngSocket, "$MyINFO $ALL Hub " & HubInfo.Interest & "$ $" & HubInfo.Speed & "$" & HubInfo.Email & "$" & HubInfo.ShareSize & "$"
    Else
        'If they are requesting their own info, send it, otherwise find appropriate nick
        If Clients(Str(lngSocket)).Nick = strNick Then
            Send lngSocket, Clients(Str(lngSocket)).InfoString
        Else
            For Each Dummy In Clients
                If Dummy.Nick = strNick Then Send lngSocket, Dummy.InfoString
                Exit For ' Unique nicknames, only one on a hub.
            Next
        End If
    End If
End Function

Function SendNickList(lngSocket As Long)
    'Send NickList + OpList to lngSocket
    Dim NickList As String
    Dim OpList As String
    Dim User As clsUser
    
    Set User = Clients(Str(lngSocket))
    
    NickList = "$NickList Hub$$" ' Make sure Hub is on list
    OpList = "$OpList Hub$$" ' And we're an op! :D
    
    For Each Dummy In Clients ' Collect all nicks + ops
        If Dummy.LogedIn = True Then
            NickList = NickList + Dummy.Nick & "$$"
            If Dummy.Operator = True Then OpList = OpList + Dummy.Nick & "$$"
        End If
    Next
                
    'Ff for some strange reason user requesting nicklist user was not added to nicklist, add them
    If InStr(NickList, User.Nick) = 0 Then NickList = NickList + User.Nick & "$$"
    If InStr(OpList, User.Nick) = 0 And User.Operator = True Then OpList = OpList + User.Nick & "$$"
    
    'Send it to the requesting client
    Send lngSocket, NickList & "|" & OpList
End Function

Private Sub txtChat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSendChat_Click
End Sub

Private Sub txtLog_Change() ' Make sure log box scrolls,
    'BUG FOUND! When form is hidden then shown again, the selstart of the box gets reset.
    '    This is annoying when showing from system tray, as you have to wait until data is recieved for it to scroll down _
    '       or scroll down yourself. GRRRR.... RSI!
    'Could fix this but it isn't needed.
    txtLog.SelStart = Len(txtLog.Text)
End Sub

Function GetPM(FromNick As String, strPM As String)
    '$To: <othernick> From: <nick> $<<nick>> <message>
    If chkPopupPM.value <> vbChecked Then Exit Function
    Dim Frm As Form
    Dim ExistPM As frmPM
    For Each Frm In Forms
        If Frm.Tag = FromNick Then
            Set ExistPM = Frm
            ExistPM.txtRecv.Text = ExistPM.txtRecv & strPM & vbCrLf
            ExistPM.Show
            Exit Function
        End If
    Next
    'Will exit if found form with nick in
    Dim NewPM As New frmPM
    NewPM.Tag = FromNick
    NewPM.Caption = "PM - " & FromNick
    NewPM.txtRecv.Text = strPM & vbCrLf
    NewPM.Show
End Function

Public Function SendPM(ToNick As String, strPM As String)
    '$To: <othernick> From: <nick> $<<nick>> <message>
    'Dont do this using sendtonick, because if it doesnt send it we want to know
    'Unless a return could be gotten from the sendtonick function if it doesnt find the nick
    For Each Dummy In Clients
        If Dummy.Nick = ToNick Then
            Send Dummy.Socket, "$To: " & ToNick & " From: Hub $<Hub> " & strPM
            Exit Function
        End If
    Next
    Call GetPM(ToNick, ToNick & " appears offline.")
End Function

Function RedirectUser(FromNick As String, strOpForceMove As String)
    '$OpForceMove $Who:<victimNick>$Where:<newIp>$Msg:<reasonMsg>
    Dim strArray() As String
    Dim VictimNick As String
    Dim NewIP As String
    Dim ReasonMSG As String
    
    strArray = Split(strOpForceMove, "$")
    If UBound(strArray) <> 4 Then Exit Function  ' contains all stuff
    
    VictimNick = Mid(strArray(2), InStr(strArray(2), ":") + 1)
    If VictimNick = "Hub" Or Trim(VictimNick) = "" Then Exit Function ' Cant redir the hub!
    
    NewIP = Mid(strArray(3), InStr(strArray(3), ":") + 1)
    If Trim(NewIP) = "" Then Exit Function
    
    ReasonMSG = Mid(strArray(4), InStr(strArray(4), ":") + 1)
    
    'Send following to client
    '$ForceMove <newIp>
    '$To: <victimNick> From: <senderNick> $<<senderNick>> You are being re-directed to <newHub> because: <reasonMsg>
    SendToNick VictimNick, "$To: " & VictimNick & " From: " & FromNick & " $<" & FromNick & "> You are being re-directed to " & NewIP & ", because: " & ReasonMSG
    SendToNick VictimNick, "$ForceMove " & NewIP
    
    For Each Dummy In Clients
        If Dummy.Nick = VictimNick Then Dummy.TimeOut = 5 ' give them five seconds for their client to do this
    Next
    
End Function
