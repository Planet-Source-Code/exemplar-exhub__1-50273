VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private dcSocket        As Long
Private dcRemoteIP      As String
Private dcLogOnTime     As Date
Private dcNick          As String
Private dcShareSize     As String
Private dcInfoString    As String
Private dcPassword      As String
Private dcOperator      As Boolean
Private dcRegistered    As Boolean
Private dcLogedIn       As Boolean
Private dcLockUsed      As String
Private dcGotKey        As Boolean
Private dcGotNick       As Boolean
Private dcTimeOut       As Integer
Private dcWaitPass      As Boolean
Private dcFlood         As Integer

'Connection related stuff
Public Property Get Socket() As Long
    Socket = dcSocket
End Property
Public Property Let Socket(nData As Long)
    dcSocket = nData
End Property

Public Property Get RemoteIP() As String
    RemoteIP = dcRemoteIP
End Property
Public Property Let RemoteIP(nData As String)
    dcRemoteIP = nData
End Property

Public Property Get LogOnTime() As Date
    LogOnTime = dcLogOnTime
End Property
Public Property Let LogOnTime(nData As Date)
    dcLogOnTime = nData
End Property

'Direct Connect related
Public Property Get Nick() As String
    Nick = dcNick
End Property
Public Property Let Nick(nData As String)
    dcNick = nData
End Property

Public Property Get ShareSize() As String
    ShareSize = dcShareSize
End Property
Public Property Let ShareSize(nData As String)
    dcShareSize = nData
End Property

'Misc
Public Property Get Operator() As Boolean
    Operator = dcOperator
End Property
Public Property Let Operator(nData As Boolean)
    dcOperator = nData
End Property

Public Property Get Registered() As Boolean
    Registered = dcRegistered
End Property
Public Property Let Registered(nData As Boolean)
    dcRegistered = nData
End Property

Public Property Get LogedIn() As Boolean
    LogedIn = dcLogedIn
End Property
Public Property Let LogedIn(nData As Boolean)
    dcLogedIn = nData
End Property

Public Property Get LockUsed() As String
    LockUsed = dcLockUsed
End Property
Public Property Let LockUsed(nData As String)
    dcLockUsed = nData
End Property

Public Property Get TimeOut() As Integer
    TimeOut = dcTimeOut
End Property
Public Property Let TimeOut(nData As Integer)
    dcTimeOut = nData
End Property

Public Property Get GotKey() As Boolean
    GotKey = dcGotKey
End Property
Public Property Let GotKey(nData As Boolean)
    dcGotKey = nData
End Property

Public Property Get GotNick() As Boolean
    GotNick = dcGotNick
End Property
Public Property Let GotNick(nData As Boolean)
    dcGotNick = nData
End Property

Public Property Get InfoString() As String
    InfoString = dcInfoString
End Property
Public Property Let InfoString(nData As String)
    dcInfoString = nData
End Property

Public Property Let WaitPass(nData As Boolean)
    dcWaitPass = nData
End Property
Public Property Get WaitPass() As Boolean
    WaitPass = dcWaitPass
End Property

Public Property Get Password() As String
    Password = dcPassword
End Property
Public Property Let Password(nData As String)
    dcPassword = nData
End Property

Public Property Get Flood() As Integer
    Flood = dcFlood
End Property
Public Property Let Flood(nData As Integer)
    dcFlood = nData
End Property
