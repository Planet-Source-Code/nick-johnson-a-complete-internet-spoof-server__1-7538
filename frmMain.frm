VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   Caption         =   "Spoof Server"
   ClientHeight    =   1425
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4650
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1425
   ScaleWidth      =   4650
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNext 
      Caption         =   ">"
      Enabled         =   0   'False
      Height          =   285
      Left            =   4320
      TabIndex        =   7
      Top             =   0
      Width           =   285
   End
   Begin VB.CommandButton cmdPrev 
      Caption         =   "<"
      Enabled         =   0   'False
      Height          =   285
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   285
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   1440
      Top             =   1440
   End
   Begin MSWinsockLib.Winsock wsServer 
      Index           =   0
      Left            =   2880
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   80
   End
   Begin MSWinsockLib.Winsock wsClient 
      Index           =   0
      Left            =   2400
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock wsListen 
      Left            =   1920
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   80
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   720
      TabIndex        =   10
      Top             =   1080
      Width           =   3855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Status:"
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label lblNumber 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0/0"
      Height          =   285
      Left            =   360
      TabIndex        =   8
      Top             =   0
      Width           =   3855
   End
   Begin VB.Label lblToPort 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3720
      TabIndex        =   5
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Port:"
      Height          =   255
      Left            =   3240
      TabIndex        =   4
      Top             =   720
      Width           =   375
   End
   Begin VB.Label lblToHost 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   720
      TabIndex        =   3
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "To:"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Width           =   615
   End
   Begin VB.Label lblFromIP 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   720
      TabIndex        =   1
      Top             =   360
      Width           =   3855
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "From:"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   615
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim intWatched As Integer

Private Sub cmdNext_Click()
    intWatched = intWatched + 1
    updateDisplay
End Sub

Private Sub cmdPrev_Click()
    intWatched = intWatched - 1
    updateDisplay
End Sub

Private Sub Form_Load()
    'Listen for connection requests
    wsListen.Listen
End Sub

Private Sub Timer1_Timer()
    'Performs garbage collection and updates UI
    'Update UI
    updateDisplay
    Dim wS As Winsock
    For Each wS In wsServer
        If (wS.State = 9 Or wS.State = 0) And wS.Index <> 0 Then
            'If the wsServer has disconnected or errored, and
            'it's not element 0, remove it (garbage collection)
            wsClient(wS.Index).Close
            wS.Close
            Unload wsClient(wS.Index)
            Unload wS
        End If
    Next
End Sub

Private Sub wsClient_Close(Index As Integer)
    'When the remote server disconnects, close the port,
    wsClient(Index).Close
    If wsServer(Index).State = 7 Then
        'If the client is still connected to our server,
        'Reset and wait for another host.
        wsClient(Index).RemoteHost = ""
        wsClient(Index).RemotePort = 0
        wsServer(Index).Tag = "1"
        wsServer(Index).SendData "Client has disconnected. Enter host:port to connect to another" + vbCrLf
    End If
End Sub

Private Sub wsClient_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    'When data arrives, if we are connected, forward it on
    Dim strData As String
    wsClient(Index).GetData strData, vbString
    If wsServer(Index).State = 7 Then
        wsServer(Index).SendData strData
    End If
End Sub

Private Sub wsClient_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    'If an error occurs and a client is connected to our server,
    'send them the error message. If it has forced the server
    'to disconnect, reset and wait for a new host:port
    If wsServer(Index).State = 7 Then
        wsServer(Index).SendData "ERR " + Str(Number) + ": " + Description + vbCrLf
        If wsClient(Index).State <> 7 Then
            wsClient(Index).Close
            wsClient(Index).RemoteHost = ""
            wsClient(Index).RemotePort = 0
            wsServer(Index).SendData "Client is not connected. Enter host:port to connect to another" + vbCrLf
            wsServer(Index).Tag = 1
        End If
    End If
    CancelDisplay = True
End Sub

Private Sub wsListen_ConnectionRequest(ByVal requestID As Long)
    'When a connection request is heard by the dedicated
    'listener, spawn a new client/server pair to handle the
    'request.
    Load wsServer(wsServer.UBound + 1)
    Load wsClient(wsClient.UBound + 1)
    wsServer(wsServer.UBound).Accept requestID
    wsServer(wsServer.UBound).SendData "Welcome! host:port to connect" & vbCrLf
    wsServer(wsServer.UBound).Tag = "1"
    updateDisplay
End Sub

Private Sub wsServer_Close(Index As Integer)
    'If the remote client disconnects, disconnect our server
    'connection and unload them both
    wsClient(Index).Close
    wsServer(Index).Close
    Unload wsClient(Index)
    Unload wsServer(Index)
End Sub

Private Sub wsServer_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    'For each letter recieved from our client, process it
    'seperately (serverData)
    Dim strData As String
    wsServer(Index).GetData strData, vbString
    For a = 1 To Len(strData)
        serverData Asc(Mid(strData, a, 1)), Index
    Next a
End Sub

Private Sub serverData(bChr As Byte, Index As Integer)
    'For each letter, first check if we are listening for a host.
    'If we are, add it, and if it is a :, listen for port.
    'If we are listening for a port, and it is a CR, then
    'make the connection. If we are already connected, then
    'simply forward everything on to the remote server.
    If wsServer(Index).Tag = "1" Then
        If bChr <> 10 Then
            If bChr = 13 Then
                wsServer(Index).SendData "Please specify both host and port, seperated by a :" + vbCrLf
                wsClient(Index).RemoteHost = ""
            ElseIf bChr = Asc(":") Then
                wsServer(Index).Tag = "2"
            Else
                If bChr <> 8 Then
                    wsClient(Index).RemoteHost = wsClient(Index).RemoteHost + Chr(bChr)
                Else
                    If Len(wsClient(Index).RemoteHost) > 0 Then
                        wsClient(Index).RemoteHost = Left(wsClient(Index).RemoteHost, Len(wsClient(Index).RemoteHost) - 1)
                    End If
                End If
            End If
        End If
    ElseIf wsServer(Index).Tag = "2" Then
        If bChr <> 10 Then
            If bChr = 13 Then
                wsServer(Index).Tag = "0"
                wsServer(Index).SendData "Connecting..." + vbCrLf
                wsClient(Index).Connect
            Else
                If bChr <> 8 Then
                    wsClient(Index).RemotePort = Val(Str(wsClient(Index).RemotePort) + Chr(bChr)) Mod 65535
                Else
                    wsClient(Index).RemotePort = Val(Left(Str(wsClient(Index).RemotePort), Len(Str(wsClient(Index).RemotePort)) - 1))
                End If
            End If
        End If
    Else
        If wsClient(Index).State = 7 Then
            wsClient(Index).SendData Chr(bChr)
        ElseIf wsClient(Index).State = 9 Or wsClient(Index).State = 0 Then
            wsClient(Index).Close
            wsClient(Index).RemoteHost = ""
            wsClient(Index).RemotePort = 0
            wsServer(Index).SendData "Remote has disconnected. Enter host:port to connect to another" + vbCrLf
            wsServer(Index).Tag = 1
        End If
    End If
End Sub

Private Sub updateDisplay()
    If intWatched >= wsServer.Count Then intWatched = wsServer.Count - 1
    If intWatched = 0 And wsServer.Count > 1 Then intWatched = 1
    lblNumber.Caption = Str(intWatched) + "/" + Str(wsServer.Count - 1)
    If intWatched = 1 Or intWatched = 0 Then cmdPrev.Enabled = False Else cmdPrev.Enabled = True
    If intWatched = wsServer.Count - 1 Then cmdNext.Enabled = False Else cmdNext.Enabled = True
    If intWatched > 0 And intWatched < wsServer.Count Then
        lblFromIP.Caption = wsServer(intWatched).RemoteHostIP
        lblToHost.Caption = wsClient(intWatched).RemoteHost
        lblToPort.Caption = wsClient(intWatched).RemotePort
        With lblStatus
            Select Case wsClient(intWatched).State
            Case 0
                .Caption = "Closed"
            Case 1
                .Caption = "Open"
            Case 2
                .Caption = "Listening"
            Case 3
                .Caption = "Connection Pending"
            Case 4
                .Caption = "Resolving Host"
            Case 5
                .Caption = "Host resolved"
            Case 6
                .Caption = "Connecting"
            Case 7
                .Caption = "Connected"
            Case 8
                .Caption = "Peer is closing connection"
            Case 9
                .Caption = "Error"
            End Select
        End With
    Else
        lblFromIP.Caption = ""
        lblToHost.Caption = ""
        lblToPort.Caption = ""
        lblStatus.Caption = ""
    End If
End Sub
