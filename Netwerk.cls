VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNetwerk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'On server: player X uses Winsock(X-2)

'# Zie Sub Main in modAlgemeen voor toevoegen van Winsocks!

'*                *
'* General Events *
'*                *
Public Event Message(SpelerNr As Integer, Message As String)
Public Event SetNaam(SpelerNr As Integer, Naam As String)
Public Event Error(Number As NetworkErrorType)

'*               *
'* Server Events *
'*               *
Public Event ClientJoined(SpelerNr As Integer, Naam As String)
Public Event ClientDisconnected(SpelerNr As Integer)
Public Event ClientLegtKaart(SpelerNr As Integer, KaartNr As Integer)
Public Event NamenRequest(Source As Integer)

'*               *
'* Client Events *
'*               *
Public Event ConnectFailed()
Public Event Connected()
'# werkt niet: Public Event SetKaarten(SpelerNr As Integer, Kaarten() As Kaart)
Public Event SetVoorspelling(SpelerNr As Integer, Voorspelling As Integer)
Public Event SetNamen(Namen() As String)
Public Event StartClientGame(EerstOpkomen As Integer)
'Public Event StopClientGame()

Dim cWinsock(0 To 2) As Winsock
Dim cWinsockListen As Winsock
Dim cClientConnected(2 To 4) As Boolean
Dim cIsServer As Boolean
Dim cLastError As Integer

Dim cInitialized As Boolean

Private Enum ClientStatusType
  'WaitForAcceptJoin
  WaitForStart
  WaitForKaarten
  WaitForVoorspelling
  WaitForLegKaart
End Enum

Public Enum NetworkErrorType
  nerGeneralError = 700
  nerNotInitialized = 701
  nerBadNetwork = 10047
  nerBadAdress = 10049
End Enum

'*                *
'* Initialisation *
'*                *
Public Sub Init(ByRef WinsockListen As Winsock, ByRef Winsock1 As Winsock, ByRef Winsock2 As Winsock, ByRef Winsock3 As Winsock)
  Set cWinsockListen = WinsockListen
  Set cWinsock(0) = Winsock1
  Set cWinsock(1) = Winsock2
  Set cWinsock(2) = Winsock3
  
  'cWinsock(0).Close
  'cWinsock(1).Close
  'cWinsock(2).Close

  cInitialized = True
End Sub

'*                           *
'* Private server procedures *
'*                           *
Private Function ServerToClientSpelerNr(ServerSpelerNr As Integer, DoelClientNr As Integer)
  'Welk nummer heeft ServerSpelerNr op DoelClientNr
  '* Server procedure *
  If Not TestServerClientError(True) Then
    ServerToClientSpelerNr = (ServerSpelerNr + 4 - DoelClientNr) Mod 4 + 1
  End If
End Function
Private Function TestServerClientError(MustBeServer As Boolean) As Boolean
  If cIsServer <> MustBeServer Then
    Err.Raise vbObjectError + 600, "clsNetwerk", "Illegal call in server/client mode"
  End If
  TestServerClientError = (cIsServer <> MustBeServer)
End Function
Private Sub ServerReceivedJoin(SpelerNr As Integer, Naam As String)
  Dim i As Integer
  'SpelerNr = Source
  
  MsgBox Naam & " joined (player " & SpelerNr & ")."

  
  'RaiseEvent SetNaam(Naam)
  For i = 2 To 4
    If i <> SpelerNr Then
      'SendNaam(
    End If
  Next i

  Stop
End Sub
Private Sub NamenRequest(Source As Integer, SpelerNr As Integer)
  Stop
End Sub

'*                           *
'* Private Client procedures *
'*                           *
Private Sub ClientGotNamen(Namen() As String)
  Dim i As Integer
  'For i = 2 To 4
  '  RaiseEvent SetNaam(Namen())
  'Next i
  Stop
End Sub


'*                           *
'* General public procedures *
'*                           *
Public Sub StartServer()
  Dim ClientNr As Integer
  
  If Not cInitialized Then
    RaiseEvent Error(nerNotInitialized)
  Else
    On Error GoTo StartServerError
    
    cIsServer = True
    
    For ClientNr = 0 To 2
      cWinsock(ClientNr).Close
    Next ClientNr
    cWinsockListen.Close
    cWinsockListen.Listen
  End If
    
  Exit Sub
  
StartServerError:
  RaiseEvent Error(Err.Number)
End Sub
Public Sub SendMessage(Message As String)
  '- Is server: verspreiden naar allen
  '- Is client: alleen naar server
  Stop
End Sub
Public Sub SendNamen(Dest As Integer, Namen() As String)
  Stop
End Sub

'*                            *
'* To be invoked by a server  *
'*                            *
Public Sub DisconnectClient(ByVal SpelerNr As Integer)
  If Not TestServerClientError(True) Then
    If SpelerNr < 2 Or SpelerNr > 4 Then
      Err.Raise vbObjectError + 602, "clsNetwerk", "Client number to disconnect must be 2, 3 or 4"
    Else
      cClientConnected(SpelerNr) = False
      'winsock.disconnect
      Stop
    End If
  End If
End Sub
Public Property Get ClientConnected(ByVal SpelerNr As Integer)
  If Not TestServerClientError(True) Then
    ClientConnected = cClientConnected(SpelerNr)
  End If
End Property
Public Sub StartGame()
  If Not TestServerClientError(True) Then
  End If
  Stop
End Sub
Public Sub SendEerstOpkomen(SpelerNr As Integer)
  If Not TestServerClientError(True) Then
  End If
  Stop
End Sub
'Public Sub SendKaarten(Kaarten() As Kaart) '#
'  If Not TestServerClientError(True) Then
'  End If
'  Stop
'End Sub
Public Sub LegKaart(SpelerNr As Integer, KaartNr As Integer)
  'Laat clients weten dat een speler een kaart legt
  If Not TestServerClientError(True) Then
  End If
  Stop
End Sub

'*                           *
'* To be invoked by a client *
'*                           *
Public Sub ConnectToServer(ByRef Address As String)
  On Error GoTo ConnectToServerError
  
  If Not cInitialized Then
    RaiseEvent Error(nerNotInitialized)
  Else
    'If Not TestServerClientError(False) Then
    cIsServer = False
    
    cWinsockListen.Close
    cWinsock(0).RemoteHost = Address
    cWinsock(0).Close
    cWinsock(0).Connect
  End If
    
  Exit Sub

ConnectToServerError:
  'cLastError = Err.Number
  
  Select Case Err.Number
    Case nerBadNetwork
      RaiseEvent Error(nerBadNetwork)
    Case nerBadAdress
      RaiseEvent Error(nerBadAdress)
    Case Else
      RaiseEvent Error(nerGeneralError)
  End Select
  Resume Next
End Sub
Public Sub ClientSendJoin(ByRef Naam As String)
  cWinsock(0).SendData "Join:" & Naam & "|"
End Sub
Public Sub ClientDisconnect()
  Dim i As Integer
  
  If Not cInitialized Then
    RaiseEvent Error(nerNotInitialized)
  Else
    For i = 0 To 2
      cWinsock(i).Close
    Next i
  End If
End Sub
Public Sub GetNaam(ByVal SpelerNr As Integer)
  Stop
End Sub
Public Sub IkLegKaart(ByVal SpelerNr As Integer, KaartNr As Integer)
  Stop
End Sub
'Public Sub ClientSendStatus(NewStatus As ClientStatusType) '#
'  'Client is ready for: ...
'  Stop
'End Sub

'*                               *
'* Generated by server's WinSock *
'*                               *
Public Sub ConnectionRequest(ByVal requestID As Long)
  'ByVal WinsockIndex As Integer,
  Dim TryClient(1 To 3) As Integer
  Dim ClientNr As Integer
  TryClient(1) = 1
  TryClient(2) = 0
  TryClient(3) = 2
  
  If cIsServer Then
    'And Not cClientConnected(WinsockIndex + 2)
    For ClientNr = 1 To 3
      If cWinsock(TryClient(ClientNr)).State = sckClosed Then
        cWinsock(TryClient(ClientNr)).Accept requestID
MsgBox "Client verbonden; nr " & ClientNr
        Exit For
      End If
    'If tcpServer.State <> sckClosed Then tcpServer.Close
        'RaiseEvent ClientJoined(SpelerNr, Naam)
    Next ClientNr
  End If
End Sub

'*                                           *
'* Generated by server's or client's WinSock *
'*                                           *
Public Sub DataArrival(ByVal WinsockIndex As Integer, ByVal bytesTotal As Long)
  Dim LinkerWoord As String
  Dim RechterWoord As String
  Dim PuntkommaLoc As Integer
  Dim Data As String
  
  cWinsock(WinsockIndex).GetData Data, vbString
  
  MsgBox "DataArrival: " & Data
  
  If Right(Data, 1) <> "|" Then
    MsgBox "Winsock error: data doesn't end with '|': " & Data
  End If
  
  PuntkommaLoc = InStr(Data, ":")
  LinkerWoord = UCase(Left(Data, PuntkommaLoc))
  RechterWoord = Mid(Data, PuntkommaLoc + 1, Len(Data) - PuntkommaLoc - 1)
  
  Select Case LinkerWoord
    Case "JOIN"
      ServerReceivedJoin WinsockIndex + 2, RechterWoord
  End Select
  'RaiseEvent GotCards(SpelerNr, Cards())
  'RaiseEvent GotVoorspelling(SpelerNr, Voorspelling)
  'RaiseEvent GotMessage(SpelerNr, Message)
  'RaiseEvent GotNaam(SpelerNr, Naam)
  ' --> Is server: alle verspreiden.
  'RaiseEvent StartClientGame()
  'RaiseEvent StopClientGame()
End Sub

'*                               *
'* Generated by server's WinSock *
'*                               *
Public Sub CloseWinsock(ByVal Index As Integer) 'Close is a reserved word
  'RaiseEvent ClientDisconnected(SpelerNr)
  Stop
End Sub

'*                               *
'* Generated by client's WinSock *
'*                               *
Public Sub Connect(WinsockIndex As Integer)
  MsgBox "Connected to server. Joining..."
  ClientSendJoin Spelers(1).Naam
End Sub

Public Property Get LastError() As Integer
  LastError = cLastError
End Property

Public Property Get Initialized() As Boolean
  Initialized = cInitialized
End Property

Private Sub Class_Initialize()
  cInitialized = False
End Sub
