VERSION 5.00
Begin VB.Form frmPuntentelling 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Puntentelling"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5055
   HelpContextID   =   301
   Icon            =   "Puntentelling.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Puntentelling"
      Height          =   2895
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4815
      Begin VB.TextBox txtStrafpuntenPerSlag 
         Height          =   285
         Left            =   1800
         MaxLength       =   1
         TabIndex        =   4
         Text            =   "3"
         ToolTipText     =   "Dit aantal strafpunten wordt afgetrokken"
         Top             =   2040
         Width           =   255
      End
      Begin VB.OptionButton optFoutVoorspeldPunten 
         Caption         =   "10 punten"
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   6
         ToolTipText     =   "Een onjuiste voorspelling levert 10 punten op, met aftrek van strafpunten"
         Top             =   1680
         Width           =   2895
      End
      Begin VB.OptionButton optFoutVoorspeldPunten 
         Caption         =   "0 punten"
         Height          =   255
         Index           =   0
         Left            =   1800
         TabIndex        =   7
         ToolTipText     =   "Een onjuiste voorspelling levert geen punten op"
         Top             =   1320
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.TextBox txtPuntenPerSlag 
         Height          =   285
         Left            =   2880
         MaxLength       =   1
         TabIndex        =   5
         Text            =   "1"
         ToolTipText     =   "Als je goed hebt voorspeld, krijg je dit aantal punten per voorspelde slag"
         Top             =   600
         Width           =   255
      End
      Begin VB.CheckBox chkNegatiefMogelijk 
         Caption         =   "Negatieve scores mogelijk"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1800
         TabIndex        =   3
         ToolTipText     =   "Toestaan dat je negatieve scores kunt halen"
         Top             =   2400
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "punten aftrek per slag verschil"
         Height          =   255
         Left            =   2160
         TabIndex        =   11
         Top             =   2040
         Width           =   2175
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000014&
         X1              =   240
         X2              =   4560
         Y1              =   1095
         Y2              =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Goed voorspeld:"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Fout voorspeld:"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "+ 10 punten +         punt(en) per slag"
         Height          =   255
         Left            =   1800
         TabIndex        =   8
         Top             =   600
         Width           =   2655
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         X1              =   240
         X2              =   4560
         Y1              =   1080
         Y2              =   1080
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK..."
      Default         =   -1  'True
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdAnnuleren 
      Cancel          =   -1  'True
      Caption         =   "&Annuleren"
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   3120
      Width           =   1215
   End
End
Attribute VB_Name = "frmPuntentelling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim OnlyOnce As Boolean

Private Sub cmdAnnuleren_Click()
  'frmMain.Winsock1.Close
  Unload Me
End Sub

Private Sub cmdOk_Click()
  Dim i As Integer
  
  Dim Ret As VbMsgBoxResult
  
  Ret = MsgBox("Deze instellingen kunnen alleen gewijzigd worden als je een nieuw spel start. Wil je nu een nieuw spel starten?", vbQuestion + vbYesNo, "Puntentelling")
  
  If Ret = vbYes Then
  
    'For i = 0 To 2
    '  If frmMain.Winsock(i).State <> sckConnected Then
    '    frmMain.Winsock(i).Close
    '  End If
    'Next i
   
    'frmMain.Winsock1.SendData "1Start"
  '  If optMaxKaarten(0).Value Then
      'MaxAantKaarten = 10
  '    'nRonden = 19
  '  ElseIf optMaxKaarten(1).Value Then
  '    MaxAantKaarten = 6
  '  End If
    'nRonden = MaxAantKaarten * 2 - 1
    'BepaalAantKaartenPerRonde
    
  '  Opties.RondmakenToegestaan = CBool(chkRondmaken.Value)
  '  Opties.IntroevenVerplicht = CBool(chkIntroeven.Value)
    
    Opties.PuntenPerSlag = Val(txtPuntenPerSlag.Text)
    Opties.FoutVoorspeldNulPunten = optFoutVoorspeldPunten(0).Value
    Opties.StrafpuntenPerSlag = Val(txtStrafpuntenPerSlag.Text)
    Opties.NegatieveScores = CBool(chkNegatiefMogelijk.Value)
    
    Me.Hide
  
    frmMain.NieuwSpel False
    
    Unload Me
  End If
End Sub

Private Sub cmdVerbinden_Click()
'** Pre: gebruiker is client
  'On Local Error GoTo cmdVerbindenFout
  
  'WinsockFout = 0

  If frmMain.Netwerk.Initialized Then
'    If cmdVerbinden.Caption = "&Verbinden" Then
'      cmdVerbinden.Caption = "&Verbreken"
'      'StatusBar1.SimpleText = "Bezig met verbinden met de server..."
'      frmMain.Netwerk.ConnectToServer txtServerIP
'    Else
'      cmdVerbinden.Caption = "&Verbinden"
'      frmMain.Netwerk.ClientDisconnect
'      txtServerIP.Enabled = True
'      UpdateStatusbar
'    End If
  Else
    MsgBox "Netwerkfunctionaliteit is uitgeschakeld in deze versie van 10 op en neer.", vbExclamation, "Netwerkfout"
  End If

'  Exit Sub
'
'cmdVerbindenFout:
'  WinsockFout = Err.Number
'  Resume Next
End Sub

Private Sub Form_Load()
  'optServerClient_Click 0
  'txtServerIP.Text = UCase(frmMain.Winsock1.LocalHostName)
  'txtServerIP.ToolTipText = frmMain.Winsock1.LocalIP
  
  optFoutVoorspeldPunten(0).Value = Opties.FoutVoorspeldNulPunten
  optFoutVoorspeldPunten(1).Value = Not Opties.FoutVoorspeldNulPunten
  txtPuntenPerSlag.Text = Opties.PuntenPerSlag
  txtStrafpuntenPerSlag.Text = Opties.StrafpuntenPerSlag
  chkNegatiefMogelijk.Value = Abs(CInt(Opties.NegatieveScores))
  
  optFoutVoorspeldPunten_Click 0 'Zodat e.e.a. op 'disabled' gaat.
End Sub

Private Sub Form_Paint()
  If OnlyOnce = False Then
    OnlyOnce = True
    KiesServerOfClient
  End If
End Sub

Private Sub optFoutVoorspeldPunten_Click(Index As Integer)
  chkNegatiefMogelijk.Enabled = optFoutVoorspeldPunten(1).Value
  If optFoutVoorspeldPunten(0).Enabled Then
    chkNegatiefMogelijk.Value = 1
  End If
  'txtStrafpuntenPerSlag.Enabled = optFoutVoorspeldPunten(1).Value
End Sub

Private Sub optServerClient_Click(Index As Integer)
  KiesServerOfClient
End Sub

Private Sub txtPuntenPerSlag_LostFocus()
  txtPuntenPerSlag = Val(txtPuntenPerSlag.Text)
End Sub

Private Sub txtServerIP_Change()
'  If Len(txtServerIP.Text) = 0 Then
'    cmdVerbinden.Enabled = False
'  Else
'    cmdVerbinden.Enabled = True
'  End If
End Sub

Private Sub txtServerIP_GotFocus()
'  txtServerIP.SelStart = 0
'  txtServerIP.SelLength = Len(txtServerIP.Text)
End Sub

Sub UpdateStatusbar()
'  If optServerClient(0).Value Then 'Server
''    If nVerbonden = 0 Then
''      If WinsockFout = 0 Then
''        StatusBar1.SimpleText = "Niet verbonden -- wacht op andere spelers of klik op 'Start'."
''      Else
''        StatusBar1.SimpleText = "Niet verbonden -- het netwerk functioneert niet goed."
''      End If
''    Else
''      StatusBar1.SimpleText = nVerbonden & " speler" & IIf(nVerbonden = 1, "", "s") & " verbonden -- klik op 'Start' om te beginnen."
''    End If
'  Else
''    If nVerbonden = 0 Then
''      'if frmmain.Winsock(0).State=StateConstants.
''      StatusBar1.SimpleText = "Niet verbonden -- typ een adres en klik op 'Verbinden'."
''    Else 'nVerbonden = 1
''      StatusBar1.SimpleText = "Verbonden -- wacht tot de spelleider het spel start."
''    End If
'  End If
'
End Sub

Sub KiesServerOfClient()
  'On Local Error GoTo optServerClientFout

  'WinsockFout = 0 '# variabele overbodig?
  
'  If optServerClient(0).Value = True Then
'    cmdStart.Visible = True
'    txtServerIP.Text = UCase(frmMain.Winsock1.LocalHostName)
'    txtServerIP.ToolTipText = frmMain.Winsock1.LocalIP
'    txtServerIP.Locked = True
'    lblServerTekst.Enabled = False
'  Else
'    cmdStart.Visible = False
'    txtServerIP.ToolTipText = ""
'    txtServerIP.Locked = False
'    lblServerTekst.Enabled = True
'  End If
  
  'frmMain.Winsock(0).Close
  'frmMain.Winsock(1).Close
  'frmMain.Winsock(2).Close
  
'  If optServerClient(0).Value Then 'Server
'    'frmMain.Winsock(0).Listen
'    'If WinsockFout = 0 Then
'    '  frmMain.Winsock(1).Listen
'    '  frmMain.Winsock(2).Listen
'    'End If
'
'    optMaxKaarten(0).Enabled = True
'    optMaxKaarten(1).Enabled = True
'    'optMaxKaarten(2).Enabled = True
'
'    frmMain.Netwerk.StartServer
'    'nVerbonden = 0
'  Else
'    optMaxKaarten(0).Enabled = False
'    optMaxKaarten(1).Enabled = False
'    'optMaxKaarten(2).Enabled = False
'    'nVerbonden = 0
'  End If
'  'UpdateStatusbar
'
'  cmdStart.Enabled = optServerClient(0).Value
'  txtServerIP.Enabled = optServerClient(1).Value
'  cmdVerbinden.Enabled = optServerClient(1).Value
'  cmdVerbinden.Caption = "&Verbinden"
'
'  Exit Sub
  
'optServerClientFout:
'  WinsockFout = Err.Number
'  Resume Next

End Sub

Private Sub txtStrafpuntenPerSlag_LostFocus()
  txtStrafpuntenPerSlag.Text = Val(txtStrafpuntenPerSlag.Text)
End Sub
