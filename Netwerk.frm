VERSION 5.00
Begin VB.Form frmSpelSpeciaal 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Nieuw spel met opties"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7695
   Icon            =   "Netwerk.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraMaxKaarten 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Opties"
      Height          =   1335
      Left            =   5040
      TabIndex        =   8
      Top             =   240
      Visible         =   0   'False
      Width           =   2535
      Begin VB.OptionButton optMaxKaarten 
         Caption         =   "6 kaarten (kort spel)"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   1815
      End
      Begin VB.OptionButton optMaxKaarten 
         Caption         =   "10 kaarten (standaard)"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Value           =   -1  'True
         Width           =   2055
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Netwerk"
      Height          =   3375
      Left            =   5040
      TabIndex        =   9
      Top             =   2160
      Visible         =   0   'False
      Width           =   2535
      Begin VB.CommandButton cmdVerbinden 
         Caption         =   "&Verbinden"
         Enabled         =   0   'False
         Height          =   285
         Left            =   240
         TabIndex        =   7
         Top             =   1680
         Width           =   975
      End
      Begin VB.OptionButton optServerClient 
         Caption         =   "Server / geen netwerk"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Value           =   -1  'True
         Width           =   2055
      End
      Begin VB.OptionButton optServerClient 
         Caption         =   "Client"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox txtServerIP 
         Height          =   285
         Left            =   240
         TabIndex        =   6
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Label lblNetwerkstatus 
         BorderStyle     =   1  'Fixed Single
         Height          =   615
         Left            =   240
         TabIndex        =   12
         Top             =   2520
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Netwerkstatus:"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label lblServerTekst 
         Caption         =   "Adres van server:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1080
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Spelregels"
      Height          =   4815
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   4815
      Begin VB.CommandButton cmdRegelsVanBoerenbridge 
         Caption         =   "Spelregels van &Boerenbridge instellen"
         Height          =   375
         Left            =   960
         TabIndex        =   25
         Top             =   3960
         Width           =   2895
      End
      Begin VB.CommandButton cmdRegelsVan10openneer 
         Caption         =   "Spelregels van &10 op en neer instellen"
         Height          =   375
         Left            =   960
         TabIndex        =   24
         Top             =   3480
         Width           =   2895
      End
      Begin VB.CheckBox chkRondmaken 
         Caption         =   "Rondmaken is toegestaan"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   2280
         Width           =   2295
      End
      Begin VB.CheckBox chkIntroeven 
         Caption         =   "Introeven is verplicht"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   2640
         Width           =   2415
      End
      Begin VB.CheckBox chkNegatiefMogelijk 
         Caption         =   "Negatieve scores mogelijk"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2040
         TabIndex        =   21
         Top             =   1680
         Width           =   2175
      End
      Begin VB.TextBox txtStrafpuntenPerSlag 
         Height          =   285
         Left            =   2880
         MaxLength       =   1
         TabIndex        =   19
         Text            =   "3"
         Top             =   1320
         Width           =   255
      End
      Begin VB.TextBox txtPuntenPerSlag 
         Height          =   285
         Left            =   2640
         MaxLength       =   1
         TabIndex        =   18
         Text            =   "1"
         Top             =   480
         Width           =   255
      End
      Begin VB.OptionButton optFoutVoorspeldPunten 
         Caption         =   "10 punten -          per slag verschil"
         Height          =   255
         Index           =   1
         Left            =   1680
         TabIndex        =   17
         Top             =   1320
         Width           =   2895
      End
      Begin VB.OptionButton optFoutVoorspeldPunten 
         Caption         =   "0 punten"
         Height          =   255
         Index           =   0
         Left            =   1680
         TabIndex        =   16
         Top             =   1080
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.Line Line3 
         X1              =   240
         X2              =   4560
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Line Line2 
         X1              =   240
         X2              =   4560
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Line Line1 
         X1              =   240
         X2              =   4560
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label Label4 
         Caption         =   "10 punten +         punt(en) per slag"
         Height          =   255
         Left            =   1680
         TabIndex        =   20
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label Label3 
         Caption         =   "Fout voorspeld:"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Goed voorspeld:"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "&Start"
      Default         =   -1  'True
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton cmdAnnuleren 
      Cancel          =   -1  'True
      Caption         =   "&Annuleren"
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   5040
      Width           =   1215
   End
End
Attribute VB_Name = "frmSpelSpeciaal"
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

Private Sub cmdRegelsVan10openneer_Click()
  txtPuntenPerSlag.Text = "1"
  optFoutVoorspeldPunten(0).Value = True
  chkRondmaken.Value = 1
  chkIntroeven.Value = 0
End Sub

Private Sub cmdRegelsVanBoerenbridge_Click()
  txtPuntenPerSlag.Text = "3"
  optFoutVoorspeldPunten(1).Value = True
  txtStrafpuntenPerSlag.Text = "3"
  chkNegatiefMogelijk.Value = 0
  chkRondmaken.Value = 0
  chkIntroeven.Value = 0
End Sub

Private Sub cmdStart_Click()
  Dim i As Integer
  
  'For i = 0 To 2
  '  If frmMain.Winsock(i).State <> sckConnected Then
  '    frmMain.Winsock(i).Close
  '  End If
  'Next i
 
  'frmMain.Winsock1.SendData "1Start"
  If optMaxKaarten(0).Value Then
    MaxAantKaarten = 10
    'nRonden = 19
  ElseIf optMaxKaarten(1).Value Then
    MaxAantKaarten = 6
  End If
  nRonden = MaxAantKaarten * 2 - 1
  BepaalAantKaartenPerRonde
  
  Opties.RondmakenToegestaan = CBool(chkRondmaken.Value)
  Opties.IntroevenVerplicht = CBool(chkIntroeven.Value)
  
  Opties.PuntenPerSlag = Val(txtPuntenPerSlag.Text)
  Opties.FoutVoorspeldNulPunten = optFoutVoorspeldPunten(0).Value
  Opties.StrafpuntenPerSlag = Val(txtStrafpuntenPerSlag.Text)
  Opties.NegatieveScores = CBool(chkNegatiefMogelijk.Value)
  
  Me.Hide

  frmMain.NieuwSpel
End Sub

Private Sub cmdVerbinden_Click()
'** Pre: gebruiker is client
  'On Local Error GoTo cmdVerbindenFout
  
  'WinsockFout = 0

  If frmMain.Netwerk.Initialized Then
    If cmdVerbinden.Caption = "&Verbinden" Then
      cmdVerbinden.Caption = "&Verbreken"
      'StatusBar1.SimpleText = "Bezig met verbinden met de server..."
      frmMain.Netwerk.ConnectToServer txtServerIP
    Else
      cmdVerbinden.Caption = "&Verbinden"
      frmMain.Netwerk.ClientDisconnect
      txtServerIP.Enabled = True
      UpdateStatusbar
    End If
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
  chkRondmaken.Value = Abs(CInt(Opties.RondmakenToegestaan))
  chkIntroeven.Value = Abs(CInt(Opties.IntroevenVerplicht))
  
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
  txtStrafpuntenPerSlag.Enabled = optFoutVoorspeldPunten(1).Value
End Sub

Private Sub optServerClient_Click(Index As Integer)
  KiesServerOfClient
End Sub

Private Sub txtPuntenPerSlag_LostFocus()
  txtPuntenPerSlag = Val(txtPuntenPerSlag.Text)
End Sub

Private Sub txtServerIP_Change()
  If Len(txtServerIP.Text) = 0 Then
    cmdVerbinden.Enabled = False
  Else
    cmdVerbinden.Enabled = True
  End If
End Sub

Private Sub txtServerIP_GotFocus()
  txtServerIP.SelStart = 0
  txtServerIP.SelLength = Len(txtServerIP.Text)
End Sub

Sub UpdateStatusbar()
  If optServerClient(0).Value Then 'Server
'    If nVerbonden = 0 Then
'      If WinsockFout = 0 Then
'        StatusBar1.SimpleText = "Niet verbonden -- wacht op andere spelers of klik op 'Start'."
'      Else
'        StatusBar1.SimpleText = "Niet verbonden -- het netwerk functioneert niet goed."
'      End If
'    Else
'      StatusBar1.SimpleText = nVerbonden & " speler" & IIf(nVerbonden = 1, "", "s") & " verbonden -- klik op 'Start' om te beginnen."
'    End If
  Else
'    If nVerbonden = 0 Then
'      'if frmmain.Winsock(0).State=StateConstants.
'      StatusBar1.SimpleText = "Niet verbonden -- typ een adres en klik op 'Verbinden'."
'    Else 'nVerbonden = 1
'      StatusBar1.SimpleText = "Verbonden -- wacht tot de spelleider het spel start."
'    End If
  End If

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
  
  If optServerClient(0).Value Then 'Server
    'frmMain.Winsock(0).Listen
    'If WinsockFout = 0 Then
    '  frmMain.Winsock(1).Listen
    '  frmMain.Winsock(2).Listen
    'End If
   
    optMaxKaarten(0).Enabled = True
    optMaxKaarten(1).Enabled = True
    'optMaxKaarten(2).Enabled = True
    
    frmMain.Netwerk.StartServer
    'nVerbonden = 0
  Else
    optMaxKaarten(0).Enabled = False
    optMaxKaarten(1).Enabled = False
    'optMaxKaarten(2).Enabled = False
    'nVerbonden = 0
  End If
  'UpdateStatusbar
  
  cmdStart.Enabled = optServerClient(0).Value
  txtServerIP.Enabled = optServerClient(1).Value
  cmdVerbinden.Enabled = optServerClient(1).Value
  cmdVerbinden.Caption = "&Verbinden"
  
  Exit Sub
  
'optServerClientFout:
'  WinsockFout = Err.Number
'  Resume Next

End Sub

Private Sub txtStrafpuntenPerSlag_LostFocus()
  txtStrafpuntenPerSlag = Val(txtStrafpuntenPerSlag.Text)
End Sub
