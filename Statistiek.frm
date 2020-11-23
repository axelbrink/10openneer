VERSION 5.00
Begin VB.Form frmStatistiek 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Statistiek"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
   HelpContextID   =   302
   Icon            =   "Statistiek.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdInfo 
      Height          =   375
      Left            =   120
      MaskColor       =   &H00FF00FF&
      Picture         =   "Statistiek.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   48
      ToolTipText     =   "Klik hier voor extra uitleg"
      Top             =   3960
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton cmdConclusies 
      Caption         =   "&Conclusies..."
      Height          =   375
      Left            =   2160
      TabIndex        =   47
      ToolTipText     =   "Klik er maar op, het is leuk"
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Wissen..."
      Height          =   375
      Left            =   3480
      TabIndex        =   37
      ToolTipText     =   "Statistiek wissen"
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Frame Frame4 
      Caption         =   "Voorspellen"
      Height          =   1575
      Left            =   120
      TabIndex        =   32
      Top             =   2280
      Width           =   2895
      Begin VB.Label lblRondgemaaktTekst 
         Caption         =   "Rondgemaakt:"
         Height          =   255
         Left            =   240
         TabIndex        =   44
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label lblRondgemaakt 
         Alignment       =   2  'Center
         Caption         =   "0"
         Height          =   255
         Left            =   2160
         TabIndex        =   43
         Top             =   1080
         Width           =   480
      End
      Begin VB.Label lblGemVoorspeldTekst 
         Caption         =   "Gemiddeld voorspeld:"
         Height          =   255
         Left            =   240
         TabIndex        =   42
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label lblGemVoorspeld 
         Alignment       =   2  'Center
         Caption         =   "0"
         Height          =   255
         Left            =   2160
         TabIndex        =   41
         Top             =   360
         Width           =   480
      End
      Begin VB.Label lblFoutAfwijking 
         Alignment       =   2  'Center
         Caption         =   "0"
         Height          =   255
         Left            =   2160
         TabIndex        =   36
         Top             =   840
         Width           =   480
      End
      Begin VB.Label lblFoutAfwijkingTekst 
         Caption         =   "Afwijking bij fout:"
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label lblAantGoed 
         Alignment       =   2  'Center
         Caption         =   "0"
         Height          =   255
         Left            =   2160
         TabIndex        =   34
         Top             =   600
         Width           =   480
      End
      Begin VB.Label lblAantGoedTekst 
         Caption         =   "Goed voorspeld:"
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   600
         Width           =   1935
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Score"
      Height          =   2055
      Left            =   120
      TabIndex        =   19
      Top             =   120
      Width           =   2895
      Begin VB.Label lblFreqProc 
         Alignment       =   2  'Center
         Caption         =   "0%"
         Height          =   255
         Index           =   3
         Left            =   2160
         TabIndex        =   31
         Top             =   1560
         Width           =   480
      End
      Begin VB.Label lblFreqProc 
         Alignment       =   2  'Center
         Caption         =   "0%"
         Height          =   255
         Index           =   2
         Left            =   2160
         TabIndex        =   30
         Top             =   1320
         Width           =   480
      End
      Begin VB.Label lblFreqProc 
         Alignment       =   2  'Center
         Caption         =   "0%"
         Height          =   255
         Index           =   1
         Left            =   2160
         TabIndex        =   29
         Top             =   1080
         Width           =   480
      End
      Begin VB.Label lblFreqProc 
         Alignment       =   2  'Center
         Caption         =   "0%"
         Height          =   255
         Index           =   0
         Left            =   2160
         TabIndex        =   28
         Top             =   840
         Width           =   480
      End
      Begin VB.Label lblPCGem 
         Alignment       =   2  'Center
         Caption         =   "0"
         Height          =   255
         Left            =   2160
         TabIndex        =   27
         Top             =   600
         Width           =   480
      End
      Begin VB.Label lblGem 
         Alignment       =   2  'Center
         Caption         =   "0"
         Height          =   255
         Left            =   2160
         TabIndex        =   26
         Top             =   360
         Width           =   480
      End
      Begin VB.Label lblGemTekst 
         Caption         =   "Gemiddelde:"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label lblFreqProcTekst 
         Caption         =   "Laatste plaats:"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   24
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label lblFreqProcTekst 
         Caption         =   "Derde plaats:"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   23
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label lblFreqProcTekst 
         Caption         =   "Tweede plaats:"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   22
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label lblFreqProcTekst 
         Caption         =   "Eerste plaats:"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   21
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label lblPCGemTekst 
         Caption         =   "PC-gemiddelde:"
         Height          =   255
         Left            =   360
         TabIndex        =   20
         Top             =   600
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tijd"
      Height          =   2055
      Left            =   3120
      TabIndex        =   8
      Top             =   120
      Width           =   2895
      Begin VB.Label Label14 
         Caption         =   "u"
         Height          =   255
         Left            =   2640
         TabIndex        =   40
         Top             =   840
         Width           =   135
      End
      Begin VB.Label Label12 
         Caption         =   "u"
         Height          =   255
         Left            =   2640
         TabIndex        =   39
         Top             =   600
         Width           =   135
      End
      Begin VB.Label Label11 
         Caption         =   "u"
         Height          =   255
         Left            =   2640
         TabIndex        =   38
         Top             =   360
         Width           =   135
      End
      Begin VB.Label lblDagenSindsEersteKeer 
         Alignment       =   2  'Center
         Caption         =   "0"
         Height          =   255
         Left            =   2160
         TabIndex        =   18
         Top             =   1080
         Width           =   480
      End
      Begin VB.Label Label21 
         Caption         =   "Dagen sinds eerste sessie:"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label lblDagenGespeeld 
         Alignment       =   2  'Center
         Caption         =   "0"
         Height          =   255
         Left            =   2160
         TabIndex        =   16
         Top             =   1320
         Width           =   480
      End
      Begin VB.Label Label19 
         Caption         =   "Aantal dagen gespeeld:"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label lblDuurPerSessie 
         Alignment       =   2  'Center
         Caption         =   "0:00"
         Height          =   255
         Left            =   2160
         TabIndex        =   14
         Top             =   840
         Width           =   480
      End
      Begin VB.Label Label17 
         Caption         =   "Gem. duur per sessie:"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label lblSessieDuur 
         Alignment       =   2  'Center
         Caption         =   "0:00"
         Height          =   255
         Left            =   2160
         TabIndex        =   12
         Top             =   360
         Width           =   480
      End
      Begin VB.Label Label15 
         Caption         =   "Duur van deze sessie:"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label lblTotaalDuur 
         Alignment       =   2  'Center
         Caption         =   "0:00"
         Height          =   255
         Left            =   2160
         TabIndex        =   10
         Top             =   600
         Width           =   480
      End
      Begin VB.Label Label13 
         Caption         =   "Totaal gespeeld:"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   600
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Spellen"
      Height          =   1575
      Left            =   3120
      TabIndex        =   1
      Top             =   2280
      Width           =   2895
      Begin VB.Label lblSpellenPerDag 
         Alignment       =   2  'Center
         Caption         =   "0"
         Height          =   255
         Left            =   2160
         TabIndex        =   45
         Top             =   1080
         Width           =   480
      End
      Begin VB.Label Label23 
         Caption         =   "Spellen per dag:"
         Height          =   255
         Left            =   240
         TabIndex        =   46
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label lblVoltooidProc 
         Alignment       =   2  'Center
         Caption         =   "0%"
         Height          =   255
         Left            =   2040
         TabIndex        =   7
         Top             =   840
         Width           =   720
      End
      Begin VB.Label lblBegonnen 
         Alignment       =   2  'Center
         Caption         =   "0"
         Height          =   255
         Left            =   2160
         TabIndex        =   6
         Top             =   600
         Width           =   480
      End
      Begin VB.Label Label8 
         Caption         =   "Spellen begonnen:"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label7 
         Caption         =   "Waarvan voltooid:"
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Aantal sessies:"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label lblSessies 
         Alignment       =   2  'Center
         Caption         =   "0"
         Height          =   255
         Left            =   2160
         TabIndex        =   2
         Top             =   360
         Width           =   480
      End
   End
   Begin VB.CommandButton cmdOk 
      Cancel          =   -1  'True
      Caption         =   "&Sluiten"
      Default         =   -1  'True
      Height          =   375
      Left            =   4800
      TabIndex        =   0
      Top             =   3960
      Width           =   1215
   End
End
Attribute VB_Name = "frmStatistiek"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdConclusies_Click()
  Dim Msg As String
  Dim DagenPerDag As Single
  Dim SpellenBegonnenPerDag As Single
  Dim SpellenVoltooidPerc As Single
  
  Msg = "Op grond van de huidige statistische gegevens, bijgehouden door Tineke, kan ze het volgende over je zeggen:" & vbCrLf & vbCrLf
  
  With Statistiek
    DagenPerDag = .DagenGespeeld / (DateDiff("d", .InstallDate, Now) + 1)
    SpellenBegonnenPerDag = .SpellenBegonnen / .DagenGespeeld
    SpellenVoltooidPerc = .SpellenVoltooid / .SpellenBegonnen
    If .SpellenVoltooid = 0 Then
      Msg = Msg & "Op dit moment is het nog niet mogelijk om conclusies te trekken over je kaartgedrag; je hebt nog geen enkel spel voltooid. "
    Else
      If SpelIs10openneer() Then
        Select Case .Gemiddelde
          Case 0 To 100
            Msg = Msg & "Je hebt duidelijk nog geen kaas gegeten van het spel. Meer oefenen is het devies. "
          Case 101 To 120
            Msg = Msg & "Je prestaties zijn matig. "
          Case 121 To 130
            Msg = Msg & "Je speelt redelijk 10 op en neer. "
          Case 131 To 140
            Msg = Msg & "Je speelt vrij aardig 10 op en neer. "
          Case Is >= 141
            Msg = Msg & "Je bent goed. "
        End Select
      End If
      
      If .FoutVoorspeld > 0 Then
        If .FoutSaldo / .FoutVoorspeld < -0.3 Then
          Msg = Msg & "Wees wat minder optimistisch bij het voorspellen. "
        ElseIf .FoutSaldo / .FoutVoorspeld > 0.3 Then
          Msg = Msg & "Wees wat minder pessimistisch bij het voorspellen. "
        End If
      End If
      
      If .SpellenVoltooid / .SpellenBegonnen < 0.5 Then
        Msg = Msg & "Veel spellen maak je niet af. "
      End If
      
      Select Case DagenPerDag
        Case 0 To 0.1
          Select Case SpellenBegonnenPerDag
            Case 1 To 2
              Msg = Msg & "Je speelt het spel zeer weinig. Je vindt het duidelijk niet leuk. Misschien wordt het tijd om 10 op en neer van de computer te verwijderen. " & vbCrLf
            Case Is > 2
              Msg = Msg & "Je speelt af en toe een aantal spelletjes. Je hebt een druk bestaan. Doe het wat rustiger aan en neem het leven niet te serieus. " & vbCrLf
          End Select
        Case 0.1 To 1
          Select Case SpellenBegonnenPerDag
            Case 1 To 2
              Msg = Msg & "Je speelt vrij vaak een spelletje 10 op en neer. Het lijkt erop dat je het leuk vindt. " & vbCrLf
            Case Is > 2
              Msg = Msg & "Je bent fan van 10 op en neer. Je vindt het zelfs zo leuk, dat je niet meer achter de computer weg te slaan bent. Een goede reden om eens een e-mailtje te schrijven aan de auteur! " & vbCrLf
          End Select
      End Select
    End If
  End With

  MsgBox Msg, vbInformation, "Observatie van Tineke"
End Sub

Private Sub cmdInfo_Click()
  MsgBox "Houd de muis stil boven een gegeven waar je meer over wilt weten. " & _
         "Dit kan alleen in de vakken Score en Voorspellen.", vbInformation, "Statistiek"
End Sub

Private Sub cmdOk_Click()
  Unload Me
End Sub

Private Sub cmdReset_Click()
  Dim Ret As Long
  
  Ret = MsgBox("Weet je zeker dat je alle statistiekgegevens wilt verwijderen?", _
        vbYesNo + vbQuestion, "Statistiek")
  
  If Ret = vbYes Then
'    IniDel "Statistiek", "DagenGespeeld" 'Kan dit
'    'IniSet "Statistiek", "DagenGespeeld", 1
    With Statistiek
      .Gemiddelde = 0
      .PCGemiddelde = 0
      .RangFreq(1) = 0
      .RangFreq(2) = 0
      .RangFreq(3) = 0
      .RangFreq(4) = 0
    
      .Sessies = 1
      .SpellenBegonnen = 1
      .SpellenVoltooid = 0
      
      .TotaalVoorspeld = 0
      .TotaalVoorspeldMax = 0
      .GoedVoorspeld = 0
      .FoutVoorspeld = 0
      .FoutSaldo = 0
      .Rondgemaakt = 0
      .NietRondgemaakt = 0
      
      .TotaalDuur = 0
      .LaatstGespeeld = Now
      .DagenGespeeld = 1
      .InstallDate = Now
    End With

    OpslaanStatistiek
    IniSave
    
    StartMoment = Now 'Wel of niet?

    Tonen
  End If
End Sub

Private Sub Form_Load()
  Tonen
End Sub
Sub Tonen()
  Dim FreqProc As Single
  Dim Sessieduur As Integer
  Dim TotaalduurNu As Long
  Dim DuurPerSessie As Integer
  Dim VerlopenDagen As Integer
  Dim i As Integer
  
  '** Score **
  
  If Statistiek.SpellenVoltooid = 0 Then
    lblGem.Caption = "-"
    lblPCGem.Caption = "-"
  Else
    lblGem.Caption = CInt(Statistiek.Gemiddelde)
    lblPCGem.Caption = CInt(Statistiek.PCGemiddelde)
  End If
  lblGem.ToolTipText = "150 is goed, 120 is redelijk, 90 is slecht (bij 10 op en neer)"
  lblGemTekst.ToolTipText = lblGem.ToolTipText
  lblPCGem.ToolTipText = "Ter vergelijking"
  lblPCGemTekst.ToolTipText = lblPCGem.ToolTipText
  
  For i = 1 To 4
    If Statistiek.SpellenVoltooid = 0 Then
      lblFreqProc(i - 1).Caption = "0%"
    Else
      FreqProc = Statistiek.RangFreq(i) / Statistiek.SpellenVoltooid
      lblFreqProc(i - 1).Caption = Format(FreqProc, "##0%")
    End If
  Next i
  lblFreqProc(0).ToolTipText = "100% = alles gewonnen"
  lblFreqProcTekst(0).ToolTipText = lblFreqProc(0).ToolTipText
  lblFreqProc(1).ToolTipText = "100% = altijd tweede geworden"
  lblFreqProcTekst(1).ToolTipText = lblFreqProc(1).ToolTipText
  lblFreqProc(2).ToolTipText = "100% = altijd derde geworden"
  lblFreqProcTekst(2).ToolTipText = lblFreqProc(2).ToolTipText
  lblFreqProc(3).ToolTipText = "100% = alles verloren"
  lblFreqProcTekst(3).ToolTipText = lblFreqProc(3).ToolTipText
  
  '** Voorspellen **
  
  If Statistiek.TotaalVoorspeldMax = 0 Then
    lblGemVoorspeld.Caption = "-"
  Else
    lblGemVoorspeld.Caption = Format(Statistiek.TotaalVoorspeld / Statistiek.TotaalVoorspeldMax, "##0%")
  End If
  lblGemVoorspeld.ToolTipText = "0% = steeds 0 voorspellen; 100% is steeds het maximum voorspellen. 25% is normaal."
  lblGemVoorspeldTekst.ToolTipText = lblGemVoorspeld.ToolTipText
  
  If Statistiek.GoedVoorspeld + Statistiek.FoutVoorspeld > 0 Then
    lblAantGoed.Caption = Format(Statistiek.GoedVoorspeld / (Statistiek.GoedVoorspeld + Statistiek.FoutVoorspeld), "##0%")
  Else
    lblAantGoed.Caption = "-"
  End If
  lblAantGoed.ToolTipText = "Hoe hoger, hoe beter"
  lblAantGoedTekst.ToolTipText = lblAantGoed.ToolTipText
  
  If Statistiek.FoutVoorspeld > 0 Then
    lblFoutAfwijking.Caption = Format(Statistiek.FoutSaldo / Statistiek.FoutVoorspeld, "+#0.00;-#,0.00;0")
    If Statistiek.FoutSaldo = 0 Then
      lblFoutAfwijking.ToolTipText = "Nul: gemiddeld evenveel gehaald als voorspeld"
    ElseIf Statistiek.FoutSaldo < 0 Then
      lblFoutAfwijking.ToolTipText = "Negatief: minder gehaald dan voorspeld; te hoog voorspeld"
    ElseIf Statistiek.FoutSaldo > 0 Then
      lblFoutAfwijking.ToolTipText = "Positief: meer gehaald dan voorspeld; te laag voorspeld"
    End If
  Else
    lblFoutAfwijking.Caption = "-"
    lblFoutAfwijking.ToolTipText = "Voorspel je te veel of te weinig"
  End If
  lblFoutAfwijkingTekst.ToolTipText = lblFoutAfwijking.ToolTipText
  
  If Statistiek.Rondgemaakt + Statistiek.NietRondgemaakt = 0 Then
    lblRondgemaakt = "-"
  Else
    lblRondgemaakt = Format(Statistiek.Rondgemaakt / (Statistiek.Rondgemaakt + Statistiek.NietRondgemaakt), "##0%")
  End If
  lblRondgemaakt.ToolTipText = "Lager is moeilijker, spannender en leuker."
  lblRondgemaaktTekst.ToolTipText = lblRondgemaakt.ToolTipText
  
  '** Tijd **
  
  Sessieduur = Int(DateDiff("n", StartMoment, Now))
  lblSessieDuur.Caption = Sessieduur \ 60 & ":" & Format(Sessieduur Mod 60, "00")
  
  TotaalduurNu = Statistiek.TotaalDuur + Sessieduur
  lblTotaalDuur.Caption = TotaalduurNu \ 60 & ":" & Format(TotaalduurNu Mod 60, "00")
  
  DuurPerSessie = TotaalduurNu \ Statistiek.Sessies
  lblDuurPerSessie = DuurPerSessie \ 60 & ":" & Format(DuurPerSessie Mod 60, "00")

  VerlopenDagen = DateDiff("d", Statistiek.InstallDate, Now) + 1
  lblDagenSindsEersteKeer.Caption = VerlopenDagen
  
  lblDagenGespeeld.Caption = Statistiek.DagenGespeeld
  
  '** Spellen **
  
  lblSessies.Caption = Statistiek.Sessies
  lblBegonnen.Caption = Statistiek.SpellenBegonnen
  
  If Statistiek.SpellenBegonnen = 0 Then
    lblVoltooidProc.Caption = "-"
  Else
    lblVoltooidProc.Caption = Statistiek.SpellenVoltooid & " (" & Format(Statistiek.SpellenVoltooid / Statistiek.SpellenBegonnen, "##0%") & ")"
  End If

  lblSpellenPerDag.Caption = Format(Statistiek.SpellenBegonnen / VerlopenDagen, "####0.0")

End Sub
