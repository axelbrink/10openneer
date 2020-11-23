VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsScore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type HiScoreItemType
  Datum As String
  GewonnenDoor As Integer
  Naam As String
  NietRondmaken As Boolean 'Dit wordt niet meer getoond in picScore
  Score As Integer
End Type

Dim HiScores(1 To 15) As HiScoreItemType
Public IemandInHiScore As Boolean '# Matig
Public HiscoreGelezen As Boolean  '# Dit ook
Dim WatInBeeld As ScoreblokType
Dim cMagInHiscore As Boolean
'Dim cNScoreRondes As Integer
Dim cToonVoorspelling As Boolean
Dim cRegelHoogte As Integer
Dim cEersteRegelY As Integer

Public Sub TestHiScore()
  Dim HiScoreNummer As Integer
  Dim NummerTemp As Integer
  Dim SpelerNr As Integer
  Dim HoogsteScore As Integer
  
  If Not HiscoreGelezen Then
    InlezenHiScore
  End If
  For SpelerNr = 1 To 4
    For HiScoreNummer = 1 To 15
      If (Spelers(SpelerNr).TotaalScore > HiScores(HiScoreNummer).Score) And _
          ((SpelerNr = 1 And cMagInHiscore) Or ((SpelerNr > 1) And Opties.ComputersInHighScore)) Then
        IemandInHiScore = True
        If HiScores(15).GewonnenDoor > 0 Then
          Spelers(HiScores(15).GewonnenDoor).HiScorePositie = 0
        End If
        For NummerTemp = 14 To HiScoreNummer Step -1
          HiScores(NummerTemp + 1) = HiScores(NummerTemp)
          If HiScores(NummerTemp).GewonnenDoor > 0 Then
            Spelers(HiScores(NummerTemp).GewonnenDoor).HiScorePositie = NummerTemp + 1
          End If
        Next NummerTemp
        With HiScores(HiScoreNummer)
          .Naam = Spelers(SpelerNr).Naam
          If Spelers(SpelerNr).Controller = ControllerComputer Then
            .Naam = .Naam & " (comp)"
          End If
          .Datum = Format(Now, "mm-yy")
          .Score = Spelers(SpelerNr).TotaalScore
          .GewonnenDoor = SpelerNr
          .NietRondmaken = Not Opties.RondmakenToegestaan
          Spelers(SpelerNr).HiScorePositie = HiScoreNummer
        End With
        Exit For
      End If
    Next HiScoreNummer
  Next SpelerNr
End Sub

Public Sub ToonHiScore()
  Dim HiScoreNummer As Integer
  Dim StartY As Integer
  Dim StartX As Integer
  Dim RegelKleur As Long
  
  StartX = 4
  StartY = 5
 
  WatInBeeld = sblHighscore
  
  If Not HiscoreGelezen Then
    InlezenHiScore
  End If
  
  With frmMain
    .picScore.Cls
    .picScore.FontStrikethru = False
 
    .picScore.CurrentX = StartX
    .picScore.CurrentY = StartY
    .picScore.FontBold = True
    SchaduwPrint .picScore, "Highscore", Voorkleur, 0
    .picScore.FontBold = False
   
    For HiScoreNummer = 1 To 15
      .picScore.ForeColor = Voorkleur
      If HiScores(HiScoreNummer).GewonnenDoor > 0 Then
        RegelKleur = vbWhite
      Else
        RegelKleur = Voorkleur
        '.picScore.ForeColor = vbWhite
      End If
    
      .picScore.CurrentX = StartX
      .picScore.CurrentY = HiScoreNummer * 16 + StartY
      SchaduwPrint .picScore, CStr(HiScoreNummer), RegelKleur, 0
    
      .picScore.CurrentX = StartX + 18
      .picScore.CurrentY = HiScoreNummer * 16 + StartY
      SchaduwPrint .picScore, HiScores(HiScoreNummer).Naam, RegelKleur, 0
    
      .picScore.CurrentX = StartX + 101
      .picScore.CurrentY = HiScoreNummer * 16 + StartY
      SchaduwPrint .picScore, HiScores(HiScoreNummer).Datum, RegelKleur, 0
    
      .picScore.CurrentX = StartX + 135
      .picScore.CurrentY = HiScoreNummer * 16 + StartY
      SchaduwPrint .picScore, CStr(HiScores(HiScoreNummer).Score), RegelKleur, 0
      '& IIf(HiScores(HiScoreNummer).NietRondmaken, " NR", ""), RegelKleur, 0
    Next HiScoreNummer
    .picScore.ToolTipText = "Highscore"
    .mnuScoreblok.Caption = "Toon score&blok"
    'frmMenus.mnuToonScoreblok.Caption = "Toon score&blok"
  End With
End Sub

Public Sub ToonRanglijst()
  'Dim Applaus As Integer
  Dim SpelerNr As Integer
  Dim RangNrNu As Integer
  'Dim Tekst As String
 
  If Spelers(1).Rang = 1 Then
    frmMain.imgGewonnen.Picture = LoadResPicture(300, vbResBitmap) 'Vrouw normaal
    frmMain.imgGewonnen.Visible = True
  End If
  
  frmMain.StatusBar.SimpleText = "Het spel is afgelopen." '#
  
  For SpelerNr = 1 To 4
    frmMain.lblRanglijst(SpelerNr - 1).Caption = ""
    frmMain.lblRanglijstScore(SpelerNr - 1).Caption = ""
  Next SpelerNr
  
  For SpelerNr = 1 To 4
    RangNrNu = Spelers(SpelerNr).Rang
    If frmMain.lblRanglijst(RangNrNu - 1).Caption = "" Then
      frmMain.lblRanglijst(RangNrNu - 1).Caption = Spelers(SpelerNr).Naam
      frmMain.lblRanglijstScore(RangNrNu - 1).Caption = Spelers(SpelerNr).TotaalScore
    Else
      frmMain.lblRanglijst(RangNrNu - 1).Caption = frmMain.lblRanglijst(RangNrNu - 1).Caption & ", " & Spelers(SpelerNr).Naam
    End If
  Next SpelerNr
    
  frmMain.fraSpelen.Visible = False
  frmMain.fraEinde.Visible = True
  frmMain.cmdNieuwSpel.SetFocus
  
  'Tineke.Zeg Tekst
End Sub

Public Sub ToonScore()
  Dim SpelerNr As Integer
  Dim ScoreRonde As Integer
  Dim i As Integer
  Dim Kleur As Long
  'Dim SpelersVoorspeld() As Boolean
 
  WatInBeeld = sblScore
  
  'SpelersVoorspeld = BepaalSpelersVoorspeld()
  
  With frmMain
    .picScore.Cls
  
    .picScore.Line (0, cRegelHoogte * Ronde + 6)-(.picScore.Width, cRegelHoogte * (Ronde + 1) + 5), vbGreen, BF
    
    .picScore.Line (0, 16)-(.picScore.Width, 16), 0
    .picScore.Line (16, 16)-(16, .picScore.ScaleHeight - 16), 0
    .picScore.Line (0, .picScore.ScaleHeight - 16)-(.picScore.Width, .picScore.ScaleHeight - 16), 0
    
    .picScore.FontStrikethru = False
    
    'Aantal kaarten voor elke ronde afdrukken
    For ScoreRonde = 1 To nRonden
      If AantKaartenRonde(ScoreRonde) >= 10 Then
        If ScoreRonde = Ronde Then
          ScorePrint CStr(AantKaartenRonde(ScoreRonde)), ScoreRonde, 0, QBColor(14)
        Else
          ScorePrint CStr(AantKaartenRonde(ScoreRonde)), ScoreRonde, 0, Voorkleur
        End If
      Else
        If ScoreRonde = Ronde Then
          ScorePrint " " & CStr(AantKaartenRonde(ScoreRonde)), ScoreRonde, 0, QBColor(14)
        Else
          ScorePrint " " & CStr(AantKaartenRonde(ScoreRonde)), ScoreRonde, 0, Voorkleur
        End If
      End If
    Next ScoreRonde
    
    For SpelerNr = 1 To 4
      .picScore.FontStrikethru = False
      .picScore.CurrentX = ((SpelerNr - 1) * 40 + 25) - .picScore.TextWidth(Spelers(SpelerNr).Naam) / 2 + .picScore.TextWidth("10") / 2
      .picScore.CurrentY = 1
      If Spelers(SpelerNr).Rang = 1 And (Opties.MeteenOptellen Or (Ronde >= nRonden And Not cToonVoorspelling)) Then
        Kleur = vbWhite
      Else
        Kleur = Voorkleur
      End If
      SchaduwPrint .picScore, Spelers(SpelerNr).Naam, Kleur, 0
      If Opties.MeteenOptellen Or (Ronde >= nRonden And Spelers(1).AantKaarten = 0) Then
        .picScore.FontStrikethru = False
        ScorePrint CStr(Spelers(SpelerNr).TotaalScore), nRonden + 1, SpelerNr, Kleur
      End If
      
      Kleur = Voorkleur
      For ScoreRonde = 1 To recRonde 'Ronde '19
        If ScoreRonde = recRonde Then 'And cToonVoorspelling Then
'        If ScoreRonde = recAantRonden And cToonVoorspelling Then
          If Spelers(SpelerNr).Voorspelling > -1 Then
'          If SpelersVoorspeld(SpelerNr) Then
            .picScore.FontStrikethru = False
'            ScorePrint CStr(Spelers(SpelerNr).Voorspelling), ScoreRonde, SpelerNr, QBColor(14)  'Kleur
            ScorePrint CStr(recVoorspellingen(ScoreRonde, SpelerNr)), ScoreRonde, SpelerNr, QBColor(14) 'Kleur
          End If
        Else
          'If Spelers(SpelerNr).Score(ScoreRonde) > -1 Then
          '.picScore.FontStrikethru = Spelers(SpelerNr).ScoreFout(ScoreRonde)
          'If ScoreRonde <= recAantRonden Then 'Or recSlagNr > 0 Then
            .picScore.FontStrikethru = (Not recVoorspellingenGoed(ScoreRonde, SpelerNr)) And (Opties.FoutVoorspeldNulPunten And Opties.StrafpuntenPerSlag = 0)
            'ScorePrint CStr(Spelers(SpelerNr).Score(ScoreRonde)), ScoreRonde, SpelerNr, Kleur
            ScorePrint CStr(recVoorspellingen(ScoreRonde, SpelerNr)), ScoreRonde, SpelerNr, Kleur
          'End If
        End If
      Next ScoreRonde
      'If cNScoreRondes < nRonden Then
'      If cToonVoorspelling Then
'      End If
    Next SpelerNr
    .picScore.ForeColor = Voorkleur
    .picScore.ToolTipText = "Scoreblok"
    .mnuScoreblok.Caption = "Toon &highscore"
    'frmMenus.mnuToonScoreblok.Caption = "Toon &highscore"
  End With
End Sub

Public Sub Update()
  If WatInBeeld = sblScore Then
    ToonScore
  End If
End Sub

Private Sub ScorePrint(Tekst As String, ByVal Regel As Integer, ByVal Kolom As Integer, Optional ByVal Kleur As Long = &H1000000)
  'Dim HoogteFactor As Single
  
  If Kleur = &H1000000 Then 'Maximale rgb-kleur + 1
    Kleur = Voorkleur
  End If
  
  With frmMain
    'HoogteFactor = (.picScore.ScaleHeight - 35) / nRonden
  
    If Kolom = 0 Then
      .picScore.CurrentX = 1
    Else
      If Len(Tekst) <= 3 Then
        Tekst = Space(3 - Len(Tekst)) & Tekst & " "
      End If
      .picScore.CurrentX = 40 * Kolom - .picScore.TextWidth(Tekst)
    End If
  
    Select Case Regel
      Case 0
        .picScore.CurrentY = 0
      Case nRonden + 1
        .picScore.CurrentY = .picScore.ScaleHeight - 15 'HoogteFactor * Regel + 8
      Case Else
        .picScore.CurrentY = cRegelHoogte * Regel + cEersteRegelY
    End Select
    
    SchaduwPrint .picScore, Tekst, Kleur, 0
  End With
End Sub

Private Sub SchaduwPrint(Voorwerp As Control, Tekst As String, ByVal TekstKleur As Long, Optional ByVal SchaduwKleur As Long = 0)
  Dim X As Integer
  Dim Y As Integer
  'Dim Kleur As Long
 
  X = Voorwerp.CurrentX
  Y = Voorwerp.CurrentY
  'Kleur = TekstKleur 'Voorwerp.ForeColor
 
  Voorwerp.CurrentX = X + 1
  Voorwerp.CurrentY = Y + 1
  Voorwerp.ForeColor = SchaduwKleur  '0
  Voorwerp.Print Tekst
  Voorwerp.CurrentX = X
  Voorwerp.CurrentY = Y
  Voorwerp.ForeColor = TekstKleur
  Voorwerp.Print Tekst

End Sub

Public Sub InlezenHiScore()
  Dim Regel As String
  Dim Bestand As String
  Dim Waarde As String
  Dim ByteNr As Integer
  Dim Nummer As Integer  'FoutCode,
 
  Bestand = DirPlusBestand(CStr(App.Path), "hiscore.dat")
  If Len(Dir(Bestand)) Then
    Open Bestand For Input As #1
    Nummer = 1
    For Nummer = 1 To 15
      If EOF(1) Then
        HiScores(Nummer).Naam = "10 op en neer"
        HiScores(Nummer).Datum = Format(Now, "mm-yy")
        HiScores(Nummer).Score = 10 * (-Nummer + 16)
        HiScores(Nummer).NietRondmaken = False
        HiScores(Nummer).GewonnenDoor = 0
      Else
        Line Input #1, HiScores(Nummer).Naam
        Line Input #1, HiScores(Nummer).Datum
        Line Input #1, Waarde
        If Right(Waarde, 1) = "*" Then
          HiScores(Nummer).NietRondmaken = True
          HiScores(Nummer).Score = CInt(Left(Waarde, Len(Waarde) - 1))
        Else
          HiScores(Nummer).NietRondmaken = False
          HiScores(Nummer).Score = CInt(Waarde)
        End If
      End If
    Next Nummer
  Else
    For Nummer = 1 To 15
      HiScores(Nummer).Naam = "10 op en neer"
      HiScores(Nummer).Datum = Format(Now, "mm-yy")
      HiScores(Nummer).Score = 10 * (-Nummer + 16)
      HiScores(Nummer).NietRondmaken = False
      HiScores(Nummer).GewonnenDoor = 0
    Next Nummer
  End If
 
  HiscoreGelezen = True
 
  Close #1
 
End Sub

Public Sub OpslaanHiScore()
  Dim Bestand As String
  Dim Nummer As Integer
  Dim ScoreString As String
 
  Bestand = DirPlusBestand(CStr(App.Path), "hiscore.dat")

  Open Bestand For Output As #1
  For Nummer = 1 To 15
    With HiScores(Nummer)
      Print #1, .Naam
      Print #1, .Datum
      ScoreString = CStr(.Score)
      If .NietRondmaken Then
        ScoreString = ScoreString & "*"
      End If
      Print #1, ScoreString
    End With
  Next Nummer
  Close #1

End Sub

Public Sub NieuwSpel()
  Dim HiScoreIndex As Integer
  Dim SpelerNr As Integer
  
  For HiScoreIndex = 1 To 15
    HiScores(HiScoreIndex).GewonnenDoor = 0
  Next HiScoreIndex
  For SpelerNr = 1 To 4
    Spelers(SpelerNr).HiScorePositie = 0
  Next SpelerNr
  IemandInHiScore = False
  'Score.NScoreRondes = 0

  cEersteRegelY = 6
  cRegelHoogte = (frmMain.picScore.ScaleHeight - 35) / nRonden
End Sub

Property Let MagInHiScore(Aan As Boolean)
  cMagInHiscore = Aan
End Property

Private Sub Class_Initialize()
  cMagInHiscore = True
End Sub

Public Sub Afdrukken()
  Dim Nr As Integer
  Dim NRKomtVoor As Boolean
  
  NRKomtVoor = False
  
  'Printer.FontName = "Courier"
  Printer.FontSize = 10
  Printer.FontBold = True
  Printer.CurrentX = 200
  Printer.Print "Highscore van 10 op en neer"
  Printer.CurrentX = 200
  Printer.Print "Afgedrukt door " & Spelers(1).Naam & " op " & Format(Now, "d mmmm yyyy")
  Printer.FontBold = False
  Printer.Print
  
  For Nr = 1 To 15
    Printer.CurrentX = 200
    Printer.CurrentY = Nr * 300 + 400
    Printer.Print Format(CStr(Nr), "@@")
  
    Printer.CurrentX = 600
    Printer.CurrentY = Nr * 300 + 400
    Printer.Print Format(HiScores(Nr).Naam, "!" & String(25, "@"))
  
    Printer.CurrentX = 2600
    Printer.CurrentY = Nr * 300 + 400
    Printer.Print Format(HiScores(Nr).Datum, String(10, "@"))
  
    Printer.CurrentX = 3800
    Printer.CurrentY = Nr * 300 + 400
    Printer.Print Format(CStr(HiScores(Nr).Score), String(4, "@"))
    
    If HiScores(Nr).NietRondmaken Then
      NRKomtVoor = True
      Printer.CurrentX = 4400
      Printer.CurrentY = Nr * 300 + 400
      Printer.Print "NR"
    End If
  Next Nr
  
  If NRKomtVoor Then
    Printer.Print
    Printer.CurrentX = 200
    Printer.Print "( NR betekent: rondmaken niet toegestaan )"
  End If
  Printer.EndDoc

End Sub

Sub UpdateSpelerNaam()
  If Spelers(1).HiScorePositie > 0 Then
    HiScores(Spelers(1).HiScorePositie).Naam = Spelers(1).Naam
  End If
End Sub

'Public Property Get NScoreRondes() As Integer
'  NScoreRondes = cNScoreRondes
'End Property
'
'Public Property Let NScoreRondes(ByVal Aantal As Integer)
'  If Aantal < 0 Or Aantal > 19 Then
'    MsgBox "Interne fout: onjuist rondenummer", vbCritical, "Fout"
'    End
'  Else
'    cNScoreRondes = Aantal
'  End If
'End Property

Public Property Let ToonVoorspelling(ByVal Tonen As Boolean)
  cToonVoorspelling = Tonen
End Property

Public Sub HighscoreWissen()
  Dim Bestand As String
  
  Bestand = DirPlusBestand(CStr(App.Path), "hiscore.dat")

  If Dir(Bestand) <> "" Then
    Kill Bestand
  End If
  InlezenHiScore
End Sub

Public Sub ToonPupupMenu()
  If WatInBeeld = sblHighscore Then
    frmMenus.PopupMenu frmMenus.mnuHighscore
  Else
    frmMenus.PopupMenu frmMenus.mnuScoreblok
  End If
End Sub

Public Function GetRondeVanMuisKlik(ByVal MouseX As Single, ByVal MouseY As Single)
  Dim HoogteFactor As Single
  Dim RondeTemp As Integer
  
  If WatInBeeld <> sblScore Then
    GetRondeVanMuisKlik = 0
  Else
    HoogteFactor = (frmMain.picScore.ScaleHeight - 35) / nRonden
    RondeTemp = (MouseY - 6 - HoogteFactor / 2) / HoogteFactor
    If RondeTemp < 1 Or RondeTemp > nRonden Then
      RondeTemp = 0
    End If
    GetRondeVanMuisKlik = RondeTemp
  End If
End Function
