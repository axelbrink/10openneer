VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTineke"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum VoorspellingType
  NietVoorspeld = -1
  MeeEens = 0
  MinderVoorspeld = 1
  MeerVoorspeld = 2
End Enum

Public Voorspelling As VoorspellingType
Public HulpTekst As String
Public HulpToonDoorgaanknop As Boolean
Public CommentaarTekst As String
Private Teksten(1 To 30) As String 'Dingen die ze kan zeggen
Private nTeksten As Integer
Private cToonPijltje As Boolean
Public GekozenTekst As Integer
Private KlikGrappigeTekstIndex As Integer

Public Sub TinekeTekst(Tekst As String)
  nTeksten = nTeksten + 1
  Teksten(nTeksten) = Tekst
End Sub

Public Sub Zeg(Gebeurtenis As String)
  Dim GekozenNummer As Integer
  
  'Randomize Timer
  
  nTeksten = 0
  
  Select Case Gebeurtenis
    Case "klik"
      TinekeTekst "Hallo!"
      TinekeTekst "Hoe gaat het ermee?"
      TinekeTekst "Hoe maak je het?"
      Select Case Hour(Now)
        Case 5 To 11
          TinekeTekst "Goedemorgen."
        Case 12 To 17
          TinekeTekst "Goedemiddag."
        Case 18 To 23
          TinekeTekst "Goedenavond."
        Case Else
          TinekeTekst "Goedenacht."
      End Select
      TinekeTekst "Wat kan ik voor je betekenen?"
      TinekeTekst "Ja zeg het maar."
      TinekeTekst "Wat is er nou?"
      TinekeTekst "Ga eens ergens anders op klikken."
      TinekeTekst "Hou eens op zeg."
      TinekeTekst "Ga nou maar kaarten."
      TinekeTekst "Heb je niks beters te doen?"
      TinekeTekst "Ik zit toch ook niet de hele tijd in jouw gezicht te klikken?"
      TinekeTekst "Je zit mijn tijd te verdoen."
      TinekeTekst "Ik zit hier ook niet voor mijn lol!"
      TinekeTekst "Denk je dat het leuk is om de hele tijd maar in een hoekje te zitten?"
      TinekeTekst "Nee dat is niet leuk."
      TinekeTekst "Daar is helemaal niks aan."
      TinekeTekst "Dus doe ons allebei een plezier en ga kaarten."
      TinekeTekst "Hup"
      TinekeTekst "Ik begin er genoeg van te krijgen."
      TinekeTekst "Hou op!"
      TinekeTekst "Ik word boos!"
      TinekeTekst "GRRRRRRRR!!!"
    Case "niet rondmaken"
      TinekeTekst "Je mag niet rondmaken. Dit betekent dat je nu niet " & SlagenOver & " slag" & IIf(SlagenOver = 1, "", "en") & " mag voorspellen."
    Case "weinig voorspeld"
      TinekeTekst "Lijkt mij een beetje weinig."
      TinekeTekst "Ik zou meer voorspellen."
      TinekeTekst "Dat is te weinig."
      TinekeTekst "Dat kan niet goed gaan."
      TinekeTekst "Je krijgt er meer."
      TinekeTekst "Pessimist."
      TinekeTekst "Dat gaat fout."
      TinekeTekst "Je bent te voorzichtig."
    Case "veel voorspeld"
      TinekeTekst "Dat lukt niet."
      TinekeTekst "Dat lijkt mij een beetje veel."
      TinekeTekst "Dat is te veel."
      TinekeTekst "Dat wordt moeilijk."
      TinekeTekst "Dat lukt vast niet."
      TinekeTekst "Je bent een optimist."
      TinekeTekst "Ja ja."
      TinekeTekst "Dat moet ik nog zien."
      TinekeTekst "Het zal mij benieuwen."
      TinekeTekst "Dat is behoorlijk veel."
    Case "te veel"
      TinekeTekst "Je hebt er te veel!"
      TinekeTekst "Foute boel."
      TinekeTekst "Pech."
      TinekeTekst "Dat zijn er te veel."
      TinekeTekst "Wat doe je nu?"
      TinekeTekst "Niet goed."
      TinekeTekst "Mis."
      TinekeTekst "Jammer."
      If Tineke.Voorspelling = 1 Then
        TinekeTekst "Zie je wel."
        TinekeTekst "Ik heb altijd gelijk."
        TinekeTekst "Haha."
        TinekeTekst "Dat dacht ik al."
        TinekeTekst "Eigen schuld."
      ElseIf Tineke.Voorspelling = 2 Then
        TinekeTekst "Vreemd."
        TinekeTekst "Dat is gek."
        TinekeTekst "Dat had ik niet gedacht."
      End If
      If SpelIs10openneer And (Spelers(1).TotaalScore / Ronde <= 3) And (Ronde >= 3) Then
        TinekeTekst "Het gaat erg slecht."
        TinekeTekst "Jij bent hier niet goed in."
        TinekeTekst "Het gaat niet best."
      End If
    Case "te weinig"
      TinekeTekst "Dat is niet leuk."
      TinekeTekst "Jammer."
      TinekeTekst "Mis."
      TinekeTekst "Pech."
      If SpelIs10openneer And (Spelers(1).TotaalScore / Ronde <= 3) And (Ronde >= 3) Then
        TinekeTekst "Het gaat erg slecht."
        TinekeTekst "Jij bent hier niet goed in."
        TinekeTekst "Het gaat niet best."
      End If
      If Tineke.Voorspelling = 1 Then
        TinekeTekst "Vreemd."
        TinekeTekst "Dat is gek."
        TinekeTekst "Dat had ik niet gedacht."
      ElseIf Tineke.Voorspelling = 2 Then
        TinekeTekst "Zie je wel."
        TinekeTekst "Ik heb altijd gelijk."
        TinekeTekst "Dat dacht ik al."
        TinekeTekst "Haha."
        TinekeTekst "Eigen schuld."
      End If
    Case "goed"
      TinekeTekst "Alweer goed."
      'TinekeTekst CStr(Spelers(1).Voorspelling + 10) & " punten erbij."
      TinekeTekst "Zo hoort het."
      TinekeTekst "Zo moet dat."
      TinekeTekst "Ga zo door."
      TinekeTekst "Oké."
      TinekeTekst "Goed."
      If SpelIs10openneer And (Spelers(1).TotaalScore / Ronde >= 8.5) And (Ronde >= 7) Then
        TinekeTekst "Het gaat erg goed."
        TinekeTekst "Jij bent hier goed in."
        TinekeTekst "Fabuleus."
      End If
      If Tineke.Voorspelling >= 1 Then
        TinekeTekst "Je had geluk."
        TinekeTekst "Wat een toeval."
        TinekeTekst "Dat had ik niet verwacht."
      End If
    Case "kleur bekennen"
      TinekeTekst "Je moet kleur bekennen. Kies een " & KleurNaam(KaartenOpTafel(NuOpkomen).Kleur) & "."
      TinekeTekst "Hela hola."
      TinekeTekst "Daar trap ik niet in."
      TinekeTekst "Opletten jij."
      TinekeTekst "Mafketel."
      TinekeTekst "Leuk geprobeerd."
      TinekeTekst "Niet valsspelen."
    Case "verplicht introeven"
      TinekeTekst "Je moet introeven. Kies een " & KleurNaam(Troef.Kleur) & "."
    Case "afgelopen", "gewonnen", "ookgewonnen"
      'If Opties.FoutVoorspeldNulPunten And Opties.PuntenPerSlag = 1 Then
      If SpelIs10openneer() Then
        '10 op en neer
        Select Case (19 * Spelers(1).TotaalScore / nRonden) 'Corrigeren vor korte potjes
          Case 0 To 49
            TinekeTekst "Wat ben je slecht!"
            TinekeTekst "Je kunt het niet."
          Case 50 To 79
            TinekeTekst "Je moet nog flink oefenen."
            TinekeTekst "Slecht."
          Case 80 To 89
            TinekeTekst "Goed, voor een beginner."
            TinekeTekst "Beter je best doen."
          Case 90 To 99
            TinekeTekst "Een magere score."
            TinekeTekst "Matig gespeeld."
          Case 100 To 109
            TinekeTekst "Het kon erger."
            TinekeTekst "Het kan beter."
          Case 110 To 119
            TinekeTekst "Niet geweldig."
            TinekeTekst "Het kan beter."
          Case 120 To 129
            TinekeTekst "Redelijk goed gespeeld."
            TinekeTekst "Een aardige score."
          Case 130 To 139
            TinekeTekst "Goed gespeeld."
            TinekeTekst "Een goede score."
          Case 140 To 149
            TinekeTekst "Een mooie score."
            TinekeTekst "Een prima score."
          Case 150 To 159
            TinekeTekst "Je weet hoe dit spel gaat."
            TinekeTekst "Goed gespeeld."
          Case 160 To 169
            TinekeTekst "Heel goed gedaan."
            TinekeTekst "Zeer goed."
          Case 170 To 179
            TinekeTekst "Uitstekend gespeeld."
            TinekeTekst "Je bent hier goed in!"
          Case 180 To 189
            TinekeTekst "Je bent een expert."
            TinekeTekst "Uitmuntend."
          Case 190 To 199
            TinekeTekst "Excellent gespeeld."
            TinekeTekst "Bijzonder goed."
          Case 200 To 209
            TinekeTekst "Reden voor feest!"
            TinekeTekst "Uitzonderlijk goed."
          Case 210 To 219
            TinekeTekst "Bel de krant!"
            TinekeTekst "Fabuleus!"
          Case 220 To 290
            TinekeTekst "Een fabelachtige score."
            TinekeTekst "Niet te geloven!"
        End Select
      Else
        TinekeTekst "Dat was het dan alweer."
        TinekeTekst "Het spel is uit."
        TinekeTekst "Het spel is afgelopen."
        TinekeTekst "Zo, dat was leuk."
      End If
    Case "highscore"
      TinekeTekst "Gefeliciteerd, je hebt de " & Spelers(1).HiScorePositie & "e plaats in de highscore behaald!"
    Case "hulp uit"
      TinekeTekst "Ik zal je niet meer helpen. Hulp kun je altijd weer inschakelen door op de knop Help te klikken."
    Case "herhalingklik"
      TinekeTekst "Je kunt nu geen kaart opleggen, omdat je een herhaling bekijkt. Druk op de spatiebalk om verder te gaan."
    Case Else
      Stop
  End Select
  
  If Gebeurtenis = "klik" Then
    KlikGrappigeTekstIndex = KlikGrappigeTekstIndex + 1
    If KlikGrappigeTekstIndex > nTeksten Then
      KlikGrappigeTekstIndex = 1
    End If
    GekozenTekst = KlikGrappigeTekstIndex
  Else
    GekozenTekst = Int(nTeksten * Rnd + 1)
  End If
  
  If Not Opties.Spelhulp Then
    If Opties.Commentaar Or Gebeurtenis = "klik" Or Left(Gebeurtenis, 10) = "commentaar" Then
      CommentaarTekst = Teksten(GekozenTekst)
      ToonPraatwolkje False
    End If
  End If
End Sub

Sub ZegHulp(Gebeurtenis As String)
  Dim Tekst As String
  
  Tekst = ""
  HulpToonDoorgaanknop = False
  
  Select Case LCase(Gebeurtenis)
    Case "kies een kaart"
      If NuOpkomen = 1 Then
        If Spelers(1).AantKaarten = 1 Then
          If Spelers(1).AantKaarten = AantKaartenRonde(Ronde) Then
            Tekst = "Je moet nu uitkomen (als eerste je kaart opleggen). Dat kun je zien aan je naam: die is nu geel. Klik met de muis op je kaart (de " & KaartNaam(Spelers(1).Kaarten(MogelijkeKaarten(1))) & ")."
          Else
            Tekst = "Jij moet nu je kaart opleggen, omdat je net een slag hebt gehaald. Klik met de muis op je kaart (de " & KaartNaam(Spelers(1).Kaarten(MogelijkeKaarten(1))) & ")."
          End If
        Else
          If Spelers(1).AantKaarten = AantKaartenRonde(Ronde) Then
            Tekst = "Je moet nu uitkomen (als eerste je kaart opleggen). Dat kun je zien aan je naam: die is nu geel. Klik met de muis op een kaart."
          Else
            Tekst = "Jij moet nu een kaart opleggen, omdat je net een slag hebt gehaald. Klik met de muis op een kaart."
          End If
        End If
      Else
        If Spelers(1).AantKaarten = 1 Then
          Tekst = "De spelers die voor je aan de beurt waren, hebben hun kaart opgelegd. Nu moet jij ook een kaart opleggen: klik met de muis op je kaart (de " & KaartNaam(Spelers(1).Kaarten(MogelijkeKaarten(1))) & ")."
        Else
          Tekst = "De eerste kaart op tafel is de " & KaartNaam(KaartenOpTafel(NuOpkomen)) & ". Nu moet jij kleur bekennen (ook een " & KleurNaam(KaartenOpTafel(NuOpkomen).Kleur) & " opleggen)."
          If KanKleurBekennen(1) Then
            Tekst = Tekst & " Klik met de muis op een " & KleurNaam(KaartenOpTafel(NuOpkomen).Kleur) & "kaart."
          Else
            Tekst = Tekst & " Je hebt geen " & KleurNaam(KaartenOpTafel(NuOpkomen).Kleur) & "; klik met de muis op een willekeurige kaart."
          End If
        End If
      End If
    Case "voorspellen"
      If Hulpniveau <= 1 Then
        Tekst = "Je moet nu het aantal slagen voorspellen. Je leert later wat dat inhoudt. Klik op een cijferknop."
      Else
        Tekst = "Je moet nu het aantal slagen voorspellen. Klik op een cijferknop. Mijn advies: kies " & CInt(TaxeerKaarten(1)) & "."
      End If
      cToonPijltje = True
    Case "voorspeld"
      If Hulpniveau <= 2 Then
        Select Case Spelers(1).Voorspelling
          Case 0
            Tekst = "Je hebt 0 slagen voorspeld. Dat kun je zien aan het grijze kruisje boven je kaarten. "
          Case 1
            Tekst = "Je hebt 1 slag voorspeld. Dat kun je zien aan het rondje boven je kaarten. "
          Case Else
            Tekst = "Je hebt " & Spelers(1).Voorspelling & " slagen voorspeld. Dat kun je zien aan de rondjes boven je kaarten. "
        End Select
        Tekst = Tekst & "Zo kun je ook zien wat de andere spelers voorspellen."
        HulpToonDoorgaanknop = True
      End If
    Case "slag naar hoogste"
      If Hulpniveau <= 10 Then
        If NuOpkomen = 1 Then
          Tekst = "Je legde als eerste een " & KaartNaam(KaartenOpTafel(NuOpkomen)) & " op. "
        Else
          Tekst = Spelers(NuOpkomen).Naam & " legde als eerste een " & KaartNaam(KaartenOpTafel(NuOpkomen)) & " op. "
        End If
        If HoogsteOpTafel = 1 Then
          If HoogsteOpTafel = NuOpkomen Then
            Tekst = Tekst & "Je krijgt nu deze vier kaarten, omdat niemand een hogere " & KleurNaam(KaartenOpTafel(HoogsteOpTafel).Kleur) & "kaart had."
          Else
            If KaartenOpTafel(HoogsteOpTafel).Kleur = Troef.Kleur And KaartenOpTafel(NuOpkomen).Kleur <> Troef.Kleur Then
              Tekst = Tekst & "Je krijgt nu deze vier kaarten, omdat je een " & KleurNaam(Troef.Kleur) & "kaart hebt opgelegd, en " & KleurNaam(Troef.Kleur) & " is troef (zie de troefkaart links van mij)."
            Else
              Tekst = Tekst & "Je krijgt nu deze vier kaarten, omdat je de hoogste " & KleurNaam(KaartenOpTafel(HoogsteOpTafel).Kleur) & " had."
            End If
          End If
        Else
          If HoogsteOpTafel = NuOpkomen Then
            Tekst = Tekst & Spelers(HoogsteOpTafel).Naam & " krijgt nu deze vier kaarten, omdat niemand een hogere " & KleurNaam(KaartenOpTafel(HoogsteOpTafel).Kleur) & "kaart had."
          Else
            If KaartenOpTafel(HoogsteOpTafel).Kleur = Troef.Kleur And KaartenOpTafel(NuOpkomen).Kleur <> Troef.Kleur Then
              Tekst = Tekst & Spelers(HoogsteOpTafel).Naam & " krijgt nu deze vier kaarten, omdat deze speler een " & KleurNaam(Troef.Kleur) & "kaart heeft opgelegd, en " & KleurNaam(Troef.Kleur) & " is troef (zie de troefkaart links van mij)."
            Else
              Tekst = Tekst & Spelers(HoogsteOpTafel).Naam & " krijgt nu deze vier kaarten, omdat deze speler de hoogste " & KleurNaam(KaartenOpTafel(HoogsteOpTafel).Kleur) & " had."
            End If
          End If
        End If
      End If
      Tineke.HulpToonDoorgaanknop = True
    Case "slag erbij"
      If Hulpniveau <= 3 Then
        If HoogsteOpTafel = 1 Then
          Tekst = "Je hebt nu een slag"
        Else
          Tekst = Spelers(HoogsteOpTafel).Naam & " heeft nu een slag"
        End If
        If Spelers(HoogsteOpTafel).AantSlagen = 1 Then
          Tekst = Tekst & "."
        Else
          Tekst = Tekst & " erbij."
        End If
      End If
      HulpToonDoorgaanknop = True
    Case "score"
      If Hulpniveau <= 5 Then
        If Spelers(1).Voorspelling = 1 Then
          Tekst = "Je had 1 slag voorspeld"
        Else
          Tekst = "Je had " & Spelers(1).Voorspelling & " slagen voorspeld"
        End If
        If Spelers(1).AantSlagen = Spelers(1).Voorspelling Then
          Tekst = Tekst & " en ook gekregen. Nu krijg je " & 10 + Spelers(1).AantSlagen & " ( = " & Spelers(1).AantSlagen & " + 10) punten erbij."
        Else
          Tekst = Tekst & ", maar dat is niet gelukt: je hebt er " & Spelers(1).AantSlagen & " gehaald. Geen punten erbij."
        End If
        If Hulpniveau <= 3 Then
          Tekst = Tekst & " De score wordt op het scoreblok boven mij bijgehouden."
        End If
        HulpToonDoorgaanknop = True
      End If
    Case "welkom"
      Tekst = "Hallo, welkom bij 10 op en neer. Ik ben Tineke en ik ga je tijdens het spelen uitleggen wat je moet doen."
      HulpToonDoorgaanknop = True
    Case Else
      Stop
  End Select
  
  If Tekst <> "" Then
    If HulpToonDoorgaanknop Then
      Tekst = Tekst & " Klik op 'Doorgaan'."
    End If
    HulpTekst = Tekst

    If Opties.Spelhulp Then
      If Ronde = 0 Then
        ToonPraatwolkje True
      Else
        ToonPraatwolkje True
      End If
    End If
  End If
End Sub

Sub ToonPraatwolkje(ByVal IsHulp As Boolean)
  Dim Tekst As String
  
  If IsHulp Then
    Tekst = Tineke.HulpTekst
  Else
    Tekst = Tineke.CommentaarTekst
  End If
  
  With frmMain.Praatwolkje
    .Visible = False
    If Len(Tekst) > 40 Then
      .Width = 2400
    Else
      .Width = 50 * Len(Tekst) + 600
    End If
  
    If IsHulp And frmMain.Praatwolkje.Width < 1455 Then
      frmMain.Praatwolkje.Width = 1455
    End If
        
    .Left = frmMain.ScaleWidth - 60 - frmMain.Praatwolkje.Width
    .Caption = Tekst
    .IsHulp = IsHulp
    .ToonDoorgaan = IsHulp And Tineke.HulpToonDoorgaanknop
    .Top = frmMain.picPraatwolkjePunt.Top - .Height + 30 '+ 30, anders past het niet (?)
    .Visible = True
  End With
  frmMain.imgVrouw.Picture = LoadResPicture(304, vbResBitmap) 'Aankijken
  
  frmMain.picPraatwolkjePunt.Visible = True
    
  frmMain.timVrouw.Enabled = False
  If IsHulp Then
    If cToonPijltje Then
      frmMain.timSpelhulp.Enabled = True
    End If
    If Tineke.HulpToonDoorgaanknop Then
      'frmMain.cmdGezien.SetFocus
      frmMain.Praatwolkje.SetFocus
      WachtOpGezien = True
      frmMain.timSpelhulp.Enabled = True
      Do
        DoEvents
      Loop While WachtOpGezien And Opties.Spelhulp
      '(Not frmMain.mnuSpelhulp.Checked)
    End If
  Else
    frmMain.timVrouw.Interval = 2400 + 80 * Len(Tekst)
    frmMain.timVrouw.Enabled = True
  End If
  'frmMain.picScore.SetFocus
End Sub

Public Sub WegPraatwolkje()
  frmMain.Praatwolkje.Visible = False
  frmMain.picPraatwolkjePunt.Visible = False
  frmMain.timSpelhulp.Enabled = False
  frmMain.timVrouw.Interval = 500
  frmMain.timVrouw.Enabled = False
  frmMain.timVrouw.Enabled = True
End Sub

Property Let ToonPijltje(Aan As Boolean)
  cToonPijltje = Aan
  frmMain.picPijltje.Visible = Aan
End Property
Property Get ToonPijltje() As Boolean
  ToonPijltje = cToonPijltje
End Property

Public Sub AfgelopenCommentaar()
  Dim Tekst As String
  Dim Applaus As String
  
  If Spelers(1).HiScorePositie > 0 Then
    Tekst = "highscore"
    Applaus = "Zeer goed gespeeld"
  ElseIf Spelers(1).Rang = 1 Then
    If Not Spelers(2).Rang = 1 And Not Spelers(3).Rang = 1 And Not Spelers(4).Rang = 1 Then
      Tekst = "gewonnen"
      Applaus = "Zeer goed gespeeld"
    Else
      Tekst = "ookgewonnen"
      Applaus = "Goed gespeeld"
    End If
  Else
    Tekst = "afgelopen"
    If SpelIs10openneer() Then
      If 19 * Spelers(1).TotaalScore / nRonden >= 190 Then
        Applaus = "Zeer goed gespeeld"
      ElseIf 19 * Spelers(1).TotaalScore / nRonden >= 180 And Applaus = "" Then
        Applaus = "Goed gespeeld"
      End If
    End If
  End If
  
  If Applaus <> "" Then
    WavPlay Applaus
  End If

  Tineke.Zeg Tekst

End Sub

Private Sub Class_Initialize()
  KlikGrappigeTekstIndex = 0
End Sub
