Attribute VB_Name = "modKaarten"
Option Explicit

Global Const MargeHWG = 0.09 'Hoge weggooien
Global Const MargeOMLT = 0.35 '0.36 'Opkomen met lage troef
Global Const MargeSNN = 0.09 '0.07 '0.08 '0.09 'Klopt dit? 'Slag niet nemen

Public Enum ControllerType
  ControllerComputer
  ControllerMens
  ControllerNetwerk
End Enum
Public Type Kaart
  Getal As Integer
  Kleur As Integer
  Legaal As Boolean
End Type
Public Type Stapel
  Kaarten(1 To 52) As Kaart
  AantKaarten As Integer
End Type
Public Type Speler
  AantKaarten As Integer
  AantSlagen As Integer
  Controller As ControllerType
  HeeftKleurNietMeer(1 To 4) As Boolean
  HiScorePositie As Integer
  IPAdres As String
  Kaarten(1 To 13) As Kaart
  Naam As String
  Rang As Integer
  Score(1 To 25) As Integer '1..13..1: 25 ronden
  ScoreFout(1 To 25) As Boolean
  Voorspelling As Integer
  Taxatie As Single
  TotaalScore As Integer
End Type
Type OptiesType
  AflopendSorteren As Boolean
  AniSnelheid As Integer
  BreedUitspreiden As Boolean
  Commentaar As Boolean
  ComputersInHighScore As Boolean
  FoutVoorspeldNulPunten As Boolean
  Geluid As Boolean
  GroteKaarten As Boolean
  IntroevenVerplicht As Boolean
  MaxAantKaarten As Integer
  MeteenOptellen As Boolean
  NegatieveScores As Boolean
  PuntenPerSlag As Integer
  RondmakenToegestaan As Boolean
  Spelhulp As Boolean
  SpelSnelheid As Integer
  StrafpuntenPerSlag As Integer
End Type

Global DeStapel As Stapel
Global Spelers(1 To 4) As Speler
Global Troef As Kaart
Global KaartenOpTafel(1 To 4) As Kaart
Global LaatsteSlag(1 To 4) As Kaart
Global LaatsteSlagOpgekomenSpeler As Integer
Global GeenKaart As Kaart

Global HoogsteOpTafel As Integer

Global WaarIsKaart(1 To 4, 1 To 14) As Integer '# Enum gebruiken
'-3 = Weg, -2 = Tafel, -1 = Troef, 0 = stapel, 1 = Speler 1, ...
Global AantKaartenGezien As Integer

Global NuVoorspellen As Integer
'Global EerstOpkomen As Integer
Global NuOpkomen As Integer
Global NuOpleggen As Integer
Global SpelerNum As Integer '1=speler die opkomt is aan de beurt, 2=volgende, ...

Global SlagIsTeHalenMetNummer(1 To 13) As Integer
Global SlagIsTeHalenMetAantal As Integer
Global SlagNietTeHalenMetNummer(1 To 13) As Integer
Global SlagNietTeHalenMetAantal As Integer
Global MogelijkeKaarten(1 To 13) As Integer
Global AantalMogelijkeKaarten As Integer
Global Troefkaarten(1 To 13) As Integer
Global AantTroefKaarten As Integer
Global LegaleTroefkaarten(1 To 13) As Integer
Global AantLegaleTroefKaarten As Integer
Global NietTroefkaarten(1 To 13) As Integer
Global AantNietTroefKaarten As Integer
Global LegaleNietTroefkaarten(1 To 13) As Integer
Global AantLegaleNietTroefKaarten As Integer

Global KleurNaam(1 To 4) As String
Global GetalNaam(1 To 14) As String
Global KaartWaarde(1 To 14) As Single

Global KleurAantalKerenGespeeld(1 To 4) As Integer '0 ook, voor als er geen troef is

Global AantKaartenRonde(1 To 25) As Integer
Global TotSlagenGok As Integer
Global Ronde As Integer
Global KaartenResterend As Integer
Global AantSpelersGegokt As Integer
Global SlagNr As Integer

Global SlagenOver As Integer

'Global Tineke As TinekeType
Global Opties As OptiesType

Global Tip As Integer
Global TipZichtbaar As Boolean

Public Function TaxeerKaarten(SpelerNr As Integer) As Single
  Dim KaartNr As Integer
  Dim GokAantalTemp As Single
  Dim Ratio As Single
  
  'Dim NormaalPerSpeler As Single
  'Dim NuPerSpeler As Single
  'Dim Factor As Single
  Dim TaxatieTemp As Single 'Nieuw
  'Dim VerschilMetGem As Single
  'Dim GokAantalTemp2 As Single
  
  Dim Weging As Single

  GokAantalTemp = 0
 
  For KaartNr = 1 To AantKaartenRonde(Ronde)
    If Spelers(SpelerNr).Kaarten(KaartNr).Getal > 0 Then
      GokAantalTemp = GokAantalTemp + Positief(BerekenKaartWaarde(SpelerNr, KaartNr, True))
    End If
  Next KaartNr
  
'  Debug.Print "Speler " & SpelerNr '& " orig: " & Format(GokAantalTemp, "0.00")
  
'  VerschilMetGem = Abs(GokAantalTemp - KaartenOverVoorDezeSpeler(SpelerNr))
'  Ratio = 0.2 * AantSpelersGegokt * VerschilMetGem '* 0.13 '0.12
'  Debug.Print "Ratio: " & Ratio
'  GokAantalTemp2 = (1 - Ratio) * GokAantalTemp + Ratio * KaartenOverVoorDezeSpeler(SpelerNr)
'
'  Debug.Print "          nu-nw: " & Format(GokAantalTemp2, "0.000")
'

  TaxatieTemp = GokAantalTemp
  If AantSpelersGegokt > 0 And AantSpelersGegokt < 4 Then
    If GokAantalTemp > Int(KaartenOverVoorDezeSpeler(SpelerNr)) Then
      'Debug.Print "         Is meer dan gem pp over"
                  ' "Int" is nieuw: naar beneden afronden
      Ratio = AantSpelersGegokt * 0.13 '0.12
      GokAantalTemp = (1 - Ratio) * GokAantalTemp + Ratio * Int(KaartenOverVoorDezeSpeler(SpelerNr))
    Else 'Wilde minder zeggen
      'Debug.Print "         Is [minder] dan gem pp over"
      Ratio = AantSpelersGegokt * 0.026 'Weinig invloed
      GokAantalTemp = (1 - Ratio) * GokAantalTemp + Ratio * Int(KaartenOverVoorDezeSpeler(SpelerNr))
    End If
  End If
'  '# Dit hierboven lijkt te sterk te werken (bij weinig kaarten)
'
'  Debug.Print "         Nuoud:" & Format(GokAantalTemp, "0.000")
'
'  NormaalPerSpeler = AantKaartenRonde(Ronde) / 4
'  NuPerSpeler = KaartenOverVoorDezeSpeler(SpelerNr)
'
'  'Factor > 1 => anderen hebben minder dan gemiddeld voorspeld
'  Factor = 1 + 0.1 * (15 - Ronde) * (NuPerSpeler - NormaalPerSpeler)
'  Debug.Print "Factor = " & Factor
'  If Factor > 1 Then
'    Debug.Print "         Anderen hebben weinig voorspeld"
'  ElseIf Factor < 1 Then
'    Debug.Print "         Anderen hebben veel voorspeld"
'  End If
'  'Debug.Print "         nups: " & Format(NuPerSpeler, "0.00") & ", norps" & Format(NormaalPerSpeler, "0.00")
'  VoorspellingTemp = Factor * GokAantalTemp2
'
'  Debug.Print "         Dus:  " & Format(VoorspellingTemp, "0.000")
'
'  GokAantalTemp = VoorspellingTemp
  
  If AantSpelersGegokt < 4 Then 'Voorspellen
    Weging = 0.07 * AantSpelersGegokt '0.07 0.06
    TaxatieTemp = Weging * KaartenOverVoorDezeSpeler(SpelerNr) + (1 - Weging) * TaxatieTemp
  Else                          'In spel
    Weging = 0.08 'Dit is iets anders dan die weging hierboven
    TaxatieTemp = TaxatieTemp + Weging * (Spelers(1).AantKaarten - SamenSlagenNodig()) 'SlagenOver
    '# KaartenOverVoorDezeSpeler heeft geen zin in spel! (wordt ingesteld als 0)
  End If
  '* (1 + 0.1 * (10 - AantKaartenRonde(Ronde)))
  'TaxatieTemp = Weging * KaartenOverVoorDezeSpeler(SpelerNr) + (1 - Weging) * TaxatieTemp
  
  'Debug.Print "      Vroeger: " & Format(GokAantalTemp, "0.00")
  'If CInt(GokAantalTemp) <> CInt(TaxatieTemp) Then
  '  Debug.Print "* verschil! *"
  'End If
  
  If Troef.Kleur = 0 Then 'Geen troef -> alles is mogelijk
    TaxatieTemp = TaxatieTemp * 2 'Uit de lucht gegrepen!
  End If
  
  If TaxatieTemp > KaartenResterend Then
    TaxatieTemp = KaartenResterend
  End If
  
  TaxatieTemp = Positief(TaxatieTemp)
  'Debug.Print "Speler " & SpelerNr & "  taxatie: " & Format(TaxatieTemp, "0.00")
  TaxeerKaarten = TaxatieTemp
End Function

Public Function BerekenKaartWaarde(ByVal SpelerNr As Integer, ByVal KaartNr As Integer, ByVal TroefIsHoog As Boolean) As Single
  Dim WaardeTemp As Single
  Dim AftrekTemp As Single
  Dim LaagsteKaartWaarde As Single
  Dim KaartNrTemp As Integer
  Dim LaagsteGetal As Integer
  Dim LaagsteNummer As Integer
  Dim MoetOpkomen As Integer
  Dim Getal As Integer
  Dim Kleur As Integer
 
  LaagsteGetal = 15
  
  If SpelerNr = NuOpkomen Then
    MoetOpkomen = True
  Else
    MoetOpkomen = False
  End If

  Kleur = Spelers(SpelerNr).Kaarten(KaartNr).Kleur
  Getal = Spelers(SpelerNr).Kaarten(KaartNr).Getal

  WaardeTemp = KaartWaarde(Getal)
  If Kleur = Troef.Kleur And TroefIsHoog Then
    WaardeTemp = WaardeTemp + 0.41 '0.39 '0.4 '0.3
  End If
  If Kleur <> Troef.Kleur And MoetOpkomen = False Then
    WaardeTemp = WaardeTemp * (1 - 0.2 * KleurAantalKerenGespeeld(Kleur))
    'WaardeTemp = WaardeTemp - 0.2 * KleurAantalKerenGespeeld(Kleur)
  ElseIf Kleur <> Troef.Kleur And MoetOpkomen = True And Troef.Kleur <> 0 Then
    WaardeTemp = WaardeTemp * (1 + 0.2 * KleurAantalKerenGespeeld(Troef.Kleur))
    'WaardeTemp = WaardeTemp + 0.2 * KleurAantalKerenGespeeld(Troef.Kleur)
  Else
    'WaardeTemp = WaardeTemp * (1 + 0.4 * KleurAantalKerenGespeeld(Kleur))
    'WaardeTemp = WaardeTemp * (1 + 0.6 * Sqr(KleurAantalKerenGespeeld(Kleur)))
    WaardeTemp = WaardeTemp * (Sqr(KleurAantalKerenGespeeld(Kleur) + 1))
    'WaardeTemp = WaardeTemp + 0.25 * KleurAantalKerenGespeeld(Kleur)
  End If
 
  For KaartNrTemp = 1 To AantKaartenRonde(Ronde)
    'Sommige kaarten daarvan zijn misschien al "weg"
    If Spelers(SpelerNr).Kaarten(KaartNrTemp).Kleur = Kleur Then
      If Spelers(SpelerNr).Kaarten(KaartNrTemp).Getal < LaagsteGetal Then
        LaagsteGetal = Spelers(SpelerNr).Kaarten(KaartNrTemp).Getal
        LaagsteNummer = KaartNrTemp
      End If
    End If
  Next KaartNrTemp
  
   If Spelers(SpelerNr).Kaarten(KaartNr).Kleur = Troef.Kleur Then
     'WaardeTemp = WaardeTemp + 0.14 * Positief(4 - KaartenResterend)
     WaardeTemp = WaardeTemp * (1 + 0.08 * Positief(4 - KaartenResterend))
   Else
     'WaardeTemp = WaardeTemp - 0.14 * Positief(4 - KaartenResterend)
     WaardeTemp = WaardeTemp * (1 - 0.18 * Positief(4 - KaartenResterend))
   End If
  
  If LaagsteNummer = KaartNr Then
  Else
    If Kleur = Troef.Kleur Then
      'WaardeTemp = WaardeTemp - (AftrekTemp / 56) * WaardeTemp
    ElseIf LaagsteNummer <> KaartNr And LaagsteGetal < 10 Then
      LaagsteKaartWaarde = BerekenKaartWaarde(SpelerNr, LaagsteNummer, True)
      'WaardeTemp = WaardeTemp - (AftrekTemp / 28) * WaardeTemp
      WaardeTemp = 0.8 * LaagsteKaartWaarde + 0.2 * WaardeTemp
    End If
  End If
  
  BerekenKaartWaarde = WaardeTemp
End Function

Public Function KaartenOverVoorDezeSpeler(SpelerNr) As Single
  If AantSpelersGegokt < 4 Then
    KaartenOverVoorDezeSpeler = (AantKaartenRonde(Ronde) - TotSlagenGok) / (4 - AantSpelersGegokt)
  Else
    KaartenOverVoorDezeSpeler = 0
  End If
End Function

Function KaartNaam(DeKaart As Kaart) As String
  KaartNaam = KleurNaam(DeKaart.Kleur) & GetalNaam(DeKaart.Getal)
End Function

Function SlagenOverTekst() As String
  Dim Tekst As String
  
  Select Case SlagenOver
    Case Is < 0
      Tekst = CStr(-SlagenOver) & " te weinig"
    Case 0
      Tekst = "Rond"
    Case Is > 0
      Tekst = CStr(SlagenOver) & " over"
  End Select
  
  If Spelers(VorigeSpeler(VorigeSpeler(NuOpkomen))).Voorspelling = -1 Then
    Tekst = "(" & Tekst & ")"
  End If
  SlagenOverTekst = Tekst
End Function

Sub ToonKaartenInfo(SpelerNr As Integer) '# Zie InfoTemp
  Dim KaartNr As Integer
  'Debug.Print "."
  For KaartNr = 1 To AantKaartenRonde(Ronde)
    If Spelers(SpelerNr).Kaarten(KaartNr).Getal > 0 Then
      Debug.Print KaartNr & ": " & KaartNaam(Spelers(SpelerNr).Kaarten(KaartNr))
    End If
  Next KaartNr

End Sub

Function LoadResConv(Id As Integer) As String
  Dim Orig As String
  Dim Nieuw As String
  Dim DollarLoc As Integer
  Dim Codewoord As String
  
  Orig = LoadResString(Id)
  DollarLoc = InStr(Orig, "$")
  Do Until DollarLoc <= 0
    Nieuw = Nieuw & Left(Orig, DollarLoc - 1)
    Orig = Mid(Orig, DollarLoc + 1)
    DollarLoc = InStr(Orig, "$")
    Codewoord = Left(Orig, DollarLoc - 1)
    
    Select Case UCase(Codewoord)
      Case "SPELER1"
        Nieuw = Nieuw & Spelers(1).Naam
      Case "SPELER2"
        Nieuw = Nieuw & Spelers(2).Naam
      Case "SPELER3"
        Nieuw = Nieuw & Spelers(3).Naam
      Case "SPELER4"
        Nieuw = Nieuw & Spelers(4).Naam
      Case "ADVIES"
        Nieuw = Nieuw & CInt(TaxeerKaarten(1))
    End Select
  
    If DollarLoc < Len(Orig) Then
      Orig = Mid(Orig, DollarLoc + 1)
      DollarLoc = InStr(Orig, "$")
    Else
      DollarLoc = 0
    End If
  Loop
  Nieuw = Nieuw & Orig
  LoadResConv = Nieuw
End Function

'Sub BepaalAantKaartenPerRonde()
'  Dim i As Integer
'
'  nRonden = 2 * Opties.MaxAantKaarten - 1
'  For i = 1 To nRonden
'    AantKaartenRonde(i) = -Abs(i - Opties.MaxAantKaarten) + Opties.MaxAantKaarten
'  Next i
'End Sub
'
Public Function KanKleurBekennen(ByVal SpelerNr As Integer) As Boolean
  '# Dit kan samen met KanIntroeven
  Dim KaartNr As Integer
 
  KanKleurBekennen = False
  If SpelerNr <> NuOpkomen Then
    For KaartNr = 1 To AantKaartenRonde(Ronde)
      If Spelers(SpelerNr).Kaarten(KaartNr).Kleur = KaartenOpTafel(NuOpkomen).Kleur Then
        KanKleurBekennen = True
      End If
    Next KaartNr
  End If

End Function

Function SamenSlagenNodig() As Integer
  Dim SpelerNr As Integer
  Dim Som As Integer
  Dim SpelerNodig As Integer
  
  Som = 0
  For SpelerNr = 1 To 4
    SpelerNodig = Spelers(SpelerNr).Voorspelling - Spelers(SpelerNr).AantSlagen
    If SpelerNodig > 0 Then
      Som = Som + SpelerNodig
    End If
  Next SpelerNr
  SamenSlagenNodig = Som
End Function

Public Function VolgendeSpeler(HuidigeSpeler As Integer) As Integer
  VolgendeSpeler = (HuidigeSpeler Mod 4) + 1
End Function

Public Function DetermineerSpel()
  'Deelt de kaarten alvast voor het hele spel
  'Dus voor het geval dat dit geen herhaling is
  'Maakt het spel deterministisch


'## Doen: WaarIsKaart
'         BepaalOpkomen

  Dim r As Integer
  Dim k As Integer
  Dim KaartNr As Integer
  Dim SpelerNr As Integer
  'Dim StapelAantKaarten As Integer

  Randomize Timer
  'If recWieBegint < 1 Then
  recWieBegint = Int(4 * Rnd + 1)
  'End If
  'EerstOpkomen = VorigeSpeler(recWieBegint) 'In NieuweRonde wordt meteen de volgende speler genomen
  
  For r = 1 To nRonden
  
    'Select Case IkBenNetSpelerNr
    '  Case 0
  
    AantKaartenRonde(r) = -Abs(r - Opties.MaxAantKaarten) + Opties.MaxAantKaarten

    '** Stapel maken **
    For KaartNr = 1 To 52
      DeStapel.Kaarten(KaartNr).Kleur = (KaartNr - 1) \ 13 + 1
      DeStapel.Kaarten(KaartNr).Getal = (KaartNr - 1) Mod 13 + 2
      'WaarIsKaart(DeStapel.Kaarten(KaartNr).Kleur, DeStapel.Kaarten(KaartNr).Getal) = 0
    Next KaartNr
    DeStapel.AantKaarten = 52

    '** Kaarten verdelen (voor elke ronde dus) **
    For KaartNr = 1 To AantKaartenRonde(r)
      For SpelerNr = 1 To 4
        k = Int(DeStapel.AantKaarten * Rnd + 1)
        recKaartenOntvangen(r, SpelerNr, KaartNr) = DeStapel.Kaarten(k)
        DeStapel.Kaarten(k) = DeStapel.Kaarten(DeStapel.AantKaarten)
        DeStapel.AantKaarten = DeStapel.AantKaarten - 1
      Next SpelerNr
    Next KaartNr
    
    '** Troef bepalen **
    If DeStapel.AantKaarten = 0 Then
      recTroeven(r) = GeenKaart
    Else
      recTroeven(r) = DeStapel.Kaarten(CInt(DeStapel.AantKaarten * Rnd + 1))
    End If
    'N.B. stapel wordt niet meer bijgewerkt; niet nodig

  Next r
  
End Function
