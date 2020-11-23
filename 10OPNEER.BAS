Attribute VB_Name = "modAlgemeen"
Option Explicit

Public Const StandaardNaam1 = "Speler"
Public Const StandaardNaam2 = "Henk"
Public Const StandaardNaam3 = "Truus"
Public Const StandaardNaam4 = "Kees"
Public Const KaartBreedte As Integer = 1065
Public Const KaartHoogte As Integer = 1440

Public Enum WachtOpType
  ComputerKaart
  ComputerVoorspelling
  MensKaart
  MensVoorspelling
  NetwerkKaart
  NetwerkVoorspelling
  Niets
End Enum

Public Enum ScoreblokType
  sblScore = 1
  sblHighscore = 2
End Enum

Public Type StatistiekType
  Gemiddelde As Single
  PCGemiddelde As Single
  RangFreq(1 To 4) As Long
  
  Sessies As Long
  SpellenBegonnen As Long
  SpellenVoltooid As Long 'nOnvoltooid = nBegonnen - nVoltooid
  
  TotaalVoorspeld As Long
  TotaalVoorspeldMax As Long
  GoedVoorspeld As Long
  FoutVoorspeld As Long
  FoutSaldo As Integer
  Rondgemaakt As Long
  NietRondgemaakt As Long
  'Sessieduur = Int(DateDiff("n", StartMoment, Now))
  'DuurPerSessie = TotaalduurNu \ nSessies
  TotaalDuur As Long   'TotaalduurNu = TotaalDuur + Sessieduur
  LaatstGespeeld As Date
  DagenGespeeld As Long
  InstallDate As Date  'VerlopenDagen = DateDiff("d", InstallDate, Now) + 1
End Type

Global Const Voorkleur = &HFF00&
Global Const Achterkleur = &H8000&
Global Const Email = "axel@fmf.nl"
Global Const Homepage = "http://www.fmf.nl/~axel/"
Global Const InnoSetupAdres As String = "http://www.jrsoftware.org/isinfo.htm"
'Global Const NietGebruikenNa As Date = "31-01-2003"
Global Const LinkColor As Long = &HFF0000
Global Const LinkActiveColor As Long = &HFF 'Rood

Global Const KeySlash = 47 '191 (nu Ascii, was Ansi)

Global WachtOp As WachtOpType
Global Statistiek As StatistiekType

Global Score As clsScore
Global Tineke As clsTineke 'New
Global NaamWijzigen As Integer
Global StartMoment As Date
Global IniBestand As String
'Global MagSpelSpelen As Boolean
Global AchterkantTitle As String
Global AchterkantBestand As String
Global KaartAfst As Integer
Global WachtOpGezien As Boolean
Global KaartAanwijzen As Integer
Global Hulpniveau As Integer
Global KaartImageGeladen(0 To 51) As Boolean
Global RondjeImageGeladen(0 To 51) As Boolean
Global OndergangGeluidGespeeld As Boolean
Global PopupMenuIndex As Integer 'Om aan te geven waarop popupmenu is gedaan
'Global MaxAantKaarten As Integer
Global nRonden As Integer
Global HiscoreGelezen As Boolean 'Highscorebestand wel of niet ingelezen
'Global HerhalingBezig As Boolean

'** Netwerk (oud) **
Global IkBenNetSpelerNr As Integer
Global ClientVraagt(2 To 4) As String

'--- Start URL or file
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Const SW_SHOWNORMAL = 1

Sub Main()
  'Randomize Timer
  StartMoment = Now
  
  App.HelpFile = DirPlusBestand(CStr(App.Path), "10opneer.hlp")
  
  IniBestand = DirPlusBestand(CStr(App.Path), "10opneer.ini")
  IniLoad
  Set Score = New clsScore
  'Score.InlezenHiScore
  'InlezenHiScore
  HiscoreGelezen = False
  frmMain.InitVars
  frmMain.InitControls
  
  If Screen.Width <= 640 * 15 Then
    frmMain.WindowState = 2
  Else
    'frmMain.Move frmMain.Left, frmMain.Top, 648 * Screen.TwipsPerPixelX, 460 * Screen.TwipsPerPixelY
    frmMain.Move frmMain.Left, frmMain.Top, 680 * Screen.TwipsPerPixelX, 490 * Screen.TwipsPerPixelY
  End If
  'frmMain.Caption = LoadResConv(1) '10 op en neer / 10 up and down
  Set frmMain.Netwerk = New clsNetwerk
'#! frmMain.Netwerk.Init frmMain.wskListen, frmMain.Winsock(0), frmMain.Winsock(1), frmMain.Winsock(2)

  frmMain.FormResize
  frmMain.Show

  With frmMain.StatusBar
    .NumPanels = 2
    .PanelText(0) = "" 'Speltype
    .PanelText(1) = "" 'Slagen over
    .PanelWidth(0) = 1200
    .PanelWidth(1) = 1200
    .PanelToolTip(0) = "Het spel dat je speelt: 10 op en neer of Boerenbridge"
    .PanelToolTip(1) = "Het aantal slagen dat over is"
    .PanelWhatsThisHelpID(0) = 1006
    .PanelWhatsThisHelpID(1) = 1007
    .PanelVisible(0) = True
    .PanelVisible(1) = True
  End With
  
  Set Tineke = New clsTineke

  WavPlay "Programma starten"
  If Opties.Spelhulp Then
    frmMain.SpelhulpAanUit
    Tineke.ZegHulp "welkom"
  End If

  frmMain.NieuwSpel False
End Sub

Function BestandVanPad(Pad As String) As String
  Dim TekenNr As Integer
  TekenNr = Len(Pad)
  Do Until Mid(Pad, TekenNr, 1) = "\"
    TekenNr = TekenNr - 1
  Loop
  BestandVanPad = Mid(Pad, TekenNr + 1)
End Function

Function DirPlusBestand(Directory As String, ByVal Bestand As String) As String
  If Right(Directory, 1) = "\" Then
    DirPlusBestand = Directory & Bestand
  Else
    DirPlusBestand = Directory & "\" & Bestand
  End If
End Function

Sub WavPlay(Gebeurtenis As String)
  Dim Bestand As String
  
  If Opties.Geluid = True Then
    Bestand = QueryValue(HKEY_CURRENT_USER, "AppEvents\Schemes\Apps\10 op en neer\" & Gebeurtenis & "\.Current", "")
    If Bestand <> "" Then
      If Len(Bestand) Then
        sndPlaySound Bestand, 1
      End If
    End If
  End If
End Sub

Public Function Positief(ByVal Waarde) 'Geen type
  If Waarde < 0 Then
    Positief = 0
  Else
    Positief = Waarde
  End If
End Function

Public Function VorigeSpeler(HuidigeSpeler As Integer) As Integer
  VorigeSpeler = (HuidigeSpeler + 2) Mod 4 + 1
End Function

'start htlm page from VB5
'on the General Declaretions section

Public Function StartURL(URL As String) As Long
  Dim Scr_hDC As Long
  Scr_hDC = GetDesktopWindow()
  StartURL = ShellExecute(Scr_hDC, "Open", URL, "", "C:\", SW_SHOWNORMAL)
  'use as Call StartDoc("http://www.microsoft.com")
End Function

Public Sub InlezenStatistiek()
  With Statistiek
    .Gemiddelde = CSng(IniGet("Statistiek", "Gemiddelde", 0))
    .PCGemiddelde = CSng(IniGet("Statistiek", "PCGemiddelde", 0))
    .RangFreq(1) = CLng(IniGet("Statistiek", "Freq1", 0))
    .RangFreq(2) = CLng(IniGet("Statistiek", "Freq2", 0))
    .RangFreq(3) = CLng(IniGet("Statistiek", "Freq3", 0))
    .RangFreq(4) = CLng(IniGet("Statistiek", "Freq4", 0))
  
    .Sessies = CLng(IniGet("Statistiek", "Sessies", 0))
    .SpellenBegonnen = CLng(IniGet("Statistiek", "SpellenBegonnen", 0))
    .SpellenVoltooid = CLng(IniGet("Statistiek", "SpellenVoltooid", 0))
    
    .TotaalVoorspeld = CLng(IniGet("Statistiek", "TotaalVoorspeld", 0))
    .TotaalVoorspeldMax = CLng(IniGet("Statistiek", "TotaalVoorspeldMax", 0))
    .GoedVoorspeld = CLng(IniGet("Statistiek", "GoedAant", 0))
    .FoutVoorspeld = CLng(IniGet("Statistiek", "FoutAant", 0))
    .FoutSaldo = CLng(IniGet("Statistiek", "FoutSaldo", 0))
    .Rondgemaakt = CLng(IniGet("Statistiek", "Rondgemaakt", 0))
    .NietRondgemaakt = CLng(IniGet("Statistiek", "NietRondgemaakt", 0))
    
    .TotaalDuur = CLng(IniGet("Statistiek", "Totaalduur", 0))
    .LaatstGespeeld = CDate(IniGet("Statistiek", "LaatstGespeeld", Format(Now, "dd-mm-yyyy")))
    .DagenGespeeld = CLng(IniGet("Statistiek", "DagenGespeeld", 1))
    .InstallDate = CDate(IniGet("Statistiek", "InstallDate", Format(Now, "dd-mm-yyyy")))
  End With
End Sub

Public Sub OpslaanStatistiek()
  'Slaat op naar Ini; daarna nog IniSave doen!
  
  With Statistiek
    IniSet "Statistiek", "Gemiddelde", .Gemiddelde
    IniSet "Statistiek", "PCGemiddelde", .PCGemiddelde
    IniSet "Statistiek", "Freq1", .RangFreq(1)
    IniSet "Statistiek", "Freq2", .RangFreq(2)
    IniSet "Statistiek", "Freq3", .RangFreq(3)
    IniSet "Statistiek", "Freq4", .RangFreq(4)
  
    IniSet "Statistiek", "Sessies", .Sessies
    IniSet "Statistiek", "SpellenBegonnen", .SpellenBegonnen
    IniSet "Statistiek", "SpellenVoltooid", .SpellenVoltooid
  
    IniSet "Statistiek", "TotaalVoorspeld", .TotaalVoorspeld
    IniSet "Statistiek", "TotaalVoorspeldMax", .TotaalVoorspeldMax
    IniSet "Statistiek", "GoedAant", .GoedVoorspeld
    IniSet "Statistiek", "FoutAant", .FoutVoorspeld
    IniSet "Statistiek", "FoutSaldo", .FoutSaldo
    IniSet "Statistiek", "Rondgemaakt", .Rondgemaakt
    IniSet "Statistiek", "NietRondgemaakt", .NietRondgemaakt
    
    IniSet "Statistiek", "Totaalduur", .TotaalDuur
    IniSet "Statistiek", "LaatstGespeeld", Format(.LaatstGespeeld, "dd-mm-yyyy")
    IniSet "Statistiek", "DagenGespeeld", .DagenGespeeld
    IniSet "Statistiek", "InstallDate", Format(.InstallDate, "dd-mm-yyyy")
  End With
End Sub

Sub InlezenOpties()
  With Opties
    .AniSnelheid = IniGet("Opties", "Animatiesnelheid", 4)
    .AflopendSorteren = IniGet("Opties", "Aflopend", False)
    .BreedUitspreiden = IniGet("Opties", "BreedUitspreiden", False)
    .Commentaar = IniGet("Opties", "Commentaar", True)
    .ComputersInHighScore = IniGet("Opties", "ComputersInHighScore", True)
    .Geluid = IniGet("Opties", "Geluid", True)
    .MaxAantKaarten = IniGet("Opties", "MaxAantKaarten", 10)
    .MeteenOptellen = IniGet("Opties", "Optellen", True)
    .RondmakenToegestaan = IniGet("Opties", "Rondmaken", True)
    .Spelhulp = IniGet("Opties", "Spelhulp", True)
    .SpelSnelheid = IniGet("Opties", "Spelsnelheid", 4)
    .IntroevenVerplicht = IniGet("Opties", "IntroevenVerplicht", False)
    .GroteKaarten = IniGet("Opties", "GroteKaarten", False)
    
    .PuntenPerSlag = IniGet("Opties", "PuntenPerSlag", 1)
    .StrafpuntenPerSlag = IniGet("Opties", "StrafpuntenPerSlag", 3)
    .FoutVoorspeldNulPunten = IniGet("Opties", "FoutVoorspeldNulPunten", True)
    .NegatieveScores = IniGet("Opties", "NegatieveScores", False)
  End With
  AchterkantTitle = IniGet("Opties", "Achterkant", "10opneer.gif")
  Spelers(1).Naam = IniGet("Opties", "Naam1", StandaardNaam1)
  Spelers(2).Naam = IniGet("Opties", "Naam2", StandaardNaam2)
  Spelers(3).Naam = IniGet("Opties", "Naam3", StandaardNaam3)
  Spelers(4).Naam = IniGet("Opties", "Naam4", StandaardNaam4)
End Sub

Sub OpslaanOpties()
  'Slaat op naar Ini; daarna nog IniSave doen!

  With Opties
    IniSet "Opties", "Animatiesnelheid", .AniSnelheid
    IniSet "Opties", "Aflopend", Abs(CInt(.AflopendSorteren))
    IniSet "Opties", "BreedUitspreiden", Abs(CInt(.BreedUitspreiden))
    IniSet "Opties", "Commentaar", Abs(CInt(.Commentaar))
    IniSet "Opties", "ComputersInHighScore", Abs(CInt(.ComputersInHighScore))
    IniSet "Opties", "Geluid", Abs(CInt(.Geluid))
    IniSet "Opties", "MaxAantKaarten", .MaxAantKaarten
    IniSet "Opties", "Optellen", Abs(CInt(.MeteenOptellen))
    IniSet "Opties", "Rondmaken", Abs(CInt(.RondmakenToegestaan))
    IniSet "Opties", "Spelhulp", Abs(CInt(.Spelhulp))
    IniSet "Opties", "Spelsnelheid", .SpelSnelheid
    IniSet "Opties", "IntroevenVerplicht", Abs(CInt(.IntroevenVerplicht))
    IniSet "Opties", "GroteKaarten", Abs(CInt(.GroteKaarten))
    
    IniSet "Opties", "PuntenPerSlag", CStr(.PuntenPerSlag)
    IniSet "Opties", "StrafpuntenPerSlag", CStr(.StrafpuntenPerSlag)
    IniSet "Opties", "FoutVoorspeldNulPunten", Abs(CInt(.FoutVoorspeldNulPunten))
    IniSet "Opties", "NegatieveScores", Abs(CInt(.NegatieveScores))
  End With
  IniSet "Opties", "Achterkant", AchterkantTitle
  IniSet "Opties", "Naam1", Spelers(1).Naam
  IniSet "Opties", "Naam2", Spelers(2).Naam
  IniSet "Opties", "Naam3", Spelers(3).Naam
  IniSet "Opties", "Naam4", Spelers(4).Naam
End Sub

Public Sub Schuif(ByVal Voorwerp As Object, ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer, ByVal TerugNaar As Integer, ByVal Bovenop As Boolean, Optional DoDoevents As Boolean = False)
  Dim StapNr As Integer
  Dim AantStap As Integer
  Dim HorAfst As Integer
  Dim VerAfst As Integer
  Dim HorStapGrootte As Integer
  Dim VerStapGrootte As Integer
 
  AantStap = 3 + 6 * (-Opties.AniSnelheid + 6) ^ 2
  HorAfst = X2 - X1
  VerAfst = Y2 - Y1
  HorStapGrootte = HorAfst \ AantStap
  VerStapGrootte = VerAfst \ AantStap
 
  'If Voorwerp.Visible = False Then
  Voorwerp.Visible = True
  'End If
  If Bovenop Then
    Voorwerp.ZOrder 0
  End If
  
  For StapNr = 0 To AantStap
    Voorwerp.Move X1 + StapNr * HorStapGrootte, Y1 + StapNr * VerStapGrootte
    If DoDoevents Then
      DoEvents
    End If
  Next StapNr

  Select Case TerugNaar
    Case 0
    Case 1
      Voorwerp.Visible = False
      Voorwerp.Move X1, Y1
    Case 2
      Voorwerp.Move X2, Y2
  End Select

End Sub

Public Sub RunFile(Filename As String)
  If Len(Dir(Filename)) Then
    Screen.MousePointer = vbArrowHourglass
    'Shell "Start " & Chr(34) & Filename & Chr(34), vbHide
    ShellExecute frmMain.hwnd, "open", Filename, vbNullString, "C:\", 1
    Screen.MousePointer = vbDefault
  Else
    MsgBox "Kan het bestand " & Filename & " niet vinden.", vbExclamation, "Fout"
  End If

End Sub

Public Sub DebugToIni(DebugText As String)
  Static TotalText As String
  
  TotalText = TotalText & ", " & DebugText
  If Len(TotalText) > 1000 Then
    TotalText = Mid(TotalText, Len(TotalText) - 1000)
  End If
  IniSet "Debug", "DebugText", TotalText
  IniSave
End Sub

Public Function SpelIs10openneer() As Boolean
  With Opties
    SpelIs10openneer = (.MaxAantKaarten = 10 And .FoutVoorspeldNulPunten And .StrafpuntenPerSlag = 0 And .RondmakenToegestaan And Not .IntroevenVerplicht And .PuntenPerSlag = 1)
  End With
End Function

Public Function SpelIsBoerenbridge() As Boolean
  With Opties
    SpelIsBoerenbridge = (.MaxAantKaarten = 13 And .StrafpuntenPerSlag > 0 And Not .RondmakenToegestaan And Not .IntroevenVerplicht)
  End With
End Function
