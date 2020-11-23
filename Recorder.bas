Attribute VB_Name = "modRecorder"
Option Explicit

Public recWieBegint As Integer
Public recRonde As Integer            'In welke ronde moet nu iets opgeslagen worden
Public recSpelerNum As Integer        'Wie staat als laatste nog in het archief; 1=die moest opkomen, 2=volgende, ...
Public recSlagNr As Integer           '0: bezig met voorspellen
Public recKaartenOntvangen(1 To 25, 1 To 4, 1 To 13) As Kaart 'Ronde, Speler, Slag
Public recTroeven(1 To 25) As Kaart                           'Ronde
Public recVoorspellingen(1 To 25, 1 To 4) As Integer          'Ronde, Speler
Public recVoorspellingenGoed(1 To 25, 1 To 4) As Boolean      'Ronde, Speler
Public recKaartenGekozen(1 To 25, 1 To 4, 1 To 13) As Integer 'Ronde, Speler, Slag
'Public recNuOpkomen(1 To 25, 1 To 13) As Integer              'Ronde, Slag
Public recWasInHerhaling As Boolean

Public Sub recOpnameWissen()
  recWieBegint = 0
  recRonde = 1
  recSpelerNum = 0
  recSlagNr = 0
'  Erase recKaartenOntvangen
'  Erase recTroeven
'  Erase recVoorspellingen
'  Erase recKaartenGekozen
'  Erase recNuOpkomen
End Sub

'Public Function recInHerhaling(ByVal VoorspellenKlaar As Boolean, ByVal SpelerNr As Integer) As Boolean
'  'True als het einde van de opname nog niet is bereikt
'
'  'recInHerhaling = Not (Ronde = recAantRonden And recSlagNr = SlagNr And VoorspellenKlaar = recVoorspellenKlaar And SpelerNr = recSpelerNr)
'  'recInHerhaling = (Ronde <= recAantRonden And SlagNr <= recSlagNr And VoorspellenKlaar >= recVoorspellenKlaar And SpelerNr <> VolgendeSpeler(recSpelerNr))
'  'recInHerhaling = Not (Ronde = recAantRonden And SlagNr = recSlagNr And VoorspellenKlaar = recVoorspellenKlaar And SpelerNr = VolgendeSpeler(recSpelerNr))
'  recInHerhaling = False '#
'End Function

Public Function recInHerhaling() As Boolean
'  return not (ronde=recAantRonden and slagnr=recSlagNr and spelernum >
  Dim inherh As Boolean
  
  If Ronde < recRonde Then
    inherh = True
  ElseIf Ronde > recRonde Then
    inherh = False
  Else  'Ronde = recRonde
    If SlagNr < recSlagNr Then
      inherh = True
    ElseIf SlagNr > recSlagNr Then
      inherh = False
    Else  'SlagNr = recSlagNr
      If SpelerNum < recSpelerNum Then
        inherh = True
      Else
        inherh = False
      End If
    End If
  End If
  If recWasInHerhaling And Not inherh Then
    MsgBox "De herhaling is beëindigd.", vbInformation + vbOKOnly, "Herhaling"
  End If
  recWasInHerhaling = inherh
  frmMain.lblHerhaling.Caption = IIf(inherh, "Herhaling", "")
'  If inherh Then Stop
  recInHerhaling = inherh
End Function

Public Function recInHerhalingTest(ByVal RondeTest As Integer, ByVal SlagNrTest As Integer, ByVal SpelerNumTest As Integer) As Boolean
  If RondeTest < recRonde Then
    recInHerhalingTest = True
    Exit Function
  ElseIf RondeTest > recRonde Then
    recInHerhalingTest = False
    Exit Function
  Else  'Ronde = recAantRonden
    If SlagNrTest < recSlagNr Then
      recInHerhalingTest = True
      Exit Function
    ElseIf SlagNrTest > recSlagNr Then
      recInHerhalingTest = False
      Exit Function
    Else  'SlagNr = recSlagNr
      If SpelerNumTest < recSpelerNum Then
        recInHerhalingTest = True
        Exit Function
      Else
        recInHerhalingTest = False
        Exit Function
      End If
    End If
  End If
End Function

Public Sub recOpslaan(Bestandsnaam As String)
  'MsgBox Bestandsnaam
  If Dir(Bestandsnaam) <> "" Then
    Kill Bestandsnaam
  End If
  
  Dim Fnum As Integer
  Fnum = FreeFile()
  
  Open Bestandsnaam For Binary As #Fnum
  WriteHeaderToFile Fnum, "10 op en neer savegame", CByte(App.Major), CByte(App.Minor), 30
  Put #Fnum, 31, CByte(recWieBegint)
  Put #Fnum, 32, CByte(recRonde)
  Put #Fnum, 33, CByte(recSpelerNum)
  Put #Fnum, 34, CByte(recSlagNr)
  
  Dim saveRonde As Integer
  Dim saveSpeler As Integer
  Dim saveSlag As Integer
  
  For saveRonde = 1 To 25
    Put #Fnum, , KaartToByte(recTroeven(saveRonde))
    For saveSpeler = 1 To 4
      Put #Fnum, , CByte(recVoorspellingen(saveRonde, saveSpeler))
      Put #Fnum, , CByte(recVoorspellingenGoed(saveRonde, saveSpeler))
      For saveSlag = 1 To 13
        Put #Fnum, , KaartToByte(recKaartenOntvangen(saveRonde, saveSpeler, saveSlag))
        Put #Fnum, , CByte(recKaartenGekozen(saveRonde, saveSpeler, saveSlag))
      Next saveSlag
    Next saveSpeler
  Next saveRonde
  
  Close Fnum
End Sub

Private Sub WriteHeaderToFile(FileNum As Integer, ID As String, VersionMajor As Byte, VersionMinor As Byte, HeaderSize As Integer)
  'Writes a header consisting of a string followed by two bytes for the version number.
  
  If HeaderSize < 4 Then Stop
  
  Dim p As Integer
  For p = 1 To HeaderSize - 4
    If ID <> "" Then
      Put #FileNum, p, CByte(Asc(Left(ID, 1)))
    Else
      Put #FileNum, p, CByte(0)
    End If
    ID = Mid(ID, 2)
  Next p
  Put #FileNum, p, CByte(13)
  Put #FileNum, p + 1, CByte(10)
  Put #FileNum, p + 2, VersionMajor
  Put #FileNum, p + 3, VersionMinor
End Sub

Private Function KaartToByte(EenKaart As Kaart) As Byte
  If EenKaart.Kleur = 0 Or EenKaart.Getal = 0 Then
    KaartToByte = 0
  Else
    KaartToByte = 13 * (EenKaart.Kleur - 1) + (EenKaart.Getal - 2) + 1
  End If
End Function

Private Function ByteToKaart(KaartByte As Byte) As Kaart
  Dim k As Kaart
  If KaartByte = 0 Then
    k.Kleur = 0
    k.Getal = 0
  Else
    k.Kleur = (KaartByte - 1) \ 13 + 1
    k.Getal = (KaartByte - 1) Mod 13 + 2
  End If
  ByteToKaart = k
End Function

Public Function recOpenen(Bestandsnaam As String)
  Dim Fnum As Integer
  Fnum = FreeFile()
  
  Open Bestandsnaam For Binary As #Fnum
  'WriteHeaderToFile Fnum, "10 op en neer savegame", CByte(App.Major), CByte(App.Minor), 30
  Dim ReadByte As Byte
  Get #Fnum, 31, ReadByte: recWieBegint = CInt(ReadByte)
  Get #Fnum, 32, ReadByte: recRonde = CInt(ReadByte)
  Get #Fnum, 33, ReadByte: recSpelerNum = CInt(ReadByte)
  Get #Fnum, 34, ReadByte: recSlagNr = CInt(ReadByte)
  
  Dim saveRonde As Integer
  Dim saveSpeler As Integer
  Dim saveSlag As Integer
  
  For saveRonde = 1 To 25
    Get #Fnum, , ReadByte: recTroeven(saveRonde) = ByteToKaart(ReadByte)
    For saveSpeler = 1 To 4
      Get #Fnum, , ReadByte: recVoorspellingen(saveRonde, saveSpeler) = CInt(ReadByte)
      Get #Fnum, , ReadByte: recVoorspellingenGoed(saveRonde, saveSpeler) = CBool(ReadByte)
      For saveSlag = 1 To 13
        Get #Fnum, , ReadByte: recKaartenOntvangen(saveRonde, saveSpeler, saveSlag) = ByteToKaart(ReadByte)
        Get #Fnum, , ReadByte: recKaartenGekozen(saveRonde, saveSpeler, saveSlag) = CInt(ReadByte)
      Next saveSlag
    Next saveSpeler
  Next saveRonde
  
  Close Fnum
End Function

'Public Function BepaalSpelersVoorspeld() As Boolean()
'  'Retourneert array(1..4): welke speler is klaar met voorspellen
'  'in laatst opgenomen ronde
'
'  Dim spv(1 To 4) As Boolean
'  Dim SpelerNr As Integer
'  Dim SpelerNum As Integer
'
'  For SpelerNr = 1 To 4
'    spv(SpelerNr) = False
'  Next SpelerNr
'
'  'klopt niet meer? --> SpelerNr = (recWieBegint + recAantRonden + 2) Mod 4 + 1
'  SpelerNum = 1
'
'  If recSlagNr = 0 And recRonde > 0 Then 'Er is een voorspelling opgeslagen
'    spv(SpelerNr) = True
'    Do Until SpelerNum >= recSpelerNum
'      SpelerNr = VolgendeSpeler(SpelerNr)
'      SpelerNum = SpelerNum + 1
'      spv(SpelerNr) = True
'    Loop
'  End If
'
'  BepaalSpelersVoorspeld = spv
'End Function
