Attribute VB_Name = "modIni"
Option Explicit
'Laatste wijziging: 23-5-99

Public Type IniVar
  Naam As String
  Waarde As String
End Type

Global IniVars(1 To 300) As IniVar
Global IniLen As Integer

Function IniGet(Section As String, Key As String, Default As Variant) As Variant
  Dim Loc As Integer

  Loc = ZoekInIni(Section, Key)
  If Loc >= 1 Then
    IniGet = IniVars(Loc).Waarde
  Else
    IniGet = Default
  End If
End Function
Sub IniSet(Section As String, Key As String, Setting As Variant)
  Dim Loc As Integer
  Dim IniVarNr As Integer
  Dim KandidaatIniVarNr As Integer
  
  Loc = ZoekInIni(Section, Key)
  If Loc >= 1 Then
    IniVars(Loc).Waarde = CStr(Setting)
  ElseIf Loc <= -1 Then
    Loc = Abs(Loc)
    KandidaatIniVarNr = Loc + 1
    For IniVarNr = Loc + 1 To IniLen
      If IniVars(IniVarNr).Naam = "[" Then
        Exit For
      End If
      If IniVars(IniVarNr).Naam <> "" Then
        KandidaatIniVarNr = IniVarNr + 1
      End If
    Next IniVarNr
    IniLen = IniLen + 1
    For IniVarNr = IniLen To KandidaatIniVarNr + 1 Step -1
      IniVars(IniVarNr) = IniVars(IniVarNr - 1)
    Next IniVarNr
    IniVars(KandidaatIniVarNr).Naam = Key
    IniVars(KandidaatIniVarNr).Waarde = CStr(Setting)
  Else
    If IniLen = 0 Then
      IniLen = 1
    Else
      IniLen = IniLen + 2
    End If
    IniVars(IniLen).Naam = "["
    IniVars(IniLen).Waarde = Section
    
    IniLen = IniLen + 1
    IniVars(IniLen).Naam = Key
    IniVars(IniLen).Waarde = CStr(Setting)
    
    KandidaatIniVarNr = IniLen
  End If
  'if
End Sub
Function ZoekInIni(Section As String, Key As String) As Integer
  Dim IniVarNr As Integer
  Dim GoedeSectie As Boolean
  Dim Gevonden As Boolean
  Dim UpSection As String
  Dim UpKey As String
  
  UpSection = UCase(Section)
  UpKey = UCase(Key)
  
  GoedeSectie = False
  ZoekInIni = 0
  For IniVarNr = 1 To IniLen
    If IniVars(IniVarNr).Naam = "[" Then
      If UCase(IniVars(IniVarNr).Waarde) = UpSection Then
        GoedeSectie = True
        ZoekInIni = -IniVarNr
      Else
        GoedeSectie = False
      End If
    ElseIf GoedeSectie And UCase(IniVars(IniVarNr).Naam) = UpKey Then
      ZoekInIni = IniVarNr
      Exit For
    End If
  Next IniVarNr

End Function

Sub IniLoad()
  Dim Regel As String
  Dim IsLoc As Integer
  
  IniLen = 0
  If Len(Dir(IniBestand)) Then
    Open IniBestand For Input As #1
    Do Until EOF(1)
      IniLen = IniLen + 1
      Line Input #1, Regel
      Regel = Trim(Regel)
      If Len(Regel) Then
        If Left(Regel, 1) = "[" Then
          IniVars(IniLen).Naam = "["
          If Len(Regel) >= 2 Then
            IniVars(IniLen).Waarde = Mid(Regel, 2)
            If Right(IniVars(IniLen).Waarde, 1) = "]" Then
              If Len(Regel) >= 3 Then
                IniVars(IniLen).Waarde = Left(IniVars(IniLen).Waarde, Len(IniVars(IniLen).Waarde) - 1)
              Else
                MsgBox "Er is een syntaxfout in het bestand " & IniBestand _
                & " op regelnummer " & IniLen & ".", vbExclamation + vbOKOnly, "Fout"
              End If
            End If
          Else
            MsgBox "Er is een syntaxfout in het bestand " & IniBestand _
            & " op regelnummer " & IniLen & ".", vbExclamation + vbOKOnly, "Fout"
          End If
        ElseIf Left(Regel, 1) = ";" Then
          IniVars(IniLen).Naam = ";"
          If Len(Regel) >= 2 Then
            IniVars(IniLen).Waarde = Mid(Regel, 2)
          End If
        Else
          IsLoc = InStr(Regel, "=")
          If IsLoc = 0 Then
            MsgBox "Er is een syntaxfout in het bestand " & IniBestand _
            & " op regelnummer " & IniLen & ": Het teken '=' ontbreekt.", vbExclamation + vbOKOnly, "Fout"
          ElseIf IsLoc = 1 Then
            MsgBox "Er is een syntaxfout in het bestand " & IniBestand _
            & " op regelnummer " & IniLen & ": De variabelenaam ontbreekt.", vbExclamation + vbOKOnly, "Fout"
          Else
            IniVars(IniLen).Naam = RTrim(Left(Regel, IsLoc - 1))
            If Len(Regel) = IsLoc Then
              IniVars(IniLen).Waarde = ""
            Else
              IniVars(IniLen).Waarde = LTrim(Mid(Regel, IsLoc + 1))
            End If
          End If
        End If
      End If
    Loop
    Close #1
  End If
End Sub
Sub IniSave()
  Dim IniVarNr As Integer
  Dim Regel As String
  
  Open IniBestand For Output As #1
  For IniVarNr = 1 To IniLen
    If IniVars(IniVarNr).Naam = "[" Then
      Regel = "[" & IniVars(IniVarNr).Waarde & "]"
    ElseIf IniVars(IniVarNr).Naam = ";" Then
      Regel = ";" & IniVars(IniVarNr).Waarde
    ElseIf IniVars(IniVarNr).Naam = "" Then
      Regel = ""
    Else
      Regel = IniVars(IniVarNr).Naam & " = " & IniVars(IniVarNr).Waarde
    End If
    Print #1, Regel
  Next IniVarNr
  Close #1
End Sub

Sub IniDel(Section As String, Key As String)
  Dim Loc As Integer
  Dim IniVarNr As Integer
  
  Loc = ZoekInIni(Section, Key)
  If Loc >= 1 Then 'Niet gevonden -> er gebeurt niets
    For IniVarNr = Loc To IniLen - 1
      IniVars(IniVarNr).Naam = IniVars(IniVarNr + 1).Naam
      IniVars(IniVarNr).Waarde = IniVars(IniVarNr + 1).Waarde
    Next IniVarNr
    IniLen = IniLen - 1
  End If
End Sub
