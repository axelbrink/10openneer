VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSpelers1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Spelers"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7230
   Icon            =   "Spelers1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picGedragKleur 
      Height          =   375
      Left            =   840
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   4
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton cmdStandaardGedrag 
      Caption         =   "&Standaard"
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   3720
      Width           =   975
   End
   Begin VB.Frame fraSpelers 
      Caption         =   "Speler 2"
      Height          =   1455
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   6015
      Begin VB.CommandButton Command1 
         Caption         =   "&Verwijderen"
         Height          =   375
         Left            =   3840
         TabIndex        =   8
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton cmdNieuwGedrag 
         Caption         =   "&Nieuw"
         Height          =   375
         Left            =   2640
         TabIndex        =   7
         Top             =   480
         Width           =   1095
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   240
         TabIndex        =   5
         Text            =   "Combo1"
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "Naam:"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton cmdAnnuleren 
      Cancel          =   -1  'True
      Caption         =   "&Annuleren"
      Height          =   375
      Left            =   3960
      TabIndex        =   1
      Top             =   6120
      Width           =   1215
   End
   Begin MSComctlLib.ImageList imlSpelers 
      Left            =   6480
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Spelers1.frx":014A
            Key             =   "Mens"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Spelers1.frx":0466
            Key             =   "Computer"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmSpelers1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim GedragTemp(2 To 4) As GedragType
Dim NamenTemp(1 To 4) As String

Private Sub cmdAlleGelijk_Click()
  Dim BronSpelerNr As Integer
  Dim SpelerNr As Integer
  
  BronSpelerNr = tabSpelers.SelectedItem.Index
  For SpelerNr = 2 To 4
    If SpelerNr <> BronSpelerNr Then
      GedragTemp(SpelerNr) = GedragTemp(BronSpelerNr)
    End If
  Next SpelerNr
  MsgBox "Gedrag gekopieerd naar de andere computerspelers.", vbInformation, "Gedrag"

End Sub

Private Sub cmdAnnuleren_Click()
  Unload Me
End Sub

Private Sub cmdOk_Click()
  Dim SpelerNr As Integer
  
  For SpelerNr = 1 To 4
    Spelers(SpelerNr).Naam = NamenTemp(SpelerNr)
    frmMain.lblNaam(SpelerNr - 1).Caption = Spelers(SpelerNr).Naam
    frmMain.lblNaamSchaduw(SpelerNr - 1).Caption = Spelers(SpelerNr).Naam
    
    IniSet "Opties", "Naam" & Trim(CStr(SpelerNr)), Spelers(SpelerNr).Naam
    If SpelerNr >= 2 Then
      Spelers(SpelerNr).Gedrag = GedragTemp(SpelerNr)
      IniSet "Opties", "Gedrag" & Trim(CStr(SpelerNr)), Gedrag2String(Spelers(SpelerNr).Gedrag)
    End If
  Next SpelerNr
  
  frmMain.ToonScoreblok False
  
  Unload Me
End Sub

Private Sub cmdStandaard_Click()
  Dim SpelerNr As Integer
  
  For SpelerNr = 2 To 4
    GedragTemp(SpelerNr).Voorspellen = 0
    GedragTemp(SpelerNr).SlagNietNemen = 0
    GedragTemp(SpelerNr).HogeWeggooien = 0
    GedragTemp(SpelerNr).OpkomenMetLageTroef = 0
    GedragTemp(SpelerNr).Variatie = 0
  Next SpelerNr
  
  tabSpelers_Click
End Sub

Private Sub cmdWillekeurig_Click()
  Dim SpelerNr As Integer
  
  Randomize Timer
  
  For SpelerNr = 2 To 4
    GedragTemp(SpelerNr).Voorspellen = Int(7 * Rnd - 3)
    GedragTemp(SpelerNr).SlagNietNemen = Int(7 * Rnd - 3)
    GedragTemp(SpelerNr).HogeWeggooien = Int(7 * Rnd - 3)
    GedragTemp(SpelerNr).OpkomenMetLageTroef = Int(7 * Rnd - 3)
    GedragTemp(SpelerNr).Variatie = Int(4 * Rnd)
  Next SpelerNr
  
  tabSpelers_Click

End Sub

Private Sub Form_Load()
  Dim SpelerNr As Integer
  
  For SpelerNr = 1 To 4
    tabSpelers.Tabs(SpelerNr).Caption = Spelers(SpelerNr).Naam
    NamenTemp(SpelerNr) = Spelers(SpelerNr).Naam
    If SpelerNr >= 2 Then
      GedragTemp(SpelerNr) = Spelers(SpelerNr).Gedrag
    End If
  Next SpelerNr
  
  tabSpelers_Click
End Sub

Private Sub sldHogeWeggooien_Click()
  GedragTemp(tabSpelers.SelectedItem.Index).HogeWeggooien = sldHogeWeggooien.Value
End Sub

Private Sub sldOpkomenMetTroef_Click()
  GedragTemp(tabSpelers.SelectedItem.Index).OpkomenMetLageTroef = sldOpkomenMetTroef.Value
End Sub

Private Sub sldSlagNietNemen_Click()
  GedragTemp(tabSpelers.SelectedItem.Index).SlagNietNemen = sldSlagNietNemen.Value
End Sub

Private Sub sldVariatie_Click()
  GedragTemp(tabSpelers.SelectedItem.Index).Variatie = sldVariatie.Value
End Sub

Private Sub sldVoorspellen_Click()
  GedragTemp(tabSpelers.SelectedItem.Index).Voorspellen = sldVoorspellen.Value
End Sub

Private Sub tabSpelers_Click()
  Dim Index As Integer
  
  Index = tabSpelers.SelectedItem.Index
  Select Case Index
    Case 1
      lblSpelerType.Caption = NamenTemp(Index) & " (Speler 1, dat ben jij)"
      picSpelerType.Picture = imlSpelers.ListImages("Mens").Picture
      fraGedrag.Visible = False
    Case 2
      lblSpelerType.Caption = NamenTemp(Index) & " (Speler 2, computer west)"
      picSpelerType.Picture = imlSpelers.ListImages("Computer").Picture
      fraGedrag.Visible = True
    Case 3
      lblSpelerType.Caption = NamenTemp(Index) & " (Speler 3, computer noord)"
      picSpelerType.Picture = imlSpelers.ListImages("Computer").Picture
      fraGedrag.Visible = True
    Case 4
      lblSpelerType.Caption = NamenTemp(Index) & " (Speler 4, computer oost)"
      picSpelerType.Picture = imlSpelers.ListImages("Computer").Picture
      fraGedrag.Visible = True
  End Select
  txtNaam.Text = NamenTemp(Index)
  
  If Index >= 2 Then
    sldVoorspellen.Value = GedragTemp(Index).Voorspellen
    sldSlagNietNemen.Value = GedragTemp(Index).SlagNietNemen
    sldHogeWeggooien.Value = GedragTemp(Index).HogeWeggooien
    sldOpkomenMetTroef.Value = GedragTemp(Index).OpkomenMetLageTroef
    sldVariatie.Value = GedragTemp(Index).Variatie
  End If
End Sub

Private Sub txtNaam_Change()
  Dim Index As Integer
  
  Index = tabSpelers.SelectedItem.Index
  NamenTemp(Index) = txtNaam.Text
  tabSpelers.Tabs(Index).Caption = NamenTemp(Index)
End Sub

Function Gedrag2Kleur(Gedrag As GedragType) As Long

End Function
