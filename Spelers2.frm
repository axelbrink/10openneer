VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSpelers 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Spelers [experimenteel]"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4935
   Icon            =   "Spelers2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraGedrag 
      Caption         =   "Gedrag van deze speler"
      Height          =   1455
      Left            =   360
      TabIndex        =   10
      Top             =   3120
      Width           =   4215
      Begin VB.CommandButton cmdGeefStandaardGedrag 
         Caption         =   "Geef standaardgedrag"
         Height          =   375
         Left            =   2160
         TabIndex        =   14
         Top             =   840
         Width           =   1815
      End
      Begin VB.CommandButton cmdGeefWillekeurigGedrag 
         Caption         =   "Geef willekeurig gedrag"
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   840
         Width           =   1815
      End
      Begin VB.PictureBox picGedragKleur 
         AutoRedraw      =   -1  'True
         Height          =   375
         Left            =   2160
         ScaleHeight     =   315
         ScaleWidth      =   1755
         TabIndex        =   11
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Gedragkleurcode:"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   420
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   2280
      TabIndex        =   0
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton cmdAnnuleren 
      Cancel          =   -1  'True
      Caption         =   "&Annuleren"
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   5040
      Width           =   1215
   End
   Begin MSComctlLib.ImageList imlSpelers 
      Left            =   4080
      Top             =   720
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
            Picture         =   "Spelers2.frx":014A
            Key             =   "Mens"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Spelers2.frx":0466
            Key             =   "Computer"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   360
      TabIndex        =   5
      Top             =   720
      Width           =   4215
      Begin VB.Label lblWieIsDat 
         Caption         =   "Dat ben jij"
         Height          =   255
         Left            =   600
         TabIndex        =   15
         Top             =   240
         Width           =   3615
      End
      Begin VB.Label lblNaam 
         Caption         =   "Axel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   6
         Top             =   0
         Width           =   3375
      End
      Begin VB.Image picSpelerType 
         Height          =   480
         Left            =   0
         Top             =   0
         Width           =   480
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Speler"
      Enabled         =   0   'False
      Height          =   1695
      Left            =   360
      TabIndex        =   3
      Top             =   1320
      Width           =   4215
      Begin VB.CommandButton Command2 
         Caption         =   "&Verwijder speler"
         Height          =   375
         Left            =   2160
         TabIndex        =   9
         Top             =   1080
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Nieuwe speler"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   1080
         Width           =   1815
      End
      Begin VB.ComboBox cboNaam 
         Height          =   315
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Width           =   3735
      End
      Begin VB.Label Label1 
         Caption         =   "Kies een speler uit de lijst of wijzig de naam:"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   3735
      End
   End
   Begin MSComctlLib.TabStrip tabSpelers 
      Height          =   4815
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   8493
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Speler 1"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Speler 2"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Speler 3"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Speler 4"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmSpelers"
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

Private Sub cmdGeefStandaardGedrag_Click()
  Dim SpelerNr As Integer
  
  SpelerNr = tabSpelers.SelectedItem.Index
    
  GedragTemp(SpelerNr).Voorspellen = 5
  GedragTemp(SpelerNr).SlagNietNemen = 5
  GedragTemp(SpelerNr).HogeWeggooien = 5
  GedragTemp(SpelerNr).OpkomenMetLageTroef = 5
  GedragTemp(SpelerNr).Variatie = 5
  
  tabSpelers_Click

End Sub

Private Sub cmdGeefWillekeurigGedrag_Click()
  Dim SpelerNr As Integer
  
  Randomize Timer
  
  SpelerNr = tabSpelers.SelectedItem.Index
  
  GedragTemp(SpelerNr).Voorspellen = Int(11 * Rnd) 'Dus 0..10
  GedragTemp(SpelerNr).SlagNietNemen = Int(11 * Rnd)
  GedragTemp(SpelerNr).HogeWeggooien = Int(11 * Rnd)
  GedragTemp(SpelerNr).OpkomenMetLageTroef = Int(11 * Rnd)
  GedragTemp(SpelerNr).Variatie = Int(11 * Rnd)
  
  tabSpelers_Click

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

Private Sub tabSpelers_Click()
  Dim Index As Integer
  
  Index = tabSpelers.SelectedItem.Index
  Select Case Index
    Case 1
      lblNaam.Caption = NamenTemp(Index)
      lblWieIsDat = "Dat ben jij"
    Case 2
      lblNaam.Caption = NamenTemp(Index)
      lblWieIsDat = "Computer west"
    Case 3
      lblNaam.Caption = NamenTemp(Index)
      lblWieIsDat = "Computer noord"
    Case 4
      lblNaam.Caption = NamenTemp(Index)
      lblWieIsDat = "Computer oost"
  End Select
  'txtNaam.Text = NamenTemp(Index)
  
  If Index = 1 Then
    fraGedrag.Visible = False
    picSpelerType.Picture = imlSpelers.ListImages("Mens").Picture
  Else
    fraGedrag.Visible = True
    picSpelerType.Picture = imlSpelers.ListImages("Computer").Picture
    'sldVoorspellen.Value = GedragTemp(Index).Voorspellen
    'sldSlagNietNemen.Value = GedragTemp(Index).SlagNietNemen
    'sldHogeWeggooien.Value = GedragTemp(Index).HogeWeggooien
    'sldOpkomenMetTroef.Value = GedragTemp(Index).OpkomenMetLageTroef
    'sldVariatie.Value = GedragTemp(Index).Variatie
    'picGedragKleur.BackColor = Gedrag2Kleur(GedragTemp(Index))
    ToonGedragKleurCode GedragTemp(Index)
  End If
End Sub

Private Sub txtNaam_Change()
  Dim Index As Integer
  
  Index = tabSpelers.SelectedItem.Index
  NamenTemp(Index) = txtNaam.Text
  tabSpelers.Tabs(Index).Caption = NamenTemp(Index)
End Sub

Private Sub ToonGedragKleurCode(Gedrag As GedragType)
  Dim R As Integer, G As Integer, B As Integer
  
  'Maximale factor is 25.5: 25.5 * 10 = 255
  'R = 13 * Gedrag.HogeWeggooien + 12.5 * Gedrag.Voorspellen
  'G = 13 * Gedrag.OpkomenMetLageTroef + 12.5 * Gedrag.Variatie
  'B = 25.5 * Gedrag.SlagNietNemen
  'picGedragKleur.BackColor = RGB(R, G, B)
  
  R = 25.5 * Gedrag.HogeWeggooien
  G = 25.5 * Gedrag.OpkomenMetLageTroef
  B = 25.5 * Gedrag.SlagNietNemen
  picGedragKleur.Line (0, 0)-(picGedragKleur.ScaleWidth \ 2, picGedragKleur.Height), RGB(R, G, B), BF
  
  R = 25.5 * Gedrag.Voorspellen
  G = 25.5 * Gedrag.Variatie
  B = 25.5 * Gedrag.SlagNietNemen
  picGedragKleur.Line (picGedragKleur.ScaleWidth \ 2, 0)-(picGedragKleur.Width, picGedragKleur.Height), RGB(R, G, B), BF
  
End Sub
