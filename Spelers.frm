VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSpelers 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Spelers"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7230
   Icon            =   "Spelers.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picGedragKleur 
      Height          =   375
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   32
      Top             =   6120
      Width           =   375
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton cmdAnnuleren 
      Cancel          =   -1  'True
      Caption         =   "&Annuleren"
      Height          =   375
      Left            =   3960
      TabIndex        =   5
      Top             =   6120
      Width           =   1215
   End
   Begin MSComctlLib.ImageList imlSpelers 
      Left            =   4920
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
            Picture         =   "Spelers.frx":014A
            Key             =   "Mens"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Spelers.frx":0466
            Key             =   "Computer"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   360
      TabIndex        =   20
      Top             =   720
      Width           =   4575
      Begin VB.Label lblSpelerType 
         Caption         =   "Speler 1 (dat ben jij)"
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
         TabIndex        =   21
         Top             =   120
         Width           =   3615
      End
      Begin VB.Image picSpelerType 
         Height          =   480
         Left            =   0
         Top             =   0
         Width           =   480
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Naam"
      Height          =   855
      Left            =   360
      TabIndex        =   18
      Top             =   1320
      Width           =   4575
      Begin VB.TextBox txtNaam 
         Height          =   285
         Left            =   1680
         TabIndex        =   0
         Text            =   "Naam2"
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label1 
         Caption         =   "Naam:"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame fraGedrag 
      Caption         =   "Gedrag"
      Height          =   3495
      Left            =   2520
      TabIndex        =   8
      Top             =   2280
      Width           =   4575
      Begin VB.CommandButton cmdAlleGelijk 
         Caption         =   "&Alle gelijk"
         Height          =   375
         Left            =   360
         TabIndex        =   27
         Top             =   2880
         Width           =   1215
      End
      Begin VB.CommandButton cmdWillekeurig 
         Caption         =   "&Willekeurig"
         Height          =   375
         Left            =   1680
         TabIndex        =   26
         Top             =   2880
         Width           =   1215
      End
      Begin VB.CommandButton cmdStandaard 
         Caption         =   "&Standaard"
         Height          =   375
         Left            =   3000
         TabIndex        =   25
         Top             =   2880
         Width           =   1215
      End
      Begin MSComctlLib.Slider sldVoorspellen 
         Height          =   255
         Left            =   1680
         TabIndex        =   1
         Top             =   360
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   1
         Min             =   -3
         Max             =   3
      End
      Begin MSComctlLib.Slider sldSlagNietNemen 
         Height          =   255
         Left            =   1680
         TabIndex        =   2
         Top             =   840
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   1
         Min             =   -3
         Max             =   3
      End
      Begin MSComctlLib.Slider sldHogeWeggooien 
         Height          =   255
         Left            =   1680
         TabIndex        =   3
         Top             =   1320
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   1
         Min             =   -3
         Max             =   3
      End
      Begin MSComctlLib.Slider sldVariatie 
         Height          =   255
         Left            =   1680
         TabIndex        =   7
         Top             =   2280
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   1
         Max             =   4
      End
      Begin MSComctlLib.Slider sldOpkomenMetTroef 
         Height          =   255
         Left            =   1680
         TabIndex        =   28
         Top             =   1800
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   1
         Min             =   -3
         Max             =   3
      End
      Begin VB.Label Label15 
         Caption         =   "Weinig"
         Height          =   255
         Left            =   1680
         TabIndex        =   30
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "Weinig"
         Height          =   255
         Left            =   1680
         TabIndex        =   16
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label16 
         Caption         =   "Opkomen met troef:"
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Caption         =   "Vaak"
         Height          =   255
         Left            =   3240
         TabIndex        =   29
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "Veel"
         Height          =   255
         Left            =   3240
         TabIndex        =   24
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label12 
         Caption         =   "Geen"
         Height          =   255
         Left            =   1680
         TabIndex        =   23
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "Variatie:"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "Hoge weggooien:"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Vaak"
         Height          =   255
         Left            =   3240
         TabIndex        =   15
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Vaak"
         Height          =   255
         Left            =   3240
         TabIndex        =   14
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Weinig"
         Height          =   255
         Left            =   1680
         TabIndex        =   13
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Slag duiken:"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Hoog"
         Height          =   255
         Left            =   3840
         TabIndex        =   11
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Laag"
         Height          =   255
         Left            =   1680
         TabIndex        =   10
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Voorspellen:"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   480
         Width           =   1455
      End
   End
   Begin MSComctlLib.TabStrip tabSpelers 
      Height          =   5895
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   10398
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
