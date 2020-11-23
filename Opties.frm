VERSION 5.00
Begin VB.Form frmAchterkant 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Kaartachterkant"
   ClientHeight    =   2655
   ClientLeft      =   1215
   ClientTop       =   1680
   ClientWidth     =   6015
   DrawWidth       =   2
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   HelpContextID   =   401
   Icon            =   "OPTIES.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2655
   ScaleWidth      =   6015
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdAnnuleren 
      Cancel          =   -1  'True
      Caption         =   "&Annuleren"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdExtra 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      MaskColor       =   &H00FF00FF&
      Picture         =   "OPTIES.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Klik hier voor extra informatie"
      Top             =   2160
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.DirListBox Dir1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   4560
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.FileListBox filAchterkant 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   4560
      Pattern         =   "*.bmp;*.gif;*.jpg;*.wmf"
      TabIndex        =   6
      Top             =   840
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.PictureBox picVenster 
      Height          =   1935
      Left            =   120
      ScaleHeight     =   125
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   381
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   120
      Width           =   5775
      Begin VB.HScrollBar hsbAchterkant 
         Height          =   255
         Left            =   0
         Max             =   5
         SmallChange     =   20
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1680
         Width           =   5640
      End
      Begin VB.PictureBox picPaneel 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00008000&
         BorderStyle     =   0  'None
         DrawWidth       =   2
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   1695
         Left            =   0
         ScaleHeight     =   113
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   376
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   0
         Width           =   5640
         Begin VB.Image imgAchterkant 
            Appearance      =   0  'Flat
            Height          =   1440
            Index           =   0
            Left            =   120
            Top             =   120
            Width           =   1065
         End
      End
   End
   Begin VB.CommandButton cmdOk 
      Appearance      =   0  'Flat
      Caption         =   "&Ok"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   0
      Top             =   2160
      Width           =   1215
   End
End
Attribute VB_Name = "frmAchterkant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AchterkantGeselecteerd As Integer
Dim nResAchterkanten As Integer

Private Sub cmdAnnuleren_Click()
  Me.Hide
End Sub

Private Sub cmdExtra_Click()
  MsgBox "Je kunt je eigen kaartachterkanten toevoegen. Teken een kaart van formaat 71x96 met Paint en sla deze op in de map " & _
         DirPlusBestand(CStr(App.Path), "Achterkant") & ".", vbInformation, "Kaartachterkanten"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyLeft Then
    If AchterkantGeselecteerd > 0 Then
      Kader AchterkantGeselecteerd, picPaneel.BackColor
      AchterkantGeselecteerd = AchterkantGeselecteerd - 1
      Kader AchterkantGeselecteerd, QBColor(10)
    End If
    SchuifInBereik AchterkantGeselecteerd
  ElseIf KeyCode = vbKeyRight Then
    If AchterkantGeselecteerd < filAchterkant.ListCount + nResAchterkanten - 1 Then
      Kader AchterkantGeselecteerd, picPaneel.BackColor
      AchterkantGeselecteerd = AchterkantGeselecteerd + 1
      Kader AchterkantGeselecteerd, QBColor(10)
    End If
    SchuifInBereik AchterkantGeselecteerd
  End If
End Sub

Sub SchuifInBereik(AchterkantNr As Integer)
  If imgAchterkant(AchterkantNr).Left - 8 < -picPaneel.Left Then
    hsbAchterkant.Value = imgAchterkant(AchterkantNr).Left - 8
  ElseIf imgAchterkant(AchterkantNr).Left + imgAchterkant(AchterkantNr).Width + 8 > -picPaneel.Left + picVenster.ScaleWidth Then
    hsbAchterkant.Value = imgAchterkant(AchterkantNr).Left + imgAchterkant(AchterkantNr).Width - picVenster.ScaleWidth + 8
  End If
End Sub
Private Sub Form_Load()
  Dim FrameNr As Integer
  Dim SpelerNr As Integer
 
  hsbAchterkant.Width = picVenster.ScaleWidth
  hsbAchterkant.Top = picVenster.ScaleHeight - hsbAchterkant.Height
  
  If Len(Dir(DirPlusBestand(CStr(App.Path), "Achterkant\NUL"))) = 0 Then
    MkDir DirPlusBestand(CStr(App.Path), "Achterkant")
  End If
  
  nResAchterkanten = 10
  
  ToonAchterkanten
End Sub

Sub ToonAchterkanten()
  Dim AchterkantDir As String
  Dim AchterkantTitleNu As String
  Dim AchterkantNr As Integer
  Dim ImageIndex As Integer
  
  frmAchterkant.MousePointer = vbHourglass
  frmAchterkant.Enabled = False
  
  AchterkantDir = DirPlusBestand(CStr(App.Path), "Achterkant")
  
  For ImageIndex = 0 To nResAchterkanten - 1
    If ImageIndex > 0 Then
      Load imgAchterkant(ImageIndex)
      imgAchterkant(ImageIndex).Left = imgAchterkant(ImageIndex - 1).Left + imgAchterkant(ImageIndex).Width + 10
    End If
    imgAchterkant(ImageIndex).Picture = LoadResPicture(200 + ImageIndex, vbResBitmap)
    imgAchterkant(ImageIndex).Tag = CStr(ImageIndex)
    imgAchterkant(ImageIndex).Visible = True
    If CStr(ImageIndex) = AchterkantTitle Then
      AchterkantGeselecteerd = ImageIndex
    End If
  Next ImageIndex
  
  If Len(Dir(AchterkantDir & "\NUL")) Then
    filAchterkant.Path = AchterkantDir
    If filAchterkant.ListCount > 0 Then
      
      For ImageIndex = nResAchterkanten To filAchterkant.ListCount - 1 + nResAchterkanten
        Load imgAchterkant(ImageIndex)
        imgAchterkant(ImageIndex).Left = imgAchterkant(ImageIndex - 1).Left + imgAchterkant(ImageIndex).Width + 10
        AchterkantTitleNu = filAchterkant.List(ImageIndex - nResAchterkanten)
        imgAchterkant(ImageIndex).Stretch = True
        imgAchterkant(ImageIndex).Picture = LoadPicture(DirPlusBestand(AchterkantDir, AchterkantTitleNu))
        imgAchterkant(ImageIndex).Tag = AchterkantTitleNu
        imgAchterkant(ImageIndex).ToolTipText = AchterkantTitleNu
        imgAchterkant(ImageIndex).Visible = True
        If LCase(AchterkantTitleNu) = LCase(AchterkantTitle) Then
          AchterkantGeselecteerd = ImageIndex
        End If
        DoEvents
      Next ImageIndex
    End If
  End If

  picPaneel.Width = 10 + (filAchterkant.ListCount + nResAchterkanten) * (imgAchterkant(0).Width + 10)
  If picPaneel.Width < picVenster.ScaleWidth Then
    picPaneel.Width = picVenster.ScaleWidth
  End If
  hsbAchterkant.Max = picPaneel.Width - picVenster.ScaleWidth
  hsbAchterkant.LargeChange = picVenster.ScaleWidth
  
  If hsbAchterkant.Max < imgAchterkant(AchterkantGeselecteerd).Left - 8 Then
    hsbAchterkant.Value = hsbAchterkant.Max
  Else
    hsbAchterkant.Value = imgAchterkant(AchterkantGeselecteerd).Left - 8
  End If

  Kader AchterkantGeselecteerd, QBColor(10)
  
  frmAchterkant.MousePointer = vbDefault
  frmAchterkant.Enabled = True

End Sub
Private Sub hsbAchterkant_Change()
  picPaneel.Left = -hsbAchterkant.Value
End Sub
Private Sub hsbAchterkant_Scroll()
  picPaneel.Left = -hsbAchterkant.Value

End Sub

Private Sub imgAchterkant_Click(Index As Integer)
  Kader AchterkantGeselecteerd, picPaneel.BackColor
  AchterkantGeselecteerd = Index
  Kader AchterkantGeselecteerd, QBColor(10)
  picPaneel.SetFocus
End Sub
Sub Kader(AchterkantNr As Integer, Kleur As Long)
  Dim X1, X2, Y1, Y2
 
  X1 = imgAchterkant(AchterkantNr).Left - 2
  X2 = X1 + imgAchterkant(AchterkantNr).Width + 4
  Y1 = imgAchterkant(AchterkantNr).Top - 2
  Y2 = Y1 + imgAchterkant(AchterkantNr).Height + 4
  picPaneel.Line (X1, Y1)-(X2, Y2), Kleur, B

End Sub
Private Sub cmdOk_Click()
  Dim i As Integer
  'Dim AchterkantTitle As String
  'Dim AchterkantBestand As String
  Dim SpelerNr As Integer
  
  AchterkantTitle = imgAchterkant(AchterkantGeselecteerd).Tag
  If IsNumeric(AchterkantTitle) Then
  Else
    AchterkantBestand = DirPlusBestand(CStr(App.Path), "Achterkant\" & AchterkantTitle)
  End If
  
  'IniSet "Opties", "Achterkant", AchterkantTitle
 
  frmMain.imgInHanden(13).Picture = imgAchterkant(AchterkantGeselecteerd).Picture
  For i = 13 To 51
    If KaartImageGeladen(i) Then
      frmMain.imgInHanden(i).Picture = frmMain.imgInHanden(13).Picture
    End If
  Next i

  frmAchterkant.Hide
End Sub

Private Sub imgAchterkant_DblClick(Index As Integer)
  cmdOk_Click
End Sub
