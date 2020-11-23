VERSION 5.00
Begin VB.UserControl ctlPraatwolkje 
   BackStyle       =   0  'Transparent
   ClientHeight    =   2295
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2775
   ScaleHeight     =   2295
   ScaleWidth      =   2775
   Begin VB.Timer timKnipper 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1800
      Top             =   480
   End
   Begin VB.CommandButton cmdStopHulp 
      Caption         =   "Stop uitleggen"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      ToolTipText     =   "Klik hier als je geen hulp meer nodig hebt"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdDoorgaan 
      Caption         =   "&Doorgaan"
      Height          =   375
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label lblTekst 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Praatwolkje"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1815
      WordWrap        =   -1  'True
   End
   Begin VB.Shape shpPraatwolkje 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   1695
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   2055
   End
End
Attribute VB_Name = "ctlPraatwolkje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event Click()
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event Doorgaan()
Public Event StopHulp()

Private cIsHulp As Boolean
Private cToonDoorgaan As Boolean

Private Sub cmdDoorgaan_Click()
  RaiseEvent Doorgaan
End Sub

Private Sub cmdStopHulp_Click()
  RaiseEvent StopHulp
End Sub

Private Sub lblTekst_Click()
  RaiseEvent Click
End Sub

Private Sub timKnipper_Timer()
  If cmdDoorgaan.BackColor = vbButtonFace Then
    cmdDoorgaan.BackColor = vbHighlight
  Else
    cmdDoorgaan.BackColor = vbButtonFace
  End If
End Sub

Private Sub UserControl_Click()
  RaiseEvent Click
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  lblTekst.Caption = PropBag.ReadProperty("Caption", "Praatwolkje")
  cIsHulp = PropBag.ReadProperty("IsHulp", False)
  cToonDoorgaan = PropBag.ReadProperty("ToonDoorgaan", False)
  MaakWolkje False
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  PropBag.WriteProperty "Caption", lblTekst.Caption, "Praatwolkje"
  PropBag.WriteProperty "IsHulp", cIsHulp, False
  PropBag.WriteProperty "ToonDoorgaan", cToonDoorgaan, False
End Sub

Private Sub UserControl_Resize()
  If Width < 240 Then
    Width = 240
  End If
  If Height < 240 Then
    Height = 240
  End If
  shpPraatwolkje.Width = Width
  'shpPraatwolkje.Height = Height
  lblTekst.Width = Width - 240
  'lblTekst.Left = (Width - lblTekst.Width) / 2
  lblTekst.Height = Height - 240
  MaakWolkje False
End Sub

Public Property Get Caption() As String
  Caption = lblTekst.Caption
End Property

Public Property Let Caption(NewCaption As String)
  lblTekst.Caption = NewCaption
  PropertyChanged "Caption"
  MaakWolkje True
End Property

Public Property Get IsHulp() As Boolean
  IsHulp = cIsHulp
End Property

Public Property Let IsHulp(ByVal NewIsHulp As Boolean)
  cIsHulp = NewIsHulp
  PropertyChanged "IsHulp"
  If cToonDoorgaan And Not cIsHulp Then
    cToonDoorgaan = False
    PropertyChanged "ToonDoorgaan"
  End If
  MaakWolkje True
End Property

Public Property Get ToonDoorgaan() As Boolean
  ToonDoorgaan = cToonDoorgaan
End Property

Public Property Let ToonDoorgaan(ByVal NewToonDoorgaan As Boolean)
  cToonDoorgaan = NewToonDoorgaan
  PropertyChanged "ToonDoorgaan"
  If cToonDoorgaan And Not cIsHulp Then
    cIsHulp = True
    PropertyChanged "IsHulp"
  End If
  timKnipper.Enabled = cToonDoorgaan
  If Not cToonDoorgaan Then
    cmdDoorgaan.BackColor = vbButtonFace
  End If
  
  MaakWolkje True
End Property

Private Sub MaakWolkje(ByVal HoogteAanpassen As Boolean)
  'Maakt het wolkje van de goede hoogte en toont knoppen.
  'Als HoogteAanpassen, dan labelhoogte bepaalt hoogte
  'Anders: hoogte bepaalt de hoogte van het label
  
  Dim Hoogte As Long
  Dim LabelHoogte As Long
  Dim KnoppenHoogte As Long
  
  KnoppenHoogte = 0
  If cIsHulp Then
    KnoppenHoogte = KnoppenHoogte + cmdStopHulp.Height + 120
  End If
  If cToonDoorgaan Then
    KnoppenHoogte = KnoppenHoogte + cmdDoorgaan.Height + 120
  End If
  
  If HoogteAanpassen Then 'hoger maken voor knoppen
    Hoogte = lblTekst.Height + KnoppenHoogte + 240
    Height = Hoogte
  Else 'label aanpassen aan hoogte
    Hoogte = Height
    LabelHoogte = Hoogte - KnoppenHoogte - 240
    If LabelHoogte < 0 Then
      Hoogte = Hoogte - LabelHoogte 'Past niet - toch hoger maken
      Height = Hoogte
      LabelHoogte = 0
    End If
    lblTekst.Height = LabelHoogte 'Hoogte - KnoppenHoogte - 240
  End If
  
  shpPraatwolkje.Height = Hoogte
  lblTekst.Left = (Width - lblTekst.Width) / 2

  cmdStopHulp.Top = Hoogte - cmdStopHulp.Height - 120
  cmdStopHulp.Left = (Width - cmdStopHulp.Width) / 2
  cmdDoorgaan.Top = cmdStopHulp.Top - cmdDoorgaan.Height - 120
  cmdDoorgaan.Left = (Width - cmdDoorgaan.Width) / 2
  
  cmdStopHulp.Visible = cIsHulp
  cmdDoorgaan.Visible = cToonDoorgaan
End Sub
