VERSION 5.00
Begin VB.UserControl ctlStatusBar 
   ClientHeight    =   1215
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6615
   ScaleHeight     =   1215
   ScaleWidth      =   6615
   Begin VB.PictureBox picPanel 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   0
      Left            =   4680
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   113
      TabIndex        =   1
      Top             =   30
      Width           =   1695
      Begin VB.Label lblPanel 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   240
         Index           =   0
         Left            =   15
         TabIndex        =   2
         Top             =   15
         Width           =   1665
      End
   End
   Begin VB.Label lblSimpleText 
      Height          =   255
      Left            =   45
      TabIndex        =   0
      Top             =   30
      Width           =   3135
   End
End
Attribute VB_Name = "ctlStatusBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim cNPanels As Integer
Public Event MouseDown(ByVal PanelIndex As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

Public Property Let NumPanels(ByVal Number As Integer)
Attribute NumPanels.VB_Description = "Dit is een test"
  Dim p As Integer
  If Number < 0 Then
    Err.Raise 0, "ctlStatusBar", "Invalid property value (NPanels = " & Number & ")"
  Else
    For p = Number To cNPanels - 1 'Delete some
      If p = 0 Then
        picPanel(0).Visible = False
      Else
        Unload picPanel(p)
        Unload lblPanel(p)
      End If
    Next p
    For p = cNPanels To Number - 1 'Add some
      If p > 0 Then
        Load picPanel(p)
        Load lblPanel(p)
        Set lblPanel(p).Container = picPanel(p)
      End If
      'DrawPanel p
    Next p
    cNPanels = Number
    PlacePanels
    For p = 0 To cNPanels - 1 'Draw them
      DrawPanel p
    Next p
    PropertyChanged "NumPanels"
  End If
End Property

Public Property Get NumPanels() As Integer
  NumPanels = cNPanels
End Property

Public Property Let PanelVisible(ByVal Index As Integer, ByVal Visible As Boolean)
  If Index < 0 Or Index > cNPanels Then
    Err.Raise 10000, "ctlStatusBar", "Invalid panel index (" & Index & ")"
  Else
    picPanel(Index).Visible = Visible
    lblPanel(Index).Visible = Visible
    PlacePanels
    PropertyChanged "PanelVisible"
  End If
End Property

Public Property Get PanelVisible(ByVal Index As Integer) As Boolean
  If Index < 0 Or Index > cNPanels Then
    Err.Raise 0, "ctlStatusBar", "Invalid panel index (" & Index & ")"
  Else
    PanelVisible = picPanel(Index).Visible
  End If
End Property

Public Property Let PanelText(ByVal Index As Integer, Text As String)
  If Index < 0 Or Index > cNPanels Then
    Err.Raise 0, "ctlStatusBar", "Invalid panel index (" & Index & ")"
  Else
    lblPanel(Index).Caption = Text
    PropertyChanged "PanelText"
  End If
End Property

Public Property Get PanelText(ByVal Index As Integer) As String
  If Index < 0 Or Index > cNPanels Then
    Err.Raise 0, "ctlStatusBar", "Invalid panel index (" & Index & ")"
  Else
    PanelText = lblPanel(Index).Caption
  End If
End Property

Public Property Let PanelWidth(ByVal Index As Integer, ByVal Width As Single)
  If Index < 0 Or Index > cNPanels Then
    Err.Raise 0, "ctlStatusBar", "Invalid panel index (" & Index & ")"
  Else
    picPanel(Index).Width = Width
    DrawPanel Index
    PlacePanels
    PropertyChanged "PanelWidth"
  End If
End Property

Public Property Get PanelWidth(ByVal Index As Integer) As Single
  If Index < 0 Or Index > cNPanels Then
    Err.Raise 0, "ctlStatusBar", "Invalid panel index (" & Index & ")"
  Else
    PanelWidth = picPanel(Index).Width
  End If
End Property

Public Property Let SimpleText(Text As String)
  lblSimpleText.Caption = Text
  PropertyChanged "SimpleText"
End Property

Public Property Get SimpleText() As String
  SimpleText = lblSimpleText.Caption
End Property

Public Property Get PanelToolTip(ByVal Index As Integer) As String
  If Index < 0 Or Index > cNPanels Then
    Err.Raise 0, "ctlStatusBar", "Invalid panel index (" & Index & ")"
  Else
    PanelToolTip = lblPanel(Index).ToolTipText
  End If
End Property

Public Property Let PanelToolTip(ByVal Index As Integer, NewToolTip As String)
  If Index < 0 Or Index > cNPanels Then
    Err.Raise 0, "ctlStatusBar", "Invalid panel index (" & Index & ")"
  Else
    lblPanel(Index).ToolTipText = NewToolTip
    PropertyChanged "PanelToolTip"
  End If
End Property

Public Property Get PanelWhatsThisHelpID(ByVal Index As Integer) As Integer
  If Index < 0 Or Index > cNPanels Then
    Err.Raise 0, "ctlStatusBar", "Invalid panel index (" & Index & ")"
  Else
    PanelWhatsThisHelpID = lblPanel(Index).WhatsThisHelpID
  End If
End Property

Public Property Let PanelWhatsThisHelpID(ByVal Index As Integer, NewWhatsThisHelpID As Integer)
  If Index < 0 Or Index > cNPanels Then
    Err.Raise 0, "ctlStatusBar", "Invalid panel index (" & Index & ")"
  Else
    lblPanel(Index).WhatsThisHelpID = NewWhatsThisHelpID
    PropertyChanged "PanelWhatsThisHelpID"
  End If
End Property

Public Sub PanelShowWhatsThis(ByVal Index As Integer)
  If Index < 0 Or Index > cNPanels Then
    Err.Raise 0, "ctlStatusBar", "Invalid panel index (" & Index & ")"
  Else
    lblPanel(Index).ShowWhatsThis 'werkt niet
  End If
End Sub

'***********
'* Private *
'***********

Private Sub PlacePanels()
  Dim p As Integer
  Dim Right As Single
  
  lblSimpleText.Top = (Height - lblSimpleText.Height) / 2
  Right = Width - 240
  For p = cNPanels - 1 To 0 Step -1
    With picPanel(p)
      .Left = Right - .Width - 60
      If .Visible Then
        Right = .Left
      End If
      .Height = Height - 60
      'lblPanel(p).Height = .ScaleHeight - 2
      lblPanel(p).Top = (.ScaleHeight - lblPanel(p).Height) / 2
      lblPanel(p).Width = .ScaleWidth - 2
    End With
  Next p
  'If cNPanels >= 0 Then
    'If picPanel(0).Left - 60 > 0 Then
      'lblSimpleText.Width = picPanel(0).Left - 60
      lblSimpleText.Width = ScaleWidth - 60
    'End If
  'End If
End Sub

Private Sub DrawPanel(ByVal Index As Integer)
  With picPanel(Index)
    .Cls
    picPanel(Index).Line (0, 0)-(.ScaleWidth - 1, 0), SystemColorConstants.vb3DShadow
    picPanel(Index).Line (0, 1)-(0, .ScaleHeight - 1), SystemColorConstants.vb3DShadow
    picPanel(Index).Line (0, .ScaleHeight - 1)-(.ScaleWidth, .ScaleHeight - 1), SystemColorConstants.vb3DHighlight
    picPanel(Index).Line (240, .ScaleHeight - 3)-(.ScaleWidth, .ScaleHeight + 3), QBColor(8), BF
    
    picPanel(Index).Line (.ScaleWidth - 1, 0)-(.ScaleWidth - 1, .ScaleHeight - 1), SystemColorConstants.vb3DHighlight
  End With
End Sub

Private Sub lblPanel_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent MouseDown(Index, Button, Shift, X, Y)
End Sub

Private Sub lblSimpleText_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent MouseDown(-1, Button, Shift, X, Y)
End Sub

Private Sub UserControl_Initialize()
  cNPanels = 0
  PlacePanels
End Sub

Private Sub UserControl_Resize()
  PlacePanels
End Sub
