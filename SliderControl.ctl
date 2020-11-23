VERSION 5.00
Begin VB.UserControl ctlSlider 
   AutoRedraw      =   -1  'True
   ClientHeight    =   375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1935
   ScaleHeight     =   25
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   129
   Begin VB.PictureBox picHandle 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   960
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   17
      TabIndex        =   0
      Top             =   0
      Width           =   255
   End
   Begin VB.Line linHor1 
      BorderColor     =   &H80000010&
      X1              =   0
      X2              =   128
      Y1              =   11
      Y2              =   11
   End
   Begin VB.Line linHor2 
      BorderColor     =   &H80000014&
      X1              =   0
      X2              =   128
      Y1              =   12
      Y2              =   12
   End
End
Attribute VB_Name = "ctlSlider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private cValue As Single
Private cInternValue As Single
Private cMin As Single
Private cMax As Single
Private cDragStartX As Single
Private cScrolling As Boolean
Private cContinuous As Boolean

Public Event Change()
Public Event Scroll()

Private Sub picHandle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 1 Then
    cDragStartX = X
  End If
End Sub

Private Sub picHandle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim NewX As Single
  
  If Button = 1 Then
    cScrolling = True
    NewX = picHandle.Left + X - cDragStartX
    If NewX < 0 Then
      NewX = 0
    ElseIf NewX > ScaleWidth - picHandle.Width Then
      NewX = ScaleWidth - picHandle.Width
    End If

    picHandle.Move NewX
    HandlePosToInternValue
    InternToExternValue
    
    If Not cContinuous Then
      MakeDiscrete 'Just in case someone would like to check the (discrete) value while scrolling
      ExternToInternValue
    End If
    
    RaiseEvent Scroll
  End If
End Sub

Private Sub picHandle_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 1 And cScrolling Then
    cScrolling = False
    InternValueToHandlePos
    PropertyChanged "Value"
    RaiseEvent Change
  End If
End Sub

Private Sub ExternToInternValue()
  If cMax - cMin = 0 Then
    cInternValue = 0
  Else
    cInternValue = (cValue - cMin) / (cMax - cMin)
  End If
End Sub

Private Sub InternToExternValue()
  cValue = cInternValue * (cMax - cMin) + cMin
End Sub

Private Sub InternValueToHandlePos()
  picHandle.Move cInternValue * (ScaleWidth - picHandle.Width)
End Sub

Private Sub HandlePosToInternValue()
  cInternValue = picHandle.Left / (ScaleWidth - picHandle.Width)
End Sub

Public Property Get Value() As Single
  Value = cValue
End Property

Public Property Let Value(ByVal NewValue As Single)
  If NewValue > cMax Then
    Err.Raise 1000, "SliderControl", "New Value is above Max"
  ElseIf NewValue < cMin Then
    Err.Raise 1000, "SliderControl", "New Value is below Min"
  Else
    cValue = NewValue
    ExternToInternValue
    InternValueToHandlePos
    PropertyChanged "Value"
    RaiseEvent Change
  End If
End Property

Public Property Get Min() As Single
  Min = cMin
End Property

Public Property Let Min(ByVal NewMin As Single)
  If NewMin > cValue Then
    Err.Raise 1000, "SliderControl", "New Min is above Value"
  ElseIf NewMin > cMax Then
    Err.Raise 1000, "SliderControl", "New Min is above Max"
  Else
    CheckRangeForDiscrete cMin, cMax
    cMin = NewMin
    ExternToInternValue
    InternValueToHandlePos
    PropertyChanged "Min"
  End If
End Property

Public Property Get Max() As Single
  Max = cMax
End Property

Public Property Let Max(ByVal NewMax As Single)
  If NewMax < cValue Then
    Err.Raise 1000, "SliderControl", "New Max is below Value"
  ElseIf NewMax < cMin Then
    Err.Raise 1000, "SliderControl", "New Max is below Min"
  Else
    CheckRangeForDiscrete cMin, cMax
    cMax = NewMax
    ExternToInternValue
    InternValueToHandlePos
    PropertyChanged "Max"
  End If
End Property

Public Property Get Continuous() As Boolean
  Continuous = cContinuous
End Property

Public Property Let Continuous(ByVal NewContinuous As Boolean)
  If cContinuous And Not NewContinuous Then
    'From continuous to discrete
    CheckRangeForDiscrete cMin, cMax 'Possibly raise error
    cContinuous = NewContinuous
    MakeDiscrete
  
    ExternToInternValue
    InternValueToHandlePos
    PropertyChanged "Value"
    RaiseEvent Change
  Else
    cContinuous = NewContinuous
  End If
  PropertyChanged "Continuous"
End Property

Private Sub CheckRangeForDiscrete(ByVal pMin As Single, ByVal pMax As Single)
  'Can the range be divided in discrete steps?
  '(it must contain at least one round number)
  If Int(pMin) = Int(pMax) And Int(pMin) <> pMin Then
    Err.Raise 1000, "ctlSlider", "Discrete range is not possible because there are no round numbers between Min and Max"
  End If
End Sub

Private Sub MakeDiscrete()
  'Makes extern value discrete
  'Assumes that it is possible
  cValue = CInt(cValue)
  If cValue < cMin Then
    cValue = cValue + 1
  ElseIf cValue > cMax Then
    cValue = cValue - 1
  End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  cMin = PropBag.ReadProperty("Min", 0)
  cMax = PropBag.ReadProperty("Max", 10)
  cValue = PropBag.ReadProperty("Value", 0)
  cContinuous = PropBag.ReadProperty("Continuous", True)
  
  ExternToInternValue
  InternValueToHandlePos
End Sub

Private Sub UserControl_Resize()
  Dim Right As Single
  Dim Bottom As Single
  
  linHor1.X2 = ScaleWidth - 1
  linHor2.X2 = ScaleWidth - 1
  linHor1.Y1 = Int(ScaleHeight / 2)
  linHor1.Y2 = Int(ScaleHeight / 2)
  linHor2.Y1 = Int(ScaleHeight / 2) + 1
  linHor2.Y2 = Int(ScaleHeight / 2) + 1
  
  picHandle.Height = ScaleHeight
  
  Right = picHandle.ScaleWidth - 1
  Bottom = picHandle.ScaleHeight - 1
  
  picHandle.Cls
  picHandle.Line (0, 0)-(Right - 1, 0), vb3DHighlight
  picHandle.Line (0, 1)-(0, Bottom - 1), vb3DHighlight
  picHandle.Line (1, Bottom - 1)-(Right, Bottom - 1), vb3DShadow
  picHandle.Line (Right - 1, 1)-(Right - 1, Bottom - 1), vb3DShadow
  
  picHandle.Line (0, Bottom)-(Right + 1, Bottom), vb3DDKShadow
  picHandle.Line (Right, 0)-(Right, Bottom), vb3DDKShadow

  cScrolling = False
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  PropBag.WriteProperty "Min", cMin, 0
  PropBag.WriteProperty "Max", cMax, 10
  PropBag.WriteProperty "Value", cValue, 0
  PropBag.WriteProperty "Continuous", cContinuous, True
End Sub
