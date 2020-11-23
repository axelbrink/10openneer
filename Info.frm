VERSION 5.00
Begin VB.Form frmInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Over 10 op en neer"
   ClientHeight    =   4935
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5055
   ClipControls    =   0   'False
   HelpContextID   =   501
   Icon            =   "Info.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   480
      Left            =   240
      Picture         =   "Info.frx":014A
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   1
      Top             =   240
      Width           =   480
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "Oké."
      Default         =   -1  'True
      Height          =   375
      Left            =   1860
      TabIndex        =   0
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Geluiden zijn afkomstig van verschillende cd's van 'Best of Select Multimedia'"
      Height          =   375
      Left            =   600
      TabIndex        =   11
      Top             =   3600
      Width           =   3975
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Met dank aan Kim van Wijngaarden voor de kaartachterkant met het schaap"
      Height          =   495
      Left            =   480
      TabIndex        =   10
      Top             =   3120
      Width           =   4095
   End
   Begin VB.Label lblInnoSetup 
      Caption         =   "InnoSetup"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   3480
      MouseIcon       =   "Info.frx":02C0
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Het installatieprogramma is gemaakt met"
      Height          =   225
      Left            =   240
      TabIndex        =   8
      Top             =   2880
      Width           =   3135
   End
   Begin VB.Label lblHomepage 
      Alignment       =   2  'Center
      Caption         =   "homepage"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1320
      MouseIcon       =   "Info.frx":0412
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   2160
      Width           =   2415
   End
   Begin VB.Label lblEmail 
      Alignment       =   2  'Center
      Caption         =   "e-mail"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1320
      MouseIcon       =   "Info.frx":0564
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   1560
      Width           =   2415
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   120
      X2              =   4920
      Y1              =   2625
      Y2              =   2625
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Kijk op Internet voor de nieuwste versie:"
      Height          =   225
      Left            =   240
      TabIndex        =   5
      Top             =   1920
      Width           =   4605
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
      Caption         =   "[Versie]"
      Height          =   225
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   4605
   End
   Begin VB.Label lblRegistratie 
      Alignment       =   2  'Center
      Caption         =   "Dit programma is gratis."
      Height          =   255
      Left            =   1560
      TabIndex        =   3
      Top             =   4080
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   960
      Picture         =   "Info.frx":06B6
      Top             =   120
      Width           =   3750
   End
   Begin VB.Label lblDescription 
      Alignment       =   2  'Center
      Caption         =   "Door Axel Brink. Ontwikkeling begon in 1997. Stuur je commentaar, opmerkingen, suggesties, vragen en fanmail naar:"
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   4605
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      BorderWidth     =   2
      X1              =   135
      X2              =   4920
      Y1              =   2640
      Y2              =   2640
   End
End
Attribute VB_Name = "frmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOk_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  Dim AppDate As Date
  Dim AppExePath As String
  
  AppExePath = DirPlusBestand(App.Path, App.EXEName & ".exe")
  If Dir(AppExePath) = "" Then
    MsgBox "Kan " & AppExePath & " niet vinden! Probeer 10 op en neer opnieuw te starten.", vbCritical, "Fout"
  Else
    AppDate = FileDateTime(AppExePath)
  End If
  
  Me.Caption = "Over " & App.Title
  lblVersion.Caption = "Versie " & App.Major & "." & Format(App.Minor, "#00")
  If App.Revision > 0 Then
    lblVersion.Caption = lblVersion.Caption & " revisie " & App.Revision
  End If
  lblVersion.Caption = lblVersion.Caption & " (" & Format(AppDate, "d mmmm yyyy") & ")"
  'lblHoudbaarheid.Caption = Format(NietGebruikenNa, "d mmmm yyyy")
  lblEmail.Caption = Email
  lblEmail.ForeColor = LinkColor
  lblHomepage.Caption = Homepage
  lblHomepage.ForeColor = LinkColor
  lblInnoSetup.ForeColor = LinkColor
End Sub

Private Sub lblEmail_Click()
  Me.MousePointer = vbHourglass
  StartURL "mailto:" & Email & "?SUBJECT=10 op en neer"
  Me.MousePointer = vbDefault
End Sub

Private Sub lblEmail_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  lblEmail.ForeColor = LinkActiveColor
End Sub

Private Sub lblEmail_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  lblEmail.ForeColor = LinkColor
End Sub

Private Sub lblHomepage_Click()
  Me.MousePointer = vbHourglass
  StartURL Homepage
  Me.MousePointer = vbDefault
End Sub

Private Sub lblHomepage_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  lblHomepage.ForeColor = LinkActiveColor
End Sub

Private Sub lblHomepage_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  lblHomepage.ForeColor = LinkColor
End Sub

Private Sub lblInnoSetup_Click()
  Me.MousePointer = vbHourglass
  StartURL Homepage
  Me.MousePointer = vbDefault
End Sub

Private Sub lblInnoSetup_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  lblInnoSetup.ForeColor = LinkActiveColor
End Sub

Private Sub lblInnoSetup_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  lblInnoSetup.ForeColor = LinkColor
End Sub
