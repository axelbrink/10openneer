VERSION 5.00
Begin VB.Form frmSnelheid 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Snelheid"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3375
   HelpContextID   =   402
   Icon            =   "Snelheid.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   3375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "Spelverloop"
      Height          =   1695
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   3135
      Begin Vb10openneer.ctlSlider sldSpelSnelheid 
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   480
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   450
         Max             =   6
         Continuous      =   0   'False
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Snel"
         Height          =   255
         Left            =   2280
         TabIndex        =   7
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Langzaam (voor langzame mensen)"
         Height          =   615
         Left            =   240
         TabIndex        =   6
         Top             =   840
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Animaties"
      Height          =   1695
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3135
      Begin Vb10openneer.ctlSlider sldAniSnelheid 
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Width           =   2655
         _ExtentX        =   2778
         _ExtentY        =   661
         Max             =   6
         Continuous      =   0   'False
      End
      Begin VB.Label Label1 
         Caption         =   "Langzaam (voor snelle computers)"
         Height          =   615
         Left            =   240
         TabIndex        =   4
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Snel"
         Height          =   255
         Left            =   2280
         TabIndex        =   3
         Top             =   840
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdAnnuleren 
      Cancel          =   -1  'True
      Caption         =   "&Annuleren"
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   3720
      Width           =   1215
   End
End
Attribute VB_Name = "frmSnelheid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAnnuleren_Click()
  Unload Me
End Sub

Private Sub cmdOk_Click()
  Opties.AniSnelheid = sldAniSnelheid.Value
  Opties.SpelSnelheid = sldSpelSnelheid.Value
  
  Unload Me
End Sub

Private Sub Form_Load()
  sldAniSnelheid.Value = Opties.AniSnelheid
  sldSpelSnelheid.Value = Opties.SpelSnelheid
End Sub
