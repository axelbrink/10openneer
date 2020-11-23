VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00008000&
   Caption         =   "10 op en neer"
   ClientHeight    =   6315
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9615
   DrawWidth       =   5
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   HelpContextID   =   101
   Icon            =   "10opneer.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6315
   ScaleWidth      =   9615
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8280
      Top             =   5280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer timLaatsteSlag 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   7560
      Top             =   3960
   End
   Begin Vb10openneer.ctlStatusBar StatusBar 
      Height          =   330
      Left            =   0
      TabIndex        =   48
      Top             =   5985
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   582
   End
   Begin VB.PictureBox picPraatwolkjePunt 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   9120
      Picture         =   "10opneer.frx":1272
      ScaleHeight     =   165
      ScaleWidth      =   165
      TabIndex        =   42
      Top             =   5040
      Visible         =   0   'False
      Width           =   165
   End
   Begin Vb10openneer.ctlPraatwolkje Praatwolkje 
      Height          =   480
      Left            =   8280
      TabIndex        =   47
      Top             =   4560
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   847
      Caption         =   "Hallo!"
   End
   Begin VB.Frame fraSpelen 
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5895
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Visible         =   0   'False
      Width           =   6770
      Begin VB.Frame fraVoorspellen 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   1335
         Left            =   2400
         TabIndex        =   43
         Top             =   2400
         Visible         =   0   'False
         WhatsThisHelpID =   1003
         Width           =   1935
         Begin VB.CommandButton cmdVoorspellen 
            Appearance      =   0  'Flat
            Caption         =   "11"
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
            Index           =   11
            Left            =   1440
            TabIndex        =   49
            ToolTipText     =   "Het aantal slagen voorspellen"
            Top             =   960
            Width           =   495
         End
         Begin VB.PictureBox picPijltje 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            BorderStyle     =   0  'None
            Height          =   150
            Left            =   0
            Picture         =   "10opneer.frx":134C
            ScaleHeight     =   150
            ScaleWidth      =   150
            TabIndex        =   46
            Top             =   0
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.CommandButton cmdVoorspellen 
            Appearance      =   0  'Flat
            Caption         =   "10"
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
            Index           =   10
            Left            =   960
            TabIndex        =   10
            ToolTipText     =   "Het aantal slagen voorspellen"
            Top             =   960
            Width           =   495
         End
         Begin VB.CommandButton cmdVoorspellen 
            Appearance      =   0  'Flat
            Caption         =   "9"
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
            Index           =   9
            Left            =   480
            TabIndex        =   9
            ToolTipText     =   "Het aantal slagen voorspellen"
            Top             =   960
            Width           =   495
         End
         Begin VB.CommandButton cmdVoorspellen 
            Appearance      =   0  'Flat
            Caption         =   "8"
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
            Index           =   8
            Left            =   0
            TabIndex        =   8
            ToolTipText     =   "Het aantal slagen voorspellen"
            Top             =   960
            Width           =   495
         End
         Begin VB.CommandButton cmdVoorspellen 
            Appearance      =   0  'Flat
            Caption         =   "7"
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
            Index           =   7
            Left            =   1440
            TabIndex        =   7
            ToolTipText     =   "Het aantal slagen voorspellen"
            Top             =   600
            Width           =   495
         End
         Begin VB.CommandButton cmdVoorspellen 
            Appearance      =   0  'Flat
            Caption         =   "6"
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
            Index           =   6
            Left            =   960
            TabIndex        =   6
            ToolTipText     =   "Het aantal slagen voorspellen"
            Top             =   600
            Width           =   495
         End
         Begin VB.CommandButton cmdVoorspellen 
            Appearance      =   0  'Flat
            Caption         =   "5"
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
            Index           =   5
            Left            =   480
            TabIndex        =   5
            ToolTipText     =   "Het aantal slagen voorspellen"
            Top             =   600
            Width           =   495
         End
         Begin VB.CommandButton cmdVoorspellen 
            Appearance      =   0  'Flat
            Caption         =   "4"
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
            Index           =   4
            Left            =   0
            TabIndex        =   4
            ToolTipText     =   "Het aantal slagen voorspellen"
            Top             =   600
            Width           =   495
         End
         Begin VB.CommandButton cmdVoorspellen 
            Appearance      =   0  'Flat
            Caption         =   "3"
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
            Index           =   3
            Left            =   1440
            TabIndex        =   3
            ToolTipText     =   "Het aantal slagen voorspellen"
            Top             =   240
            Width           =   495
         End
         Begin VB.CommandButton cmdVoorspellen 
            Appearance      =   0  'Flat
            Caption         =   "2"
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
            Index           =   2
            Left            =   960
            TabIndex        =   2
            ToolTipText     =   "Het aantal slagen voorspellen"
            Top             =   240
            Width           =   495
         End
         Begin VB.CommandButton cmdVoorspellen 
            Appearance      =   0  'Flat
            Caption         =   "1"
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
            Index           =   1
            Left            =   480
            TabIndex        =   1
            ToolTipText     =   "Het aantal slagen voorspellen"
            Top             =   240
            Width           =   495
         End
         Begin VB.CommandButton cmdVoorspellen 
            Appearance      =   0  'Flat
            Caption         =   "0"
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
            Index           =   0
            Left            =   0
            TabIndex        =   0
            ToolTipText     =   "Het aantal slagen voorspellen"
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lblHoeveelSlagen 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00008000&
            BackStyle       =   0  'Transparent
            Caption         =   "Hoeveel slagen?"
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
            Height          =   255
            Left            =   240
            TabIndex        =   44
            Top             =   0
            Width           =   1455
         End
         Begin VB.Label lblHoeveelSlagenSchaduw 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00008000&
            Caption         =   "Hoeveel slagen?"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   255
            TabIndex        =   45
            Top             =   15
            Width           =   1455
         End
      End
      Begin VB.TextBox txtNaamWijzig 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   1920
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lblHerhaling 
         BackColor       =   &H00008000&
         Caption         =   "Herhaling"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   50
         Top             =   120
         Width           =   1335
      End
      Begin VB.Image imgInHanden 
         Height          =   1440
         Index           =   26
         Left            =   2880
         Stretch         =   -1  'True
         Top             =   120
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Image imgInHanden 
         Height          =   1440
         Index           =   39
         Left            =   5640
         Stretch         =   -1  'True
         Top             =   2400
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Image imgInHanden 
         Height          =   1440
         Index           =   13
         Left            =   120
         Stretch         =   -1  'True
         Top             =   2400
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Image imgInHanden 
         Height          =   1440
         Index           =   0
         Left            =   2880
         Stretch         =   -1  'True
         Top             =   4440
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Label lblNaam 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   "Naam4"
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
         Height          =   255
         Index           =   3
         Left            =   5640
         MousePointer    =   3  'I-Beam
         TabIndex        =   22
         ToolTipText     =   "Speler 4"
         Top             =   3840
         Width           =   1095
      End
      Begin VB.Image imgKaartAanwijs 
         Height          =   75
         Left            =   2880
         Picture         =   "10opneer.frx":141E
         ToolTipText     =   "Druk op spatie om deze kaart op te gooien"
         Top             =   4380
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.Label lblNaam 
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   "Naam2"
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
         Height          =   255
         Index           =   1
         Left            =   120
         MousePointer    =   3  'I-Beam
         TabIndex        =   17
         ToolTipText     =   "Speler 2"
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Image imgSlagen 
         Height          =   180
         Index           =   0
         Left            =   3300
         Top             =   4200
         WhatsThisHelpID =   1005
         Width           =   180
      End
      Begin VB.Image imgSlagen 
         Height          =   180
         Index           =   13
         Left            =   1275
         Top             =   3060
         WhatsThisHelpID =   1005
         Width           =   180
      End
      Begin VB.Image imgSlagen 
         Height          =   180
         Index           =   26
         Left            =   3300
         Top             =   1620
         WhatsThisHelpID =   1005
         Width           =   180
      End
      Begin VB.Image imgSlagen 
         Height          =   180
         Index           =   39
         Left            =   5385
         Top             =   3060
         WhatsThisHelpID =   1005
         Width           =   180
      End
      Begin VB.Image imgOpTafel 
         Appearance      =   0  'Flat
         Height          =   1440
         Index           =   3
         Left            =   3495
         Stretch         =   -1  'True
         Top             =   2295
         Width           =   1065
      End
      Begin VB.Image imgOpTafel 
         Appearance      =   0  'Flat
         Height          =   1440
         Index           =   2
         Left            =   2895
         Stretch         =   -1  'True
         Top             =   1905
         Width           =   1065
      End
      Begin VB.Image imgOpTafel 
         Appearance      =   0  'Flat
         Height          =   1440
         Index           =   1
         Left            =   2295
         Stretch         =   -1  'True
         Top             =   2295
         Width           =   1065
      End
      Begin VB.Image imgOpTafel 
         Appearance      =   0  'Flat
         Height          =   1440
         Index           =   0
         Left            =   2895
         Stretch         =   -1  'True
         Top             =   2685
         Width           =   1065
      End
      Begin VB.Label lblNaam 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   "Naam1"
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
         Height          =   255
         Index           =   0
         Left            =   1800
         MousePointer    =   3  'I-Beam
         TabIndex        =   16
         ToolTipText     =   "Speler 1"
         Top             =   5640
         Width           =   1095
      End
      Begin VB.Label lblNaam 
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   "Naam3"
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
         Height          =   255
         Index           =   2
         Left            =   3960
         MousePointer    =   3  'I-Beam
         TabIndex        =   15
         ToolTipText     =   "Speler 3"
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label lblNaamSchaduw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00008000&
         Caption         =   "Naam1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   1815
         TabIndex        =   23
         ToolTipText     =   "Speler 1"
         Top             =   5655
         Width           =   1095
      End
      Begin VB.Label lblNaamSchaduw 
         BackColor       =   &H00008000&
         Caption         =   "Naam2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   135
         TabIndex        =   24
         ToolTipText     =   "Speler 1"
         Top             =   2175
         Width           =   1095
      End
      Begin VB.Label lblNaamSchaduw 
         BackColor       =   &H00008000&
         Caption         =   "Naam3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   3975
         TabIndex        =   25
         ToolTipText     =   "Speler 1"
         Top             =   135
         Width           =   1095
      End
      Begin VB.Label lblNaamSchaduw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00008000&
         Caption         =   "Naam4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   3
         Left            =   5655
         TabIndex        =   26
         ToolTipText     =   "Speler 1"
         Top             =   3855
         Width           =   1095
      End
   End
   Begin VB.TextBox txtNetwerkStatus 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   6000
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   40
      Text            =   "10opneer.frx":14B4
      Top             =   1800
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Timer timResize 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   7440
      Top             =   3960
   End
   Begin VB.Timer timSpelhulp 
      Interval        =   1000
      Left            =   7320
      Top             =   3960
   End
   Begin VB.Timer timTimeout 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   7200
      Top             =   3960
   End
   Begin VB.Timer timVrouw 
      Interval        =   3000
      Left            =   7080
      Top             =   3960
   End
   Begin VB.PictureBox picScore 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
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
      Height          =   4095
      Left            =   6840
      ScaleHeight     =   273
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   177
      TabIndex        =   21
      Top             =   120
      WhatsThisHelpID =   1001
      Width           =   2655
      Begin VB.Line linScore4 
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   272
      End
      Begin VB.Line linScore1 
         X1              =   0
         X2              =   176
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line linScore2 
         BorderColor     =   &H0000FF00&
         X1              =   176
         X2              =   176
         Y1              =   0
         Y2              =   272
      End
      Begin VB.Line linScore3 
         BorderColor     =   &H0000FF00&
         X1              =   176
         X2              =   0
         Y1              =   272
         Y2              =   272
      End
   End
   Begin VB.Frame fraTroef 
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1575
      Left            =   6840
      TabIndex        =   18
      Top             =   4320
      Visible         =   0   'False
      Width           =   1275
      Begin VB.PictureBox picTroef 
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         ForeColor       =   &H80000008&
         Height          =   1440
         Left            =   0
         ScaleHeight     =   94
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   69
         TabIndex        =   20
         TabStop         =   0   'False
         ToolTipText     =   "Dit is de troef"
         Top             =   120
         WhatsThisHelpID =   1004
         Width           =   1065
      End
      Begin VB.PictureBox picTroefTekst 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1080
         Picture         =   "10opneer.frx":14EE
         ScaleHeight     =   375
         ScaleWidth      =   150
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   1185
         Width           =   150
      End
   End
   Begin VB.Frame fraEinde 
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   6015
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Visible         =   0   'False
      Width           =   6735
      Begin VB.CommandButton cmdNieuwSpel 
         Caption         =   "&Nieuw Spel"
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
         Left            =   4920
         TabIndex        =   12
         Top             =   5040
         Width           =   1455
      End
      Begin VB.CommandButton cmdAfsluiten 
         Cancel          =   -1  'True
         Caption         =   "&Afsluiten"
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
         Left            =   4920
         TabIndex        =   13
         Top             =   5520
         Width           =   1455
      End
      Begin VB.Label lblRanglijstScore 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00008000&
         Caption         =   "50"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Index           =   3
         Left            =   5760
         TabIndex        =   39
         Top             =   3480
         Width           =   615
      End
      Begin VB.Label lblRanglijstScore 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00008000&
         Caption         =   "100"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Index           =   2
         Left            =   5760
         TabIndex        =   38
         Top             =   2880
         Width           =   615
      End
      Begin VB.Label lblRanglijstScore 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00008000&
         Caption         =   "150"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Index           =   1
         Left            =   5760
         TabIndex        =   37
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label lblRanglijstScore 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00008000&
         Caption         =   "200"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   0
         Left            =   5760
         TabIndex        =   36
         Top             =   1680
         Width           =   615
      End
      Begin VB.Line linPodium 
         BorderColor     =   &H00E0E0E0&
         X1              =   480
         X2              =   6360
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label Label1 
         BackColor       =   &H00008000&
         Caption         =   "Het podium"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   480
         TabIndex        =   41
         Top             =   600
         Width           =   1695
      End
      Begin VB.Image imgGewonnen 
         Height          =   1440
         Left            =   480
         Picture         =   "10opneer.frx":1638
         Stretch         =   -1  'True
         Top             =   4440
         Visible         =   0   'False
         Width           =   1440
      End
      Begin VB.Label lblRanglijst 
         BackColor       =   &H00008000&
         Caption         =   "Tara"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Index           =   3
         Left            =   960
         TabIndex        =   35
         Top             =   3480
         Width           =   4695
      End
      Begin VB.Label lblRanglijst 
         BackColor       =   &H00008000&
         Caption         =   "Frank"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Index           =   2
         Left            =   960
         TabIndex        =   34
         Top             =   2880
         Width           =   4695
      End
      Begin VB.Label lblRanglijst 
         BackColor       =   &H00008000&
         Caption         =   "Robbert"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Index           =   1
         Left            =   960
         TabIndex        =   33
         Top             =   2280
         Width           =   4695
      End
      Begin VB.Label lblRanglijst 
         BackColor       =   &H00008000&
         Caption         =   "Axel"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   0
         Left            =   960
         TabIndex        =   32
         Top             =   1680
         Width           =   4695
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00008000&
         Caption         =   "4."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   360
         TabIndex        =   31
         Top             =   3480
         Width           =   375
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00008000&
         Caption         =   "3."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   360
         TabIndex        =   30
         Top             =   2880
         Width           =   375
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00008000&
         Caption         =   "2."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   360
         TabIndex        =   29
         Top             =   2280
         Width           =   375
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00008000&
         Caption         =   "1."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   360
         TabIndex        =   28
         Top             =   1680
         Width           =   375
      End
   End
   Begin VB.Image imgVrouw 
      Height          =   570
      Left            =   8880
      ToolTipText     =   "Ik ben Tineke, en ik hou je in de gaten!"
      Top             =   5280
      WhatsThisHelpID =   1002
      Width           =   570
   End
   Begin VB.Menu mnuSpel 
      Caption         =   "&Spel"
      Begin VB.Menu mnuNieuwSpel 
         Caption         =   "&Nieuw spel"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuSpelNieuwSpel10openneer 
         Caption         =   "Nieuw spel 10 op en neer..."
      End
      Begin VB.Menu mnuSpelNieuwSpelBoerenbridge 
         Caption         =   "Nieuw spel Boerenbridge..."
      End
      Begin VB.Menu mnuSpelStreep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSpelOpenen 
         Caption         =   "&Openen..."
      End
      Begin VB.Menu mnuSpelOpslaanAls 
         Caption         =   "Opslaan &als..."
      End
      Begin VB.Menu mnuSpelStreep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSpelLaatsteSlagTonen 
         Caption         =   "&Laatste slag tonen"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuSpelHerhalen 
         Caption         =   "Spel &herhalen..."
      End
      Begin VB.Menu mnuSpelStreep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSpelWebpagina 
         Caption         =   "Naar de &webpagina"
      End
      Begin VB.Menu mnuVerwijderen 
         Caption         =   "10 op en neer verwijderen..."
      End
      Begin VB.Menu mnuSpelStreep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAfsluiten 
         Caption         =   "&Afsluiten"
      End
   End
   Begin VB.Menu mnuScore 
      Caption         =   "S&core"
      Begin VB.Menu mnuScorePuntentelling 
         Caption         =   "Puntentelling..."
      End
      Begin VB.Menu mnuStreep10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuScoreblok 
         Caption         =   "Toon &highscore"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuTotaal 
         Caption         =   "Toon &tussenstand"
      End
      Begin VB.Menu mnuStreep 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuSpelHighscoreAfdrukken 
         Caption         =   "Highscore &afdrukken"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuScoreHighscoreWissen 
         Caption         =   "Highscore &wissen..."
      End
      Begin VB.Menu mnuStreep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStatistiek 
         Caption         =   "Statistie&k..."
         Shortcut        =   ^K
      End
   End
   Begin VB.Menu mnuOpties 
      Caption         =   "&Opties"
      Begin VB.Menu mnuOmgekeerdSorteren 
         Caption         =   "&Omgekeerd sorteren"
      End
      Begin VB.Menu mnuBreedUitspreiden 
         Caption         =   "&Kaarten breed uitspreiden"
      End
      Begin VB.Menu mnuOptiesGroot 
         Caption         =   "G&rote kaarten"
      End
      Begin VB.Menu mnuKaarten 
         Caption         =   "Kaart&achterkant..."
      End
      Begin VB.Menu mnuStreep7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGeluid 
         Caption         =   "&Geluid aan"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuGeluiden 
         Caption         =   "G&eluiden..."
      End
      Begin VB.Menu mnuStreep8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCommentaar 
         Caption         =   "&Commentaar aan"
      End
      Begin VB.Menu mnuOptiesComputersInHighscore 
         Caption         =   "&Computers in highscore aan"
      End
      Begin VB.Menu mnuSnelheid 
         Caption         =   "&Snelheid..."
      End
      Begin VB.Menu mnuOptiesNamen 
         Caption         =   "&Namen..."
      End
      Begin VB.Menu mnuStreep9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptiesRondmaken 
         Caption         =   "Rondmaken toestaan..."
      End
      Begin VB.Menu mnuOptiesVerplichtIntroeven 
         Caption         =   "Verplicht introeven..."
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuSpelhulp 
         Caption         =   "&Hoe moet dit?"
         Shortcut        =   +{F1}
      End
      Begin VB.Menu mnuHelpHelponderwerpen 
         Caption         =   "Help-&onderwerpen"
         HelpContextID   =   101
      End
      Begin VB.Menu mnuStreep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpNieuw 
         Caption         =   "&Nieuw in deze versie"
      End
      Begin VB.Menu mnuInfo 
         Caption         =   "&Info..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public WithEvents Netwerk As clsNetwerk
Attribute Netwerk.VB_VarHelpID = -1

Private Sub BerekenScore()
  'Aanname: niet in herhaling.
  
  Dim SpelerNr As Integer
   
  For SpelerNr = 1 To 4
    If Spelers(SpelerNr).AantSlagen = Spelers(SpelerNr).Voorspelling Then
      'Goed
      Spelers(SpelerNr).Score(Ronde) = 10 + Opties.PuntenPerSlag * Spelers(SpelerNr).Voorspelling
      Spelers(SpelerNr).TotaalScore = Spelers(SpelerNr).TotaalScore + Spelers(SpelerNr).Score(Ronde)
      If SpelerNr = 1 Then
        Statistiek.GoedVoorspeld = Statistiek.GoedVoorspeld + 1
      End If
      recVoorspellingenGoed(Ronde, SpelerNr) = True
    Else
      'Fout
      If Opties.FoutVoorspeldNulPunten And Opties.StrafpuntenPerSlag = 0 Then
        Spelers(SpelerNr).Score(Ronde) = Spelers(SpelerNr).Voorspelling
        Spelers(SpelerNr).ScoreFout(Ronde) = True
      Else
        If Opties.FoutVoorspeldNulPunten Then
          Spelers(SpelerNr).Score(Ronde) = 0
        Else
          Spelers(SpelerNr).Score(Ronde) = 10
        End If
        Spelers(SpelerNr).Score(Ronde) = Spelers(SpelerNr).Score(Ronde) - Opties.StrafpuntenPerSlag * Abs(Spelers(SpelerNr).AantSlagen - Spelers(SpelerNr).Voorspelling)
        If Not Opties.NegatieveScores And Spelers(SpelerNr).Score(Ronde) < 0 Then
          Spelers(SpelerNr).Score(Ronde) = 0
        End If
        Spelers(SpelerNr).TotaalScore = Spelers(SpelerNr).TotaalScore + Spelers(SpelerNr).Score(Ronde)
        Spelers(SpelerNr).ScoreFout(Ronde) = False 'Geen streep erdoor
      End If
      If SpelerNr = 1 Then
        Statistiek.FoutVoorspeld = Statistiek.FoutVoorspeld + 1
        Statistiek.FoutSaldo = Statistiek.FoutSaldo + Spelers(SpelerNr).AantSlagen - Spelers(SpelerNr).Voorspelling
      End If
      recVoorspellingenGoed(Ronde, SpelerNr) = False
    End If
  Next SpelerNr

End Sub

Private Sub Delen()
  Dim GetalGekozen As Integer
  Dim KleurGekozen As Integer
  Dim SpelerNr As Integer
  Dim KaartNr As Integer
 
  'Randomize Timer
 
  StatusBar.SimpleText = Spelers(VorigeSpeler(NuOpkomen)).Naam & " deelt de kaarten."
 
  'DoEvents
  SchikKaarten 1
  SchikKaarten 2
  SchikKaarten 3
  SchikKaarten 4
  
'  If IkBenNetSpelerNr <= 1 Then
    For KaartNr = 1 To AantKaartenRonde(Ronde)
      For SpelerNr = 1 To 4
'        ** Onderstaande wordt gedetermineerd in DetermineerSpel
'        If recKaartenOntvangen(Ronde, SpelerNr, KaartNr).Kleur = 0 Then
'        recKaartenOntvangen(Ronde, SpelerNr, KaartNr) = DeStapel.Kaarten(DeStapel.AantKaarten)
'        End If
        Spelers(SpelerNr).Kaarten(KaartNr) = recKaartenOntvangen(Ronde, SpelerNr, KaartNr)
        WaarIsKaart(Spelers(SpelerNr).Kaarten(KaartNr).Kleur, Spelers(SpelerNr).Kaarten(KaartNr).Getal) = SpelerNr
        'DeStapel.AantKaarten = DeStapel.AantKaarten - 1
   
        If SpelerNr = 1 Then
          LaadKaart imgInHanden((SpelerNr - 1) * 13 + (KaartNr - 1)), Spelers(SpelerNr).Kaarten(KaartNr)
        End If
        imgInHanden((SpelerNr - 1) * 13 + (KaartNr - 1)).Visible = True
      Next SpelerNr
    Next KaartNr
    
'    If IkBenNetSpelerNr = 1 Then
'      txtNetwerkStatus.Text = txtNetwerkStatus.Text & vbCrLf & "Wacht tot client kaarten wil hebben"
'
'      Do Until ClientVraagt(3) = "GeefKaarten"
'        DoEvents
'      Loop
'      txtNetwerkStatus.Text = txtNetwerkStatus.Text & vbCrLf & "Client wil kaarten hebben"
'
'    End If
'  Else
'    txtNetwerkStatus.Text = txtNetwerkStatus.Text & vbCrLf & "Wacht op kaarten van de server..."
'    'Winsock1.SendData CStr(IkBenNetSpelerNr) & "GeefKaarten"
'    Do
'      DoEvents
'    Loop
'  End If
 
  If recTroeven(Ronde).Kleur = 0 Then
    Troef = GeenKaart
    picTroef.Picture = LoadPicture("")
    picTroef.BorderStyle = 1
    picTroef.ToolTipText = "Er is in deze ronde geen troef; alle kaarten zijn op"
  Else
    picTroef.BorderStyle = 0
    Troef = recTroeven(Ronde)
    DeStapel.AantKaarten = DeStapel.AantKaarten - 1
    WaarIsKaart(Troef.Kleur, Troef.Getal) = -1
    
    AantKaartenGezien = AantKaartenGezien + 1
  
    LaadKaart picTroef, Troef
    picTroef.ToolTipText = KleurNaam(Troef.Kleur) & " is troef"
    fraTroef.Visible = True
  End If
 
'  If DeStapel.AantKaarten = 0 Then
'    Troef = GeenKaart
'    picTroef.Picture = LoadPicture("")
'    picTroef.BorderStyle = 1
'    picTroef.ToolTipText = "Er is in deze ronde geen troef; alle kaarten zijn op"
'  Else
'    picTroef.BorderStyle = 0
'    If recTroeven(Ronde).Kleur = 0 Then
'      recTroeven(Ronde) = DeStapel.Kaarten(DeStapel.AantKaarten)
'    End If
'    Troef = recTroeven(Ronde)
'    DeStapel.AantKaarten = DeStapel.AantKaarten - 1
'    WaarIsKaart(Troef.Kleur, Troef.Getal) = -1
'
'    AantKaartenGezien = AantKaartenGezien + 1
'
'    LaadKaart picTroef, Troef
'    picTroef.ToolTipText = KleurNaam(Troef.Kleur) & " is troef"
'  End If

  For SpelerNr = 1 To 4
    Spelers(SpelerNr).AantKaarten = AantKaartenRonde(Ronde)
  Next SpelerNr
  Sorteren 1

End Sub

Function NuSlagenOverVoorAnderen(SpelerNr As Integer) As Integer
 Dim SpelerTemp As Integer
 Dim TotGevraagd As Integer
 Dim TotBinnen As Integer
 Dim TotNodig As Integer
  
 TotGevraagd = TotSlagenGok - Spelers(SpelerNr).Voorspelling
 
 SpelerTemp = VolgendeSpeler(SpelerNr)
 Do
  'TotGevraagd = TotGevraagd + SlagenGok(SpelerTemp)
  TotBinnen = TotBinnen + Spelers(SpelerTemp).AantSlagen
  SpelerTemp = VolgendeSpeler(SpelerTemp)
 Loop Until SpelerTemp = SpelerNr
 
 TotNodig = TotGevraagd - TotBinnen
 NuSlagenOverVoorAnderen = KaartenResterend - TotNodig
  'AantalVanDezeSpeler(VorigeSpeler(NuOpkomen)) - TotNodig
  'VorigeSpeler(NuOpkomen)) heeft de meeste kaarten in hand;
  'Zoveel slagen kunnen nog gehaald worden
End Function

Sub ToonOpkomen()
  Dim SpelerNr As Integer
  
  For SpelerNr = 1 To 4
    If SpelerNr = NuOpkomen Then
      lblNaam(SpelerNr - 1).ForeColor = QBColor(14)
    Else
      lblNaam(SpelerNr - 1).ForeColor = Voorkleur
    End If
  Next SpelerNr
End Sub
Private Function GooiWeg(SpelerNr As Integer, RangNummer As Integer) As Integer
 'RangNummer: 1=laagste, 2=eennalaagste, -1=hoogste, -2=eennahoogste, enz.
  Dim VolgordeWaarde(1 To 13) As Single
  Dim VolgordeNummer(1 To 13) As Integer
  Dim VolgordeIndex As Integer
  Dim VolgordeIndex2 As Integer
  Dim VolgordeIngevoegd As Integer
  Dim AantalVolgordeKaarten As Integer
  Dim WaardeBerekening1 As Single
  Dim WaardeBerekening2 As Single
  Dim WaardeNu As Single
  Dim DuikMetVolgordeNummer As Integer
  'Dim DuikMetVolgordeNummerPlus As Integer
  Dim Stap As Integer
 
  Dim KaartNr As Integer
      
  If SlagNietTeHalenMetAantal = 0 Then
    MsgBox "Fatale fout: GooiWeg met SlagNietTeHalenMetAantal = 0", vbCritical, "Fatale fout"
    Stop
  End If

  AantalVolgordeKaarten = 0
 
  For KaartNr = 1 To AantKaartenRonde(Ronde)
    If Spelers(SpelerNr).Kaarten(KaartNr).Getal > 0 Then
      VolgordeIngevoegd = False
    
      WaardeBerekening1 = BerekenKaartWaarde(SpelerNr, KaartNr, True)
      WaardeBerekening2 = (25 - KaartenErboven(SpelerNr, KaartNr)) / 25
      WaardeNu = 0.05 * WaardeBerekening1 + 0.95 * WaardeBerekening2
      
      'WaardeNu = 52 - KaartenErboven(SpelerNr, KaartNr)
      'BerekenKaartWaarde(SpelerNr, KaartNr, True)
      
      For VolgordeIndex = 1 To AantalVolgordeKaarten + 1
        If WaardeNu <= VolgordeWaarde(VolgordeIndex) Or VolgordeIndex > AantalVolgordeKaarten Then
          For VolgordeIndex2 = AantalVolgordeKaarten To VolgordeIndex Step -1  '9
            VolgordeWaarde(VolgordeIndex2 + 1) = VolgordeWaarde(VolgordeIndex2)
            VolgordeNummer(VolgordeIndex2 + 1) = VolgordeNummer(VolgordeIndex2)
          Next VolgordeIndex2
          VolgordeWaarde(VolgordeIndex) = WaardeNu
          VolgordeNummer(VolgordeIndex) = KaartNr
          AantalVolgordeKaarten = AantalVolgordeKaarten + 1
          VolgordeIngevoegd = True
          Exit For
        End If
      Next VolgordeIndex
      If VolgordeIngevoegd = False Then Stop 'Dat moet niet
    End If
  Next KaartNr
 
  If RangNummer > 0 Then
    DuikMetVolgordeNummer = RangNummer
    Stap = 1
  ElseIf RangNummer < 0 Then
    DuikMetVolgordeNummer = AantalVolgordeKaarten + RangNummer + 1
    Stap = -1
  Else
    MsgBox "Fatale fout in GooiWeg", vbCritical, "Fatale fout"
    Stop
  End If
  
  Do Until (HoogsteKaart(KaartenOpTafel(HoogsteOpTafel), Spelers(SpelerNr).Kaarten(VolgordeNummer(DuikMetVolgordeNummer))) = 1) And Spelers(SpelerNr).Kaarten(VolgordeNummer(DuikMetVolgordeNummer)).Legaal
    'Zoek een kaart waarmee je de slag niet haalt
    DuikMetVolgordeNummer = DuikMetVolgordeNummer + Stap
    If DuikMetVolgordeNummer > AantalVolgordeKaarten Then
      DuikMetVolgordeNummer = RangNummer - 1 '-1 is nieuw!
      Stap = -1
    End If
    If DuikMetVolgordeNummer < 1 Then
      DuikMetVolgordeNummer = AantalVolgordeKaarten + RangNummer + 2
      Stap = 1
    End If
  Loop
  
  'If RangNummer > 1 And DuikMetVolgordeNummer = AantalVolgordeKaarten Then
  '  '** Bijv duiken met eennalaatste. Liever met allerlaagste dan met hoogste
  '  DuikMetVolgordeNummer = DuikMetVolgordeNummer - 1
  '  Stop
  'End If
  'If DuikMetVolgordeNummer > AantalVolgordeKaarten Then
  '  DuikMetVolgordeNummer = AantalVolgordeKaarten '-1??
  '  Stop
  'End If
  'If DuikMetVolgordeNummer < 1 Then
  '  DuikMetVolgordeNummer = 1
  '  Stop
  'End If

  'OokTroef
  If VolgordeNummer(DuikMetVolgordeNummer) = 0 Then Stop 'Dan had deze functie niet aangeroepen mogen worden
 
  GooiWeg = VolgordeNummer(DuikMetVolgordeNummer)

End Function

Private Sub LaadKaart(Voorwerp As Object, DeKaart As Kaart)
  Dim KaartNr As Integer
 
  KaartNr = 13 * (DeKaart.Kleur - 1) + DeKaart.Getal + 99
  Voorwerp.Picture = LoadResPicture(KaartNr, vbResBitmap)
End Sub

Private Sub cmdAfsluiten_Click()
  Einde
End Sub

Private Sub cmdHerhaling_Click()
  NieuwSpel True
End Sub

Private Sub cmdNieuwSpel_Click()
  NieuwSpel False
End Sub

Private Sub cmdVoorspellen_Click(Index As Integer)
  Spelers(1).Voorspelling = Index
  VoorspeldMens
End Sub

Private Sub cmdVoorspellen_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then
    PopupMenu frmMenus.mnuVoorspellen
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If Not (TypeOf Screen.ActiveControl Is TextBox) And Not recInHerhaling() Then
    Select Case KeyCode
      Case vbKeyRight
        If WachtOp = MensKaart Then
          ZoekAanTeWijzenKaart "r"
          imgKaartAanwijs.Left = imgInHanden(KaartAanwijzen - 1).Left
          imgKaartAanwijs.Visible = True
        End If
      Case vbKeyLeft
        If WachtOp = MensKaart Then
          ZoekAanTeWijzenKaart "l"
          imgKaartAanwijs.Left = imgInHanden(KaartAanwijzen - 1).Left
          imgKaartAanwijs.Visible = True
        End If
    End Select
  End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If Not (TypeOf Screen.ActiveControl Is TextBox) Then
    Select Case KeyAscii
      Case vbKeySpace
        If WachtOp = MensKaart And Not (TypeOf Screen.ActiveControl Is CommandButton) Then
          If recInHerhaling() Then
            KaartGekozenMens recKaartenGekozen(Ronde, 1, SlagNr)
          Else
            KaartGekozenMens KaartAanwijzen
          End If
        'Else
          'If WachtOp = MensVoorspelling Then
          '  StatusBar.simpletext = "Je moet eerst het aantal slagen voorspellen."
          'Else
          '  StatusBar.simpletext = "Je bent nog niet aan de beurt."
          'End If
        End If
      Case vbKey0 To vbKey9
        If WachtOp = MensVoorspelling Then
          If KeyAscii - 48 >= 0 And KeyAscii - 48 <= AantKaartenRonde(Ronde) Then
            Spelers(1).Voorspelling = KeyAscii - 48
            VoorspeldMens
          End If
        End If
      Case KeySlash
        If Not TipZichtbaar Then
          If WachtOp = MensKaart Then
            If Tip <= 0 Then
              Tip = KaartKiezenComputer(1)
            End If
            imgInHanden(Tip - 1).Top = imgInHanden(Tip - 1).Top - 240
          ElseIf WachtOp = MensVoorspelling Then
            If Tip < 0 Then
              Tip = LegaleVoorspelling(TaxeerKaarten(1))
            End If
            cmdVoorspellen(Tip).FontUnderline = True
          End If
          TipZichtbaar = True
        End If
    End Select
  End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  If TipZichtbaar Then
    If WachtOp = MensKaart Then
      imgInHanden(Tip - 1).Top = imgInHanden(Tip - 1).Top + 240
    ElseIf WachtOp = MensVoorspelling Then
      cmdVoorspellen(Tip).FontUnderline = False
    End If
    TipZichtbaar = False
  End If
  If KeyCode = 49 And Shift = vbShiftMask + vbCtrlMask Then
    DebugWisseltruc
  ElseIf KeyCode = 50 And Shift = vbShiftMask + vbCtrlMask Then
    Debug.Print "Speler 2":
    ToonKaartenInfo (2)
    Debug.Print "Speler 3":
    ToonKaartenInfo (3)
    Debug.Print "Speler 4":
    ToonKaartenInfo (4)
  End If
End Sub

Private Sub Form_Load()
  'Init
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Einde
End Sub

Private Function HogeEnLageAftrek(SpelerNr) As Single
  Dim KaartNr As Integer
  Dim i As Integer
  Dim AftrekTemp As Single
  Dim HeeftHoge(1 To 4) As Integer
  Dim HeeftLage(1 To 4) As Integer
  
  'For i = 1 To 4
  '  HeeftLage(i) = False
  '  HeeftHoge(i) = False
  'Next i
  Erase HeeftLage
  Erase HeeftHoge
  
  For KaartNr = 1 To AantKaartenRonde(Ronde)
    With Spelers(SpelerNr).Kaarten(KaartNr)
      If .Kleur <> Troef.Kleur Then 'Nieuw: is nvt bij troeven
        If Spelers(SpelerNr).Kaarten(KaartNr).Getal < 7 Then
          HeeftLage(Spelers(SpelerNr).Kaarten(KaartNr).Kleur) = True
        ElseIf Spelers(SpelerNr).Kaarten(KaartNr).Getal > 7 Then
          HeeftHoge(Spelers(SpelerNr).Kaarten(KaartNr).Kleur) = True
        End If
      End If
    End With
  Next KaartNr
  For i = 1 To 4
    If HeeftLage(i) = True And HeeftHoge(i) = True Then
      AftrekTemp = AftrekTemp + 0.3
    End If
  Next i
  '** Dubieus!
  HogeEnLageAftrek = AftrekTemp
End Function

Private Function HoogsteKaart(Kaart1 As Kaart, Kaart2 As Kaart) As Integer
  If Kaart2.Getal > Kaart1.Getal And Kaart1.Kleur = Kaart2.Kleur Then
    HoogsteKaart = 2
  ElseIf Kaart2.Kleur = Troef.Kleur And Kaart1.Kleur <> Troef.Kleur Then
    HoogsteKaart = 2
  Else
    HoogsteKaart = 1
  End If
End Function

Private Sub KaartenWeg()
  Dim i As Integer, j As Integer
  Dim SpelerNr As Integer
 
  For SpelerNr = 1 To 4
    Spelers(SpelerNr).Voorspelling = -1
    KleurAantalKerenGespeeld(SpelerNr) = 0 'Is niet afhankelijk van spelernummer,
  Next SpelerNr                            'maar van kleurnummer, maar ach...
 
  For i = 0 To 51
    If KaartImageGeladen(i) Then
      imgInHanden(i).Visible = False
      'If i >= 10 Then
      '  imgInHanden(i).ToolTipText = "De kaarten van " & Spelers(Int(i / 10) + 1).Naam
      'End If
    End If
  Next i
  For i = 1 To 4
    Spelers(i).AantSlagen = 0
    imgOpTafel(i - 1).Visible = False '?Deze regel overbodig ivm picTafel.visible=false?
  Next i

  ToonRondjes 1
  ToonRondjes 2
  ToonRondjes 3
  ToonRondjes 4

  'StatusBar.simpletext = "Het spel is afgelopen."
  fraVoorspellen.Visible = False   'Deze drie i.v.m. herhaling bekijken
  frmMain.fraEinde.Visible = False
  frmMain.fraSpelen.Visible = True
  
  imgKaartAanwijs.Visible = False
End Sub

Private Sub KaartKiezen()
  Dim AantalSlagenNogNodig As Integer

  'WachtOp = MensKaart 'Speler mag al van tevoren een kaart kiezen
  If Not recInHerhaling Then
    recSpelerNum = SpelerNum
    recSlagNr = SlagNr
  End If

  If Spelers(NuOpleggen).Controller = ControllerComputer Then
    StatusBar.SimpleText = "Wacht op de beurt van " & Spelers(NuOpleggen).Naam & "..."
    If Not recInHerhaling Then
      recKaartenGekozen(Ronde, NuOpleggen, SlagNr) = KaartKiezenComputer(NuOpleggen)
      
      'recSpelerNr = NuOpleggen
      'recSpelerNum = SpelerNum
      'recSlagNr = SlagNr
      'recVoorspellenKlaar = True
    Else ': Stop
    End If
    KaartLeggen recKaartenGekozen(Ronde, NuOpleggen, SlagNr)
  ElseIf Spelers(NuOpleggen).Controller = ControllerMens Then
    Tip = -1
    
    ZoekMogelijkeKaarten NuOpleggen 'Wordt ook aangeroepen als de speler een tip wil, maar ja
    TestSlagTeHalen 1
        
    AantalSlagenNogNodig = Spelers(NuOpleggen).Voorspelling - Spelers(NuOpleggen).AantSlagen
    If AantalSlagenNogNodig = Spelers(NuOpleggen).AantKaarten And SlagIsTeHalenMetAantal = 0 Then
      'Gaat te weinig krijgen
      OndergangGeluidGespeeld = True
      WavPlay "Onvoorkombare ondergang" '"Hartenjagen stuk.wav"
    ElseIf AantalSlagenNogNodig = 0 And SlagNietTeHalenMetAantal = 0 And NuOpkomen = VolgendeSpeler(NuOpleggen) Then
      'Gaat te veel krijgen
      OndergangGeluidGespeeld = True
      WavPlay "Onvoorkombare ondergang" '"Hartenjagen stuk.wav"
    End If
    
    If recInHerhaling Then
      imgKaartAanwijs.Left = imgInHanden(recKaartenGekozen(Ronde, 1, SlagNr) - 1).Left
      imgKaartAanwijs.Visible = True
      StatusBar.SimpleText = "Je bekijkt een herhaling. Druk op de spatiebalk om verder te gaan."
    Else
      KaartAanwijzen = MogelijkeKaarten(1)
      StatusBar.SimpleText = "Kies een kaart."
      'imgKaartAanwijs.Visible = False
    End If

    Tineke.ZegHulp "kies een kaart"
    WachtOp = MensKaart

  ElseIf Spelers(NuOpleggen).Controller = Netwerk Then
    StatusBar.SimpleText = "Wacht op de beurt van " & Spelers(NuOpleggen).Naam & "..."
    Stop '#
  End If

End Sub

Function KomOpMetLageTroef(SpelerNr As Integer, HogeIsOokGoed As Boolean) As Integer
  Static VolgordeGetal(13) As Single
  Static VolgordeNummer(13) As Integer
  Static VolgordeWaarde(13) As Single
  Dim VolgordeIndex As Integer, VolgordeIndex2 As Integer
  Dim AantalVolgordeKaarten As Integer
  Dim WaardeNu As Single
  Dim SlagIsTeHalenMetIndex As Integer
  Dim NeemMetNummer As Integer
  Dim TeHoogMarge As Single
  
  TeHoogMarge = 0.6 'Maximumwaarde troef om op te leggen
       
  For SlagIsTeHalenMetIndex = 1 To SlagIsTeHalenMetAantal
    'KleurNu = SpelerKaarten(SpelerNr, SlagIsTeHalenMetNummer(SlagIsTeHalenMetIndex)).Kleur
    'GetalNu = SpelerKaarten(SpelerNr, SlagIsTeHalenMetNummer(SlagIsTeHalenMetIndex)).Getal
    If Spelers(SpelerNr).Kaarten(SlagIsTeHalenMetNummer(SlagIsTeHalenMetIndex)).Kleur = Troef.Kleur Then
      WaardeNu = BerekenKaartWaarde(SpelerNr, SlagIsTeHalenMetNummer(SlagIsTeHalenMetIndex), True)
  
      For VolgordeIndex = 1 To AantalVolgordeKaarten + 1
        If WaardeNu <= VolgordeWaarde(VolgordeIndex) Or VolgordeIndex > AantalVolgordeKaarten Then
          For VolgordeIndex2 = 12 To VolgordeIndex Step -1
            VolgordeWaarde(VolgordeIndex2 + 1) = VolgordeWaarde(VolgordeIndex2)
            VolgordeNummer(VolgordeIndex2 + 1) = VolgordeNummer(VolgordeIndex2)
          Next VolgordeIndex2
          VolgordeWaarde(VolgordeIndex) = WaardeNu
          VolgordeNummer(VolgordeIndex) = SlagIsTeHalenMetNummer(SlagIsTeHalenMetIndex)
          AantalVolgordeKaarten = AantalVolgordeKaarten + 1
          Exit For
        End If
      Next VolgordeIndex
    End If
  Next SlagIsTeHalenMetIndex
  If AantalVolgordeKaarten > 0 Then
    'Stop
    If VolgordeWaarde(1) < TeHoogMarge Or HogeIsOokGoed Then
      KomOpMetLageTroef = VolgordeNummer(1)
    Else
      KomOpMetLageTroef = NeemSlag(SpelerNr, 1, False)
    End If
  Else
    KomOpMetLageTroef = NeemSlag(SpelerNr, 1, False)
  End If
End Function

Private Sub Form_Resize()
  timResize.Enabled = False
  timResize.Enabled = True
End Sub

Private Sub imgKaartAanwijs_Click()
  imgKaartAanwijs.Visible = False
End Sub

Private Sub imgOpTafel_Click(Index As Integer)
  If timLaatsteSlag.Enabled Then
    ToonHuidigeSlag
  End If
End Sub

Private Sub imgSlagen_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then
    PopupMenu frmMenus.mnuRondjes
  End If
End Sub

Private Sub imgVrouw_DblClick()
  imgVrouw_MouseDown 1, 0, 0, 0
End Sub

Private Sub imgVrouw_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 1 Then
    If Praatwolkje.Visible Then
      Tineke.WegPraatwolkje
    Else
      Tineke.Zeg "klik"
    End If
  ElseIf Button = 2 Then
    If Tineke.CommentaarTekst = "" Then
      frmMenus.mnuHerhalen.Enabled = False
    Else
      frmMenus.mnuHerhalen.Enabled = True
    End If
    
    PopupMenu frmMenus.mnuTineke
  End If
End Sub

'Private Sub lblIntroeven_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'  If Button = 2 Then
'    PopupMenu frmMenus.mnuIntroeven
'  End If
'End Sub

Private Sub lblNaam_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 1 Then
    StartNaamWijzigen Index + 1
  ElseIf Button = 2 Then
    PopupMenuIndex = Index + 1
    PopupMenu frmMenus.mnuSpelerNaam
  End If

End Sub

Private Sub mnuAfsluiten_Click()
  Einde
End Sub

Private Sub mnuBreedUitspreiden_Click()
  Opties.BreedUitspreiden = Not Opties.BreedUitspreiden
  mnuBreedUitspreiden.Checked = Opties.BreedUitspreiden
  frmMenus.mnuBreedUitspreiden.Checked = Opties.BreedUitspreiden
  SchikKaarten 1
End Sub

Private Sub mnuCommentaar_Click()
  Opties.Commentaar = Not Opties.Commentaar
  mnuCommentaar.Checked = Opties.Commentaar
  frmMenus.mnuCommentaar.Checked = Opties.Commentaar
End Sub

Private Sub mnuGeluid_Click()
  Opties.Geluid = Not Opties.Geluid
  mnuGeluid.Checked = Opties.Geluid
End Sub

Private Sub mnuGeluiden_Click()
  Dim Ret As Long
  'Rundll32.exe shell32,Control_RunDLL "C:\WINDOWS\SYSTEM\MMSYS.CPL",Geluiden
  'Ret = Shell("Rundll32.exe shell32,Control_RunDLL " & Chr(34) & "C:\WINDOWS\SYSTEM\MMSYS.CPL" & Chr(34) & ",Geluiden")
  Ret = Shell("rundll32.exe shell32.dll,Control_RunDLL mmsys.cpl @1,1")
  'If Ret Then 'Succes
  '  SendKeys "10{PGDN}{PGUP}"
  'End If
End Sub

Private Sub mnuHelpHelponderwerpen_Click()
  RunFile App.HelpFile
End Sub

Private Sub mnuHelpNieuw_Click()
  RunFile DirPlusBestand(App.Path, "Whatsnew.txt")
End Sub

Private Sub mnuInfo_Click()
  frmInfo.Show 1
End Sub

Private Sub mnuKaarten_Click()
  frmAchterkant.Show 1
End Sub

Private Sub mnuNieuwSpel_Click()
  Dim Ret As VbMsgBoxResult
  
  If Ronde >= nRonden / 2 Then
    If Spelers(1).TotaalScore / Ronde >= 9 Then
      Ret = MsgBox("Het gaat erg goed. Weet je zeker dat je een nieuw spel wilt beginnen?", vbYesNo + vbQuestion, "Nieuw spel")
      If Ret = vbYes Then
        NieuwSpel False
      End If
    Else
      NieuwSpel False
    End If
  Else
    NieuwSpel False
  End If
End Sub

Private Sub mnuOptiesComputersInHighscore_Click()
  Opties.ComputersInHighScore = Not Opties.ComputersInHighScore
  mnuOptiesComputersInHighscore.Checked = Opties.ComputersInHighScore
  frmMenus.mnuHighscoreComputersInHighscore.Checked = Opties.ComputersInHighScore
End Sub

Private Sub mnuOptiesGroot_Click()
  Opties.GroteKaarten = Not Opties.GroteKaarten
  mnuOptiesGroot.Checked = Opties.GroteKaarten
  frmMenus.mnuKaartenGroteKaarten.Checked = Opties.GroteKaarten
  KaartgrootteInstellen
End Sub

Private Sub mnuOptiesNamen_Click()
  MsgBox "Klik op de naam van een speler om deze te veranderen.", vbInformation, "Namen wijzigen"
End Sub

Private Sub mnuOptiesRondmaken_Click()
  Dim Ret As VbMsgBoxResult
  
  Ret = MsgBox("Deze instelling kan alleen gewijzigd worden als je een nieuw spel start. Wil je nu een nieuw spel starten?", vbQuestion + vbYesNo, "Rondmaken toestaan")
  
  If Ret = vbYes Then
    mnuOptiesRondmaken.Checked = Not mnuOptiesRondmaken.Checked
    Opties.RondmakenToegestaan = mnuOptiesRondmaken.Checked
    NieuwSpel False
  End If
End Sub

Private Sub mnuOptiesVerplichtIntroeven_Click()
  Dim Ret As VbMsgBoxResult
  
  Ret = MsgBox("Deze instelling kan alleen gewijzigd worden als je een nieuw spel start. Wil je nu een nieuw spel starten?", vbQuestion + vbYesNo, "Verplicht introeven")
  
  If Ret = vbYes Then
    mnuOptiesVerplichtIntroeven.Checked = Not mnuOptiesVerplichtIntroeven.Checked
    Opties.IntroevenVerplicht = mnuOptiesVerplichtIntroeven.Checked
    NieuwSpel False
  End If
End Sub

Private Sub mnuScoreblok_Click()
  ScoreblokMenuKlik
End Sub

Private Sub mnuScoreHighscoreWissen_Click()
  MenuHighscoreWissen
End Sub

Private Sub mnuScorePuntentelling_Click()
  frmPuntentelling.Show 1
End Sub

Private Sub mnuSnelheid_Click()
  frmSnelheid.Show 1
End Sub

Private Sub mnuOmgekeerdSorteren_Click()
  Opties.AflopendSorteren = Not Opties.AflopendSorteren
  mnuOmgekeerdSorteren.Checked = Opties.AflopendSorteren
  frmMenus.mnuOmgekeerdSorteren.Checked = Opties.AflopendSorteren
  'IniSet "Opties", "Aflopend", Opties.AflopendSorteren
  Sorteren 1
End Sub

Private Sub mnuSpelHerhalen_Click()
  Dim Ret As VbMsgBoxResult
  
  Ret = MsgBox("Weet je zeker dat je een herhaling van het spel wilt bekijken?", vbQuestion + vbOKCancel, "Herhaling")
  If Ret = vbOK Then
    'NieuwSpel True
    Ronde = 0
    NieuweRonde
  End If
End Sub

Private Sub mnuSpelHighscoreAfdrukken_Click()
  MenuAfdrukken
End Sub

Private Sub mnuSpelhulp_Click()
  Opties.Spelhulp = Not Opties.Spelhulp
  SpelhulpAanUit
End Sub

Private Sub mnuSpelLaatsteSlagTonen_Click()
  If mnuSpelLaatsteSlagTonen.Checked Then
    ToonHuidigeSlag
  Else
    ToonLaatsteSlag
  End If
End Sub

Private Sub mnuSpelNieuwSpel10openneer_Click()
  Dim Ret As VbMsgBoxResult
  
  Ret = MsgBox("Hiermee stel je de spelregels en puntentelling van 10 op en neer in, en start je een nieuw spel. Wil je doorgaan?", vbQuestion + vbYesNo, "Nieuw spel 10 op en neer")
  
  If Ret = vbYes Then
    Opties.PuntenPerSlag = 1
    Opties.FoutVoorspeldNulPunten = True
    Opties.StrafpuntenPerSlag = 0
    Opties.RondmakenToegestaan = True
    Opties.IntroevenVerplicht = False
    Opties.MaxAantKaarten = 10
    mnuOptiesRondmaken.Checked = Opties.RondmakenToegestaan
    mnuOptiesVerplichtIntroeven.Checked = Opties.IntroevenVerplicht
    NieuwSpel False
  End If
End Sub

Private Sub mnuSpelNieuwSpelBoerenbridge_Click()
  Dim Ret As VbMsgBoxResult
  
  Ret = MsgBox("Hiermee stel je de meest gangbare spelregels en de puntentelling van Boerenbridge in, en start je een nieuw spel. Wil je doorgaan?", vbQuestion + vbYesNo, "Nieuw spel Boerenbridge")
  
  If Ret = vbYes Then
    Opties.PuntenPerSlag = 3
    Opties.FoutVoorspeldNulPunten = False
    Opties.StrafpuntenPerSlag = 3
    Opties.NegatieveScores = False
    Opties.RondmakenToegestaan = False
    Opties.IntroevenVerplicht = False
    Opties.MaxAantKaarten = 13
    mnuOptiesRondmaken.Checked = Opties.RondmakenToegestaan
    mnuOptiesVerplichtIntroeven.Checked = Opties.IntroevenVerplicht
    NieuwSpel False
  End If

End Sub

Private Sub mnuSpelOpenen_Click()
  Dim CancelPressed As Boolean

'  DestFolder = GetSpecialFolder(Me.hwnd, CSIDL_PERSONAL)
  With frmMain.CommonDialog1
    .DefaultExt = ".10s"
    .Filter = "10 op en neer-spellen|*.10s|Alle bestanden|*.*"
    .FilterIndex = 0
    .CancelError = True
    On Error GoTo OpenCancelError
    .ShowOpen
    On Error GoTo 0
    If Not CancelPressed Then
      recOpenen .Filename
      Ronde = 0
      NieuweRonde
    End If
  End With
  
  Exit Sub
OpenCancelError:
  CancelPressed = True
  Resume Next

End Sub

Private Sub mnuSpelOpslaanAls_Click()
  Dim DestFolder As String
  Dim CancelPressed As Boolean
  Dim Ret As VbMsgBoxResult
  Dim Klaar As Boolean
  Dim OpslaanOk As Boolean
  
  'DestFolder = GetSpecialFolder(Me.hwnd, CSIDL_PERSONAL)
  With frmMain.CommonDialog1
    'If .Filename = "" And Dir(DestFolder & "NUL") <> "" Then
    '  .Filename = DestFolder & Spelers(1).Naam & ".10s"
    'End If
    .Filename = Spelers(1).Naam & ".10s"
    .DefaultExt = ".10s"
    .Filter = "10 op en neer-spellen|*.10s|Alle bestanden|*.*"
    .FilterIndex = 0
    Klaar = False
    OpslaanOk = True
    Do Until Klaar
      CancelPressed = False
      .CancelError = True
      On Error GoTo SaveCancelError
      .ShowSave
      On Error GoTo 0
      If CancelPressed Then
        OpslaanOk = False
        Klaar = True
      Else
        If Dir(.Filename) <> "" Then
          '** Bestand bestaat al **
          Ret = MsgBox("Het bestand " & .Filename & " bestaat al. Wil je het overschrijven?", vbQuestion + vbYesNoCancel, "Opslaan als")
          If Ret = vbNo Then
            OpslaanOk = False
            Klaar = False
            'Nog een keer proberen
          ElseIf Ret = vbCancel Then
            OpslaanOk = False
            Klaar = True
          ElseIf Ret = vbYes Then
            OpslaanOk = True
            Klaar = True
          End If
        Else
          '** Bestand bestaat nog niet **
          OpslaanOk = True
          Klaar = True
        End If
        If OpslaanOk Then
          recOpslaan .Filename
        End If
      End If
    Loop
  End With
  
  Exit Sub
SaveCancelError:
  CancelPressed = True
  Resume Next
End Sub

Private Sub mnuSpelWebpagina_Click()
  Me.MousePointer = vbHourglass
  StartURL Homepage
  Me.MousePointer = vbDefault
End Sub

Private Sub mnuStatistiek_Click()
  frmStatistiek.Show 1
End Sub

Private Function NeemSlag(SpelerNr As Integer, RangNummer As Integer, TroefIsOokGoed As Boolean) As Integer
  '-1 = hoogste; -2 = eennahoogste, 1 = laagste, 2 = eennalaagste, enz.
  
  Static VolgordeWaarde(1 To 13) As Single
  Static VolgordeNummer(1 To 13) As Integer
  Dim VolgordeIndex As Integer
  Dim VolgordeIndex2 As Integer
  Dim VolgordeIngevoegd As Integer
  Dim AantalVolgordeKaarten As Integer
  Dim WaardeBerekening1 As Single
  Dim WaardeBerekening2 As Single
  Dim WaardeBerekeningVolgende As Single
  Dim WaardeNu As Single
  Dim KleurNu As Integer
  Dim GetalNu As Integer
  Dim SlagIsTeHalenMetIndex As Integer
  Dim NeemMetNummer As Integer
  Dim KiesUitKaarten(1 To 13) As Integer
  Dim AantKiesUitKaarten As Integer
  Dim LegaleNietTroefKaartIndex As Integer
  Dim KiesUitIndex As Integer
  Dim GevondenVolgendeKaartVanKleur As Integer
  
  AantKiesUitKaarten = 0
  If TroefIsOokGoed Or AantLegaleNietTroefKaarten = 0 Then
    For SlagIsTeHalenMetIndex = 1 To SlagIsTeHalenMetAantal
      KiesUitKaarten(SlagIsTeHalenMetIndex) = SlagIsTeHalenMetNummer(SlagIsTeHalenMetIndex)
    Next SlagIsTeHalenMetIndex
    AantKiesUitKaarten = SlagIsTeHalenMetAantal
  Else
    For LegaleNietTroefKaartIndex = 1 To AantLegaleNietTroefKaarten
      KiesUitKaarten(LegaleNietTroefKaartIndex) = LegaleNietTroefkaarten(LegaleNietTroefKaartIndex)
    Next LegaleNietTroefKaartIndex
    AantKiesUitKaarten = AantLegaleNietTroefKaarten
  End If
       
  VolgordeIngevoegd = False
  For KiesUitIndex = 1 To AantKiesUitKaarten
    KleurNu = Spelers(SpelerNr).Kaarten(KiesUitKaarten(KiesUitIndex)).Kleur
    GetalNu = Spelers(SpelerNr).Kaarten(KiesUitKaarten(KiesUitIndex)).Getal
    'Debug.Print Format(WaardeNu / 25, "0.00") & " <-> " & Format(BerekenKaartWaarde(SpelerNr, KiesUitKaarten(KiesUitIndex), TroefIsOokGoed), "0.00") & " (" & KleurNaam(KleurNu) & GetalNaam(GetalNu)
    WaardeBerekening1 = BerekenKaartWaarde(SpelerNr, KiesUitKaarten(KiesUitIndex), TroefIsOokGoed) 'Is dit goed?
    WaardeBerekening2 = (25 - KaartenErboven(SpelerNr, KiesUitKaarten(KiesUitIndex))) / 25
    WaardeNu = 0.05 * WaardeBerekening1 + 0.95 * WaardeBerekening2
    If RangNummer = 1 Then 'Liefst een kale lage kaart opgooien
      GevondenVolgendeKaartVanKleur = VolgendeKaartVanKleur(SpelerNr, KiesUitKaarten(KiesUitIndex))
      If GevondenVolgendeKaartVanKleur >= 1 Then
        WaardeBerekeningVolgende = (25 - KaartenErboven(SpelerNr, GevondenVolgendeKaartVanKleur)) / 25
        WaardeNu = 0.65 * WaardeNu + 0.35 * WaardeBerekeningVolgende '0.85, 0.2
      End If
    End If
   
    For VolgordeIndex = 1 To AantalVolgordeKaarten + 1
      If WaardeNu <= VolgordeWaarde(VolgordeIndex) Or VolgordeIndex > AantalVolgordeKaarten Then
        For VolgordeIndex2 = 12 To VolgordeIndex Step -1
          VolgordeWaarde(VolgordeIndex2 + 1) = VolgordeWaarde(VolgordeIndex2)
          VolgordeNummer(VolgordeIndex2 + 1) = VolgordeNummer(VolgordeIndex2)
        Next VolgordeIndex2
        VolgordeWaarde(VolgordeIndex) = WaardeNu
        VolgordeNummer(VolgordeIndex) = KiesUitKaarten(KiesUitIndex)
        VolgordeIngevoegd = True
        AantalVolgordeKaarten = AantalVolgordeKaarten + 1
        Exit For
      End If
    Next VolgordeIndex
  Next KiesUitIndex
  If VolgordeIngevoegd = False Then Stop 'Dat moet niet
 
  If RangNummer > 0 Then
    NeemMetNummer = RangNummer
  ElseIf RangNummer < 0 Then
    NeemMetNummer = AantalVolgordeKaarten + RangNummer + 1
  Else                                  '+ want is negatief
    Stop 'Mag niet 0 zijn
  End If
  If NeemMetNummer > AantalVolgordeKaarten Then
    NeemMetNummer = AantalVolgordeKaarten
  End If
  If NeemMetNummer < 1 Then
    NeemMetNummer = 1
  End If

  'If SpelerNr = 1 Then
  '  Debug.Print "NeemSlag met nr " & RangNummer & " : kaart nr. " & VolgordeNummer(NeemMetNummer)
  'End If
  
  NeemSlag = VolgordeNummer(NeemMetNummer)

End Function

Private Sub imgInHanden_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 1 Then
    If Index < 13 Then
      If recInHerhaling() Then
        'StatusBar.SimpleText = "Je kunt nu geen kaart opleggen, omdat je een herhaling bekijkt."
        Tineke.Zeg "herhalingklik"
      Else
        KaartGekozenMens Index + 1
      End If
    Else
      StatusBar.SimpleText = "Deze kaarten zijn niet van jou. Jouw kaarten zie je onderin."
    End If
  ElseIf Button = 2 Then
    PopupMenu frmMenus.mnuKaarten
  End If
End Sub

Sub ToonRondjes(SpelerNr As Integer)
  Dim RondjeNr As Integer
  Dim Index As Integer
  Dim RondjesAfstand As Integer
  Dim nRondjes As Integer
  
  'If Spelers(SpelerNr).voorspelling > Spelers(SpelerNr).AantSlagen Then
  '  If Spelers(SpelerNr).voorspelling = 0 Then
  '    nRondjes = 1
  '  Else
  '    nRondjes = Spelers(SpelerNr).voorspelling
  '  End If
  'Else
  '  nRondjes = Spelers(SpelerNr).AantSlagen
  'End If
  
  If Spelers(SpelerNr).Voorspelling = 0 Then
    RondjesAfstand = imgSlagen(0).Width
  Else
    RondjesAfstand = Spelers(SpelerNr).Voorspelling * imgSlagen(0).Width
  End If 'Hoogte v/d rondjes = breedte
  
  If Ronde = 0 Then 'Opruimen
    For RondjeNr = 1 To Opties.MaxAantKaarten
      Index = 13 * (SpelerNr - 1) + RondjeNr - 1
      If RondjeImageGeladen(Index) Then
        imgSlagen(Index).Picture = LoadPicture("")
      End If
    Next RondjeNr
  Else
    For RondjeNr = 1 To AantKaartenRonde(Ronde)
      Index = 13 * (SpelerNr - 1) + RondjeNr - 1
      
      If Not RondjeImageGeladen(Index) Then
        Load imgSlagen(Index)
      End If
      
      Select Case SpelerNr
        Case 1
          imgSlagen(Index).Move (fraSpelen.Width - RondjesAfstand) / 2 + (RondjeNr - 1) * imgSlagen(0).Width, imgInHanden(0).Top - imgSlagen(0).Height - 30
        Case 2
          imgSlagen(Index).Move imgInHanden(13).Left + imgInHanden(13).Width + 30, (fraSpelen.Height - RondjesAfstand) / 2 + (RondjeNr - 1) * imgSlagen(0).Height
        Case 3
          imgSlagen(Index).Move (fraSpelen.Width + RondjesAfstand) / 2 - (RondjeNr) * imgSlagen(0).Width, imgInHanden(26).Top + imgInHanden(26).Height + 30
        Case 4
          imgSlagen(Index).Move imgInHanden(39).Left - imgSlagen(0).Width - 30, (fraSpelen.Height + RondjesAfstand) / 2 - (RondjeNr) * imgSlagen(0).Height
      End Select
      
      If RondjeNr <= Spelers(SpelerNr).Voorspelling Then  'Was voorspeld
        If RondjeNr <= Spelers(SpelerNr).AantSlagen Then
          imgSlagen(Index).Picture = LoadResPicture(401, vbResBitmap) 'Rondje vol
          imgSlagen(Index).ToolTipText = "Een voorspelde en gehaalde slag"
        ElseIf RondjeNr - Spelers(SpelerNr).AantSlagen <= KaartenResterend Then
          imgSlagen(Index).Picture = LoadResPicture(400, vbResBitmap) 'Rondje leeg
          imgSlagen(Index).ToolTipText = "Een voorspelde slag"
        Else 'Kan niet meer halen
          imgSlagen(Index).Picture = LoadResPicture(402, vbResBitmap) 'Rondje leeg fout
          imgSlagen(Index).ToolTipText = "Een onhaalbare voorspelde slag"
        End If
      Else 'Niet voorspeld
        If RondjeNr <= Spelers(SpelerNr).AantSlagen Then  'Is binnen
          imgSlagen(Index).Picture = LoadResPicture(403, vbResBitmap) 'Rondje fout
          imgSlagen(Index).ToolTipText = "Een onvoorspelde slag"
        Else 'Is ook niet binnen
          If RondjeNr = 1 And Spelers(SpelerNr).Voorspelling = 0 Then  'Wil 0 slagen
            imgSlagen(Index).Picture = LoadResPicture(404, vbResBitmap) 'NulSlagen
            imgSlagen(Index).ToolTipText = "0 slagen voorspeld"
          Else 'Wil meer dan 0
            imgSlagen(Index).Picture = LoadPicture("")
            imgSlagen(Index).ToolTipText = ""
          End If
        End If
      End If
      If Not RondjeImageGeladen(Index) Then
        imgSlagen(Index).Visible = True
        RondjeImageGeladen(Index) = True
      End If
    Next RondjeNr
    If AantKaartenRonde(Ronde) < Opties.MaxAantKaarten Then
      Index = 13 * (SpelerNr - 1) + AantKaartenRonde(Ronde) '+1 -1
      If RondjeImageGeladen(Index) Then
        imgSlagen(Index).Picture = LoadPicture("")
      End If
    End If
  End If

End Sub
Public Sub Sorteren(SpelerNr As Integer)
  Dim i As Integer
  Dim KaartNr As Integer
  Dim Temp As Kaart
  Dim Laagste As Integer
  Dim KaartRang(13) As Integer
  Dim Index As Integer
  
  Erase KaartRang
  For KaartNr = 1 To AantKaartenRonde(Ronde)
    If Spelers(SpelerNr).Kaarten(KaartNr).Getal <= 0 Then
      KaartRang(KaartNr) = 99 'Als het maar hoog is
    Else
      If Opties.AflopendSorteren Then
        KaartRang(KaartNr) = 13 * (Spelers(SpelerNr).Kaarten(KaartNr).Kleur - 1) + (14 - Spelers(SpelerNr).Kaarten(KaartNr).Getal)
      Else
        KaartRang(KaartNr) = 13 * (Spelers(SpelerNr).Kaarten(KaartNr).Kleur - 1) + Spelers(SpelerNr).Kaarten(KaartNr).Getal
      End If
    End If
  Next KaartNr
  
  For KaartNr = 1 To AantKaartenRonde(Ronde) - 1
    Laagste = KaartNr
    For i = KaartNr + 1 To AantKaartenRonde(Ronde)
      If KaartRang(i) < KaartRang(Laagste) Then
        Laagste = i
      End If
    Next i
    SwapKaarten Spelers(SpelerNr).Kaarten(KaartNr), Spelers(SpelerNr).Kaarten(Laagste)
    KaartRang(Laagste) = KaartRang(KaartNr) 'Omgekeerd hoeft niet
  Next KaartNr
 
  If SpelerNr = 1 Then
    For KaartNr = 1 To AantKaartenRonde(Ronde)
      Index = (SpelerNr - 1) * 13 + (KaartNr - 1)
      If Spelers(SpelerNr).Kaarten(KaartNr).Getal <= 0 Then
        frmMain.imgInHanden(Index).Visible = False
      Else
        frmMain.imgInHanden(Index).Visible = True
        LaadKaart frmMain.imgInHanden(Index), Spelers(SpelerNr).Kaarten(KaartNr)
      End If
    Next KaartNr
  End If
End Sub

Private Sub TestSlagTeHalen(SpelerNr)
  Dim SpelerTemp As Integer
  Dim i As Integer
  Dim GetalTemp As Integer
  Dim KleurTemp As Integer
 
  SlagIsTeHalenMetAantal = 0
  SlagNietTeHalenMetAantal = 0

  If SpelerNr = NuOpkomen Then
    SlagIsTeHalenMetAantal = 0
    For i = 1 To AantalMogelijkeKaarten
      If Spelers(SpelerNr).Kaarten(MogelijkeKaarten(i)).Getal > 0 Then
        SlagIsTeHalenMetAantal = SlagIsTeHalenMetAantal + 1
        SlagIsTeHalenMetNummer(SlagIsTeHalenMetAantal) = MogelijkeKaarten(i)
      End If
    Next i
  Else
    SlagIsTeHalenMetAantal = 0

    For i = 1 To AantalMogelijkeKaarten
      If HoogsteKaart(KaartenOpTafel(HoogsteOpTafel), Spelers(SpelerNr).Kaarten(MogelijkeKaarten(i))) = 2 Then
        SlagIsTeHalenMetAantal = SlagIsTeHalenMetAantal + 1
        SlagIsTeHalenMetNummer(SlagIsTeHalenMetAantal) = MogelijkeKaarten(i)
      Else
        SlagNietTeHalenMetAantal = SlagNietTeHalenMetAantal + 1
        SlagNietTeHalenMetNummer(SlagNietTeHalenMetAantal) = MogelijkeKaarten(i)
      End If
    Next i
  End If
End Sub

Private Sub Wacht(Duur As Single)
  Dim StartTimer As Single
  
  StartTimer = Timer
  Do
    If StartTimer + Duur - Timer > 0.25 Then '0.15
      DoEvents
    End If
  Loop Until Timer - StartTimer >= Duur Or Timer < StartTimer
End Sub

Private Sub ZoekMogelijkeKaarten(SpelerNr As Integer)
  Dim KaartNr As Integer
  Dim KanKleurBekennenTemp As Boolean
  Dim KanIntroevenTemp As Boolean
  'Als hij niet kan bedienen en wel troeven heeft,
  ' denkt hij dat hij alleen mag troeven. Waarom?

  AantTroefKaarten = 0
  AantLegaleTroefKaarten = 0
  AantNietTroefKaarten = 0
  AantLegaleNietTroefKaarten = 0
  AantalMogelijkeKaarten = 0
 
  KanKleurBekennenTemp = KanKleurBekennen(SpelerNr)
  KanIntroevenTemp = KanIntroeven(SpelerNr)

  For KaartNr = 1 To AantKaartenRonde(Ronde)
    With Spelers(SpelerNr).Kaarten(KaartNr)
      If .Getal > 0 Then
        If KanKleurBekennenTemp Then
          If .Kleur = KaartenOpTafel(NuOpkomen).Kleur Then
            .Legaal = True
          Else
            .Legaal = False
          End If
        Else 'Kan geen kleur bekennen
          If KanIntroevenTemp And Opties.IntroevenVerplicht Then 'Moet introeven
            If .Kleur = Troef.Kleur Then
              .Legaal = True
            Else
              .Legaal = False
            End If
          Else 'Introeven kan of hoeft niet
            .Legaal = True
          End If
        End If
        If .Legaal Then
          AantalMogelijkeKaarten = AantalMogelijkeKaarten + 1
          MogelijkeKaarten(AantalMogelijkeKaarten) = KaartNr
        End If
        If .Kleur = Troef.Kleur Then
          AantTroefKaarten = AantTroefKaarten + 1
          Troefkaarten(AantTroefKaarten) = KaartNr
          If .Legaal Then
            AantLegaleTroefKaarten = AantLegaleTroefKaarten + 1
            LegaleTroefkaarten(AantLegaleTroefKaarten) = KaartNr
          End If
        Else
          AantNietTroefKaarten = AantNietTroefKaarten + 1
          NietTroefkaarten(AantNietTroefKaarten) = KaartNr
          If .Legaal Then
            AantLegaleNietTroefKaarten = AantLegaleNietTroefKaarten + 1
            LegaleNietTroefkaarten(AantLegaleNietTroefKaarten) = KaartNr
          End If
        End If
      Else
        .Legaal = False
      End If
    End With
  Next KaartNr
 
End Sub

Private Sub mnuTotaal_Click()
  MeteenOptellenKlik
End Sub

Private Sub mnuVerwijderen_Click()
  Dim ShellRet As Double
  Dim MsgRet As VbMsgBoxResult
  
  MsgRet = MsgBox("Kies '10 op en neer' en klik op 'Toevoegen/Verwijderen...' in het volgende venster om 10 op en neer van deze computer te verwijderen.", vbInformation + vbOKCancel, "10 op en neer verwijderen")
  
  If MsgRet = vbOK Then
    'ShellRet = Shell("Rundll32.exe shell32,Control_RunDLL " & Chr(34) & "C:\WINDOWS\SYSTEM\APPWIZ.CPL" & Chr(34) & ",Software")
    ShellRet = Shell("rundll32.exe shell32.dll,Control_RunDLL appwiz.cpl,,0")
    SendKeys "{TAB}", True
    End
  End If
End Sub

Private Sub Netwerk_Error(Number As NetworkErrorType)
'  If WinsockFout = 0 Then
'    IkBenNetwerkClient = optServerClient(1).Value '#Alleen als het gelukt is
'    txtServerIP.Enabled = False
'  Else
  
  Select Case Number
    Case nerNotInitialized
      'Netwerk uitgeschakeld
      'If frmSpelSpeciaal.Visible Then
      '  frmSpelSpeciaal.lblNetwerkstatus.Caption = "Netwerkfunctionaliteit is uitgeschakeld."
      'End If
    Case nerBadNetwork
      'MsgBox "Kan geen verbinding maken. Het netwerk functioneert niet goed.", vbCritical, "Netwerkfout"
      'If frmSpelSpeciaal.Visible Then
      '  frmSpelSpeciaal.lblNetwerkstatus.Caption = "Kan geen verbinding maken. Het netwerk functioneert niet goed."
      'End If
    Case nerBadAdress
      MsgBox "Kan geen verbinding maken. Controleer de computernaam of het IP-adres van de server.", vbCritical, "Netwerkfout"
    Case Else
      MsgBox "Er is een netwerkfout opgetreden: " & vbCrLf & Error(Number) & vbCrLf & "(Foutcode " & Number & ")"
  End Select

'    frmMain.Winsock(0).Close
'    cmdVerbinden.Caption = "&Verbinden"
'    'StatusBar1.SimpleText = "Niet verbonden."
'    UpdateStatusbar
'  End If

End Sub

Private Sub picPraatwolkjePunt_Click()
  If Not Opties.Spelhulp Then
    Tineke.WegPraatwolkje
  End If
End Sub

Private Sub picScore_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim RondeKlik As Integer
  Dim Ret As VbMsgBoxResult
  
  If Button = 1 Then
    RondeKlik = Score.GetRondeVanMuisKlik(X, Y)
    If RondeKlik > 0 And recInHerhalingTest(RondeKlik, 0, 1) Then
      Ret = MsgBox("Wil je een herhaling bekijken vanaf deze ronde?", vbQuestion + vbOKCancel, "Herhaling")
      If Ret = vbOK Then
        Ronde = RondeKlik - 1
        NieuweRonde
      End If
    End If
  End If
End Sub

Private Sub picScore_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 1 Then
    'ToonScoreblok True
  ElseIf Button = 2 Then
    'PopupMenu frmMenus.mnuScoreblok
    Score.ToonPupupMenu
  End If
End Sub

Private Sub picTroef_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then
    PopupMenu frmMenus.mnuTroef
  End If
End Sub

Private Sub Praatwolkje_Click()
  If Not Opties.Spelhulp Then
    Tineke.WegPraatwolkje
  End If
End Sub

Private Sub Praatwolkje_Doorgaan()
  WachtOpGezien = False
  Tineke.WegPraatwolkje
End Sub

Private Sub Praatwolkje_StopHulp()
  Opties.Spelhulp = False
  SpelhulpAanUit
End Sub

Private Sub StatusBar_MouseDown(ByVal PanelIndex As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'Werkt niet
'  If Button = 2 Then
'    If PanelIndex = 0 Then
'      PopupMenu frmMenus.mnuSpeltype
'    ElseIf PanelIndex = 1 Then
'      PopupMenu frmMenus.mnuSlagenOver
'    End If
'  End If
End Sub

Private Sub timLaatsteSlag_Timer()
  ToonHuidigeSlag
End Sub

Private Sub timResize_Timer()
  FormResize
End Sub

Private Sub timSpelhulp_Timer()
  If Praatwolkje.ToonDoorgaan Then
'    If cmdGezien.BackColor = vbButtonFace Then
'      cmdGezien.BackColor = vbHighlight
'    Else
'      cmdGezien.BackColor = vbButtonFace
'    End If
  ElseIf Tineke.ToonPijltje And Opties.Spelhulp Then
    picPijltje.Visible = Not picPijltje.Visible
  Else
    timSpelhulp.Enabled = False
  End If
End Sub

Private Sub timTimeout_Timer()
' VerbindingStatus = 5
End Sub

Private Sub timVrouw_Timer()
  Dim Willekeurig As Integer
  
  Praatwolkje.Visible = False
  picPraatwolkjePunt.Visible = False

  Willekeurig = Int(100 * Rnd + 1)
 
  Select Case Willekeurig
    Case 1       'scheel
      imgVrouw.Picture = LoadResPicture(303, vbResBitmap)
      timVrouw.Interval = 2000
    Case 2 To 25 'knipper
      imgVrouw.Picture = LoadResPicture(302, vbResBitmap)
      timVrouw.Interval = 150
    Case 26 To 88 'normaal
      imgVrouw.Picture = LoadResPicture(300, vbResBitmap)
      timVrouw.Interval = 1000 + Int(1500 * Rnd)
    Case 89 To 100 'rechts
      imgVrouw.Picture = LoadResPicture(301, vbResBitmap)
      timVrouw.Interval = 1000 + Int(1500 * Rnd)
  End Select
End Sub

Sub InitVars()
  Dim i As Integer
  
  'MagSpelSpelen = True
  KaartImageGeladen(0) = True
  KaartImageGeladen(13) = True
  KaartImageGeladen(26) = True
  KaartImageGeladen(39) = True
  RondjeImageGeladen(0) = True
  RondjeImageGeladen(13) = True
  RondjeImageGeladen(26) = True
  RondjeImageGeladen(39) = True

  InlezenOpties
  InlezenStatistiek
  Statistiek.Sessies = Statistiek.Sessies + 1
  If DateDiff("d", Statistiek.LaatstGespeeld, Now) <> 0 Then
    Statistiek.LaatstGespeeld = Now 'Format(Now, "dd-mm-yyyy")
    Statistiek.DagenGespeeld = Statistiek.DagenGespeeld + 1
  End If

  'MaxAantKaarten = 10
  'nRonden = 19
  'BepaalAantKaartenPerRonde
  
  AchterkantBestand = DirPlusBestand(CStr(App.Path), "Achterkant\" & AchterkantTitle)
 
  'Scoreblok = sblScore

  Spelers(1).Controller = ControllerMens
  Spelers(2).Controller = ControllerComputer
  Spelers(3).Controller = ControllerComputer
  Spelers(4).Controller = ControllerComputer
 
  KleurNaam(1) = "klaver"
  KleurNaam(2) = "ruiten"
  KleurNaam(3) = "schoppen"
  KleurNaam(4) = "harten"
 
  GetalNaam(2) = "twee"
  GetalNaam(3) = "drie"
  GetalNaam(4) = "vier"
  GetalNaam(5) = "vijf"
  GetalNaam(6) = "zes"
  GetalNaam(7) = "zeven"
  GetalNaam(8) = "acht"
  GetalNaam(9) = "negen"
  GetalNaam(10) = "tien"
  GetalNaam(11) = "boer"
  GetalNaam(12) = "vrouw"
  GetalNaam(13) = "heer"
  GetalNaam(14) = "aas"
  
  For i = 2 To 14
    KaartWaarde(i) = 0.7 * ((i / 14) ^ 2.8)
    'If i > 7 Then
    ' KaartWaarde(i) = (i - 7) / 10
    'End If
  Next i
 
  GeenKaart.Getal = 0
  GeenKaart.Kleur = 0
 
  KaartAfst = 225
  'BasisPoort = 90
  WachtOpGezien = False

  Hulpniveau = 0
  Score.MagInHiScore = True
  
  LaatsteSlag(1).Legaal = False
  LaatsteSlag(2).Legaal = False
  LaatsteSlag(3).Legaal = False
  LaatsteSlag(4).Legaal = False
End Sub
Sub ResetVars()
  Dim SpelerNr As Integer
  Dim RondeNrLokaal As Integer
  
  'BepaalAantKaartenPerRonde
  nRonden = 2 * Opties.MaxAantKaarten - 1
  
  Score.NieuwSpel

  For SpelerNr = 1 To 4
    Spelers(SpelerNr).Rang = 0
    Spelers(SpelerNr).TotaalScore = 0
    For RondeNrLokaal = 1 To nRonden
      Spelers(SpelerNr).Score(RondeNrLokaal) = -1
      Spelers(SpelerNr).ScoreFout(RondeNrLokaal) = False
    Next RondeNrLokaal
  Next SpelerNr
  
  'HerhalingBezig = False
  recWasInHerhaling = False
  Ronde = 0
  WachtOp = Niets
  
End Sub
Sub Voorspellen()
  Dim KnopNr As Integer
  
  StatusBar.PanelText(1) = SlagenOverTekst()
  If Not recInHerhaling Then
    recSpelerNum = SpelerNum
    recSlagNr = 0
  End If
  
  If Spelers(NuVoorspellen).Controller = ControllerMens Then
    'Tip = -1
    If Not recInHerhaling() Then
      WachtOp = MensVoorspelling
      Tineke.ZegHulp "voorspellen"
  
      StatusBar.SimpleText = "Klik op een knop om het aantal slagen te voorspellen."
  
      For KnopNr = 0 To 11 'Meer past niet, jammer dan
        If KnopNr <= AantKaartenRonde(Ronde) Then
          cmdVoorspellen(KnopNr).Visible = True
        Else
          cmdVoorspellen(KnopNr).Visible = False
        End If
      Next KnopNr
  
      fraVoorspellen.ZOrder 0
      fraVoorspellen.Visible = True
      cmdVoorspellen(0).SetFocus
      picScore.TabStop = False
          
      Spelers(NuVoorspellen).Taxatie = TaxeerKaarten(NuVoorspellen)
      If NuVoorspellen = 1 Then 'Moet haast wel
        'Als we toch al de taxatie weten
        Tip = LegaleVoorspelling(Spelers(NuVoorspellen).Taxatie)
      End If
    Else
      Spelers(NuVoorspellen).Voorspelling = recVoorspellingen(Ronde, NuVoorspellen)
      Voorspeld
    End If
  ElseIf Spelers(NuVoorspellen).Controller = ControllerComputer Then
    StatusBar.SimpleText = "Wacht op de beurt van " & Spelers(NuVoorspellen).Naam & "..."
    Spelers(NuVoorspellen).Taxatie = TaxeerKaarten(NuVoorspellen)
    If recInHerhaling() Then
      'recVoorspellingen(Ronde, NuVoorspellen) = LegaleVoorspelling(Spelers(NuVoorspellen).Taxatie)
      Spelers(NuVoorspellen).Voorspelling = recVoorspellingen(Ronde, NuVoorspellen)
    Else
      Spelers(NuVoorspellen).Voorspelling = LegaleVoorspelling(Spelers(NuVoorspellen).Taxatie)
    End If
    
    Voorspeld
  End If

End Sub
Sub ResetRondeVars()
  Dim SpelerNr As Integer
  Dim KaartNr As Integer
  'Dim Temp As Kaart
  Dim i As Integer
  Dim kl As Integer
  Dim ge As Integer
  
  'Randomize Timer
  
  TotSlagenGok = 0
  AantSpelersGegokt = 0
  AantKaartenGezien = 0
  SlagNr = 0
  SlagenOver = AantKaartenRonde(Ronde) 'Dit lijkt wel erg veel op elkaar...
  KaartenResterend = AantKaartenRonde(Ronde)
  'SpatieGedrukt = False
  OndergangGeluidGespeeld = False

  Score.NieuwSpel
  For SpelerNr = 1 To 4
    Spelers(SpelerNr).Voorspelling = -1
    'PositieInHiScore(SpelerNr) = 0
    KaartenOpTafel(SpelerNr) = GeenKaart
  Next SpelerNr

  For KaartNr = 1 To 52
'    DeStapel.Kaarten(KaartNr).Kleur = (KaartNr - 1) \ 13 + 1
'    DeStapel.Kaarten(KaartNr).Getal = (KaartNr - 1) Mod 13 + 2
    kl = (KaartNr - 1) \ 13 + 1
    ge = (KaartNr - 1) Mod 13 + 2
    WaarIsKaart(kl, ge) = 0
  Next KaartNr
'
'  For KaartNr = 1 To 52 'Schudden
'    'i = Int(52 * Rnd + 1)                  '--> i = 1..52
'    i = Int((53 - KaartNr) * Rnd + KaartNr) '--> i = KaartNr..52; eerlijker
'    Temp = DeStapel.Kaarten(KaartNr)
'    DeStapel.Kaarten(KaartNr) = DeStapel.Kaarten(i)
'    DeStapel.Kaarten(i) = Temp
'  Next KaartNr
'  DeStapel.AantKaarten = 52

  'Erase KaartGezien
  Erase Spelers(1).HeeftKleurNietMeer
  Erase Spelers(2).HeeftKleurNietMeer
  Erase Spelers(3).HeeftKleurNietMeer
  Erase Spelers(4).HeeftKleurNietMeer

End Sub
Sub KaartLeggen(KaartNr As Integer)
  Dim KleurGekozen As Integer
  Dim GetalGekozen As Integer
  Dim i As Integer
  Dim SchuifX1 As Integer, SchuifY1 As Integer
  Dim SchuifX2 As Integer, SchuifY2 As Integer
  
  If KaartNr = 0 Then
    Stop
  End If

  If NuOpleggen = 1 Then 'And Spelers(1).AantKaarten = 0 Then
    imgKaartAanwijs.Visible = False
  End If

  KaartenOpTafel(NuOpleggen) = Spelers(NuOpleggen).Kaarten(KaartNr)
  KaartenOpTafel(NuOpleggen).Legaal = True
  
  imgInHanden((NuOpleggen - 1) * 13 + (KaartNr - 1)).Visible = False
  LaadKaart imgOpTafel(NuOpleggen - 1), Spelers(NuOpleggen).Kaarten(KaartNr)
  imgOpTafel(NuOpleggen - 1).ToolTipText = "De kaart die " & Spelers(NuOpleggen).Naam & " heeft opgelegd"
 
  SchuifX1 = imgInHanden((NuOpleggen - 1) * 13 + (KaartNr - 1)).Left
  SchuifY1 = imgInHanden((NuOpleggen - 1) * 13 + (KaartNr - 1)).Top
  SchuifX2 = imgOpTafel(NuOpleggen - 1).Left
  SchuifY2 = imgOpTafel(NuOpleggen - 1).Top
  Schuif imgOpTafel(NuOpleggen - 1), SchuifX1, SchuifY1, SchuifX2, SchuifY2, 2, True

  Spelers(NuOpleggen).Kaarten(KaartNr) = GeenKaart

  WaarIsKaart(KaartenOpTafel(NuOpleggen).Kleur, KaartenOpTafel(NuOpleggen).Getal) = -2
  Spelers(NuOpleggen).AantKaarten = Spelers(NuOpleggen).AantKaarten - 1
  If KaartenOpTafel(NuOpleggen).Kleur <> KaartenOpTafel(NuOpkomen).Kleur Then
    Spelers(NuOpleggen).HeeftKleurNietMeer(KaartenOpTafel(NuOpkomen).Kleur) = True
    If Opties.IntroevenVerplicht And KaartenOpTafel(NuOpleggen).Kleur <> Troef.Kleur Then
      If Troef.Kleur <> 0 Then
        Spelers(NuOpleggen).HeeftKleurNietMeer(Troef.Kleur) = True
      End If
    End If
  End If

  If NuOpleggen = NuOpkomen Then
    KleurAantalKerenGespeeld(KaartenOpTafel(NuOpleggen).Kleur) = KleurAantalKerenGespeeld(KaartenOpTafel(NuOpleggen).Kleur) + 1
  End If
  AantKaartenGezien = AantKaartenGezien + 1

  If NuOpleggen = NuOpkomen Then
    HoogsteOpTafel = NuOpleggen
  ElseIf HoogsteKaart(KaartenOpTafel(HoogsteOpTafel), KaartenOpTafel(NuOpleggen)) = 2 Then
    HoogsteOpTafel = NuOpleggen
  End If
  
  If Opties.SpelSnelheid <= 1 Then
    Wacht (-Opties.SpelSnelheid + 6) / 5
  End If
  'frmMenus.mnuIntroevenVerplicht.Enabled = False
  
  NuOpleggen = VolgendeSpeler(NuOpleggen)
  
  If NuOpleggen = NuOpkomen Then
    SlagNaarHoogste
    SlagNr = SlagNr + 1
    If KaartenResterend = 0 Then
      'Score.NScoreRondes = Ronde
      If Not recInHerhaling Then
        BerekenScore
        BepaalRangen
        Score.ToonVoorspelling = False
        Score.ToonScore
      End If
      Tineke.ZegHulp "score"
      If Ronde = nRonden Then
        SpelAfgelopen
      Else
        'EerstOpkomen = VolgendeSpeler(EerstOpkomen)
        'EerstOpkomen = (recWieBegint + (Ronde + 1) + 2) Mod 4 + 1
        NieuweRonde
      End If
    Else
'      If Not recInHerhaling Then
'        recNuOpkomen(Ronde, SlagNr) = HoogsteOpTafel
'      End If
      
      NuOpkomen = HoogsteOpTafel 'recNuOpkomen(Ronde, SlagNr)
      NuOpleggen = NuOpkomen
      
'      SlagNr = SlagNr + 1
      SpelerNum = 1
      KaartKiezen
    End If
  Else
    SpelerNum = SpelerNum + 1
    KaartKiezen
  End If
End Sub
Sub SlagNaarHoogste()
  Dim SpelerTemp As Integer
  Dim SchuifX1 As Integer, SchuifY1 As Integer
  Dim SchuifX2 As Integer, SchuifY2 As Integer
  Dim EenAnderNatGegaan As Boolean
  Dim SpelerNr As Integer
  
  'Randomize Timer
  
  KaartenResterend = KaartenResterend - 1
'  SlagNr = SlagNr + 1
  
  StatusBar.SimpleText = "De slag is voor " & Spelers(HoogsteOpTafel).Naam & "."
  
  Spelers(HoogsteOpTafel).AantSlagen = Spelers(HoogsteOpTafel).AantSlagen + 1
  
  Tineke.ZegHulp "Slag naar hoogste"
  
  Wacht (-Opties.SpelSnelheid + 6) / 5 + 0.1
  ToonRondjes 1
  ToonRondjes 2
  ToonRondjes 3
  ToonRondjes 4

  For SpelerNr = 2 To 4
    With Spelers(SpelerNr)
      If HoogsteOpTafel = SpelerNr And .AantSlagen = .Voorspelling + 1 Then
        EenAnderNatGegaan = True
      ElseIf .Voorspelling - .AantSlagen = .AantKaarten + 1 Then
        EenAnderNatGegaan = True
      End If
    End With
  Next SpelerNr
  If EenAnderNatGegaan Then
    WavPlay "Medespeler nat"
  End If
  
  If HoogsteOpTafel = 1 And Spelers(1).AantSlagen = 1 + Spelers(1).Voorspelling Then
    If Not OndergangGeluidGespeeld Then
      WavPlay "Slag te veel"
    End If
    If Int(3 * Rnd) = 0 Then
      Tineke.Zeg "te veel"
    End If
  ElseIf HoogsteOpTafel <> 1 And Spelers(1).Voorspelling - Spelers(1).AantSlagen = Spelers(1).AantKaarten + 1 Then
    If Not OndergangGeluidGespeeld Then
      WavPlay "Slag te weinig"
    End If
    If Int(3 * Rnd) = 0 Then
      Tineke.Zeg "te weinig"
    End If
  ElseIf Spelers(1).Voorspelling = Spelers(1).AantSlagen And Spelers(1).AantKaarten = 0 Then
    WavPlay "Goed voorspeld"
    If Int(4 * Rnd) = 0 Then
      Tineke.Zeg "goed"
    End If
  End If
  
  'If SpelerKaartGekozen = -1 Then 'Speler is niet ongeduldig
  Wacht (-Opties.SpelSnelheid + 7) / 6
  'End If
  
  If timLaatsteSlag.Enabled Then
    ToonHuidigeSlag
  End If
  
  Select Case HoogsteOpTafel
    Case 1
      SchuifX2 = fraSpelen.Width / 2 - imgInHanden(0).Width
      SchuifY2 = fraSpelen.Height
    Case 2
      SchuifX2 = -imgInHanden(0).Width
      SchuifY2 = (fraSpelen.Height - imgInHanden(0).Height) / 2
    Case 3
      SchuifX2 = fraSpelen.Width / 2 - imgInHanden(0).Width
      SchuifY2 = -imgInHanden(0).Height
    Case 4
      SchuifX2 = fraSpelen.Width
      SchuifY2 = (fraSpelen.Height - imgInHanden(0).Height) / 2
  End Select
  SpelerTemp = VorigeSpeler(NuOpkomen)
  Do
    LaatsteSlag(SpelerTemp) = KaartenOpTafel(SpelerTemp)
    SchuifX1 = imgOpTafel(SpelerTemp - 1).Left
    SchuifY1 = imgOpTafel(SpelerTemp - 1).Top
   
    Schuif imgOpTafel(SpelerTemp - 1), SchuifX1, SchuifY1, SchuifX2, SchuifY2, 1, False
    
    imgOpTafel(SpelerTemp - 1).Visible = False
    WaarIsKaart(KaartenOpTafel(SpelerTemp).Kleur, KaartenOpTafel(SpelerTemp).Getal) = -3
    KaartenOpTafel(SpelerTemp).Getal = 0
    KaartenOpTafel(SpelerTemp).Kleur = 0
    KaartenOpTafel(SpelerTemp).Legaal = False
    SpelerTemp = VorigeSpeler(SpelerTemp)
    'DoEvents
  Loop Until SpelerTemp = VorigeSpeler(NuOpkomen)
  LaatsteSlagOpgekomenSpeler = NuOpkomen
  mnuSpelLaatsteSlagTonen.Enabled = True
 
  Tineke.ZegHulp "slag erbij"
  If mnuSpelhulp.Checked Then
    Hulpniveau = Hulpniveau + 1
  End If
 
End Sub
Sub ResetControls()
  Dim RondjeNr As Integer
  
  fraEinde.Visible = False
  fraTroef.Visible = True
  fraSpelen.Visible = True
  imgGewonnen.Visible = False
  mnuSpelLaatsteSlagTonen.Enabled = False
  fraVoorspellen.Visible = False
  'frmMenus.mnuIntroevenVerplicht.Enabled = True
  'ToonRondmakenToegestaan
  'ToonIntroevenVerplicht
  ToonSpelType
  ToonRondjes 1
  ToonRondjes 2
  ToonRondjes 3
  ToonRondjes 4
  
End Sub
Sub InitControls()
  Dim SpelerNr As Integer
  Dim KaartNr As Integer
  Dim Foutcode As Integer
  
  imgVrouw.Picture = LoadResPicture(300, vbResBitmap)
  
  'frmMain.MMControl1.DeviceType = "WaveAudio"
   
  For SpelerNr = 0 To 3
    lblNaam(SpelerNr).Caption = Spelers(SpelerNr + 1).Naam
    lblNaamSchaduw(SpelerNr).Caption = Spelers(SpelerNr + 1).Naam
  Next SpelerNr
 
  fraVoorspellen.BackColor = Achterkleur
  picTroef.BorderStyle = vbTransparent
  'ToonHiScore

  If IsNumeric(AchterkantTitle) Then
    imgInHanden(13).Picture = LoadResPicture(200 + CInt(AchterkantTitle), vbResBitmap)
  Else
    If Len(Dir(AchterkantBestand)) And Len(Dir(AchterkantBestand & "NUL")) = 0 Then 'Het is geen map
      imgInHanden(13).Picture = LoadPicture(AchterkantBestand)
    Else
      imgInHanden(13).Picture = LoadResPicture(200, vbResBitmap) 'Standaardachterkant
    End If
  End If
  imgInHanden(26).Picture = imgInHanden(13).Picture
  imgInHanden(39).Picture = imgInHanden(13).Picture

  'ToonRondmakenToegestaan
  'picTroef.BorderStyle = 0

  mnuGeluid.Checked = Opties.Geluid
  mnuOmgekeerdSorteren.Checked = Opties.AflopendSorteren
  frmMenus.mnuOmgekeerdSorteren.Checked = Opties.AflopendSorteren
  mnuTotaal.Checked = Opties.MeteenOptellen
  frmMenus.mnuTotaalscore.Checked = Opties.MeteenOptellen
  mnuBreedUitspreiden.Checked = Opties.BreedUitspreiden
  frmMenus.mnuBreedUitspreiden.Checked = Opties.BreedUitspreiden
  mnuSpelhulp.Checked = Opties.Spelhulp
  frmMenus.mnuSpelhulp.Checked = Opties.Spelhulp
  mnuCommentaar.Checked = Opties.Commentaar
  frmMenus.mnuCommentaar.Checked = Opties.Commentaar
  mnuOptiesComputersInHighscore.Checked = Opties.ComputersInHighScore
  frmMenus.mnuHighscoreComputersInHighscore.Checked = Opties.ComputersInHighScore
  mnuOptiesGroot.Checked = Opties.GroteKaarten
  mnuOptiesRondmaken.Checked = Opties.RondmakenToegestaan
  mnuOptiesVerplichtIntroeven.Checked = Opties.IntroevenVerplicht
  
  frmMenus.mnuKaartenGroteKaarten.Checked = Opties.GroteKaarten
  If Opties.GroteKaarten Then
    KaartgrootteInstellen
  End If
  
End Sub

Private Sub txtNaamWijzig_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    Spelers(NaamWijzigen).Naam = txtNaamWijzig.Text
    lblNaam(NaamWijzigen - 1).Caption = Spelers(NaamWijzigen).Naam
    lblNaamSchaduw(NaamWijzigen - 1).Caption = Spelers(NaamWijzigen).Naam
    txtNaamWijzig.Visible = False
    Score.UpdateSpelerNaam
    Score.Update
    KeyAscii = 0
  ElseIf KeyAscii = vbKeyEscape Then
    txtNaamWijzig.Visible = False
    Score.UpdateSpelerNaam
    'Score.Update
    KeyAscii = 0
  End If
End Sub

Sub ToonNetwerkSpelers()
 Dim SpelerNr As Integer
 Dim Tekst As String
 
 For SpelerNr = 1 To 4
  If Spelers(SpelerNr).Controller = "netwerk" Then
   If Len(Tekst) Then
    Tekst = Tekst & ", "
   End If
   Tekst = Tekst & Spelers(SpelerNr).Naam
  End If
 Next SpelerNr
 If Len(Tekst) Then
  StatusBar.SimpleText = "Spelers: " & Tekst
 End If

End Sub
Function NetNaarLokaalSpeler(NetSpeler As Integer) As Integer
  NetNaarLokaalSpeler = (NetSpeler - IkBenNetSpelerNr + 4) Mod 4 + 1
End Function

Private Sub SwapKaarten(Kaart1 As Kaart, Kaart2 As Kaart)
  Dim Temp As Kaart
  Temp = Kaart1
  Kaart1 = Kaart2
  Kaart2 = Temp
End Sub

Function KaartenErboven(SpelerNr As Integer, KaartNr As Integer) As Integer
  'Hier kunnen we de pc laten valsspelen...
  Dim Getal As Integer
  Dim ErbovenTemp As Integer
  'Dim MaxErboven As Integer
  Dim SpelerNrTemp As Integer
  Dim DeKaart As Kaart
  Dim nSpelersZekerNietHoger As Integer
  
  DeKaart = Spelers(SpelerNr).Kaarten(KaartNr)
  
  ErbovenTemp = 0
  'MaxErboven = 0
  
  For Getal = DeKaart.Getal + 1 To 14
    If WaarIsKaart(DeKaart.Kleur, Getal) >= 0 And _
            WaarIsKaart(DeKaart.Kleur, Getal) <= 4 And _
            WaarIsKaart(DeKaart.Kleur, Getal) <> SpelerNr Then
      'Om 't even of de kaart op de stapel ligt of in iemands handen
      ErbovenTemp = ErbovenTemp + 1
    End If
  Next Getal
  If DeKaart.Kleur <> Troef.Kleur And Troef.Kleur <> 0 Then
    For Getal = 2 To 14
      If WaarIsKaart(Troef.Kleur, Getal) >= 0 And _
              WaarIsKaart(Troef.Kleur, Getal) <= 4 And _
              WaarIsKaart(Troef.Kleur, Getal) <> SpelerNr Then
       'Om 't even of de kaart op de stapel ligt of in iemands handen
        ErbovenTemp = ErbovenTemp + 1
      End If
    Next Getal
  End If
  
  'Maximum begrensd door aantal kaarten en door HeeftKleurNietMeer
'  For SpelerNrTemp = 1 To 4
'    If SpelerNrTemp <> SpelerNr Then
'      If DeKaart.Kleur = Troef.Kleur Then
'        If Spelers(SpelerNrTemp).HeeftKleurNietMeer(DeKaart.Kleur) Then
'          MaxErboven = MaxErboven + 0
'        Else
'          MaxErboven = MaxErboven + Spelers(SpelerNrTemp).AantKaarten
'        End If
'      Else
'        If Spelers(SpelerNrTemp).HeeftKleurNietMeer(DeKaart.Kleur) And _
'           Spelers(SpelerNrTemp).HeeftKleurNietMeer(Troef.Kleur) Then
'          MaxErboven = MaxErboven + 0
'        Else
'          MaxErboven = MaxErboven + Spelers(SpelerNrTemp).AantKaarten
'        End If
'      End If
'    End If
'  Next SpelerNrTemp
  
'  If ErbovenTemp > MaxErboven Then
'    Debug.Print "Oud: " & ErbovenTemp & " Nieuw: " & MaxErboven
'    Stop
'    ErbovenTemp = MaxErboven
'  End If
  
  'Maximum alleen begrensd door HeeftKleurNietMeer
  nSpelersZekerNietHoger = 0
  For SpelerNrTemp = 1 To 4
    If SpelerNrTemp <> SpelerNr Then
      If DeKaart.Kleur = Troef.Kleur Then
        If Spelers(SpelerNrTemp).HeeftKleurNietMeer(DeKaart.Kleur) Then
          nSpelersZekerNietHoger = nSpelersZekerNietHoger + 1
        End If
      Else
        If Troef.Kleur <> 0 Then
          If Spelers(SpelerNrTemp).HeeftKleurNietMeer(DeKaart.Kleur) And _
            Spelers(SpelerNrTemp).HeeftKleurNietMeer(Troef.Kleur) Then
            nSpelersZekerNietHoger = nSpelersZekerNietHoger + 1
          End If
        End If
      End If
    End If
  Next SpelerNrTemp
  If nSpelersZekerNietHoger = 3 Then
    ErbovenTemp = 0
  End If
  
  KaartenErboven = ErbovenTemp
End Function
Function KaartenEronder(SpelerNr As Integer, KaartNr As Integer) As Integer
  Dim Getal As Integer
  Dim KleurNr As Integer
  Dim EronderTemp As Integer
  
  EronderTemp = 0
  
  KleurNr = Spelers(SpelerNr).Kaarten(KaartNr).Kleur
  For Getal = 2 To Spelers(SpelerNr).Kaarten(KaartNr).Getal - 1
    If WaarIsKaart(KleurNr, Getal) >= 0 And WaarIsKaart(KleurNr, Getal) <= 4 And WaarIsKaart(KleurNr, Getal) <> SpelerNr Then
      'Om 't even of de kaart op de stapel ligt of in iemands handen
      EronderTemp = EronderTemp + 1
    End If
  Next Getal
  
  KaartenEronder = EronderTemp
End Function

Function KaartWaardeNieuw(SpelerNr As Integer, KaartNr As Integer, AlsIkHemNuOpleg As Boolean) As Single
  Dim KaartenOver As Integer
  Dim LagereOver As Integer
  Dim SpelerNrNu As Integer
  Dim Teller As Long 'Voor kansrekening
  Dim Noemer As Long 'Voor kansrekening
  Dim VanSpeler As Integer
  Dim NaarSpeler As Integer
  
  If SpelerNr = NuOpkomen Then
    VanSpeler = VolgendeSpeler(SpelerNr)
    NaarSpeler = VorigeSpeler(SpelerNr)
  ElseIf KaartenOpTafel(NuOpkomen).Getal = 0 Then
    'Niemand heeft nog opgelegd
    VanSpeler = VolgendeSpeler(SpelerNr)
    NaarSpeler = VorigeSpeler(SpelerNr)
  Else
    'Anderen hebben al opgelegd
    If AlsIkHemNuOpleg And HoogsteKaart(KaartenOpTafel(HoogsteOpTafel), Spelers(SpelerNr).Kaarten(KaartNr)) = 1 Then
      KaartWaardeNieuw = 0 'Lager dan de hoogste kaart op tafel
      VanSpeler = 0
      NaarSpeler = 0
    Else
      'Kaart is hoger dan de hoogste op tafel
      If AlsIkHemNuOpleg And VolgendeSpeler(SpelerNr) = NuOpkomen Then
        KaartWaardeNieuw = 1
        VanSpeler = 0
        NaarSpeler = 0
      Else
        VanSpeler = VolgendeSpeler(SpelerNr)
        NaarSpeler = VorigeSpeler(NuOpkomen)
      End If
    End If
  End If
    
  If VanSpeler > 0 Then
    KaartenOver = DeStapel.AantKaarten
    'Zo nemen we andere kleuren ook mee (zijn niet hoger of lager)
    SpelerNrNu = VanSpeler
  
    Do
      KaartenOver = KaartenOver + Spelers(SpelerNrNu).AantKaarten
      SpelerNrNu = VolgendeSpeler(SpelerNrNu)
    Loop Until SpelerNrNu = VolgendeSpeler(NaarSpeler)
  
    'Dit is bij het kaart bijleggen
    'Bij opkomen: LagereOver=kaarteneronder
    '** Ook niet helemaal goed
    If SpelerNr = NuOpkomen Or Not AlsIkHemNuOpleg Then
      LagereOver = KaartenEronder(SpelerNr, KaartNr)
    Else
      LagereOver = KaartenOver - KaartenErboven(SpelerNr, KaartNr)
    End If
    
    Teller = 1
    Noemer = 1
  
    SpelerNrNu = VanSpeler
    
    Do
      If Spelers(SpelerNrNu).HeeftKleurNietMeer(Spelers(SpelerNr).Kaarten(KaartNr).Kleur) Then
      Else
        'Klopt niet: andere spelers hoeven niet een lagere kaart van dezelfde kleur te hebben
        Teller = Teller * LagereOver
        Noemer = Noemer * KaartenOver
        LagereOver = LagereOver - 1
        KaartenOver = KaartenOver - 1
      End If
      SpelerNrNu = VolgendeSpeler(SpelerNrNu)
    Loop Until SpelerNrNu = VolgendeSpeler(NaarSpeler)
    KaartWaardeNieuw = Teller / Noemer
  End If
End Function

Sub InfoTemp()
  Dim KaartNr As Integer '# procedure weg
  Debug.Print
  For KaartNr = 1 To AantKaartenRonde(Ronde)
    If Spelers(1).Kaarten(KaartNr).Getal > 0 Then
      'Debug.Print KaartNaam(Spelers(1).Kaarten(KaartNr)) & ": Kans: " & CInt(100 * KaartWaardeNieuw(1, KaartNr)) & "%. " & KaartenErboven(1, KaartNr) & " boven, " & KaartenEronder(1, KaartNr) & " onder."
    End If
  Next KaartNr
End Sub
Sub SpelhulpAanUit()
  mnuSpelhulp.Checked = Opties.Spelhulp
  frmMenus.mnuSpelhulp.Checked = Opties.Spelhulp
  'cmdHelp.Visible = Not Opties.Spelhulp
  'IniSet "Opties", "Spelhulp", Opties.Spelhulp

  Hulpniveau = 1
  If Opties.Spelhulp Then
    If Tineke.HulpTekst <> "" Then
      If Ronde = 0 Then
        Tineke.ToonPraatwolkje True
      Else
        Tineke.ToonPraatwolkje True
      End If
    End If
    'If Opties.SpelSnelheid > 2 Then
    '  Opties.SpelSnelheid = 2
    'End If
  Else
    WachtOpGezien = False
    Praatwolkje.Visible = False
    picPraatwolkjePunt.Visible = False
    picPijltje.Visible = False
    Tineke.Zeg "hulp uit"
  End If
End Sub

Sub ZoekAanTeWijzenKaart(Richting As String)
  Dim i As Integer
  
  If Richting = "r" Then
    i = KaartAanwijzen + 1
    Do Until i > AantKaartenRonde(Ronde)
      If Spelers(1).Kaarten(i).Legaal Then
        KaartAanwijzen = i
        Exit Do
      End If
      i = i + 1
    Loop
  Else
    i = KaartAanwijzen - 1
    Do Until i < 1
      If Spelers(1).Kaarten(i).Legaal Then
        KaartAanwijzen = i
        Exit Do
      End If
      i = i - 1
    Loop
  End If

End Sub
Sub SchikKaarten(SpelerNr As Integer)
  Dim KaartNr As Integer
  Dim Index As Integer
  Dim EersteKaartLeft As Integer
  Dim EersteKaartTop As Integer
  Dim KaartAfstSpeler As Integer

  If Ronde <= 0 Then Exit Sub
  
  If Opties.BreedUitspreiden Then
    KaartAfstSpeler = 300
  Else
    KaartAfstSpeler = KaartAfst
  End If
  
  Select Case SpelerNr
    Case 1
      EersteKaartLeft = fraSpelen.Width / 2 - ((AantKaartenRonde(Ronde) - 1) * KaartAfstSpeler + imgInHanden(0).Width) / 2
      EersteKaartTop = fraSpelen.Height - imgInHanden(0).Height - 120
      For KaartNr = 1 To AantKaartenRonde(Ronde)
        Index = (SpelerNr - 1) * 13 + (KaartNr - 1)
        If Not KaartImageGeladen(Index) Then
          Load imgInHanden(Index)
          KaartImageGeladen(Index) = True
          imgInHanden(Index).Picture = imgInHanden(13).Picture
          imgInHanden(Index).ZOrder 0
        End If
        imgInHanden(Index).Move EersteKaartLeft + (KaartNr - 1) * KaartAfstSpeler, EersteKaartTop
      Next KaartNr
      lblNaam(0).Move imgInHanden(0).Left - lblNaam(0).Width - 15, imgInHanden(0).Top + imgInHanden(0).Height - lblNaam(0).Height
      lblNaamSchaduw(0).Move lblNaam(0).Left + 15, lblNaam(0).Top + 15
    Case 2
      EersteKaartLeft = 120
      EersteKaartTop = fraSpelen.Height / 2 - ((AantKaartenRonde(Ronde) - 1) * KaartAfst + imgInHanden(0).Height) / 2
      For KaartNr = 1 To AantKaartenRonde(Ronde)
        Index = (SpelerNr - 1) * 13 + (KaartNr - 1)
        If Not KaartImageGeladen(Index) Then
          Load imgInHanden(Index)
          KaartImageGeladen(Index) = True
          imgInHanden(Index).Picture = imgInHanden(13).Picture
          imgInHanden(Index).ZOrder 0
        End If
        imgInHanden(Index).Move EersteKaartLeft, EersteKaartTop + (KaartNr - 1) * KaartAfst
      Next KaartNr
      lblNaam(1).Move imgInHanden(13).Left, imgInHanden(13).Top - lblNaam(1).Height
      lblNaamSchaduw(1).Move lblNaam(1).Left + 15, lblNaam(1).Top + 15
    Case 3
      EersteKaartLeft = fraSpelen.Width / 2 + ((AantKaartenRonde(Ronde) - 1) * KaartAfst - imgInHanden(0).Width) / 2
      EersteKaartTop = 120
      For KaartNr = 1 To AantKaartenRonde(Ronde)
        Index = (SpelerNr - 1) * 13 + (KaartNr - 1)
        If Not KaartImageGeladen(Index) Then
          Load imgInHanden(Index)
          KaartImageGeladen(Index) = True
          imgInHanden(Index).Picture = imgInHanden(13).Picture
          imgInHanden(Index).ZOrder 0
        End If
        imgInHanden(Index).Move EersteKaartLeft - (KaartNr - 1) * KaartAfst, EersteKaartTop
      Next KaartNr
      lblNaam(2).Move imgInHanden(26).Left + imgInHanden(26).Width, imgInHanden(26).Top
      lblNaamSchaduw(2).Move lblNaam(2).Left + 15, lblNaam(2).Top + 15
    Case 4
      EersteKaartLeft = fraSpelen.Width - imgInHanden(0).Width - 120
      EersteKaartTop = fraSpelen.Height / 2 + ((AantKaartenRonde(Ronde) - 1) * KaartAfst - imgInHanden(0).Height) / 2
      For KaartNr = 1 To AantKaartenRonde(Ronde)
        Index = (SpelerNr - 1) * 13 + (KaartNr - 1)
        If Not KaartImageGeladen(Index) Then
          Load imgInHanden(Index)
          KaartImageGeladen(Index) = True
          imgInHanden(Index).Picture = imgInHanden(13).Picture
          imgInHanden(Index).ZOrder 0
        End If
        imgInHanden(Index).Move EersteKaartLeft, EersteKaartTop - (KaartNr - 1) * KaartAfst
      Next KaartNr
      lblNaam(3).Move imgInHanden(39).Left + imgInHanden(39).Width - lblNaam(3).Width - 15, imgInHanden(39).Top + imgInHanden(39).Height
      lblNaamSchaduw(3).Move lblNaam(3).Left + 15, lblNaam(3).Top + 15
  End Select
End Sub
Sub BepaalRangen()
  Dim SpelerNr As Integer
  Dim SpelerNr2 As Integer
  Dim SortSpeler(1 To 4) As Integer
  Dim SpelerTemp As Integer
  
  For SpelerNr = 1 To 4
    SortSpeler(SpelerNr) = SpelerNr
  Next SpelerNr
  
  For SpelerNr = 1 To 3
    For SpelerNr2 = SpelerNr + 1 To 4
      If Spelers(SortSpeler(SpelerNr2)).TotaalScore > Spelers(SortSpeler(SpelerNr)).TotaalScore Then
        SpelerTemp = SortSpeler(SpelerNr)
        SortSpeler(SpelerNr) = SortSpeler(SpelerNr2)
        SortSpeler(SpelerNr2) = SpelerTemp
      End If
    Next SpelerNr2
  Next SpelerNr
  
  Spelers(SortSpeler(1)).Rang = 1
  For SpelerNr = 2 To 4
    If Spelers(SortSpeler(SpelerNr)).TotaalScore = Spelers(SortSpeler(SpelerNr - 1)).TotaalScore Then
      Spelers(SortSpeler(SpelerNr)).Rang = Spelers(SortSpeler(SpelerNr - 1)).Rang
    Else
      Spelers(SortSpeler(SpelerNr)).Rang = SpelerNr
    End If
  Next SpelerNr
 
End Sub

Sub SpelAfgelopen()
  Dim TotaalTemp As Long
  Dim SpelerNr As Integer
  Dim AantPCs As Integer
  Dim NieuwNaam As String
  
  KaartenWeg
  fraTroef.Visible = False
  StatusBar.PanelText(1) = ""

  AantPCs = 0
  TotaalTemp = 0
  For SpelerNr = 1 To 4
    If Spelers(SpelerNr).Controller = ControllerComputer Then
      AantPCs = AantPCs + 1
      TotaalTemp = TotaalTemp + Spelers(SpelerNr).TotaalScore
    End If
  Next SpelerNr

  'If nRonden = 19 Then
    If AantPCs > 0 Then
      Statistiek.PCGemiddelde = (Statistiek.PCGemiddelde * Statistiek.SpellenVoltooid + TotaalTemp) / (AantPCs + Statistiek.SpellenVoltooid)
    End If
    Statistiek.Gemiddelde = (Statistiek.Gemiddelde * Statistiek.SpellenVoltooid + Spelers(1).TotaalScore) / (Statistiek.SpellenVoltooid + 1)
  'End If

  BepaalRangen
  Statistiek.RangFreq(Spelers(1).Rang) = Statistiek.RangFreq(Spelers(1).Rang) + 1
  Statistiek.SpellenVoltooid = Statistiek.SpellenVoltooid + 1
  Score.TestHiScore
  Score.ToonScore
  Tineke.AfgelopenCommentaar
  'BerekenStatistiek
  If Spelers(1).HiScorePositie > 0 Then
    'Tineke.Zeg "highscore"
    If Spelers(1).Naam = StandaardNaam1 Then
      NieuwNaam = InputBox("Je hebt een plaats in de highscore behaald." & vbCrLf & "Typ hier je naam.", "Highscore", Spelers(1).Naam)
      If NieuwNaam <> "" Then
        Spelers(1).Naam = NieuwNaam
        lblNaam(0).Caption = Spelers(1).Naam
        lblNaamSchaduw(0).Caption = Spelers(1).Naam
        Score.UpdateSpelerNaam
      End If
    End If
  End If

  Score.ToonRanglijst
  If Score.IemandInHiScore Then
    Score.ToonHiScore
    Score.OpslaanHiScore
  End If
  
  mnuSpelLaatsteSlagTonen.Checked = False
  'frmMenus.mnuRondmakenToestaan.Enabled = True
  
  OpslaanStatistiek
  OpslaanOpties
  IniSave
End Sub

Sub Einde()
  Dim TotaalDuur As Integer
  
  WavPlay "Programma sluiten"
  If DateDiff("n", StartMoment, Now) > 0 Then
    Statistiek.TotaalDuur = Statistiek.TotaalDuur + Int(DateDiff("n", StartMoment, Now))
  End If
  OpslaanStatistiek
  OpslaanOpties
  IniSave
'  If Statistiek.Sessies = 1 Then
'    MsgBox "De volgende keer kun je dit spel weer starten " & _
'           "door te klikken op het menu Start, " & _
'           "dan Programma's, en dan '10 op en neer'.", vbInformation, "Informatie"
'  End If
  End
End Sub

Private Sub Winsock1_Close()
  'frmSpelSpeciaal.StatusBar1.SimpleText = "Verbinding verbroken."
  'frmSpelSpeciaal.cmdStart.Enabled = False
  'WavPlay "Verbinding verbroken" '"Ir_end.wav"
End Sub

Private Sub Winsock1_Connect()
  'IkBenNetSpelerNr = 3

  'frmSpelSpeciaal.StatusBar1.SimpleText = "Verbonden."
  'frmSpelSpeciaal.cmdVerbinden.Caption = "&Verbreken"
  WavPlay "Verbonden" '"Ir_begin.wav"

  txtNetwerkStatus.Visible = True
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
'  frmNetwerk.StatusBar1.SimpleText = "Oproep beantwoorden..."
'  frmMain.Winsock1.Close
'  frmMain.Winsock1.Accept requestID
'  frmNetwerk.StatusBar1.SimpleText = "Verbonden."
'
'  frmNetwerk.cmdStart.Enabled = True
'
'  WavPlay "Verbonden" '"Ir_begin.wav"
'
'  IkBenNetSpelerNr = 1
'
'  txtNetwerkStatus.Visible = True
'
'  frmMain.Winsock1.SendData "1Naam:" & Spelers(1).Naam
'  DoEvents
'
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
'  Dim Regel As String
'  Dim SpelerNr As Integer
'
'  Winsock1.GetData Regel
'
'  SpelerNr = CInt(Left(Regel, 1))
'  Regel = Mid(Regel, 2)
'
'  If SpelerNr = 1 Then
'    If Regel = "Start" Then
'      frmNetwerk.Hide
'      NieuwSpel
'    ElseIf Left(Regel, 8) = "EerstOpkomen:" Then
'      txtNetwerkStatus.Text = txtNetwerkStatus.Text & vbCrLf & "Server zegt wie eerst op moet komen"
'      EerstOpkomen = NetNaarLokaalSpeler(CInt(Mid(Regel, 9)))
'      MsgBox "1: Speler " & EerstOpkomen & " moet eerst opkomen"
'    ElseIf Left(Regel, 5) = "Naam:" Then
'      Spelers(NetNaarLokaalSpeler(1)).Naam = Mid(Regel, 6)
'      lblNaam(NetNaarLokaalSpeler(1) - 1).Caption = Spelers(NetNaarLokaalSpeler(1)).Naam
'      lblNaamSchaduw(NetNaarLokaalSpeler(1) - 1).Caption = Spelers(NetNaarLokaalSpeler(1)).Naam
'
'      frmMain.Winsock1.SendData CStr(IkBenNetSpelerNr) & "Naam:" & Spelers(1).Naam
'    Else
'      MsgBox "Hier kan ik niks mee: " & Regel
'    End If
'  Else
'    If Regel = "EerstOpkomen" Then
'      ClientVraagt(SpelerNr) = Regel
'    ElseIf Regel = "GeefKaarten" Then
'      ClientVraagt(SpelerNr) = Regel
'      txtNetwerkStatus.Text = txtNetwerkStatus.Text & vbCrLf & "(1) Client " & SpelerNr & "wil kaarten hebben"
'
'    ElseIf Left(Regel, 5) = "Naam:" Then
'      Spelers(SpelerNr).Naam = Mid(Regel, 6)
'      lblNaam(SpelerNr - 1).Caption = Spelers(SpelerNr).Naam
'      lblNaamSchaduw(SpelerNr - 1).Caption = Spelers(SpelerNr).Naam
'    Else
'      MsgBox "Hier kan ik niks mee: " & Regel
'    End If
'  End If
End Sub
'Sub BepaalEerstOpkomen()
'  Randomize Timer
'
'  'Select Case IkBenNetSpelerNr
'  '  Case 0
'  If recWieBegint < 1 Then
'    recWieBegint = Int(4 * Rnd + 1)
'  End If
'  EerstOpkomen = VorigeSpeler(recWieBegint) 'In NieuweRonde wordt meteen de volgende speler genomen
'  '  Case 1
'  '    EerstOpkomen = Int(4 * Rnd + 1)
'  '    Do Until ClientVraagt(1) = "EerstOpkomen" Or Spelstatus <> "Bezig"
'  '      DoEvents
'  '    Loop
'      'Winsock1.SendData "1Opkomen:" & CStr(EerstOpkomen)
'      'MsgBox "Hee Frank, speler " & EerstOpkomen & " moet eerst opkomen. Leuk he. Snel tegen Axel zeggen."
'  '  Case Is >= 2
'      'VraagData = "Opkomen"
'      'Winsock1.SendData CStr(IkBenNetSpelerNr) & "EerstOpkomen"
'  '    Do Until EerstOpkomen >= 1 Or Spelstatus <> "Bezig"
'  '      DoEvents
'  '    Loop
'  '    MsgBox "2: Speler " & EerstOpkomen & "moet eerst opkomen"
'  'End Select
'End Sub
Sub TestMagSpelSpelen()
'  If MagSpelSpelen Then 'Als het eerst niet mocht mag het nu ook niet (ivm geknoei met datum)
'    If DateDiff("d", Now, NietGebruikenNa) < 0 Then
'      MagSpelSpelen = False
'    Else
'      MagSpelSpelen = True
'    End If
'  End If
End Sub

Sub MeteenOptellenKlik()
  Opties.MeteenOptellen = Not Opties.MeteenOptellen
  mnuTotaal.Checked = Opties.MeteenOptellen
  frmMenus.mnuTotaalscore.Checked = Opties.MeteenOptellen
  'IniSet "Opties", "Optellen", CStr(Opties.MeteenOptellen)
  Score.Update
  'Score.ToonScore
End Sub

Sub NieuwSpel(ByVal IsHerhaling As Boolean)
  Dim SpelerNr As Integer
  Dim SlagNr As Integer
 
  ResetVars
  ResetControls
  
  If Not IsHerhaling Then
    Statistiek.SpellenBegonnen = Statistiek.SpellenBegonnen + 1
    recOpnameWissen
    DetermineerSpel
  End If
  
  'BepaalEerstOpkomen
  Score.ToonScore
  NieuweRonde
 
End Sub

Sub NieuweRonde()
  Ronde = Ronde + 1
  recRonde = Ronde
  ResetRondeVars
  KaartenWeg '# Kan beter?
  Delen
  'EerstOpkomen = (recWieBegint + Ronde + 2) Mod 4 + 1
  NuOpkomen = (recWieBegint + Ronde + 2) Mod 4 + 1 'EerstOpkomen
  NuOpleggen = NuOpkomen
  NuVoorspellen = NuOpkomen
  SpelerNum = 1
  
  KaartenResterend = AantKaartenRonde(Ronde)
  ToonOpkomen
  Score.ToonVoorspelling = True
  Score.ToonScore '# Eigenlijk dubbelop, maar nu staat Ronde goed en wordt de juiste ronde geel gemarkeerd
  Voorspellen
End Sub

Sub Voorspeld()
  'If Not recInHerhaling() Then
    recVoorspellingen(Ronde, NuVoorspellen) = Spelers(NuVoorspellen).Voorspelling
    'recSpelerNr = NuVoorspellen
  '  recSpelerNum = SpelerNum
    'recAantRonden = Ronde
  '  recSlagNr = 0
  'End If
  
  TotSlagenGok = TotSlagenGok + Spelers(NuVoorspellen).Voorspelling
  SlagenOver = SlagenOver - Spelers(NuVoorspellen).Voorspelling
  AantSpelersGegokt = AantSpelersGegokt + 1

  Score.ToonScore
  StatusBar.PanelText(1) = SlagenOverTekst()
  ToonRondjes NuVoorspellen
  If NuVoorspellen = 1 Then
    Tineke.ZegHulp "voorspeld"
  End If
  NuVoorspellen = VolgendeSpeler(NuVoorspellen)
  If NuVoorspellen = NuOpkomen Then
    SlagNr = 1
    SpelerNum = 1
    KaartKiezen
  Else
    SpelerNum = SpelerNum + 1
    Voorspellen
  End If
End Sub

Sub VoorspeldMens()
  Dim TotaalVoorspeldIni As Integer
  Dim Rondgemaakt As Boolean
  
  Praatwolkje.Visible = False
  picPraatwolkjePunt.Visible = False
  Tineke.ToonPijltje = False
  
  'frmMenus.mnuRondmakenToestaan.Enabled = False

  If Spelers(NuVoorspellen).Voorspelling = SlagenOver And NuOpkomen = 2 Then
    Rondgemaakt = True
  Else
    Rondgemaakt = False
  End If
  
  If Rondgemaakt And Not Opties.RondmakenToegestaan Then
    Tineke.Zeg "niet rondmaken"
  Else
    picScore.TabStop = True
    fraVoorspellen.Visible = False
    WachtOp = Niets
    
    'If nRonden = 19 Then
      Statistiek.TotaalVoorspeld = Statistiek.TotaalVoorspeld + Spelers(1).Voorspelling
      Statistiek.TotaalVoorspeldMax = Statistiek.TotaalVoorspeldMax + AantKaartenRonde(Ronde)
    'End If
    If NuOpkomen = 2 Then 'Als laatste voorspeld (NuVoorspellen=1)
      If Rondgemaakt Then
        Statistiek.Rondgemaakt = Statistiek.Rondgemaakt + 1
      Else
        Statistiek.NietRondgemaakt = Statistiek.NietRondgemaakt + 1
      End If
    End If
    
    If Spelers(1).Voorspelling < Spelers(1).Taxatie - 1.4 Then
      Tineke.Voorspelling = 1
      Tineke.Zeg "weinig voorspeld"
    ElseIf Spelers(1).Voorspelling > Spelers(1).Taxatie + 1.5 Then
      Tineke.Voorspelling = 2
      Tineke.Zeg "veel voorspeld"
    Else
      Tineke.Voorspelling = 0
    End If
    Voorspeld
  End If
End Sub

Function KaartKiezenComputer(SpelerNr As Integer) As Integer
  '"de computer" onderscheidt 27 situaties
  
  Dim NummerGekozen As Integer 'Het *nummer*, niet 'getal'
  Dim AantalSlagenNogNodig As Integer
  Dim OverschotInHanden As Single
  Dim NuSlagenOverschot As Integer '>0 -> iedereen wil duiken
  Dim IsNat As Boolean

  AantalSlagenNogNodig = Spelers(SpelerNr).Voorspelling - Spelers(SpelerNr).AantSlagen

  ZoekMogelijkeKaarten SpelerNr

  If AantalMogelijkeKaarten = 1 Then
    NummerGekozen = MogelijkeKaarten(1)
  Else 'Keus uit meerdere legale kaarten
    TestSlagTeHalen SpelerNr
    OverschotInHanden = TaxeerKaarten(SpelerNr) - AantalSlagenNogNodig
    NuSlagenOverschot = Spelers(SpelerNr).AantKaarten - TotaalSlagenNodig()
    If OverschotInHanden > 0 Then
      OverschotInHanden = OverschotInHanden * (1 + (NuSlagenOverschot / AantKaartenRonde(Ronde)))
    End If
    
    If AantalSlagenNogNodig < 0 Then
      IsNat = True 'Heeft er al te veel
    'ElseIf AantalSlagenNogNodig = 0 And SlagNietTeHalenMetAantal = 0 And NuOpkomen = VolgendeSpeler(NuOpleggen) Then
    ElseIf AantalSlagenNogNodig = 0 And SlagNietTeHalenMetAantal = 0 And NuOpkomen = VolgendeSpeler(SpelerNr) Then
      IsNat = True 'Krijgt er te veel
    ElseIf AantalSlagenNogNodig = Spelers(SpelerNr).AantKaarten And SlagIsTeHalenMetAantal = 0 Then
      IsNat = True 'Krijgt er te weinig
    ElseIf AantalSlagenNogNodig > Spelers(SpelerNr).AantKaarten Then
      IsNat = True 'Krijgt er te weinig
    Else
      IsNat = False
    End If
    
    If IsNat And Opties.StrafpuntenPerSlag = 0 Then
      'MsgBox Spelers(SpelerNr).Naam & " gaat dwarsliggen", vbInformation
      If NuSlagenOverVoorAnderen(SpelerNr) > 0 Then 'Slagen over: duiken
        If SlagNietTeHalenMetAantal > 0 Then 'Kan duiken
          NummerGekozen = GooiWeg(SpelerNr, -1)
        ElseIf SpelerNr = NuOpkomen Then
          NummerGekozen = NeemSlag(SpelerNr, 1, True)
        ElseIf VolgendeSpeler(SpelerNr) <> NuOpkomen Then
          NummerGekozen = NeemSlag(SpelerNr, 1, True)
        Else
          NummerGekozen = NeemSlag(SpelerNr, -1, True)
        End If
      Else 'Slagen te weinig of rond: alles nemen
        If SlagIsTeHalenMetAantal > 0 Then
          If SpelerNr = NuOpkomen Then
            NummerGekozen = NeemSlag(SpelerNr, -1, True)
          ElseIf VolgendeSpeler(SpelerNr) <> NuOpkomen Then
            NummerGekozen = NeemSlag(SpelerNr, -1, True)
          Else
            NummerGekozen = NeemSlag(SpelerNr, 1, True)
          End If
        Else
          NummerGekozen = GooiWeg(SpelerNr, 1)
        End If
      End If
    
    Else 'Is niet nat -> proberen voorspelling te halen
    
      If AantalSlagenNogNodig > 0 Then
        '** Wil meer slagen **
        If SlagIsTeHalenMetAantal > 0 Then
          '** De slag is te halen **
          If AantalSlagenNogNodig = Spelers(SpelerNr).AantKaarten Then
            '** Moet de rest
            If SpelerNr = NuOpkomen Then
              NummerGekozen = NeemSlag(SpelerNr, -1, True)
            ElseIf KaartenOpTafel(NuOpkomen).Kleur <> Troef.Kleur And Not KanKleurBekennen(SpelerNr) Then
              'Speler kan introeven
              If VolgendeSpelerMoetNuIntroeven(SpelerNr) Then
                'Er zijn meer spelers die alles willen: hoog introeven
                NummerGekozen = NeemSlag(SpelerNr, -1, True)
              Else
                'Rustig introeven: liefst met lage troef
                NummerGekozen = NeemSlag(SpelerNr, 1, True)
              End If
            Else 'Niet opkomen en moet bedienen
              If VolgendeSpeler(SpelerNr) = NuOpkomen Then
                'In de achterhand: met lage nemen
                NummerGekozen = NeemSlag(SpelerNr, 1, True)
              Else 'Niet in achterhand
                If VolgendeSpelerWilMeer(SpelerNr) Then
                  'Volgende spelers willen ook nog slagen; met hoogste nemen
                  NummerGekozen = NeemSlag(SpelerNr, -1, True)
                Else
                  'Volgende spelers hoeven niet meer; met laagste nemen
                  NummerGekozen = NeemSlag(SpelerNr, 1, True)
                End If
              End If
            End If
          ElseIf OverschotInHanden > MargeSNN Then
            '** Heeft zat hoge kaarten **
            If SpelerNr = NuOpkomen Then
              If OverschotInHanden > MargeOMLT Then
                NummerGekozen = KomOpMetLageTroef(SpelerNr, False)
              Else
                NummerGekozen = NeemSlag(SpelerNr, 1, False)
              End If
            ElseIf SlagNietTeHalenMetAantal > 0 Then
              '** Moet niet opkomen **
              '** Kan duiken **
              NummerGekozen = GooiWeg(SpelerNr, CInt(-AantalSlagenNogNodig - 1))
            Else
              '** Moet niet opkomen **
              '** Niet te duiken... **
              NummerGekozen = NeemSlag(SpelerNr, CInt(-AantalSlagenNogNodig), True)
            End If
          Else
            '** Wil meer slagen **
            '** Kan meer slagen halen **
            '** Heeft geen overschot **
            If Spelers(SpelerNr).AantSlagen + 1 = Spelers(SpelerNr).Voorspelling Then
              '** Hoeft er nog maar 1 **
              If SpelerNr = NuOpkomen Then
                NummerGekozen = NeemSlag(SpelerNr, -1, False)
              Else
                '** Moet niet opkomen **
                NummerGekozen = NeemSlag(SpelerNr, -1, True)
              End If
            Else
              '** Moet er nog meer dan 1 **
              If SpelerNr = NuOpkomen Then
                NummerGekozen = NeemSlag(SpelerNr, -1, False)
              Else
                NummerGekozen = NeemSlag(SpelerNr, CInt(-AantalSlagenNogNodig), True)
              End If
            End If '"hoeft er maar 1?"
          End If '"heeft overschot?"
        Else
          '** Slag is niet te halen **
          If OverschotInHanden > MargeHWG Then
            '** Heeft zat hoge kaarten **
            NummerGekozen = GooiWeg(SpelerNr, CInt(-AantalSlagenNogNodig - 1))
          Else
            If Spelers(SpelerNr).AantKaarten > 4 Then
              NummerGekozen = GooiWeg(SpelerNr, 2)
            Else 'Weinig kaarten over dus toch maar laagste weggooien
              NummerGekozen = GooiWeg(SpelerNr, 1)
            End If
          End If
        End If
      'ElseIf AantalSlagenNogNodig < 0 Then
      '** IsNat -> zie boven **
      Else
        '** Hoeft geen slagen meer **
        If SpelerNr = NuOpkomen Then
          NummerGekozen = KomOpMetLageTroef(SpelerNr, True)
        Else
          If SlagNietTeHalenMetAantal = 0 Then
            '** Kan niet duiken **
            
            '** Hier volgt onzin **
            '** (deze If hieronder kan niet voorkomen: valt onder IsNat) **
'            If SpelerNr = VorigeSpeler(NuOpkomen) Then
'              '** Zit in de achterhand
'              If Opties.StrafpuntenPerSlag = 0 Then
'                '** Meer slagen halen kost geen extra punten: dwarsliggen
'                NummerGekozen = NeemSlag(SpelerNr, 1, True)
'              Else
'                '** Meer slagen halen kost extra punten: schade beperken
'                NummerGekozen = NeemSlag(SpelerNr, -1, True)
'              End If
'            Else
'              '** Zit niet in de achterhand
'            If Opties.StrafpuntenPerSlag = 0 Then
'              '** Meer slagen halen kost geen extra punten: dwarsliggen
'              NummerGekozen = NeemSlag(SpelerNr, -1, True)
'            Else
'              '** Meer slagen halen kost extra punten: schade beperken

              '** Hopen dat de volgende speler nog hoger heeft **
            NummerGekozen = NeemSlag(SpelerNr, 1, True)
'            End If
'            End If
          Else
            '** Kan duiken **
            NummerGekozen = GooiWeg(SpelerNr, -1)
          End If
        End If
      End If
    End If 'nl.: If IsNat .. Else
  End If 'nl.: If AantalMogelijkeKaarten = 1

  If Not Spelers(SpelerNr).Kaarten(NummerGekozen).Legaal Then
    MsgBox "Speler " & SpelerNr & " speelt vals!", vbCritical, "Fatale fout"
    Stop
  End If
  
  KaartKiezenComputer = NummerGekozen

End Function

Sub KaartGekozenMens(KaartNr As Integer)
  If WachtOpGezien Then
    StatusBar.SimpleText = "Lees eerst de instructies van Tineke en klik dan op 'Doorgaan'."
  Else
    If WachtOp = MensKaart Then
      Praatwolkje.Visible = False
      picPraatwolkjePunt.Visible = False
      If Spelers(1).Kaarten(KaartNr).Legaal = False Then
        If KanKleurBekennen(1) And Spelers(1).Kaarten(KaartNr).Kleur <> KaartenOpTafel(NuOpkomen).Kleur Then
          StatusBar.SimpleText = "Je moet kleur bekennen. Kies een " & KleurNaam(KaartenOpTafel(NuOpkomen).Kleur) & "."
          Beep
          Tineke.Zeg "kleur bekennen"
        ElseIf Opties.IntroevenVerplicht And Spelers(1).Kaarten(KaartNr).Kleur <> Troef.Kleur Then
          StatusBar.SimpleText = "Je moet introeven. Kies een " & KleurNaam(Troef.Kleur) & "."
          Beep
          Tineke.Zeg "verplicht introeven"
        End If
      Else
        WachtOp = Niets
        'If recKaartenGekozen(Ronde, 1, SlagNr) = 0 Then
        If Not recInHerhaling() Then
          recKaartenGekozen(Ronde, 1, SlagNr) = KaartNr
'          Debug.Print "r" & Ronde & " no" & 1 & " sl" & SlagNr & " ";
'          Debug.Print KaartNaam(Spelers(NuOpleggen).Kaarten(recKaartenGekozen(Ronde, 1, SlagNr)))
          'recSpelerNr = 1
          recSpelerNum = SpelerNum
          recSlagNr = SlagNr
          'recVoorspellenKlaar = True
        End If
        KaartLeggen KaartNr
      End If
    'ElseIf WachtOp = MensVoorspelling Then
    '  StatusBar.simpletext = "Je moet eerst het aantal slagen voorspellen. Klik op een knop."
    'Else
    '  StatusBar.simpletext = "Je bent nog niet aan de beurt, je zult wat meer geduld moeten hebben."
    End If
  End If
End Sub

Sub StartNaamWijzigen(SpelerNr As Integer)
  If lblNaam(SpelerNr - 1).Alignment = vbLeftJustify Then
    txtNaamWijzig.Move lblNaam(SpelerNr - 1).Left - 45, lblNaam(SpelerNr - 1).Top - 45, lblNaam(SpelerNr - 1).Width, lblNaam(SpelerNr - 1).Height
  Else
    txtNaamWijzig.Move lblNaam(SpelerNr - 1).Left + 60, lblNaam(SpelerNr - 1).Top - 45, lblNaam(SpelerNr - 1).Width, lblNaam(SpelerNr - 1).Height
  End If
  txtNaamWijzig.Text = Spelers(SpelerNr).Naam
  txtNaamWijzig.Alignment = lblNaam(SpelerNr - 1).Alignment
  txtNaamWijzig.SelStart = 0
  txtNaamWijzig.SelLength = Len(txtNaamWijzig.Text)
  txtNaamWijzig.Visible = True
  txtNaamWijzig.SetFocus
  
  NaamWijzigen = SpelerNr
End Sub

Sub FormResize()
  Dim Breedte As Integer
  Dim Hoogte As Integer
  Dim picScoreScaleHeight As Integer
  Dim KaartWidth As Integer
  Dim KaartHeight As Integer
  Dim TafelWidth As Integer
  Dim TafelHeight As Integer
  
  timResize.Enabled = False
  
  If frmMain.ScaleWidth >= 8200 Then
    Breedte = frmMain.ScaleWidth
  Else
    Breedte = 8200
  End If
  
  If frmMain.ScaleHeight - StatusBar.Height >= 5895 Then
    Hoogte = frmMain.ScaleHeight - StatusBar.Height
  Else
    Hoogte = 5895
  End If
  
'  picStatusBar.Top = Hoogte '- picStatusBar.Height
'  picStatusBar.Width = Breedte
'  picSlagenOver.Left = picStatusBar.ScaleWidth - picSlagenOver.Width - 20
'  picRondmaken.Left = picSlagenOver.Left - picRondmaken.Width - 3
'  picIntroeven.Left = picRondmaken.Left - picIntroeven.Width - 3
'  lblStatus.Width = picIntroeven.Left - 3
  
  StatusBar.Top = Hoogte '- picStatusBar.Height
  StatusBar.Width = Breedte
  
  picScore.Left = Breedte - picScore.Width - 120
  picScore.Height = Hoogte - fraTroef.Height - 240
  fraTroef.Left = picScore.Left
  fraTroef.Top = Hoogte - fraTroef.Height - 120
  
  imgVrouw.Left = Breedte - imgVrouw.Width - 120
  imgVrouw.Top = Hoogte - imgVrouw.Height - 120
  'cmdHelp.Left = Breedte - cmdHelp.Width - 120
  'cmdHelp.Top = imgVrouw.Top - cmdHelp.Height - 180
  picPraatwolkjePunt.Move Breedte - 600, imgVrouw.Top - 180
  Praatwolkje.Move Breedte - Praatwolkje.Width - 120, frmMain.picPraatwolkjePunt.Top - Praatwolkje.Height + 30 '+ 30, anders past het niet (?)

  picScoreScaleHeight = picScore.ScaleHeight
  linScore2.Y2 = picScoreScaleHeight - 1
  linScore3.Y1 = picScoreScaleHeight - 1
  linScore3.Y2 = picScoreScaleHeight - 1
  linScore4.Y2 = picScoreScaleHeight - 1
  
  fraSpelen.Move 0, 0, fraTroef.Left, Hoogte
  fraEinde.Move 0, 0, fraTroef.Left, Hoogte
  
  imgGewonnen.Move 480, Hoogte - imgGewonnen.Height - 120
  cmdAfsluiten.Move fraTroef.Left - cmdAfsluiten.Width - 480, Hoogte - cmdAfsluiten.Height - 120
  cmdNieuwSpel.Move cmdAfsluiten.Left, cmdAfsluiten.Top - cmdNieuwSpel.Height - 120
  linPodium.X2 = fraEinde.Width - 480
  lblRanglijstScore(0).Move fraEinde.Width - lblRanglijstScore(0).Width - 480
  lblRanglijstScore(1).Move fraEinde.Width - lblRanglijstScore(1).Width - 480
  lblRanglijstScore(2).Move fraEinde.Width - lblRanglijstScore(2).Width - 480
  lblRanglijstScore(3).Move fraEinde.Width - lblRanglijstScore(3).Width - 480
  
  KaartWidth = imgOpTafel(0).Width
  KaartHeight = imgOpTafel(0).Height
  TafelWidth = fraSpelen.Width
  TafelHeight = fraSpelen.Height
  imgOpTafel(0).Move (TafelWidth - KaartWidth) / 2, TafelHeight / 2 - 290
  imgOpTafel(1).Move TafelWidth / 2 - KaartWidth - 60, (TafelHeight - KaartHeight) / 2
  imgOpTafel(2).Move (TafelWidth - KaartWidth) / 2, TafelHeight / 2 - KaartHeight + 290
  imgOpTafel(3).Move TafelWidth / 2 + 60, (TafelHeight - KaartHeight) / 2
  
  If fraSpelen.Visible Then
    SchikKaarten 1
    SchikKaarten 2
    SchikKaarten 3
    SchikKaarten 4
    ToonRondjes 1
    ToonRondjes 2
    ToonRondjes 3
    ToonRondjes 4
  End If
  
  imgKaartAanwijs.Top = imgInHanden(0).Top - imgKaartAanwijs.Height
  fraVoorspellen.Move (fraSpelen.Width - fraVoorspellen.Width) / 2, (fraSpelen.Height - fraVoorspellen.Height) / 2
  'If Scoreblok = sblScore Then
  '  ToonScore
  'End If
  Score.Update
End Sub

Private Sub Winsock_Close(Index As Integer)
  'nVerbonden = nVerbonden - 1
  'If frmSpelSpeciaal.Visible Then
    'If nVerbonden = 0 Then
    '  frmSpelSpeciaal.StatusBar1.SimpleText = "Niet verbonden."
    'Else
    '  frmSpelSpeciaal.StatusBar1.SimpleText = nVerbonden & " speler" & IIf(nVerbonden = 1, "", "s") & " verbonden."
    'End If
  '  frmSpelSpeciaal.UpdateStatusbar
  'End If
End Sub

Private Sub Winsock_Connect(Index As Integer)
  Netwerk.Connect Index
End Sub

Private Sub Winsock_ConnectionRequest(Index As Integer, ByVal requestID As Long)
  'Moet niet voorkomen
  
  'Winsock(Index).Accept requestID
  'nVerbonden = nVerbonden + 1
  'If frmSpelSpeciaal.Visible Then
  '  frmSpelSpeciaal.StatusBar1.SimpleText = nVerbonden & " speler" & IIf(nVerbonden = 1, "", "s") & " verbonden."
  'End If
End Sub

Private Sub Winsock_DataArrival(Index As Integer, ByVal bytesTotal As Long)
  Netwerk.DataArrival Index, bytesTotal
End Sub

Private Sub wskListen_ConnectionRequest(ByVal requestID As Long)
  Netwerk.ConnectionRequest requestID
End Sub

Private Function LegaleVoorspelling(ByVal Taxatie As Single)
  Dim GewensteVoorspelling As Integer
  
  GewensteVoorspelling = CInt(Taxatie)
  If Not Opties.RondmakenToegestaan And GewensteVoorspelling = SlagenOver And VolgendeSpeler(NuVoorspellen) = NuOpkomen Then
    If Taxatie > GewensteVoorspelling Then
      If GewensteVoorspelling + 1 <= AantKaartenRonde(Ronde) Then
        LegaleVoorspelling = GewensteVoorspelling + 1
      Else
        LegaleVoorspelling = GewensteVoorspelling - 1
      End If
    Else
      If GewensteVoorspelling - 1 >= 0 Then
        LegaleVoorspelling = GewensteVoorspelling - 1
      Else
        LegaleVoorspelling = GewensteVoorspelling + 1
      End If
    End If
  Else
    LegaleVoorspelling = GewensteVoorspelling
  End If
End Function

Private Function VolgendeKaartVanKleur(SpelerNr As Integer, KaartNr As Integer) As Integer
  Dim i As Integer 'Index die kaarten langs gaat
  Dim Gevonden As Boolean
  
  i = KaartNr + 1
  Do Until i > Opties.MaxAantKaarten Or Gevonden
    With Spelers(SpelerNr)
      If .Kaarten(i).Getal >= 2 Then 'Het is een kaart
        If .Kaarten(i).Kleur = .Kaarten(KaartNr).Kleur Then
          VolgendeKaartVanKleur = i
        Else
          VolgendeKaartVanKleur = -1
        End If
        Gevonden = True
      End If
      i = i + 1
    End With
  Loop
End Function

Private Function TotaalSlagenNodig() As Integer
  'Hoeveel slagen willen de anderen samen
  Dim splr As Integer
  Dim Som As Integer
  
  Som = 0
  For splr = 1 To 4
    With Spelers(splr)
      If .AantSlagen < .Voorspelling Then
        Som = Som + (.Voorspelling - .AantSlagen)
      End If
    End With
  Next splr
  
  TotaalSlagenNodig = Som
End Function

'Public Sub ToonRondmakenToegestaan()
'  If Opties.RondmakenToegestaan Then
'    lblRondmaken.Caption = "R"
'    lblRondmaken.ToolTipText = "Rondmaken is toegestaan"
'  Else
'    lblRondmaken.Caption = "NR"
'    lblRondmaken.ToolTipText = "Rondmaken is niet toegestaan"
'  End If
'  frmMenus.mnuRondmakenToestaan.Checked = Opties.RondmakenToegestaan
'End Sub
'
'Public Sub ToonIntroevenVerplicht()
'  If Opties.IntroevenVerplicht Then
'    lblIntroeven.Caption = "VI"
'    lblIntroeven.ToolTipText = "Introeven is verplicht"
'  Else
'    lblIntroeven.Caption = "I"
'    lblIntroeven.ToolTipText = "Introeven is niet verplicht"
'  End If
'  frmMenus.mnuIntroevenVerplicht.Checked = Opties.IntroevenVerplicht
'End Sub

Public Sub ToonSpelType()
  Dim Regel As String
  With Opties
    If SpelIs10openneer() Then
    'If .MaxAantKaarten = 10 And .FoutVoorspeldNulPunten And .StrafpuntenPerSlag = 0 And .RondmakenToegestaan And Not .IntroevenVerplicht And .PuntenPerSlag = 1 Then
      StatusBar.PanelText(0) = "10 op en neer"
      StatusBar.PanelToolTip(0) = "Je speelt nu 10 op en neer"
    'ElseIf .MaxAantKaarten = 13 And .StrafpuntenPerSlag > 0 And Not .RondmakenToegestaan And Not .IntroevenVerplicht Then
    ElseIf SpelIsBoerenbridge() Then
      StatusBar.PanelText(0) = "Boerenbridge"
      StatusBar.PanelToolTip(0) = "Je speelt nu Boerenbridge"
    Else
      StatusBar.PanelText(0) = "Hybride (" & IIf(.RondmakenToegestaan, "r", "nr") & IIf(.IntroevenVerplicht, ", iv", "") & ")"
      If .RondmakenToegestaan Then
        Regel = Regel & "Rondmaken is toegestaan"
      Else
        Regel = Regel & "Rondmaken is niet toegestaan"
      End If
      If .IntroevenVerplicht Then
        Regel = Regel & " en introeven is verplicht"
      End If
      StatusBar.PanelToolTip(0) = Regel
    End If
  End With
End Sub

Public Sub DebugWisseltruc() '# Om te testen
  Dim StapelWill As Integer
  Dim HandWill As Integer
  Dim TempKaart As Kaart
  
  If DeStapel.AantKaarten >= 1 And Spelers(1).AantKaarten >= 1 Then
    
    StapelWill = Int(DeStapel.AantKaarten * Rnd) + 1
    Do
      HandWill = Int(Spelers(1).AantKaarten * Rnd) + 1
    Loop Until Spelers(1).Kaarten(HandWill).Getal <> 0
  '-3 = Weg, -2 = Tafel, -1 = Troef, 0 = stapel, 1 = Speler 1, ...
    
    TempKaart = DeStapel.Kaarten(StapelWill)
    DeStapel.Kaarten(StapelWill) = Spelers(1).Kaarten(HandWill)
    Spelers(1).Kaarten(HandWill) = TempKaart
    
    WaarIsKaart(Spelers(1).Kaarten(HandWill).Kleur, Spelers(1).Kaarten(HandWill).Getal) = 1
    WaarIsKaart(DeStapel.Kaarten(StapelWill).Kleur, DeStapel.Kaarten(StapelWill).Getal) = 0
    
    LaadKaart frmMain.imgInHanden(HandWill - 1), Spelers(1).Kaarten(HandWill)
    Tip = -1
    Score.MagInHiScore = False
  End If
End Sub

Public Sub ScoreblokMenuKlik()
  If mnuScoreblok.Caption = "Toon &highscore" Then
    Score.ToonHiScore
    'mnuScoreblok.Caption = "Toon score&blok"
    'frmMenus.mnuToonScoreblok.Caption = "Toon score&blok"
    'picScore.ToolTipText = "Highscore"
  Else
    Score.ToonScore
'    mnuScoreblok.Caption = "Toon &highscore"
'    frmMenus.mnuToonScoreblok.Caption = "Toon &highscore"
    'picScore.ToolTipText = "Scoreblok"
  End If
End Sub

Private Sub txtNaamWijzig_LostFocus()
  txtNaamWijzig_KeyPress vbKeyReturn
End Sub

Public Sub MenuAfdrukken()
  Dim StatusTekstOud As String
  
  StatusTekstOud = StatusBar.SimpleText
  StatusBar.SimpleText = "Bezig met afdrukken..."
  DoEvents
  Score.Afdrukken
  StatusBar.SimpleText = StatusTekstOud
End Sub

Public Sub MenuHighscoreWissen()
  Dim Ret As VbMsgBoxResult
  
  Ret = MsgBox("Weet je zeker dat je de highscore wilt wissen?", vbYesNo + vbQuestion, "Highscore wissen")
  If Ret = vbYes Then
    Score.HighscoreWissen
    Score.ToonHiScore
  End If
End Sub

Public Sub KaartgrootteInstellen()
  Dim Index As Integer
  Dim Factor As Integer
  
  If Opties.GroteKaarten Then
    Factor = 2
  Else
    Factor = 1
  End If
  
  For Index = 0 To 51
    If KaartImageGeladen(Index) Then
      imgInHanden(Index).Width = Factor * KaartBreedte
      imgInHanden(Index).Height = Factor * KaartHoogte
    End If
  Next Index
  For Index = 0 To 3
    imgOpTafel(Index).Width = Factor * KaartBreedte
    imgOpTafel(Index).Height = Factor * KaartHoogte
  Next Index
'  picTroef.Width = Factor * KaartBreedte
'  picTroef.Height = Factor * KaartHoogte 'Past niet
  
  FormResize
End Sub

Private Sub ToonLaatsteSlag()
  Dim sp As Integer
  
  timLaatsteSlag.Enabled = False
  timLaatsteSlag.Enabled = True
  
  sp = LaatsteSlagOpgekomenSpeler
  Do
    If LaatsteSlag(sp).Legaal Then
      LaadKaart imgOpTafel(sp - 1), LaatsteSlag(sp)
      imgOpTafel(sp - 1).Visible = True
      imgOpTafel(sp - 1).ZOrder 0
    Else
      imgOpTafel(sp - 1).Visible = False
    End If
    sp = VolgendeSpeler(sp)
  Loop Until sp = LaatsteSlagOpgekomenSpeler
  
  fraVoorspellen.Visible = False
  mnuSpelLaatsteSlagTonen.Checked = True
End Sub

Private Sub ToonHuidigeSlag()
  Dim sp As Integer
  
  timLaatsteSlag.Enabled = False
  sp = NuOpkomen
  Do
    If KaartenOpTafel(sp).Legaal Then
      LaadKaart imgOpTafel(sp - 1), KaartenOpTafel(sp)
      imgOpTafel(sp - 1).Visible = True
      imgOpTafel(sp - 1).ZOrder 0
    Else
      imgOpTafel(sp - 1).Visible = False
    End If
    sp = VolgendeSpeler(sp)
  Loop Until sp = NuOpkomen

  If WachtOp = MensVoorspelling Then
    fraVoorspellen.Visible = True
  End If

  mnuSpelLaatsteSlagTonen.Checked = False
End Sub

Public Function KanIntroeven(ByVal SpelerNr As Integer) As Boolean
  '# Dit kan samen met KanKleurBekennen
  Dim KaartNr As Integer
 
  KanIntroeven = False
  If SpelerNr <> NuOpkomen Then
    For KaartNr = 1 To AantKaartenRonde(Ronde)
      If Spelers(SpelerNr).Kaarten(KaartNr).Kleur = Troef.Kleur Then
        KanIntroeven = True
      End If
    Next KaartNr
  End If

End Function

Public Function VolgendeSpelerMoetNuIntroeven(ByVal SpelerNr As Integer) As Boolean
  'Als een van de volgende spelers de rest van de slagen moet hebben,
  'moet die nu introeven. Als die speler ook troeven zou kunnen hebben,
  'levert deze functie True op.
  
  Dim sp As Integer
  
  VolgendeSpelerMoetNuIntroeven = False
  If Troef.Kleur > 0 Then
    sp = VolgendeSpeler(SpelerNr)
    Do Until sp = NuOpkomen
      If Spelers(sp).AantSlagen = Spelers(sp).Voorspelling - Spelers(sp).AantKaarten Then
        'Deze speler moet nu alles zien te halen
        If Not Spelers(sp).HeeftKleurNietMeer(Troef.Kleur) Then
          VolgendeSpelerMoetNuIntroeven = True
          Exit Do
        End If
      End If
      sp = VolgendeSpeler(sp)
    Loop
  End If
End Function

Public Function VolgendeSpelerWilMeer(ByVal SpelerNr As Integer) As Boolean
  'Als een van de volgende spelers nog slagen wil en zou kunnen halen,
  'levert deze functie True op.
  
  Dim sp As Integer
  
  VolgendeSpelerWilMeer = False
  'If Troef.Kleur > 0 Then
  sp = VolgendeSpeler(SpelerNr)
  Do Until sp = NuOpkomen
    If Spelers(sp).AantSlagen < Spelers(sp).Voorspelling Then
      'Deze speler wil meer slagen
      If Spelers(sp).AantSlagen >= Spelers(sp).Voorspelling - Spelers(sp).AantKaarten Then
        'Er zijn nog genoeg slagen in het spel
        If Not Spelers(sp).HeeftKleurNietMeer(KaartenOpTafel(NuOpkomen).Kleur) Then
          'De speler zou een hogere kaart kunnen hebben
          'Hierbij wordt niet geteld welke kaarten er nog zijn
          VolgendeSpelerWilMeer = True
          Exit Do
        ElseIf Troef.Kleur <> 0 Then
          If Not Spelers(sp).HeeftKleurNietMeer(Troef.Kleur) Then
            VolgendeSpelerWilMeer = True
            Exit Do
          End If
        End If
      End If
    End If
    sp = VolgendeSpeler(sp)
  Loop
End Function

