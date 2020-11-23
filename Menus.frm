VERSION 5.00
Begin VB.Form frmMenus 
   Caption         =   "Menu's"
   ClientHeight    =   1695
   ClientLeft      =   165
   ClientTop       =   1155
   ClientWidth     =   7095
   LinkTopic       =   "Form1"
   ScaleHeight     =   1695
   ScaleWidth      =   7095
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuTineke 
      Caption         =   "Tineke"
      Begin VB.Menu mnuSpelhulp 
         Caption         =   "&Begeleidende hulp"
         Shortcut        =   +{F1}
      End
      Begin VB.Menu mnuCommentaar 
         Caption         =   "&Commentaar aan"
      End
      Begin VB.Menu mnuHerhalen 
         Caption         =   "&Herhaal commentaar"
      End
      Begin VB.Menu mnuStreep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTinekeWatIsDit 
         Caption         =   "&Wat is dit?"
      End
   End
   Begin VB.Menu mnuScoreblok 
      Caption         =   "Scoreblok"
      Begin VB.Menu mnuScoreToonHighscore 
         Caption         =   "Toon &highscore"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuStreep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTotaalscore 
         Caption         =   "Toon &tussenstand"
      End
      Begin VB.Menu mnuStreep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuScoreWatIsDit 
         Caption         =   "&Wat is dit?"
         HelpContextID   =   101
      End
   End
   Begin VB.Menu mnuHighscore 
      Caption         =   "Highscore"
      Begin VB.Menu mnuHighscoreToonScore 
         Caption         =   "Toon &scoreblok"
      End
      Begin VB.Menu mnuStreep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHighscoreAfdrukken 
         Caption         =   "Highscore &afdrukken"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuHighscoreWissen 
         Caption         =   "Highscore &wissen..."
      End
      Begin VB.Menu mnuStreep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHighscoreComputersInHighscore 
         Caption         =   "&Computers in highscore aan"
      End
      Begin VB.Menu mnuStreep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHighscoreWatIsDit 
         Caption         =   "Wat is dit?"
      End
   End
   Begin VB.Menu mnuKaarten 
      Caption         =   "Kaarten"
      Begin VB.Menu mnuOmgekeerdSorteren 
         Caption         =   "&Omgekeerd sorteren"
      End
      Begin VB.Menu mnuBreedUitspreiden 
         Caption         =   "&Breed uitspreiden"
      End
      Begin VB.Menu mnuKaartenGroteKaarten 
         Caption         =   "&Grote kaarten"
      End
      Begin VB.Menu mnuAchterkant 
         Caption         =   "Kaart&achterkant..."
      End
   End
   Begin VB.Menu mnuSpelerNaam 
      Caption         =   "Spelernaam"
      Begin VB.Menu mnuNaamWijzigen 
         Caption         =   "&Naam wijzigen"
      End
   End
   Begin VB.Menu mnuVoorspellen 
      Caption         =   "Voorspellen"
      Begin VB.Menu mnuVoorspellenWatIsDit 
         Caption         =   "&Wat is dit?"
      End
   End
   Begin VB.Menu mnuTroef 
      Caption         =   "Troef"
      Begin VB.Menu mnuTroefWatIsDit 
         Caption         =   "Wat is dit?"
      End
   End
   Begin VB.Menu mnuRondjes 
      Caption         =   "Rondjes"
      Begin VB.Menu mnuRondjesWatIsDit 
         Caption         =   "Wat is dit?"
      End
   End
   Begin VB.Menu mnuSpeltype 
      Caption         =   "Speltype"
      Begin VB.Menu mnuSpeltypeWatIsDit 
         Caption         =   "Wat is dit?"
      End
   End
   Begin VB.Menu mnuSlagenOver 
      Caption         =   "Slagen over"
      Begin VB.Menu mnuSlagenOverWatIsDit 
         Caption         =   "Wat is dit?"
      End
   End
End
Attribute VB_Name = "frmMenus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub mnuAchterkant_Click()
  frmAchterkant.Show 1
End Sub

Private Sub mnuBreedUitspreiden_Click()
  Opties.BreedUitspreiden = Not Opties.BreedUitspreiden
  mnuBreedUitspreiden.Checked = Opties.BreedUitspreiden
  frmMenus.mnuBreedUitspreiden.Checked = Opties.BreedUitspreiden
  'IniSet "Opties", "BreedUitspreiden", Opties.BreedUitspreiden
  frmMain.SchikKaarten 1
End Sub

Private Sub mnuCommentaar_Click()
  Opties.Commentaar = Not Opties.Commentaar
  mnuCommentaar.Checked = Opties.Commentaar
  frmMain.mnuCommentaar.Checked = Opties.Commentaar
  'IniSet "Opties", "Commentaar", Opties.Commentaar
End Sub

Private Sub mnuHerhalen_Click()
  If Ronde = 0 Then
    Tineke.ToonPraatwolkje False
  Else
    Tineke.ToonPraatwolkje False
  End If
End Sub

Private Sub mnuHighscoreAfdrukken_Click()
  frmMain.MenuAfdrukken
End Sub

Private Sub mnuHighscoreComputersInHighscore_Click()
  Opties.ComputersInHighScore = Not Opties.ComputersInHighScore
  frmMain.mnuOptiesComputersInHighscore.Checked = Opties.ComputersInHighScore
  mnuHighscoreComputersInHighscore.Checked = Opties.ComputersInHighScore
End Sub

Private Sub mnuHighscoreToonScore_Click()
  'frmMain.ScoreblokMenuKlik
  Score.ToonScore
End Sub

Private Sub mnuHighscoreWatIsDit_Click()
  frmMain.picScore.ShowWhatsThis
End Sub

Private Sub mnuHighscoreWissen_Click()
  frmMain.MenuHighscoreWissen
End Sub

'Private Sub mnuIntroevenVerplicht_Click()
'  Opties.IntroevenVerplicht = Not Opties.IntroevenVerplicht
'  frmMain.ToonIntroevenVerplicht
'End Sub

Private Sub mnuKaartenGroteKaarten_Click()
  Opties.GroteKaarten = Not Opties.GroteKaarten
  frmMain.mnuOptiesGroot.Checked = Opties.GroteKaarten
  mnuKaartenGroteKaarten.Checked = Opties.GroteKaarten
  frmMain.KaartgrootteInstellen
End Sub

Private Sub mnuNaamWijzigen_Click()
  frmMain.StartNaamWijzigen PopupMenuIndex
End Sub

Private Sub mnuOmgekeerdSorteren_Click()
  Opties.AflopendSorteren = Not Opties.AflopendSorteren
  frmMain.mnuOmgekeerdSorteren.Checked = Opties.AflopendSorteren
  mnuOmgekeerdSorteren.Checked = Opties.AflopendSorteren
  'IniSet "Opties", "Aflopend", Opties.AflopendSorteren
  frmMain.Sorteren 1
End Sub

Private Sub mnuRondjesWatIsDit_Click()
  frmMain.imgSlagen(0).ShowWhatsThis
End Sub

Private Sub mnuScoreToonHighscore_Click()
  Score.ToonHiScore
End Sub

'Private Sub mnuRondmakenToestaan_Click()
'  Opties.RondmakenToegestaan = Not Opties.RondmakenToegestaan
'  frmMain.ToonRondmakenToegestaan
'End Sub

Private Sub mnuScoreWatIsDit_Click()
  frmMain.picScore.ShowWhatsThis
End Sub

Private Sub mnuSlagenOverWatIsDit_Click()
  frmMain.StatusBar.PanelShowWhatsThis 1
End Sub

Private Sub mnuSpelhulp_Click()
  Opties.Spelhulp = Not Opties.Spelhulp
  frmMain.SpelhulpAanUit
End Sub

Private Sub mnuSpeltypeWatIsDit_Click()
  frmMain.StatusBar.PanelShowWhatsThis 0
End Sub

Private Sub mnuTinekeWatIsDit_Click()
  frmMain.imgVrouw.ShowWhatsThis
End Sub

Private Sub mnuTotaalscore_Click()
  frmMain.MeteenOptellenKlik
End Sub

Private Sub mnuTroefWatIsDit_Click()
  frmMain.picTroef.ShowWhatsThis
End Sub

Private Sub mnuVoorspellenWatIsDit_Click()
  frmMain.fraVoorspellen.ShowWhatsThis
End Sub
