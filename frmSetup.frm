VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSetup 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Configuration de l'horloge AnalogoDigitale"
   ClientHeight    =   4110
   ClientLeft      =   7680
   ClientTop       =   4665
   ClientWidth     =   4005
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   4005
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Affichage digital"
      Height          =   795
      Left            =   285
      TabIndex        =   12
      Top             =   1485
      Width           =   3465
      Begin VB.CheckBox Check3 
         Alignment       =   1  'Right Justify
         Height          =   270
         Left            =   2700
         TabIndex        =   14
         Top             =   345
         Value           =   1  'Checked
         Width           =   330
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   345
         Value           =   1  'Checked
         Width           =   495
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Chiffres encadrés"
         Height          =   195
         Left            =   1365
         TabIndex        =   15
         Top             =   360
         Width           =   1230
      End
   End
   Begin VB.CommandButton btnAppliquer 
      Caption         =   "Appliquer"
      Height          =   375
      Left            =   2310
      TabIndex        =   11
      Top             =   3540
      Width           =   1455
   End
   Begin VB.CheckBox Check2 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   2190
      TabIndex        =   9
      Top             =   2460
      Value           =   1  'Checked
      Width           =   495
   End
   Begin VB.CommandButton btnCouleur 
      Height          =   375
      Left            =   3000
      Picture         =   "frmSetup.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2820
      Width           =   735
   End
   Begin VB.PictureBox couleur 
      Height          =   375
      Left            =   2430
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   7
      Top             =   2820
      Width           =   375
   End
   Begin MSComDlg.CommonDialog cmd1 
      Left            =   3570
      Top             =   15
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ComboBox cbVitesse 
      Height          =   315
      Left            =   2430
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   1080
      Width           =   735
   End
   Begin VB.ComboBox cbDiamètre 
      Height          =   315
      Left            =   2430
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   600
      Width           =   735
   End
   Begin VB.CommandButton btnAnnuler 
      Caption         =   "Annuler"
      Height          =   375
      Left            =   390
      TabIndex        =   1
      Top             =   3540
      Width           =   975
   End
   Begin VB.CommandButton IDP_ABOUT 
      Caption         =   "À propos"
      Height          =   315
      Left            =   1335
      TabIndex        =   0
      Top             =   135
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "Anneau lumineux"
      Height          =   390
      Left            =   390
      TabIndex        =   10
      Top             =   2490
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Couleur de fond"
      Height          =   255
      Left            =   390
      TabIndex        =   6
      Top             =   2940
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Vitesse de déplacement"
      Height          =   375
      Left            =   390
      TabIndex        =   5
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Diamètre"
      Height          =   375
      Left            =   390
      TabIndex        =   3
      Top             =   720
      Width           =   1575
   End
End
Attribute VB_Name = "frmSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim clickOk As Boolean
Private Sub Command1_Click()

End Sub

Private Sub btnAnnuler_Click()


    End
    
End Sub


Private Sub btnAppliquer_Click()

    SaveSetting "HorloGhis", "Apparence", "Diamètre", rayon * 2
    SaveSetting "HorloGhis", "Apparence", "Vitesse", vitesse
    SaveSetting "HorloGhis", "Apparence", "Digital", digital
    SaveSetting "HorloGhis", "Apparence", "Anneau", anneau
    SaveSetting "HorloGhis", "Apparence", "CouleurDeFond", couleurDeFond
    SaveSetting "HorloGhis", "Apparence", "caseDigitale", caseDigitale
    
    End
    
End Sub

Private Sub btnCouleur_Click()

    
    cmd1.ShowColor
    
    couleur.BackColor = cmd1.Color

    couleurDeFond = cmd1.Color
    
    frmMain.Form_Resize

End Sub

Private Sub cbDiamètre_Click()

    If Not clickOk Then Exit Sub
    
    rayon = cbDiamètre / 2
    
    frmMain.Form_Resize
    
End Sub


Private Sub cbVitesse_Click()

    If Not clickOk Then Exit Sub
    
    vitesse = cbVitesse
    
    If vitesse = 0 Then frmMain.Form_Resize
    
End Sub


Private Sub Check1_Click()

    If Not clickOk Then Exit Sub
    
    digital = Check1
    
   
End Sub

Private Sub Check2_Click()

    If Not clickOk Then Exit Sub
    
    anneau = Check2
    
End Sub

Private Sub Check3_Click()

    If Not clickOk Then Exit Sub
    
    caseDigitale = Check3
    
End Sub

Private Sub Form_Load()

    Dim z As Integer, i As Integer
    
    funct = SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_SHOWWINDOW Or SWP_NOSIZE Or SWP_NOMOVE)

        
    For z = 200 To 0.8 * Abs(frmMain.ScaleHeight) Step 50
        cbDiamètre.AddItem Str(z)
        If i = 0 And z >= 2 * rayon Then i = cbDiamètre.NewIndex
    Next
    
    cbDiamètre.ListIndex = i
    
    For z = 0 To 5
    
        cbVitesse.AddItem z
        If z = vitesse Then i = cbVitesse.NewIndex
        
    Next
    
    cbVitesse.ListIndex = i
    
    frmMain.Show
    
    couleur.BackColor = frmMain.BackColor
    cmd1.Color = frmMain.BackColor
    
    Check1.Value = Abs(digital)
    Check2.Value = Abs(anneau)
    Check3.Value = Abs(caseDigitale)
    
    clickOk = True

End Sub

Private Sub IDP_ABOUT_Click()

    frmAbout.Show
    
End Sub

