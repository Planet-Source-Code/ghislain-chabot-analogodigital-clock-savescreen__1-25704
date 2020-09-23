VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "À propos ..."
   ClientHeight    =   4140
   ClientLeft      =   3570
   ClientTop       =   2190
   ClientWidth     =   4350
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   4350
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "Horloge AnalogoDigitale"
      Height          =   1455
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   3855
      Begin VB.Label Label4 
         Caption         =   "chaghis@hotmail.com"
         Height          =   240
         Left            =   1860
         TabIndex        =   7
         Top             =   990
         Width           =   1740
      End
      Begin VB.Label Label3 
         Caption         =   "Par Ghislain Chabot:"
         Height          =   255
         Left            =   285
         TabIndex        =   6
         Top             =   705
         Width           =   3225
      End
      Begin VB.Label Label1 
         Caption         =   "Un essai de sauve-écran utilitaire..."
         Height          =   285
         Left            =   285
         TabIndex        =   5
         Top             =   390
         Width           =   3150
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Merci à Igguk pour la base de ce sauve-écran"
      Height          =   1335
      Left            =   240
      TabIndex        =   1
      Top             =   1920
      Width           =   3855
      Begin VB.Label lblDeco1 
         BackStyle       =   0  'Transparent
         Caption         =   "Screen Saver 1.0"
         Height          =   240
         Left            =   150
         TabIndex        =   3
         Top             =   270
         Width           =   3390
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Minimum code to create a screen saver using Visual Basic This example was created by Igguk E-Mail : pcharles@swing.be"
         Height          =   720
         Left            =   150
         TabIndex        =   2
         Top             =   510
         Width           =   3375
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   510
      TabIndex        =   0
      Top             =   3585
      Width           =   3390
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdOK_Click()

    ' Unload and deallocate the about box.
    Unload frmAbout
    Set frmAbout = Nothing
    If RunMode = rmScreenSaver Then
        LockOn frmMain
    End If

End Sub

