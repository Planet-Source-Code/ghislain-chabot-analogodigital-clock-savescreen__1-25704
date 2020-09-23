VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   2535
   ClientLeft      =   3855
   ClientTop       =   3345
   ClientWidth     =   3210
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Small Fonts"
      Size            =   6
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H000000FF&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   169
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   214
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox p1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   90
      ScaleHeight     =   89
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   125
      TabIndex        =   0
      Top             =   60
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.Timer tmrAnimate 
      Interval        =   10
      Left            =   2115
      Top             =   150
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function VerifyScreenSavePwd Lib "password.cpl" (ByVal hwnd&) As Boolean

Const Pi = 3.14159

Dim centreX!, centreY!, cx0!, cy0!
Dim nx!, ny!, nx0!, ny0!, nx1!, ny1!, nx2!, ny2!, nx3!, ny3!, nx4!, ny4!, i!, cx!, cy!
Dim heures As Single, minutes As Single, secondes As Single
Dim heureX As Integer, minuteX As Integer, secondeX As Integer

Dim angle As Single, distance As Single
Sub dessineAiguille(longueur As Integer, largeur As Integer, cas As String, couleur As Long)

    Dim valeur As String, distanceValeur As Single
    
    Select Case cas
    
    Case "secondes"
        i = secondes * 6
    
    Case "minutes"
        i = minutes * 6
    
    Case "heures"
        i = heures * 30
    
    Case Else
    End Select
    
    nx1 = cx + Cos(Pi / 180 * (90 - i)) * longueur
    ny1 = cy + Sin(Pi / 180 * (90 - i)) * longueur
    
    nx2 = cx + Cos(Pi / 180 * (90 + 90 - i)) * largeur
    ny2 = cy + Sin(Pi / 180 * (90 + 90 - i)) * largeur

    nx3 = cx + Cos(Pi / 180 * (90 - 90 - i)) * largeur
    ny3 = cy + Sin(Pi / 180 * (90 - 90 - i)) * largeur

    nx4 = cx - Cos(Pi / 180 * (90 - i)) * 0.11 * rayon
    ny4 = cy - Sin(Pi / 180 * (90 - i)) * 0.11 * rayon
    
    p1.Line (nx4, ny4)-(nx2, ny2), QBColor(1)
    p1.Line (nx2, ny2)-(nx1, ny1), QBColor(1)
    p1.Line (nx1, ny1)-(nx3, ny3), QBColor(1)
    p1.Line (nx3, ny3)-(nx4, ny4), QBColor(1)
    
    p1.FillColor = couleur
    p1.FillStyle = 0

    funct = ExtFloodFill(p1.hdc, cx, Abs(p1.ScaleHeight + cy), QBColor(1), 0)
    
    p1.Line (nx4, ny4)-(nx2, ny2), QBColor(15)
    p1.Line (nx2, ny2)-(nx1, ny1), QBColor(15)
    p1.Line (nx1, ny1)-(nx3, ny3), QBColor(15)
    p1.Line (nx3, ny3)-(nx4, ny4), QBColor(15)
        
    If digital Then
        
        Select Case cas
        
        Case "secondes"
            i = secondeX * 6
            valeur = Format(Trim(secondeX), " 00 s ")
            distanceValeur = 0.92 * rayon
    
        Case "minutes"
            i = minutes * 6
            valeur = Format(Trim(minuteX), " 00 m ")
            distanceValeur = 0.71 * rayon
            
        Case "heures"
            i = heures * 30
            valeur = Format(Trim(heureX), " 00 h ")
            distanceValeur = 0.52 * rayon
        
        Case Else
        End Select
        
        nx0 = cx + Cos(Pi / 180 * (90 - i)) * distanceValeur
        ny0 = cy + Sin(Pi / 180 * (90 - i)) * distanceValeur
    
        If caseDigitale Then dessineRectangle nx0, ny0, QBColor(10), 1.2 * p1.TextHeight(valeur) / 2, p1.TextWidth(valeur) / 2, 2
              
        p1.CurrentX = nx0 - p1.TextWidth(valeur) / 2
        p1.CurrentY = ny0 - p1.TextHeight(valeur) / 2
                    
        If caseDigitale Then p1.ForeColor = QBColor(0) Else p1.ForeColor = QBColor(10)
        p1.Print valeur
        
  
  End If
    
End Sub

Sub dessineRectangle(x As Single, y As Single, couleur As Long, hauteur As Integer, longueur As Single, bordure As Integer)

    p1.DrawWidth = bordure
    p1.Line (x - longueur, y - hauteur)-(x + longueur, y + hauteur), couleur, BF
    p1.Line (x - longueur, y - hauteur)-(x - longueur, y + hauteur), QBColor(7)
    p1.Line -(x + longueur, y + hauteur), QBColor(7)
    p1.Line -(x + longueur, y - hauteur), QBColor(15)
    p1.Line -(x - longueur, y - hauteur), QBColor(15)
    p1.DrawWidth = 1

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If RunMode = rmScreenSaver Then
        Unload Me
        
    End If
    
End Sub
Private Sub Form_Load()


    If Mid$(Command, 1, 2) <> "/P" And Mid$(Command, 1, 2) <> "/C" Then
        LockOn Me
    Else
        ShowCursor True
    End If
    
    'Type your initialization code here
        

    rayon = GetSetting("HorloGhis", "Apparence", "Diamètre", 350) / 2
    vitesse = GetSetting("HorloGhis", "Apparence", "Vitesse", 2)
    digital = GetSetting("HorloGhis", "Apparence", "Digital", True)
    anneau = GetSetting("HorloGhis", "Apparence", "Anneau", True)
    couleurDeFond = GetSetting("HorloGhis", "Apparence", "CouleurDeFond", QBColor(8))
    caseDigitale = GetSetting("HorloGhis", "Apparence", "caseDigitale", True)

    Show
    DoEvents
        
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    If RunMode = rmScreenSaver Then
        Unload Me
        End
    End If
    
End Sub

Public Sub Form_Resize()
    
    If Me.WindowState = 1 Then Exit Sub
    
    Dim z As Integer
    
    Me.Cls
    p1.Picture = LoadPicture()
    
    p1.BackColor = couleurDeFond
    Me.BackColor = couleurDeFond

    Me.Scale (0, Abs(Me.ScaleHeight))-(Me.ScaleWidth, 0)

    p1.Width = 2 * rayon + 20
    p1.Height = 2 * rayon + 20
    
    p1.Scale (0, p1.Height)-(p1.Width, 0)
    
    centreX = Abs(Me.ScaleWidth) / 2 - p1.Width / 2
    centreY = Abs(Me.ScaleHeight) / 2 - p1.Height / 2
        
    cx = Abs(p1.ScaleWidth) / 2
    cy = Abs(p1.ScaleHeight) / 2
    

    For z = 0 To 359
    
        If z Mod 30 = 0 Then
            nx = cx + Cos(Pi / 180 * (90 - z)) * 0.83 * rayon
            ny = cy + Sin(Pi / 180 * (90 - z)) * 0.83 * rayon
            dessineRectangle nx, ny, QBColor(10), 0.017 * rayon, 0.017 * rayon, 1
        End If
        
        If z Mod 6 = 0 Then
            nx = cx + Cos(Pi / 180 * (90 - z)) * 0.83 * rayon
            ny = cy + Sin(Pi / 180 * (90 - z)) * 0.83 * rayon
            p1.PSet (nx, ny), QBColor(10)
        End If
        
    Next
    

    
    p1.Picture = p1.Image
    
    z = 0.035 * rayon
    
    If z < 8 Then p1.FontName = "Small Fonts" Else p1.FontName = "Arial"
    
    p1.FontSize = z
    distance = 0
    
    Randomize (Timer)
    angle = Int((360 * Rnd) + 1)
    

End Sub

Private Sub Form_Unload(Cancel As Integer)

    'Put the focus away from the screensaver
    If RunMode = rmScreenSaver Then
        LockOff Me
    End If

    funct = ShowCursor(True)
    
    End
    
End Sub
Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    
    Static Count As Integer
    
    Count = Count + 1 ' Give enough time for program to run
    
    If Count > 10 Then
        If RunMode = rmScreenSaver Then
            Unload Me
            End
        End If
    End If
    
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Cancel = False

    'If Windows is shut down close this application too
    If UnloadMode = vbAppWindows Then
        ShowCursor True
        Exit Sub
    End If
    
    'if a password is beeing used ask for it and check its validity
    If RunMode = rmScreenSaver And UsePassword Then
        ShowCursor True
        If (VerifyScreenSavePwd(p1.hwnd)) = False Then
            Cancel = True
        End If
        ShowCursor False
    End If

End Sub

Private Sub déplaceHorloge()

    Dim cxTemp As Single, cyTemp As Single
    
    distance = distance + vitesse

    cxTemp = centreX + Cos(Pi / 180 * (90 - angle)) * distance
    cyTemp = centreY + Sin(Pi / 180 * (90 - angle)) * distance


    If cxTemp > Me.ScaleWidth - p1.Width + 10 Or cxTemp < -10 Then

        If angle = 180 Then angle = 182 Else angle = Abs(angle - 360)
        
        distance = 0
        centreX = cx0
        centreY = cy0

    ElseIf cyTemp < -10 Or cyTemp > Abs(Me.ScaleHeight) - p1.Height + 10 Then

        Select Case angle
        Case 0, 90, 180, 270
            angle = angle + 2
        Case 1 To 89
            angle = 90 + 90 - angle
        Case 181 To 269
            angle = 270 + 270 - angle
        Case 91 To 179
            angle = Abs(angle - 180)
        Case 271 To 359
            angle = 180 + 360 - angle
        Case Else
        End Select

        distance = 0
        centreX = cx0
        centreY = cy0

    Else

        cx0 = cxTemp
        cy0 = cyTemp

    End If


funct = BitBlt(frmMain.hdc, cx0, cy0, p1.ScaleWidth, Abs(p1.ScaleHeight), p1.hdc, 0, 0, &HCC0020)

End Sub

Private Sub tmrAnimate_Timer()
        
    
    Dim z As Integer, maintenant As Single, h As Integer
    Static k As Integer
    
    maintenant = Timer
    
    heures = Int(maintenant) / 3600
    minutes = Int(maintenant - Int(heures) * 3600) / 60
    secondes = (maintenant - Int(heures) * 3600 - Int(minutes) * 60)
    
    heureX = Int(heures)
    minuteX = Int(minutes)
    secondeX = Int(secondes)
                            
    p1.Cls

    dessineAiguille 0.67 * rayon, 0 * rayon, "secondes", QBColor(15)

    dessineAiguille 0.61 * rayon, 0.04 * rayon, "minutes", QBColor(7)
    
    dessineAiguille 0.43 * rayon, 0.04 * rayon, "heures", QBColor(7)
    
    dessineRectangle cx, cy, QBColor(10), 0.017 * rayon, 0.017 * rayon, 1
    
    If anneau Then
    
        If k = 0 Then k = 2 Else k = 0
                    
        For z = 0 To 359 Step 3
                
            nx = cx + Cos(Pi / 180 * (90 - z + k)) * rayon
            ny = cy + Sin(Pi / 180 * (90 - z + k)) * rayon
            p1.PSet (nx, ny), QBColor(10)
            
        Next
        
    End If
              
    déplaceHorloge
    
End Sub
Function calculLongueur(x!, y!, xx!, yy!) As Single

    calculLongueur = Sqr((y! - yy!) ^ 2 + (x! - xx!) ^ 2)

End Function

